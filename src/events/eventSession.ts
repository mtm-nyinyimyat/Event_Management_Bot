import { ILogger } from "@microsoft/teams.common";
import { clearRagIndexAsync, ensureWorkbookIndexed, getRagConfig } from "../rag";
import { getPostgresPool } from "../storage/postgres";
import { clearQueryCaches } from "../utils/queryCache";
import {
  clearConversationWorkbook,
  clearWorkbookCache,
  getConversationWorkbook,
  loadWorkbookFromBuffer,
  setConversationWorkbook,
  type CachedWorkbook,
  type SheetData,
} from "./excelStore";
import { downloadExcelBinaryFromShareUrl } from "./graphExcelClient";

export type EventSessionStatus = "idle" | "pending" | "active";

export interface EventSession {
  conversationId: string;
  status: EventSessionStatus;
  pendingUrls: string[];
  activeSource?: string | null;
  activeFileName?: string | null;
  startedBy?: string | null;
  startedAt?: string | null;
  updatedAt: string;
}

let schemaReady: Promise<void> | null = null;

async function ensureSchema(): Promise<void> {
  if (!schemaReady) {
    schemaReady = (async () => {
      const pool = getPostgresPool();
      await pool.query(`
        CREATE TABLE IF NOT EXISTS event_sessions (
          conversation_id TEXT PRIMARY KEY,
          status TEXT NOT NULL CHECK (status IN ('idle', 'pending', 'active')),
          pending_urls JSONB NOT NULL DEFAULT '[]'::jsonb,
          active_source TEXT,
          active_file_name TEXT,
          started_by TEXT,
          started_at TIMESTAMPTZ,
          updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
        );
      `);
    })();
  }
  await schemaReady;
}

function normalizeCommandText(text: string): string {
  return text
    .replace(/<\/?at>/gi, " ")
    .replace(/<at>[^<]*<\/at>/gi, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

export function isStartEventCommand(text: string): boolean {
  const t = normalizeCommandText(text);
  return (
    t === "/start" ||
    t === "/startevent" ||
    t === "startevent" ||
    t === "start event"
  );
}

export function isEndEventCommand(text: string): boolean {
  const t = normalizeCommandText(text);
  return (
    t === "/end" ||
    t === "/endevent" ||
    t === "endevent" ||
    t === "end event"
  );
}

export function extractSharePointUrls(text?: string): string[] {
  if (!text) {
    return [];
  }
  const matches = text.match(/https?:\/\/[^\s"'<>]+/gi) || [];
  const cleaned = matches.map((url) => url.replace(/&amp;/g, "&").replace(/[),.;]+$/g, ""));
  return [
    ...new Set(
      cleaned.filter((url) => {
        const lower = url.toLowerCase();
        return (
          lower.includes("sharepoint.com") ||
          lower.includes("onedrive") ||
          lower.includes("/:x:/") ||
          lower.includes(".xlsx") ||
          lower.includes("doc.aspx")
        );
      })
    ),
  ];
}

function mapRow(row: {
  conversation_id: string;
  status: EventSessionStatus;
  pending_urls: string[] | string;
  active_source: string | null;
  active_file_name: string | null;
  started_by: string | null;
  started_at: Date | string | null;
  updated_at: Date | string;
}): EventSession {
  const pending =
    typeof row.pending_urls === "string"
      ? (JSON.parse(row.pending_urls) as string[])
      : row.pending_urls || [];
  return {
    conversationId: row.conversation_id,
    status: row.status,
    pendingUrls: pending,
    activeSource: row.active_source,
    activeFileName: row.active_file_name,
    startedBy: row.started_by,
    startedAt: row.started_at ? String(row.started_at) : null,
    updatedAt: String(row.updated_at),
  };
}

export async function getEventSession(conversationId: string): Promise<EventSession> {
  await ensureSchema();
  const result = await getPostgresPool().query(
    `SELECT conversation_id, status, pending_urls, active_source, active_file_name,
            started_by, started_at, updated_at
     FROM event_sessions WHERE conversation_id = $1`,
    [conversationId]
  );
  if (!result.rows[0]) {
    return {
      conversationId,
      status: "idle",
      pendingUrls: [],
      updatedAt: new Date().toISOString(),
    };
  }
  return mapRow(result.rows[0]);
}

export async function addPendingShareUrls(
  conversationId: string,
  urls: string[]
): Promise<EventSession> {
  await ensureSchema();
  const current = await getEventSession(conversationId);
  const merged = [...new Set([...current.pendingUrls, ...urls])];
  const status: EventSessionStatus = current.status === "active" ? "active" : "pending";

  await getPostgresPool().query(
    `INSERT INTO event_sessions (conversation_id, status, pending_urls, updated_at)
     VALUES ($1, $2, $3::jsonb, NOW())
     ON CONFLICT (conversation_id) DO UPDATE SET
       status = CASE
         WHEN event_sessions.status = 'active' THEN 'active'
         ELSE 'pending'
       END,
       pending_urls = $3::jsonb,
       updated_at = NOW()`,
    [conversationId, status, JSON.stringify(merged)]
  );

  return getEventSession(conversationId);
}

async function downloadAndMergeWorkbooks(
  urls: string[],
  logger?: ILogger
): Promise<{ workbook: CachedWorkbook; files: string[] }> {
  const allSheets: SheetData[] = [];
  const files: string[] = [];
  const errors: string[] = [];

  for (const url of urls) {
    try {
      logger?.debug(`📥 Downloading SharePoint Excel: ${url}`);
      const { buffer, fileName } = await downloadExcelBinaryFromShareUrl(url);
      const sheets = loadWorkbookFromBuffer(buffer, fileName);
      files.push(fileName);
      allSheets.push(...sheets);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      errors.push(`${url} → ${message}`);
      logger?.warn(`SharePoint download failed: ${message}`);
    }
  }

  if (!files.length) {
    throw new Error(
      `Could not download any Excel workbook from the pending SharePoint URL(s).\n${errors.join("\n")}`
    );
  }

  return {
    files,
    workbook: {
      source: urls.join(" | "),
      sourceType: "upload",
      fileName: files.join(", "),
      sheets: allSheets,
      loadedAt: Date.now(),
    },
  };
}

/**
 * /start — download pending SharePoint URLs, activate workbook for this conversation, index RAG.
 */
export async function startEventSession(options: {
  conversationId: string;
  startedBy: string;
  logger?: ILogger;
}): Promise<string> {
  const { conversationId, startedBy, logger } = options;
  const session = await getEventSession(conversationId);

  if (!session.pendingUrls.length && session.status !== "active") {
    return (
      "No SharePoint Excel URL is pending.\n" +
      "1) Paste a SharePoint/OneDrive Excel link\n" +
      "2) Send /start (or startevent) to process it"
    );
  }

  const urls =
    session.pendingUrls.length > 0
      ? session.pendingUrls
      : session.activeSource
        ? session.activeSource.split(" | ").map((part) => part.trim()).filter(Boolean)
        : [];

  if (!urls.length) {
    return "No SharePoint URL available to start. Paste a link first, then send /start.";
  }

  const { workbook, files } = await downloadAndMergeWorkbooks(urls, logger);
  setConversationWorkbook(conversationId, workbook);

  await ensureWorkbookIndexed(
    workbook.sheets,
    { source: workbook.source, loadedAt: workbook.loadedAt, conversationId },
    getRagConfig()
  );

  clearQueryCaches();

  await getPostgresPool().query(
    `INSERT INTO event_sessions (
       conversation_id, status, pending_urls, active_source, active_file_name,
       started_by, started_at, updated_at
     ) VALUES ($1, 'active', '[]'::jsonb, $2, $3, $4, NOW(), NOW())
     ON CONFLICT (conversation_id) DO UPDATE SET
       status = 'active',
       pending_urls = '[]'::jsonb,
       active_source = EXCLUDED.active_source,
       active_file_name = EXCLUDED.active_file_name,
       started_by = EXCLUDED.started_by,
       started_at = NOW(),
       updated_at = NOW()`,
    [conversationId, workbook.source, workbook.fileName || files.join(", "), startedBy]
  );

  const details = files.map((name) => `• ${name}`).join("\n");
  return (
    `Event started.\nLoaded:\n${details}\n\n` +
    `I will answer questions from this workbook until someone sends /end (or endevent).`
  );
}

/**
 * /end — clear active workbook, RAG vectors, query caches, and session row for this conversation.
 */
export async function endEventSession(options: {
  conversationId: string;
  logger?: ILogger;
}): Promise<string> {
  const { conversationId, logger } = options;
  const session = await getEventSession(conversationId);

  clearConversationWorkbook(conversationId);
  clearWorkbookCache();
  clearQueryCaches();
  await clearRagIndexAsync({ persist: true, conversationId });

  await ensureSchema();
  await getPostgresPool().query(`DELETE FROM event_sessions WHERE conversation_id = $1`, [
    conversationId,
  ]);

  logger?.debug(`🧹 Event session ended for ${conversationId}`);

  if (session.status === "idle" && !session.pendingUrls.length) {
    return "No active event to end. Paste a SharePoint URL and send /start when ready.";
  }

  return (
    "Event ended.\n" +
    "Cleared workbook cache, RAG vectors, answer cache, and pending SharePoint URLs for this chat.\n" +
    "Paste a new SharePoint link and send /start to begin another event."
  );
}

export async function requireActiveEvent(conversationId: string): Promise<EventSession | null> {
  const session = await getEventSession(conversationId);
  return session.status === "active" ? session : null;
}

/**
 * If Postgres says the event is active but memory was cleared (process restart),
 * re-download from the saved SharePoint source URL(s).
 */
export async function ensureActiveWorkbookInMemory(
  conversationId: string,
  logger?: ILogger
): Promise<CachedWorkbook | null> {
  const existing = getConversationWorkbook(conversationId);
  if (existing) {
    return existing;
  }

  const session = await getEventSession(conversationId);
  if (session.status !== "active" || !session.activeSource) {
    return null;
  }

  const urls = session.activeSource
    .split(" | ")
    .map((part) => part.trim())
    .filter((part) => /^https?:\/\//i.test(part));

  if (!urls.length) {
    return null;
  }

  logger?.debug(`♻️ Rehydrating active event workbook for ${conversationId}`);
  const { workbook } = await downloadAndMergeWorkbooks(urls, logger);
  setConversationWorkbook(conversationId, workbook);

  await ensureWorkbookIndexed(
    workbook.sheets,
    { source: workbook.source, loadedAt: workbook.loadedAt, conversationId },
    getRagConfig()
  );

  return workbook;
}
