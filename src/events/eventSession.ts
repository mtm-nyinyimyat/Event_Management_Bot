import { ILogger } from "@microsoft/teams.common";
import {
  clearRagIndexAsync,
  createRagDocument,
  deleteRagDocument,
  ensureDocumentSchema,
  ensureWorkbookIndexed,
  getRagConfig,
  hashContent,
  updateRagDocumentStatus,
  upsertSharepointSource,
  type DocumentSourceType,
} from "../rag";
import { getPostgresPool } from "../storage/postgres";
import { clearEventLookupMemory } from "./excel/eventLookupMemory";
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
import { clearActiveEventExport, setActiveEventExport } from "./activeEventFile";
import {
  downloadExcelBinaryFromShareUrl,
  getDeltaChanges,
  resolveSharePointExcelItem,
} from "./graphExcelClient";

export type EventSessionStatus = "idle" | "pending" | "active" | "ended";
export type EventIngestMode = "upload" | "sharepoint";

export interface EventSession {
  conversationId: string;
  documentId?: string | null;
  status: EventSessionStatus;
  pendingUrls: string[];
  activeSource?: string | null;
  activeFileName?: string | null;
  startedBy?: string | null;
  startedAt?: string | null;
  updatedAt: string;
}

export interface PendingUploadFile {
  fileName: string;
  buffer: Buffer;
  uploadedAt: number;
}

/** In-memory staged Excel uploads (permission-free testing). Cleared on /end. */
const pendingUploads = new Map<string, PendingUploadFile[]>();

let schemaReady: Promise<void> | null = null;

/**
 * Testing default: chat file uploads.
 * Set EVENTS_INGEST_MODE=sharepoint later to re-enable SharePoint URL flow.
 */
export function getEventIngestMode(): EventIngestMode {
  const raw = (process.env.EVENTS_INGEST_MODE || "upload").trim().toLowerCase();
  return raw === "sharepoint" || raw === "url" ? "sharepoint" : "upload";
}

export function isUploadIngestMode(): boolean {
  return getEventIngestMode() === "upload";
}

async function ensureSchema(): Promise<void> {
  if (!schemaReady) {
    schemaReady = ensureDocumentSchema();
  }
  await schemaReady;
}

function normalizeCommandText(text: string): string {
  return (
    String(text || "")
      // Teams mentions: <at>name</at>, <at id="28:...">name</at>, nested markup
      .replace(/<at\b[^>]*>[\s\S]*?<\/at>/gi, " ")
      .replace(/<\/?at\b[^>]*>/gi, " ")
      // Graph HTML body wrappers (<div>, <p>, <span>, …)
      .replace(/<[^>]+>/g, " ")
      .replace(/&nbsp;/gi, " ")
      .replace(/&#39;/g, "'")
      .replace(/&quot;/g, '"')
      .replace(/\u200b|\ufeff|\u00a0/g, " ")
      .replace(/\s+/g, " ")
      .trim()
      .toLowerCase()
  );
}

export function isStartEventCommand(text: string): boolean {
  const t = normalizeCommandText(text);
  if (t === "/start" || t === "start event") {
    return true;
  }
  // Mention leftovers / extra punctuation: treat a standalone /start token as the command.
  if (/(?:^|\s)\/start(?:$|[\s.,!?])/.test(t) && !/(?:^|\s)\/end/.test(t)) {
    return true;
  }
  return false;
}

export function isEndEventCommand(text: string): boolean {
  const t = normalizeCommandText(text);
  if (t === "/end" || t === "/endevent" || t === "endevent" || t === "end event") {
    return true;
  }
  return /(?:^|\s)\/end(?:event)?(?:$|[\s.,!?])/.test(t);
}

function findSharePointUrlsInText(text: string): string[] {
  const matches = text.match(/https?:\/\/[^\s"'<>]+/gi) || [];
  const cleaned = matches.map((url) => url.replace(/&amp;/g, "&").replace(/[),.;]+$/g, ""));
  return [
    ...new Set(
      cleaned.filter((url) => {
        const lower = url.toLowerCase();
        return (
          lower.includes("sharepoint.com") ||
          lower.includes("sharepoint") ||
          lower.includes("onedrive") ||
          lower.includes("1drv.") ||
          lower.includes("office.com") ||
          lower.includes("teams.microsoft.com") ||
          lower.includes("/:x:/") ||
          lower.includes("/:u:/") ||
          lower.includes(".xlsx") ||
          lower.includes(".xls") ||
          lower.includes("doc.aspx") ||
          lower.includes("download.aspx")
        );
      })
    ),
  ];
}

export function extractSharePointUrls(text?: string): string[] {
  if (!text) {
    return [];
  }
  return findSharePointUrlsInText(text);
}

function collectStringsFromContent(value: unknown, out: string[], depth = 0): void {
  if (depth > 8 || value == null) {
    return;
  }
  if (typeof value === "string") {
    out.push(value);
    return;
  }
  if (Array.isArray(value)) {
    for (const item of value) {
      collectStringsFromContent(item, out, depth + 1);
    }
    return;
  }
  if (typeof value === "object") {
    for (const item of Object.values(value as Record<string, unknown>)) {
      collectStringsFromContent(item, out, depth + 1);
    }
  }
}

function isShareLinkUrl(url: string): boolean {
  return /\/:[xuws]:\//i.test(url);
}

function directPathKey(url: string): string {
  try {
    const parsed = new URL(url);
    return `${parsed.host.toLowerCase()}${decodeURIComponent(parsed.pathname).toLowerCase()}`;
  } catch {
    return url.split("?")[0].toLowerCase();
  }
}

/**
 * One Teams link preview often yields contentUrl + downloadUrl + :x:/ href.
 * Prefer share links when present; otherwise collapse same-path variants.
 */
function collapseShareUrlVariants(urls: string[]): string[] {
  if (urls.length <= 1) {
    return urls;
  }
  const unique = [...new Set(urls)];
  const shareLinks = unique.filter(isShareLinkUrl);
  if (shareLinks.length > 0) {
    return [...new Set(shareLinks)];
  }

  const byPath = new Map<string, string>();
  for (const url of unique) {
    const key = directPathKey(url);
    const existing = byPath.get(key);
    if (!existing) {
      byPath.set(key, url);
      continue;
    }
    if (!/[?&]download=/i.test(url) && /[?&]download=/i.test(existing)) {
      byPath.set(key, url);
    }
  }
  return [...byPath.values()];
}

/**
 * Collect SharePoint/Excel URLs from message text AND link-preview/file attachments.
 * Teams link previews often leave activity.text empty or title-only; the real URL
 * lives in text/html content, contentUrl, or file.download.info JSON.
 */
export function extractSharePointUrlsFromActivity(activity: {
  text?: string;
  attachments?: Array<{
    contentType?: string;
    contentUrl?: string;
    content?: unknown;
    name?: string;
    thumbnailUrl?: string;
  }>;
  entities?: unknown[];
}): string[] {
  const parts: string[] = [];
  if (activity.text) {
    parts.push(activity.text);
  }
  for (const attachment of activity.attachments || []) {
    if (attachment.contentUrl) {
      parts.push(attachment.contentUrl);
    }
    if (attachment.thumbnailUrl) {
      parts.push(attachment.thumbnailUrl);
    }
    collectStringsFromContent(attachment.content, parts);
    // File cards sometimes only put the workbook name in text + a URL on the card body.
    if (attachment.name && /\.xlsx?$/i.test(attachment.name)) {
      collectStringsFromContent(attachment.content, parts);
    }
  }
  // Link cards sometimes live only on entities.
  collectStringsFromContent(activity.entities, parts);
  return collapseShareUrlVariants(findSharePointUrlsInText(parts.join("\n")));
}

/** URLs from a Graph Teams message (channel/chat) body + attachments. */
export function extractSharePointUrlsFromGraphMessage(message?: {
  body?: { content?: string };
  attachments?: Array<{
    contentType?: string;
    contentUrl?: string;
    content?: string;
    name?: string;
    webUrl?: string;
  }>;
}): string[] {
  if (!message) {
    return [];
  }
  const parts: string[] = [];
  if (message.body?.content) {
    parts.push(message.body.content);
  }
  for (const attachment of message.attachments || []) {
    if (attachment.contentUrl) {
      parts.push(attachment.contentUrl);
    }
    if (attachment.webUrl) {
      parts.push(attachment.webUrl);
    }
    if (typeof attachment.content === "string") {
      parts.push(attachment.content);
    }
  }
  return collapseShareUrlVariants(findSharePointUrlsInText(parts.join("\n")));
}

/** Merge BF + Graph URL candidates, then collapse preview variants. */
export function mergeSharePointUrls(...groups: string[][]): string[] {
  return collapseShareUrlVariants(groups.flat());
}

function mapRow(row: {
  conversation_id: string;
  document_id?: string | null;
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
    documentId: row.document_id || null,
    status: row.status,
    pendingUrls: pending,
    activeSource: row.active_source,
    activeFileName: row.active_file_name,
    startedBy: row.started_by,
    startedAt: row.started_at ? String(row.started_at) : null,
    updatedAt: String(row.updated_at),
  };
}

const SESSION_SELECT = `conversation_id, document_id, status, pending_urls, active_source,
            active_file_name, started_by, started_at, updated_at`;

export async function getEventSession(conversationId: string): Promise<EventSession> {
  await ensureSchema();
  const result = await getPostgresPool().query(
    `SELECT ${SESSION_SELECT}
     FROM event_sessions WHERE conversation_id = $1`,
    [conversationId]
  );
  if (!result.rows[0]) {
    return {
      conversationId,
      documentId: null,
      status: "idle",
      pendingUrls: [],
      updatedAt: new Date().toISOString(),
    };
  }
  return mapRow(result.rows[0]);
}

/** Active event from any chat (DM and group share one started event / document). */
export async function findActiveEventSession(): Promise<EventSession | null> {
  await ensureSchema();
  const result = await getPostgresPool().query(
    `SELECT ${SESSION_SELECT}
     FROM event_sessions
     WHERE status = 'active'
     ORDER BY started_at DESC NULLS LAST
     LIMIT 1`
  );
  if (!result.rows[0]) {
    return null;
  }
  return mapRow(result.rows[0]);
}

export async function isAnyEventActive(): Promise<boolean> {
  return Boolean(await findActiveEventSession());
}

export function getPendingUploads(conversationId: string): PendingUploadFile[] {
  return [...(pendingUploads.get(conversationId) || [])];
}

export function clearPendingUploads(conversationId: string): void {
  pendingUploads.delete(conversationId);
}

/**
 * Stage Excel file bytes from a chat upload. Not indexed until /start.
 */
export async function addPendingUploads(
  conversationId: string,
  files: Array<{ fileName: string; buffer: Buffer }>
): Promise<{ session: EventSession; pendingCount: number; fileNames: string[] }> {
  await ensureSchema();
  const existing = pendingUploads.get(conversationId) || [];
  const added: PendingUploadFile[] = files.map((file) => ({
    fileName: file.fileName,
    buffer: file.buffer,
    uploadedAt: Date.now(),
  }));
  const merged = [...existing, ...added];
  pendingUploads.set(conversationId, merged);

  const markers = merged.map((file) => `chat-upload://${file.fileName}`);
  const current = await getEventSession(conversationId);
  const status: EventSessionStatus = current.status === "active" ? "active" : "pending";

  await getPostgresPool().query(
    `INSERT INTO event_sessions (conversation_id, document_id, status, pending_urls, updated_at)
     VALUES ($1, NULL, $2, $3::jsonb, NOW())
     ON CONFLICT (conversation_id) DO UPDATE SET
       status = CASE
         WHEN event_sessions.status = 'active' THEN 'active'
         ELSE 'pending'
       END,
       pending_urls = $3::jsonb,
       updated_at = NOW()`,
    [conversationId, status, JSON.stringify(markers)]
  );

  return {
    session: await getEventSession(conversationId),
    pendingCount: merged.length,
    fileNames: added.map((file) => file.fileName),
  };
}

/** @deprecated Prefer upload mode for testing. Kept for SharePoint URL flow. */
export async function addPendingShareUrls(
  conversationId: string,
  urls: string[]
): Promise<EventSession> {
  await ensureSchema();
  const current = await getEventSession(conversationId);
  const merged = [...new Set([...current.pendingUrls, ...urls])];
  const status: EventSessionStatus = current.status === "active" ? "active" : "pending";

  await getPostgresPool().query(
    `INSERT INTO event_sessions (conversation_id, document_id, status, pending_urls, updated_at)
     VALUES ($1, NULL, $2, $3::jsonb, NOW())
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

function workbookFromUploadBuffers(
  files: PendingUploadFile[]
): { workbook: CachedWorkbook; fileNames: string[] } {
  const allSheets: SheetData[] = [];
  const fileNames: string[] = [];

  for (const file of files) {
    const sheets = loadWorkbookFromBuffer(file.buffer, file.fileName);
    fileNames.push(file.fileName);
    allSheets.push(...sheets);
  }

  if (!fileNames.length) {
    throw new Error("No Excel upload buffers available to start the event.");
  }

  return {
    fileNames,
    workbook: {
      source: `chat-upload://${fileNames.join(" | ")}`,
      sourceType: "upload",
      fileName: fileNames.join(", "),
      sheets: allSheets,
      loadedAt: Date.now(),
    },
  };
}

/**
 * SharePoint URL download path — used when EVENTS_INGEST_MODE=sharepoint.
 * Left intact for later; not used in upload test mode.
 */
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
      sourceType: "graph",
      fileName: files.join(", "),
      sheets: allSheets,
      loadedAt: Date.now(),
    },
  };
}

function resolveDocumentSourceType(workbook: CachedWorkbook): DocumentSourceType {
  if (workbook.sourceType === "upload" || workbook.source.startsWith("chat-upload://")) {
    return "upload";
  }
  if (workbook.sourceType === "graph") {
    return "graph";
  }
  if (/sharepoint\.com|onedrive/i.test(workbook.source)) {
    return "sharepoint";
  }
  return "local";
}

async function activateWorkbook(options: {
  conversationId: string;
  startedBy: string;
  workbook: CachedWorkbook;
  fileNames: string[];
  exportBuffer?: Buffer;
}): Promise<string> {
  const { conversationId, startedBy, workbook, fileNames, exportBuffer } = options;
  await ensureSchema();

  const prior = await getEventSession(conversationId);
  if (prior.documentId && prior.status === "active") {
    await clearRagIndexAsync({ persist: true, documentId: prior.documentId });
    await deleteRagDocument(prior.documentId);
  }

  const contentHash = exportBuffer
    ? hashContent(exportBuffer)
    : hashContent(`${workbook.source}::${workbook.loadedAt}::${workbook.sheets.length}`);

  const document = await createRagDocument({
    sourceType: resolveDocumentSourceType(workbook),
    sourceUri: workbook.source,
    fileName: workbook.fileName || fileNames.join(", "),
    contentHash,
    status: "syncing",
  });

  if (document.sourceType === "sharepoint" || document.sourceType === "graph") {
    const url = (workbook.source.split(" | ")[0] || workbook.source).trim();
    if (url && /^https?:\/\//i.test(url)) {
      try {
        const resolved = await resolveSharePointExcelItem(url);
        const delta = await getDeltaChanges({
          driveId: resolved.driveId,
          itemId: resolved.itemId,
        });
        const baselineEtag =
          resolved.eTag ||
          delta.items.find((item) => item.id === resolved.itemId)?.eTag ||
          null;
        await upsertSharepointSource({
          documentId: document.id,
          driveId: resolved.driveId,
          itemId: resolved.itemId,
          webUrl: resolved.webUrl || url,
          etag: baselineEtag,
          lastModifiedAt: resolved.lastModifiedDateTime || null,
          deltaLink: delta.deltaLink,
          lastDeltaSyncAt: new Date().toISOString(),
        });
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        console.warn(
          `SharePoint delta seed failed (webUrl only): ${message}`
        );
        await upsertSharepointSource({
          documentId: document.id,
          webUrl: url,
          lastDeltaSyncAt: new Date().toISOString(),
        });
      }
    }
  }

  setConversationWorkbook(conversationId, workbook);
  setActiveEventExport({ workbook, conversationId, originalBuffer: exportBuffer });

  await ensureWorkbookIndexed(
    workbook.sheets,
    { source: workbook.source, loadedAt: workbook.loadedAt, documentId: document.id },
    getRagConfig()
  );

  await updateRagDocumentStatus(document.id, "ready", {
    contentHash,
    fileName: workbook.fileName || fileNames.join(", "),
  });

  clearQueryCaches();
  clearPendingUploads(conversationId);

  await getPostgresPool().query(
    `INSERT INTO event_sessions (
       conversation_id, document_id, status, pending_urls, active_source, active_file_name,
       started_by, started_at, updated_at
     ) VALUES ($1, $2, 'active', '[]'::jsonb, $3, $4, $5, NOW(), NOW())
     ON CONFLICT (conversation_id) DO UPDATE SET
       document_id = EXCLUDED.document_id,
       status = 'active',
       pending_urls = '[]'::jsonb,
       active_source = EXCLUDED.active_source,
       active_file_name = EXCLUDED.active_file_name,
       started_by = EXCLUDED.started_by,
       started_at = NOW(),
       updated_at = NOW()`,
    [
      conversationId,
      document.id,
      workbook.source,
      workbook.fileName || fileNames.join(", "),
      startedBy,
    ]
  );

  const details = fileNames.map((name) => `• ${name}`).join("\n");
  return (
    `We're live 🙂\nLoaded:\n${details}\n\n` +
    `Ask me anything about the event — send /end when you're done.`
  );
}

/**
 * /start — activate pending Excel (chat uploads in test mode, or SharePoint URLs later).
 */
export async function startEventSession(options: {
  conversationId: string;
  startedBy: string;
  logger?: ILogger;
}): Promise<string> {
  const { conversationId, startedBy, logger } = options;
  const mode = getEventIngestMode();
  const session = await getEventSession(conversationId);

  const existingActive = await findActiveEventSession();
  if (existingActive && existingActive.conversationId !== conversationId) {
    return (
      "A previous event is still in progress in another chat and has not been ended yet.\n" +
      "Send /end (or endevent) there (or here) to finish it before starting another event."
    );
  }

  // --- Upload test mode (default) ---
  if (mode === "upload") {
    const uploads = getPendingUploads(conversationId);
    if (!uploads.length) {
      return (
        "I don't have an Excel yet.\n" +
        "1) Upload an .xlsx via paperclip → Upload from this device\n" +
        "2) Send /start"
      );
    }

    logger?.debug(`📎 Starting event from ${uploads.length} pending chat upload(s)`);
    const { workbook, fileNames } = workbookFromUploadBuffers(uploads);
    const exportBuffer =
      uploads.length === 1 ? uploads[0].buffer : undefined;
    return activateWorkbook({ conversationId, startedBy, workbook, fileNames, exportBuffer });
  }

  // --- SharePoint URL mode (kept for later; enable with EVENTS_INGEST_MODE=sharepoint) ---
  logger?.debug(
    `startEventSession sharepoint: status=${session.status} pending=${session.pendingUrls.length} urls=${JSON.stringify(session.pendingUrls)}`
  );
  if (!session.pendingUrls.length && session.status !== "active") {
    return (
      "I don't have a SharePoint link yet.\n" +
      "1) Paste a SharePoint/OneDrive Excel link\n" +
      "2) Send /start"
    );
  }

  const urls =
    session.pendingUrls.length > 0
      ? session.pendingUrls.filter((item) => /^https?:\/\//i.test(item))
      : session.activeSource
        ? session.activeSource.split(" | ").map((part) => part.trim()).filter(Boolean)
        : [];

  if (!urls.length) {
    return "I need a SharePoint link first — paste one, then send /start.";
  }

  const { workbook, files } = await downloadAndMergeWorkbooks(urls, logger);
  return activateWorkbook({ conversationId, startedBy, workbook, fileNames: files });
}

/**
 * /end — clear active workbook, RAG vectors, query caches, pending uploads/URLs.
 * Works from DM or group: clears the currently active event wherever it was started.
 */
export async function endEventSession(options: {
  conversationId: string;
  logger?: ILogger;
}): Promise<string> {
  const { conversationId, logger } = options;
  const localSession = await getEventSession(conversationId);
  const active = await findActiveEventSession();
  const targetId = active?.conversationId || conversationId;
  const session = active || localSession;
  const hadUploads =
    getPendingUploads(conversationId).length > 0 || getPendingUploads(targetId).length > 0;

  clearPendingUploads(conversationId);
  clearPendingUploads(targetId);
  clearActiveEventExport();
  clearConversationWorkbook(conversationId);
  clearConversationWorkbook(targetId);
  clearWorkbookCache();
  clearQueryCaches();
  clearEventLookupMemory(conversationId);
  clearEventLookupMemory(targetId);

  const documentIds = new Set<string>();
  if (session.documentId) {
    documentIds.add(session.documentId);
  }
  if (localSession.documentId) {
    documentIds.add(localSession.documentId);
  }

  for (const documentId of documentIds) {
    await clearRagIndexAsync({ persist: true, documentId });
    await deleteRagDocument(documentId);
  }

  await ensureSchema();
  await getPostgresPool().query(
    `DELETE FROM event_sessions WHERE conversation_id = $1 OR status = 'active'`,
    [targetId]
  );

  logger?.debug(`🧹 Event session ended for ${targetId} (requested from ${conversationId})`);

  if (session.status === "idle" && !session.pendingUrls.length && !hadUploads) {
    return isUploadIngestMode()
      ? "Nothing active right now — upload an .xlsx and send /start when you're ready."
      : "Nothing active right now — paste a SharePoint URL and send /start when you're ready.";
  }

  return (
    "All set — I cleared the event data.\n" +
    (isUploadIngestMode()
      ? "Upload a new .xlsx and send /start whenever you want to start again."
      : "Paste a new SharePoint link and send /start whenever you want to start again.")
  );
}

/**
 * Rehydrate after process restart when possible.
 * SharePoint URL sessions can re-download; chat-upload sessions cannot (buffers are in-memory only).
 * Pass force:true to re-download even when an in-memory copy exists (TTL refresh).
 */
export async function ensureActiveWorkbookInMemory(
  conversationId: string,
  logger?: ILogger,
  options?: { force?: boolean }
): Promise<CachedWorkbook | null> {
  const existing = getConversationWorkbook(conversationId);
  if (existing && !options?.force) {
    return existing;
  }

  const session = await getEventSession(conversationId);
  if (session.status !== "active" || !session.activeSource) {
    return null;
  }

  // Upload-mode sources cannot be rehydrated after restart
  if (session.activeSource.startsWith("chat-upload://")) {
    logger?.warn(
      `Active upload workbook for ${conversationId} was lost after restart. Re-upload and /start again.`
    );
    return null;
  }

  // SharePoint rehydrate (EVENTS_INGEST_MODE=sharepoint)
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

  if (session.documentId) {
    await ensureWorkbookIndexed(
      workbook.sheets,
      { source: workbook.source, loadedAt: workbook.loadedAt, documentId: session.documentId },
      getRagConfig()
    );
  }

  return workbook;
}
