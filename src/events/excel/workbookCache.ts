import fs from "fs";
import path from "path";
import type { CachedWorkbook } from "./types";

const UPLOAD_PERSIST_PATH = path.join(process.cwd(), "data", ".last-chat-upload.json");
const PARSE_VERSION = 5;

let workbookCache: CachedWorkbook | null = null;
let sharedUploadedWorkbook: CachedWorkbook | null = null;
const conversationWorkbooks = new Map<string, CachedWorkbook>();
const userWorkbooks = new Map<string, CachedWorkbook>();

export function cacheTtlMs(): number {
  const raw = Number(process.env.EVENTS_CACHE_TTL_MS || 5 * 60 * 1000);
  return Number.isFinite(raw) && raw >= 0 ? raw : 5 * 60 * 1000;
}

export function getEventsSource(): "graph" | "local" {
  const source = (process.env.EVENTS_SOURCE || "local").trim().toLowerCase();
  return source === "graph" || source === "teams" || source === "sharepoint" ? "graph" : "local";
}

export function isEventSessionMode(): boolean {
  return !["0", "false", "no", "off"].includes(
    (process.env.EVENTS_SESSION_MODE || "1").trim().toLowerCase()
  );
}

export function resolveExcelPath(): string {
  const configured = process.env.EVENTS_EXCEL_PATH?.trim();
  if (configured) {
    return path.isAbsolute(configured) ? configured : path.resolve(process.cwd(), configured);
  }
  return path.resolve(path.join(process.cwd(), "data", "events.xlsx"));
}

function tryRestorePersistedUpload(): void {
  if (sharedUploadedWorkbook || !fs.existsSync(UPLOAD_PERSIST_PATH)) {
    return;
  }
  try {
    const raw = JSON.parse(fs.readFileSync(UPLOAD_PERSIST_PATH, "utf8")) as CachedWorkbook;
    if (raw?.sheets?.length && raw.sourceType === "upload" && raw.parseVersion === PARSE_VERSION) {
      sharedUploadedWorkbook = raw;
    }
  } catch {
    // ignore corrupt cache
  }
}

export function persistSharedUpload(workbook: CachedWorkbook): void {
  try {
    const dir = path.dirname(UPLOAD_PERSIST_PATH);
    if (!fs.existsSync(dir)) {
      fs.mkdirSync(dir, { recursive: true });
    }
    const payload: CachedWorkbook = { ...workbook, parseVersion: PARSE_VERSION };
    fs.writeFileSync(UPLOAD_PERSIST_PATH, JSON.stringify(payload));
  } catch {
    // non-fatal
  }
}

export function getWorkbookCache(): CachedWorkbook | null {
  return workbookCache;
}

export function setWorkbookCache(workbook: CachedWorkbook | null): void {
  workbookCache = workbook;
}

export function resolveUploadedWorkbook(
  conversationId?: string,
  userId?: string
): CachedWorkbook | undefined {
  tryRestorePersistedUpload();
  if (conversationId) {
    const byConversation = conversationWorkbooks.get(conversationId);
    if (byConversation) {
      return byConversation;
    }
  }
  if (userId) {
    const byUser = userWorkbooks.get(userId);
    if (byUser) {
      return byUser;
    }
  }
  return sharedUploadedWorkbook || undefined;
}

export function setConversationWorkbook(conversationId: string, workbook: CachedWorkbook): void {
  conversationWorkbooks.set(conversationId, workbook);
  sharedUploadedWorkbook = workbook;
  persistSharedUpload(workbook);
}

export function getConversationWorkbook(conversationId: string): CachedWorkbook | undefined {
  return conversationWorkbooks.get(conversationId);
}

export function clearConversationWorkbook(conversationId: string): void {
  conversationWorkbooks.delete(conversationId);
  if (sharedUploadedWorkbook && workbookCache?.sourceType === "upload") {
    const stillReferenced = [...conversationWorkbooks.values()].some(
      (wb) => wb.source === sharedUploadedWorkbook?.source
    );
    if (!stillReferenced) {
      sharedUploadedWorkbook = null;
      try {
        if (fs.existsSync(UPLOAD_PERSIST_PATH)) {
          fs.unlinkSync(UPLOAD_PERSIST_PATH);
        }
      } catch {
        // ignore
      }
    }
  }
}

export function getSharedUploadedWorkbook(): CachedWorkbook | undefined {
  tryRestorePersistedUpload();
  return sharedUploadedWorkbook || undefined;
}

export function clearUploadedWorkbooks(): void {
  sharedUploadedWorkbook = null;
  conversationWorkbooks.clear();
  userWorkbooks.clear();
  try {
    if (fs.existsSync(UPLOAD_PERSIST_PATH)) {
      fs.unlinkSync(UPLOAD_PERSIST_PATH);
    }
  } catch {
    // ignore
  }
}

export function clearWorkbookMemoryCache(): void {
  workbookCache = null;
}
