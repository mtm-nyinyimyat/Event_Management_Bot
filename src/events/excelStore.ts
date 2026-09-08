import fs from "fs";
import path from "path";
import * as XLSX from "xlsx";
import {
  assertGraphExcelConfig,
  describeGraphExcelConfig,
  readAllowedExcelWorkbooksFromGraph,
} from "./graphExcelClient";

export type EventRecord = Record<string, string>;

export interface SheetData {
  sheet: string;
  rows: EventRecord[];
}

export interface WorkbookSearchResult {
  source: string;
  source_type: "graph" | "local" | "upload";
  file_name?: string;
  total_rows: number;
  match_count: number;
  sheets: SheetData[];
}

export interface CachedWorkbook {
  source: string;
  sourceType: "graph" | "local" | "upload";
  fileName?: string;
  sheets: SheetData[];
  loadedAt: number;
}

let workbookCache: CachedWorkbook | null = null;
/** Last successful chat upload — available in every chat until replaced. */
let sharedUploadedWorkbook: CachedWorkbook | null = null;
/** Per-conversation copy of an upload (same chat). */
const conversationWorkbooks = new Map<string, CachedWorkbook>();
/** Per-user copy so the uploader can reuse across chats even if shared is replaced. */
const userWorkbooks = new Map<string, CachedWorkbook>();

const UPLOAD_PERSIST_PATH = path.join(process.cwd(), "data", ".last-chat-upload.json");

function cacheTtlMs(): number {
  const raw = Number(process.env.EVENTS_CACHE_TTL_MS || 5 * 60 * 1000);
  return Number.isFinite(raw) && raw >= 0 ? raw : 5 * 60 * 1000;
}

export function getEventsSource(): "graph" | "local" | "chat" {
  const source = (process.env.EVENTS_SOURCE || "chat").trim().toLowerCase();
  if (source === "graph" || source === "teams" || source === "sharepoint") {
    return "graph";
  }
  if (source === "local" || source === "file") {
    return "local";
  }
  // Default: chat uploads only (no silent fallback to data/*.xlsx).
  return "chat";
}

function tryRestorePersistedUpload(): void {
  if (sharedUploadedWorkbook || !fs.existsSync(UPLOAD_PERSIST_PATH)) {
    return;
  }
  try {
    const raw = JSON.parse(fs.readFileSync(UPLOAD_PERSIST_PATH, "utf8")) as CachedWorkbook;
    if (raw?.sheets?.length && raw.sourceType === "upload") {
      sharedUploadedWorkbook = raw;
    }
  } catch {
    // ignore corrupt cache
  }
}

function persistSharedUpload(workbook: CachedWorkbook): void {
  try {
    const dir = path.dirname(UPLOAD_PERSIST_PATH);
    if (!fs.existsSync(dir)) {
      fs.mkdirSync(dir, { recursive: true });
    }
    fs.writeFileSync(UPLOAD_PERSIST_PATH, JSON.stringify(workbook));
  } catch {
    // non-fatal
  }
}

/** Resolve uploaded workbook: this chat → this user → shared (any chat) → none. */
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

/** Store an uploaded workbook for this chat, this user, and globally for every chat. */
export function setUploadedWorkbook(
  workbook: CachedWorkbook,
  options?: { conversationId?: string; userId?: string }
): void {
  sharedUploadedWorkbook = workbook;
  persistSharedUpload(workbook);
  if (options?.conversationId) {
    conversationWorkbooks.set(options.conversationId, workbook);
  }
  if (options?.userId) {
    userWorkbooks.set(options.userId, workbook);
  }
}

export function setConversationWorkbook(conversationId: string, workbook: CachedWorkbook): void {
  conversationWorkbooks.set(conversationId, workbook);
  sharedUploadedWorkbook = workbook;
  persistSharedUpload(workbook);
}

export function getConversationWorkbook(conversationId: string): CachedWorkbook | undefined {
  return resolveUploadedWorkbook(conversationId);
}

export function clearConversationWorkbook(conversationId: string): void {
  conversationWorkbooks.delete(conversationId);
}

export function getSharedUploadedWorkbook(): CachedWorkbook | undefined {
  tryRestorePersistedUpload();
  return sharedUploadedWorkbook || undefined;
}

export function resolveExcelPath(): string {
  if (process.env.EVENTS_EXCEL_PATH) {
    return path.resolve(process.env.EVENTS_EXCEL_PATH);
  }

  const dataDir = path.join(process.cwd(), "data");
  if (fs.existsSync(dataDir)) {
    const candidates = fs
      .readdirSync(dataDir)
      .filter((name) => name.toLowerCase().endsWith(".xlsx") && !name.startsWith("~$"))
      .map((name) => {
        const fullPath = path.join(dataDir, name);
        return { name, fullPath, mtime: fs.statSync(fullPath).mtimeMs };
      })
      .sort((a, b) => b.mtime - a.mtime);

    if (candidates.length > 0) {
      return candidates[0].fullPath;
    }
  }

  return path.resolve(path.join(process.cwd(), "data", "events.xlsx"));
}

function cellToString(value: unknown): string {
  if (value === null || value === undefined) {
    return "";
  }
  return String(value).replace(/\r\n/g, "\n").trim();
}

function cleanRecord(row: Record<string, unknown>): EventRecord | null {
  const record: EventRecord = {};
  for (const [key, value] of Object.entries(row)) {
    const column = key.trim();
    if (!column || column.startsWith("__EMPTY")) {
      continue;
    }
    const text = cellToString(value);
    if (text) {
      record[column] = text;
    }
  }
  return Object.keys(record).length > 0 ? record : null;
}

function parseWorkbookBuffer(buffer: Buffer, sheetPrefix?: string): SheetData[] {
  const workbook = XLSX.read(buffer, { type: "buffer" });
  if (!workbook.SheetNames.length) {
    throw new Error("No worksheets found in Excel workbook");
  }

  return workbook.SheetNames.map((sheetName) => {
    const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(workbook.Sheets[sheetName], {
      defval: "",
      raw: false,
    });

    return {
      sheet: sheetPrefix ? `${sheetPrefix} / ${sheetName}` : sheetName,
      rows: rows.map(cleanRecord).filter((row): row is EventRecord => row !== null),
    };
  });
}

/** Parse an uploaded .xlsx/.xls buffer into sheet rows. */
export function loadWorkbookFromBuffer(buffer: Buffer, fileName?: string): SheetData[] {
  return parseWorkbookBuffer(buffer, fileName);
}

function loadLocalWorkbook(): CachedWorkbook {
  const filePath = resolveExcelPath();
  if (!fs.existsSync(filePath)) {
    throw new Error(
      `Excel file not found at ${filePath}. Upload an .xlsx in chat with the bot, place one in data/, or set EVENTS_SOURCE=graph for Teams Files.`
    );
  }

  const buffer = fs.readFileSync(filePath);
  return {
    source: filePath,
    sourceType: "local",
    fileName: path.basename(filePath),
    sheets: parseWorkbookBuffer(buffer),
    loadedAt: Date.now(),
  };
}

async function loadGraphWorkbook(): Promise<CachedWorkbook> {
  assertGraphExcelConfig();
  const workbooks = await readAllowedExcelWorkbooksFromGraph();
  const multi = workbooks.length > 1;
  const sheets: SheetData[] = workbooks.flatMap((workbook) =>
    workbook.sheets.map((sheet) => ({
      sheet: multi ? `${workbook.fileName} / ${sheet.sheet}` : sheet.sheet,
      rows: sheet.rows,
    }))
  );

  return {
    source: workbooks.map((workbook) => workbook.webUrl || workbook.source).join(" | "),
    sourceType: "graph",
    fileName: workbooks.map((workbook) => workbook.fileName).join(", "),
    sheets,
    loadedAt: Date.now(),
  };
}

export interface WorkbookLookupOptions {
  conversationId?: string;
  userId?: string;
}

export async function loadWorkbook(
  forceRefresh = false,
  conversationIdOrOptions?: string | WorkbookLookupOptions
): Promise<SheetData[]> {
  const options: WorkbookLookupOptions =
    typeof conversationIdOrOptions === "string"
      ? { conversationId: conversationIdOrOptions }
      : conversationIdOrOptions || {};

  const uploaded = resolveUploadedWorkbook(options.conversationId, options.userId);
  if (uploaded) {
    return uploaded.sheets;
  }

  const source = getEventsSource();
  if (source === "chat") {
    throw new Error(
      "No Excel workbook is loaded yet. Attach an .xlsx file in chat (the file name should appear in my confirmation). I will not use the local data/ folder unless EVENTS_SOURCE=local."
    );
  }

  const ttl = cacheTtlMs();

  if (
    !forceRefresh &&
    workbookCache &&
    workbookCache.sourceType === source &&
    Date.now() - workbookCache.loadedAt < ttl
  ) {
    return workbookCache.sheets;
  }

  workbookCache = source === "graph" ? await loadGraphWorkbook() : loadLocalWorkbook();
  return workbookCache.sheets;
}

export async function loadEvents(
  conversationIdOrOptions?: string | WorkbookLookupOptions
): Promise<EventRecord[]> {
  const sheets = await loadWorkbook(false, conversationIdOrOptions);
  return sheets.flatMap((sheet) =>
    sheet.rows.map((row) => ({
      Sheet: sheet.sheet,
      ...row,
    }))
  );
}

function normalizeText(value: string): string {
  return value.toLocaleLowerCase("my").normalize("NFC");
}

function matchesQuery(haystack: string, query: string): boolean {
  const normalizedHaystack = normalizeText(haystack);
  const normalizedQuery = normalizeText(query).trim();
  if (!normalizedQuery) {
    return true;
  }

  if (normalizedHaystack.includes(normalizedQuery)) {
    return true;
  }

  const terms = normalizedQuery
    .split(/\s+/)
    .map((term) => term.trim())
    .filter((term) => term.length > 0);

  return terms.length > 0 && terms.every((term) => normalizedHaystack.includes(term));
}

export async function searchEvents(
  query: string,
  maxResults = 20,
  conversationIdOrOptions?: string | WorkbookLookupOptions
): Promise<EventRecord[]> {
  const events = await loadEvents(conversationIdOrOptions);
  const matches = !query.trim()
    ? events
    : events.filter((event) =>
        matchesQuery([...Object.keys(event), ...Object.values(event)].join(" "), query)
      );

  return matches.slice(0, Math.max(1, Math.min(maxResults, 50)));
}

export async function searchWorkbook(
  query: string,
  maxResults = 20,
  conversationIdOrOptions?: string | WorkbookLookupOptions
): Promise<WorkbookSearchResult> {
  const options: WorkbookLookupOptions =
    typeof conversationIdOrOptions === "string"
      ? { conversationId: conversationIdOrOptions }
      : conversationIdOrOptions || {};

  const sheetsData = await loadWorkbook(false, options);
  const uploaded = resolveUploadedWorkbook(options.conversationId, options.userId);
  const sourceType = uploaded?.sourceType || workbookCache?.sourceType || getEventsSource();
  const source =
    uploaded?.source ||
    workbookCache?.source ||
    (sourceType === "graph" ? describeGraphExcelConfig() : resolveExcelPath());
  const fileName = uploaded?.fileName || workbookCache?.fileName;
  const totalRows = sheetsData.reduce((sum, sheet) => sum + sheet.rows.length, 0);
  const limit = Math.max(1, Math.min(maxResults, 50));

  if (!query.trim()) {
    return {
      source,
      source_type: sourceType,
      file_name: fileName,
      total_rows: totalRows,
      match_count: totalRows,
      sheets: sheetsData.map((sheet) => ({
        sheet: sheet.sheet,
        rows: sheet.rows.slice(0, Math.min(10, limit)),
      })),
    };
  }

  const sheets: SheetData[] = [];
  let matchCount = 0;

  for (const sheet of sheetsData) {
    const rows = sheet.rows.filter((row) =>
      matchesQuery([...Object.keys(row), ...Object.values(row), sheet.sheet].join(" "), query)
    );
    if (rows.length === 0) {
      continue;
    }
    matchCount += rows.length;
    sheets.push({
      sheet: sheet.sheet,
      rows: rows.slice(0, limit),
    });
  }

  return {
    source,
    source_type: sourceType,
    file_name: fileName,
    total_rows: totalRows,
    match_count: matchCount,
    sheets,
  };
}

export function clearWorkbookCache(): void {
  workbookCache = null;
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
