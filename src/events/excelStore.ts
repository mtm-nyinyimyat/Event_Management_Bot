import fs from "fs";
import path from "path";
import * as XLSX from "xlsx";
import { clearRagIndex, ensureWorkbookIndexed, retrieveHybrid } from "../rag";
import { getRagConfig } from "../rag/config";
import {
  assertGraphExcelConfig,
  describeGraphExcelConfig,
  readAllowedExcelWorkbooksFromGraph,
} from "./graphExcelClient";

export type EventRecord = Record<string, string>;

export type SheetSummary = Record<
  string,
  string | number | Record<string, number> | Record<string, string>
>;

export interface SheetData {
  sheet: string;
  rows: EventRecord[];
  numbered_item_count?: number;
  summary?: SheetSummary;
}

export interface WorkbookSearchResult {
  source: string;
  source_type: "graph" | "local" | "upload";
  file_name?: string;
  total_rows: number;
  match_count: number;
  answer_hint?: string;
  retrieval?: "hybrid" | "lexical" | "overview";
  rag?: {
    enabled: boolean;
    provider: string;
    top_k: number;
    hit_count: number;
  };
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
let sharedUploadedWorkbook: CachedWorkbook | null = null;
const conversationWorkbooks = new Map<string, CachedWorkbook>();
const userWorkbooks = new Map<string, CachedWorkbook>();

const UPLOAD_PERSIST_PATH = path.join(process.cwd(), "data", ".last-chat-upload.json");
const DEFAULT_MAX_RESULTS = 80;

function cacheTtlMs(): number {
  const raw = Number(process.env.EVENTS_CACHE_TTL_MS || 5 * 60 * 1000);
  return Number.isFinite(raw) && raw >= 0 ? raw : 5 * 60 * 1000;
}

export function getEventsSource(): "graph" | "local" {
  const source = (process.env.EVENTS_SOURCE || "local").trim().toLowerCase();
  return source === "graph" || source === "teams" || source === "sharepoint" ? "graph" : "local";
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

function uniqueHeaders(headers: string[]): string[] {
  const used = new Map<string, number>();
  return headers.map((header, index) => {
    const base = header || `Column_${index + 1}`;
    const count = used.get(base) || 0;
    used.set(base, count + 1);
    return count === 0 ? base : `${base}_${count + 1}`;
  });
}

function rowHasValue(row: unknown[]): boolean {
  return (row || []).some((cell) => cellToString(cell));
}

function looksLikeHeaderRow(row: unknown[]): boolean {
  const hints = new Set([
    "no",
    "no.",
    "start time",
    "end time",
    "period",
    "detail",
    "name",
    "description",
    "remark",
    "location",
    "日本語",
    "event participate",
    "gps point",
    "karaoke",
    "beverage",
    "volunteer",
    "pm",
  ]);
  const cells = (row || []).map((cell) => cellToString(cell).toLocaleLowerCase("my"));
  const hits = cells.filter((cell) => hints.has(cell)).length;
  return hits >= 2 || cells[0] === "no." || cells[0] === "no";
}

function isNumbered(value: string | undefined): boolean {
  return /^\d+$/.test((value || "").trim());
}

function countByField(rows: EventRecord[], field: string): Record<string, number> {
  const counts: Record<string, number> = {};
  for (const row of rows) {
    const value = (row[field] || "").trim();
    if (!value) {
      continue;
    }
    counts[value] = (counts[value] || 0) + 1;
  }
  return counts;
}

function primaryTableWidth(headerRow: unknown[]): number {
  let lastNamed = -1;
  (headerRow || []).forEach((cell, index) => {
    if (cellToString(cell)) {
      lastNamed = index;
    }
  });
  return Math.max(lastNamed + 1, 1);
}

function mergeHeaderRows(top: unknown[], bottom: unknown[], width: number): string[] {
  const merged: string[] = [];
  for (let i = 0; i < width; i += 1) {
    const a = cellToString(top[i]);
    const b = cellToString(bottom[i]);
    if (a && b && !/^ferry info/i.test(a)) {
      merged.push(`${a} / ${b}`);
    } else if (b && (!a || /^ferry info/i.test(a) || a.startsWith("Column_"))) {
      merged.push(b);
    } else if (a) {
      merged.push(a.startsWith("Ferry Info") ? "Ferry Info" : a);
    } else {
      merged.push("");
    }
  }

  // Normalize common Participants ferry columns
  return uniqueHeaders(
    merged.map((header) => {
      const lower = header.toLocaleLowerCase("my");
      if (lower.includes("office to event")) return "Office to Event";
      if (lower.includes("event place to home") || lower.includes("event to home")) {
        return "Event to Home";
      }
      if (lower.includes("ferry") && lower.includes("မှတ်တိုင်")) return "Ferry Note";
      if (header === "NAME") return "Name";
      return header;
    })
  );
}

function parseSidebarSummary(values: unknown[][], startCol: number): SheetSummary {
  const summary: SheetSummary = {};
  const counts: Record<string, number> = {};
  let section = "";

  for (const row of values) {
    const cells = (row || [])
      .slice(startCol)
      .map(cellToString)
      .filter(Boolean);
    if (cells.length === 0) {
      continue;
    }

    if (cells.length === 1 && cells[0].endsWith(":")) {
      section = cells[0].replace(/:$/, "");
      continue;
    }

    if (cells.length >= 2 && /^-?\d+(\.\d+)?$/.test(cells[1])) {
      const key = section ? `${section}: ${cells[0]}` : cells[0];
      counts[key] = Number(cells[1]);
      continue;
    }

    if (cells[0].endsWith(":") && cells.length >= 2) {
      summary[cells[0].replace(/:$/, "")] = cells.slice(1).join(" ");
      continue;
    }

    if (cells.length >= 2 && /^(Place|Date|Estimate Time)$/i.test(cells[0].replace(/:$/, ""))) {
      summary[cells[0].replace(/:$/, "")] = cells.slice(1).join(" ");
    }
  }

  if (Object.keys(counts).length > 0) {
    summary.sidebar_counts = counts;
  }
  return summary;
}

function parseParticipantsSheet(values: unknown[][]): { rows: EventRecord[]; summary: SheetSummary } {
  const headerIndex = values.findIndex((row) => looksLikeHeaderRow(row));
  if (headerIndex < 0) {
    return { rows: matrixToRecords(values), summary: {} };
  }

  const width = Math.min(primaryTableWidth(values[headerIndex] || []), 10);
  const next = values[headerIndex + 1] || [];
  const hasSubheader = next
    .slice(0, width)
    .some((cell, index) => !cellToString((values[headerIndex] || [])[index]) && !!cellToString(cell));
  const headers = hasSubheader
    ? mergeHeaderRows(values[headerIndex] || [], next, width)
    : uniqueHeaders((values[headerIndex] || []).slice(0, width).map(cellToString));
  const dataStart = headerIndex + (hasSubheader ? 2 : 1);

  const rows: EventRecord[] = [];
  for (let i = dataStart; i < values.length; i += 1) {
    const row = values[i] || [];
    const no = cellToString(row[0]);
    if (!isNumbered(no)) {
      if (rows.length > 0 && !row.slice(0, width).some((cell) => cellToString(cell))) {
        continue;
      }
      if (rows.length > 0) {
        // stop once numbered people end
        const maybeName = cellToString(row[1]);
        if (!maybeName) {
          continue;
        }
      }
      continue;
    }

    const record: EventRecord = {};
    headers.forEach((header, column) => {
      if (!header || header.startsWith("Column_")) {
        return;
      }
      const text = cellToString(row[column]);
      if (text) {
        record[header] = text;
      }
    });
    if (Object.keys(record).length > 0) {
      rows.push(record);
    }
  }

  const sidebar = parseSidebarSummary(values, width + 1);
  const byStatus = countByField(rows, "Event Participate");
  const byBeverage = countByField(rows, "Beverage");
  const summary: SheetSummary = {
    total_people: rows.length,
    participate:
      Object.entries(byStatus).find(([key]) => /^participate$/i.test(key))?.[1] ||
      byStatus.Participate ||
      0,
    cannot_participate:
      Object.entries(byStatus)
        .filter(([key]) => /can'?t|cannot|not participate/i.test(key))
        .reduce((sum, [, n]) => sum + n, 0) || 0,
    by_event_participate: byStatus,
    by_beverage: byBeverage,
    ...sidebar,
  };

  return { rows, summary };
}

function parseFerryDriverName(driverCell: string): string {
  // "Driver  -    Ko Hein Htoo  ( 09 - 254481458 )" -> "Ko Hein Htoo"
  return driverCell
    .replace(/^driver\s*[-–—:]?\s*/i, "")
    .replace(/\(\s*0?9[\d\s\-–—]*\s*\)/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function parseFerryCarPlate(ferryTitle: string): string {
  // "4N-9173 Ferry List" -> "4N-9173"
  const match = ferryTitle.match(/([A-Z0-9]+-[A-Z0-9]+)/i);
  return match ? match[1].toUpperCase() : "";
}

function parseFerrySheet(values: unknown[][]): { rows: EventRecord[]; summary: SheetSummary } {
  const rows: EventRecord[] = [];
  let ferryName = "";
  let driver = "";
  let driverName = "";
  let carPlate = "";
  let ferryNo = 0;
  let headers: string[] = [];

  for (let i = 0; i < values.length; i += 1) {
    const row = values[i] || [];
    const first = cellToString(row[0]);
    const joined = row.map(cellToString).filter(Boolean).join(" | ");

    if (/ferry\s*list/i.test(joined) || /ferry\s*list/i.test(first)) {
      ferryNo += 1;
      ferryName = first || joined;
      const driverCell = row.map(cellToString).find((cell) => /driver/i.test(cell)) || "";
      driver = driverCell;
      driverName = parseFerryDriverName(driverCell);
      carPlate = parseFerryCarPlate(ferryName);
      headers = [];
      continue;
    }

    if (looksLikeHeaderRow(row) || ((first === "No" || first === "No.") && row.map(cellToString).includes("Name"))) {
      headers = uniqueHeaders(
        row
          .map(cellToString)
          .map((header) => (header === "No" ? "No." : header))
          .filter((header, index, all) => header || all.slice(0, index + 1).some(Boolean))
      );
      // trim trailing empties
      while (headers.length && !headers[headers.length - 1]) {
        headers.pop();
      }
      headers = uniqueHeaders(headers.map((h) => h || "Column"));
      continue;
    }

    if (!headers.length || !isNumbered(first)) {
      if (joined && !headers.length) {
        rows.push({ Section: "Notice", Content: joined });
      }
      continue;
    }

    const record: EventRecord = {
      Ferry_No: String(ferryNo),
      Ferry: ferryName,
      ...(carPlate ? { Car_Plate: carPlate } : {}),
      ...(driver ? { Driver: driver } : {}),
      ...(driverName ? { Driver_Name: driverName } : {}),
    };
    headers.forEach((header, column) => {
      if (!header || header.startsWith("Column")) {
        const text = cellToString(row[column]);
        if (text && header) {
          record[header] = text;
        }
        return;
      }
      const text = cellToString(row[column]);
      if (text) {
        record[header] = text;
      }
    });
    rows.push(record);
  }

  const people = rows.filter((row) => isNumbered(row["No."] || row.No));
  const byFerry = countByField(people, "Ferry");
  const ferryIndex: Record<string, string> = {};
  for (const row of people) {
    if (row.Ferry_No && row.Ferry && !ferryIndex[row.Ferry_No]) {
      const label = [
        `Ferry ${row.Ferry_No}`,
        row.Car_Plate || "",
        row.Driver_Name || row.Driver || "",
      ]
        .filter(Boolean)
        .join(" / ");
      ferryIndex[row.Ferry_No] = label;
    }
  }

  return {
    rows,
    summary: {
      total_passengers: people.length,
      ferry_count: Object.keys(byFerry).length,
      by_ferry: byFerry,
      ferry_index: ferryIndex,
    },
  };
}

const TABLE_DRINK_PATTERN = /^(juice|beer|cocktail)$/i;
const TABLE_LABEL_PATTERN = /^(table[-\s]?\d+|vip)$/i;
const NON_PERSON_TABLE_TEXT = /^(stage|dinner menu|appetizer|salad|soup|main course|dessert|juice|beer|cocktail|all)$/i;

function isTablePersonName(text: string): boolean {
  if (!text || TABLE_LABEL_PATTERN.test(text) || TABLE_DRINK_PATTERN.test(text)) {
    return false;
  }
  if (NON_PERSON_TABLE_TEXT.test(text)) {
    return false;
  }
  if (/^※/.test(text) || text.endsWith(":")) {
    return false;
  }
  // Seat labels are mostly Latin names / Guest-N; skip pure Burmese menu lines
  if (/^[A-Za-z]/.test(text) || /^guest/i.test(text)) {
    return text.length >= 2;
  }
  return false;
}

function drinkNearCell(values: unknown[][], row: number, col: number): string {
  const neighbors: Array<[number, number]> = [
    [row - 1, col],
    [row + 1, col],
    [row, col - 1],
    [row, col + 1],
    [row - 1, col - 1],
    [row - 1, col + 1],
    [row + 1, col - 1],
    [row + 1, col + 1],
  ];
  for (const [r, c] of neighbors) {
    if (r < 0 || c < 0 || r >= values.length) {
      continue;
    }
    const text = cellToString((values[r] || [])[c]);
    if (TABLE_DRINK_PATTERN.test(text)) {
      return text.toUpperCase();
    }
  }
  return "";
}

function parseMenuItems(menuCells: string[]): {
  rows: EventRecord[];
  byCategory: Record<string, string>;
} {
  const rows: EventRecord[] = [];
  const byCategory: Record<string, string> = {};
  let category = "";

  for (const raw of menuCells) {
    const text = raw.trim();
    if (!text || /^dinner\s*menu:?$/i.test(text)) {
      continue;
    }

    // Category headers look like "Appetizer :" / "Main Course:"
    if (/^.+:\s*$/.test(text) || /^(appetizer|salad|soup|main\s*course|dessert)\s*:?\s*$/i.test(text)) {
      category = text.replace(/:\s*$/, "").trim();
      continue;
    }

    if (!category) {
      continue;
    }

    rows.push({
      Section: "Menu",
      Category: category,
      Dish: text,
    });

    byCategory[category] = byCategory[category]
      ? `${byCategory[category]} | ${text}`
      : text;
  }

  return { rows, byCategory };
}

function parseTableLayoutSheet(values: unknown[][]): { rows: EventRecord[]; summary: SheetSummary } {
  const rows: EventRecord[] = [];
  const namesByTable: Record<string, string[]> = {};
  const drinksByTable: Record<string, Record<string, number>> = {};
  const menuCells: string[] = [];
  const beverageSummary: EventRecord[] = [];

  // Menu column on the right
  for (let r = 0; r < values.length; r += 1) {
    for (let c = 14; c < (values[r] || []).length; c += 1) {
      const text = cellToString((values[r] || [])[c]);
      if (text) {
        menuCells.push(text);
      }
    }
  }

  // Drink count summary grid near the bottom (Juice / Cocktail / Beer headers)
  let summaryHeader = -1;
  for (let r = values.length - 1; r >= 0; r -= 1) {
    const texts = (values[r] || []).map(cellToString).filter(Boolean);
    const hasJuice = texts.some((text) => /^juice$/i.test(text));
    const hasCocktail = texts.some((text) => /^cocktail$/i.test(text));
    const hasBeer = texts.some((text) => /^beer$/i.test(text));
    // Seating-chart drink rows also contain those words; the summary row is short
    if (hasJuice && hasCocktail && hasBeer && texts.length <= 4) {
      summaryHeader = r;
      break;
    }
  }
  if (summaryHeader >= 0) {
    for (let r = summaryHeader + 1; r < values.length; r += 1) {
      const row = values[r] || [];
      const label = cellToString(row[2]) || cellToString(row[0]);
      if (!TABLE_LABEL_PATTERN.test(label)) {
        continue;
      }
      beverageSummary.push({
        Table: label,
        Juice: cellToString(row[3]) || "0",
        Cocktail: cellToString(row[4]) || "0",
        Beer: cellToString(row[5]) || "0",
      });
    }
  }

  // Cheers / notes under the seating chart
  for (let r = 0; r < values.length; r += 1) {
    const text = cellToString((values[r] || [])[2]);
    if (text.startsWith("※") || text.includes("Beer, Cocktail")) {
      rows.push({ Section: "Notice", Content: text });
    }
  }

  type LabelHit = { table: string; row: number; col: number };
  const labels: LabelHit[] = [];
  for (let r = 0; r < values.length; r += 1) {
    const row = values[r] || [];
    for (let c = 0; c < row.length; c += 1) {
      const text = cellToString(row[c]);
      if (!TABLE_LABEL_PATTERN.test(text)) {
        continue;
      }
      // Skip the beverage-summary labels at the bottom (same labels appear twice)
      if (summaryHeader >= 0 && r > summaryHeader) {
        continue;
      }
      const normalized = /^vip$/i.test(text)
        ? "VIP"
        : (() => {
            const match = text.match(/table[-\s]?(\d+)/i);
            return match ? `Table-${match[1]}` : text;
          })();
      labels.push({ table: normalized, row: r, col: c });
    }
  }

  for (const label of labels) {
    // Each table is a ~5-column block with names 2 rows above through 2-3 rows below the label
    const minCol = label.col;
    const maxCol = label.col + 4;
    const minRow = Math.max(0, label.row - 2);
    const maxRow = Math.min(values.length - 1, label.row + 3);

    namesByTable[label.table] = namesByTable[label.table] || [];
    drinksByTable[label.table] = drinksByTable[label.table] || {};

    for (let r = minRow; r <= maxRow; r += 1) {
      for (let c = minCol; c <= maxCol; c += 1) {
        const text = cellToString((values[r] || [])[c]);
        if (!isTablePersonName(text)) {
          continue;
        }
        if (namesByTable[label.table].includes(text)) {
          continue;
        }
        namesByTable[label.table].push(text);
        const drink = drinkNearCell(values, r, c);
        if (drink) {
          drinksByTable[label.table][drink] = (drinksByTable[label.table][drink] || 0) + 1;
        }
        rows.push({
          Table: label.table,
          Name: text,
          ...(drink ? { Beverage: drink } : {}),
        });
      }
    }
  }

  const menu = parseMenuItems(menuCells);
  rows.push(...menu.rows);
  for (const drinkRow of beverageSummary) {
    rows.push({
      Section: "Beverage Summary",
      Table: drinkRow.Table,
      Juice: drinkRow.Juice,
      Cocktail: drinkRow.Cocktail,
      Beer: drinkRow.Beer,
    });
  }

  return {
    rows,
    summary: {
      table_count: Object.keys(namesByTable).length,
      seated_names: Object.values(namesByTable).reduce((sum, names) => sum + names.length, 0),
      by_table: Object.fromEntries(
        Object.entries(namesByTable).map(([table, names]) => [table, names.length])
      ),
      people_by_table: Object.fromEntries(
        Object.entries(namesByTable).map(([table, names]) => [table, names.join(", ")])
      ),
      drinks_by_table: drinksByTable,
      ...(Object.keys(menu.byCategory).length
        ? { menu_by_category: menu.byCategory }
        : {}),
    },
  };
}

function matrixToRecords(values: unknown[][]): EventRecord[] {
  if (!values?.length) {
    return [];
  }

  let headerIndex = values.findIndex((row) => looksLikeHeaderRow(row));
  if (headerIndex < 0) {
    headerIndex = values.findIndex((row) => rowHasValue(row));
  }
  if (headerIndex < 0) {
    return [];
  }

  const records: EventRecord[] = [];

  for (let i = 0; i < headerIndex; i += 1) {
    const text = (values[i] || []).map(cellToString).filter(Boolean).join(" ").trim();
    if (text) {
      records.push({ Section: "Notice", Content: text });
    }
  }

  const width = primaryTableWidth(values[headerIndex] || []);
  const headers = uniqueHeaders(
    (values[headerIndex] || []).slice(0, width).map((cell) => cellToString(cell))
  );
  let emptyStreak = 0;

  for (let i = headerIndex + 1; i < values.length; i += 1) {
    const row = values[i] || [];
    if (!row.slice(0, width).some((cell) => cellToString(cell))) {
      emptyStreak += 1;
      if (emptyStreak >= 8) {
        break;
      }
      continue;
    }
    emptyStreak = 0;

    // New section header inside sheet (e.g. ferry) — keep as notice and continue with same headers
    if (looksLikeHeaderRow(row)) {
      continue;
    }

    const record: EventRecord = {};
    headers.forEach((header, column) => {
      if (!header || header.startsWith("__EMPTY") || header.startsWith("Column_")) {
        return;
      }
      const text = cellToString(row[column]);
      if (text) {
        record[header] = text;
      }
    });

    if (Object.keys(record).length > 0) {
      records.push(record);
    }
  }

  return records;
}

function defaultSummaryForRows(sheetName: string, rows: EventRecord[]): SheetSummary {
  const numbered = rows.filter((row) => isNumbered(row["No."] || row.No));
  const summary: SheetSummary = {
    total_rows: rows.length,
    numbered_item_count: numbered.length,
  };

  if (/agenda/i.test(sheetName)) {
    summary.agenda_items = numbered.length;
  }
  if (/volunteer/i.test(sheetName)) {
    summary.volunteer_roles = numbered.length;
  }
  return summary;
}

function parseSheet(sheetName: string, values: unknown[][]): SheetData {
  if (/participant/i.test(sheetName)) {
    const parsed = parseParticipantsSheet(values);
    return {
      sheet: sheetName,
      rows: parsed.rows,
      numbered_item_count: parsed.rows.filter((row) => isNumbered(row["No."] || row.No)).length,
      summary: parsed.summary,
    };
  }

  if (/ferry/i.test(sheetName) || /အပြန်/.test(sheetName)) {
    const parsed = parseFerrySheet(values);
    return {
      sheet: sheetName,
      rows: parsed.rows,
      numbered_item_count: parsed.rows.filter((row) => isNumbered(row["No."] || row.No)).length,
      summary: parsed.summary,
    };
  }

  if (/table\s*layout/i.test(sheetName)) {
    const parsed = parseTableLayoutSheet(values);
    return {
      sheet: sheetName,
      rows: parsed.rows,
      numbered_item_count: 0,
      summary: parsed.summary,
    };
  }

  const rows = matrixToRecords(values);
  return {
    sheet: sheetName,
    rows,
    numbered_item_count: rows.filter((row) => isNumbered(row["No."] || row.No)).length,
    summary: defaultSummaryForRows(sheetName, rows),
  };
}

function parseWorkbookBuffer(buffer: Buffer, sheetPrefix?: string): SheetData[] {
  const workbook = XLSX.read(buffer, { type: "buffer" });
  if (!workbook.SheetNames.length) {
    throw new Error("No worksheets found in Excel workbook");
  }

  return workbook.SheetNames.map((sheetName) => {
    const values = XLSX.utils.sheet_to_json<unknown[]>(workbook.Sheets[sheetName], {
      header: 1,
      defval: "",
      raw: false,
      blankrows: false,
    });

    const parsed = parseSheet(sheetName, values);
    return {
      ...parsed,
      sheet: sheetPrefix ? `${sheetPrefix} / ${sheetName}` : sheetName,
    };
  });
}

export function loadWorkbookFromBuffer(buffer: Buffer, fileName?: string): SheetData[] {
  return parseWorkbookBuffer(buffer, fileName);
}

function loadLocalWorkbook(): CachedWorkbook {
  const filePath = resolveExcelPath();
  if (!fs.existsSync(filePath)) {
    throw new Error(
      `Excel file not found at ${filePath}. Place an .xlsx in data/ or set EVENTS_EXCEL_PATH.`
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
  _conversationIdOrOptions?: string | WorkbookLookupOptions
): Promise<SheetData[]> {
  const source = getEventsSource();
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

  const ragConfig = getRagConfig();
  if (ragConfig.enabled) {
    try {
      await ensureWorkbookIndexed(workbookCache.sheets, {
        source: workbookCache.source,
        loadedAt: workbookCache.loadedAt,
      }, ragConfig);
    } catch (error) {
      // Lexical search still works if indexing fails
      console.warn(
        `RAG index failed: ${error instanceof Error ? error.message : String(error)}`
      );
    }
  }

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

const SEARCH_STOP_WORDS = new Set([
  "a",
  "an",
  "all",
  "are",
  "count",
  "for",
  "give",
  "how",
  "in",
  "is",
  "item",
  "items",
  "list",
  "many",
  "me",
  "of",
  "please",
  "show",
  "the",
  "them",
  "total",
  "what",
  "who",
  "about",
]);

const SHEET_QUERY_HINTS: Array<{ sheetIncludes: string; hints: string[] }> = [
  {
    sheetIncludes: "agenda",
    hints: ["agenda", "program", "schedule", "အစီအစဉ်", "အစီအစဥ်"],
  },
  { sheetIncludes: "volunteer", hints: ["volunteer", "စေတနာ"] },
  { sheetIncludes: "ferry", hints: ["ferry", "ဖယ်ရီ", "အပြန်", "driver"] },
  {
    sheetIncludes: "participant",
    hints: ["participant", "participants", "attendee", "attendees", "member", "members", "မန်ဘာ", "ပါဝင်"],
  },
  {
    sheetIncludes: "table layout",
    hints: [
      "table",
      "seat",
      "seating",
      "layout",
      "menu",
      "appetizer",
      "salad",
      "soup",
      "dessert",
      "main course",
      "dinner",
      "dish",
      "စားပွဲ",
      "ဟင်း",
      "မီနူး",
    ],
  },
];

function significantTerms(query: string): string[] {
  return normalizeText(query)
    .split(/\s+/)
    .map((term) => term.trim())
    .filter((term) => term.length > 0 && !SEARCH_STOP_WORDS.has(term));
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

  const terms = significantTerms(query);
  if (terms.length === 0) {
    return false;
  }

  return terms.every((term) => normalizedHaystack.includes(term));
}

function countNumberedItems(rows: EventRecord[]): number {
  return rows.filter((row) => isNumbered(row["No."] || row["No"])).length;
}

function isCountOrSummaryQuery(query: string): boolean {
  const q = normalizeText(query);
  return (
    /\b(how many|count|total|summary|overview|statistics|stats)\b/.test(q) ||
    q.includes("စုစုပေါင်း") ||
    q.includes("ဘယ်နှစ်") ||
    q.includes("ဦးရေ") ||
    q.includes("ဦးရေ") ||
    q.includes("အရေအတွက်")
  );
}

function isListQuery(query: string): boolean {
  const q = normalizeText(query);
  return /\b(list|show|all|full|every|entire)\b/.test(q) || q.includes("စာရင်း");
}

function isMenuQuery(query: string): boolean {
  const q = normalizeText(query);
  return (
    /\b(menu|appetizer|salad|soup|dessert|main\s*course|dinner\s*menu|dish|dishes)\b/.test(q) ||
    q.includes("မီနူး") ||
    q.includes("ဟင်း") ||
    q.includes("အာလူး") ||
    q.includes("သုပ်") ||
    q.includes("အသီး")
  );
}

function isFerryQuery(query: string): boolean {
  const q = normalizeText(query);
  return (
    /\b(ferry|driver|drop-?off|pickup|route)\b/.test(q) ||
    q.includes("ဖယ်ရီ") ||
    q.includes("အပြန်") ||
    q.includes("လမ်း") ||
    q.includes("မှတ်တိုင်")
  );
}

function filterMenuRows(rows: EventRecord[], query: string): EventRecord[] {
  const menuRows = rows.filter((row) => row.Section === "Menu" && row.Dish);
  if (!menuRows.length) {
    return [];
  }

  const q = normalizeText(query);
  const categoryHints: Array<{ category: RegExp; terms: RegExp }> = [
    { category: /appetizer/i, terms: /\bappetizer\b/ },
    { category: /salad/i, terms: /\bsalad\b|သုပ်/ },
    { category: /soup/i, terms: /\bsoup\b|ဟင်းချို/ },
    { category: /main/i, terms: /\bmain\b|main\s*course/ },
    { category: /dessert/i, terms: /\bdessert\b|အသီး/ },
  ];

  for (const hint of categoryHints) {
    if (hint.terms.test(q)) {
      const matched = menuRows.filter((row) => hint.category.test(row.Category || ""));
      if (matched.length) {
        return matched;
      }
    }
  }

  const filtered = menuRows.filter((row) =>
    matchesQuery([row.Category, row.Dish, row.Section].filter(Boolean).join(" "), query)
  );
  return filtered.length ? filtered : menuRows;
}

function queryTargetsSheet(sheetName: string, query: string): boolean {
  const name = normalizeText(sheetName);
  const q = normalizeText(query).trim();
  if (!q) {
    return false;
  }

  if (q.includes(name) || name.includes(q)) {
    return true;
  }

  const terms = significantTerms(query);
  if (terms.some((term) => name.includes(term) || term.includes(name))) {
    return true;
  }

  return SHEET_QUERY_HINTS.some(
    (alias) =>
      name.includes(alias.sheetIncludes) &&
      alias.hints.some((hint) => q.includes(normalizeText(hint)))
  );
}

function isNoiseColumnKey(key: string): boolean {
  if (!key || key.startsWith("__")) {
    return true;
  }
  if (key.startsWith("Column_") || key.startsWith("__EMPTY")) {
    return true;
  }
  return false;
}

/** Column order discovered from workbook rows (first-seen key order). */
function fieldOrderFromRows(rows: EventRecord[]): string[] {
  const order: string[] = [];
  const seen = new Set<string>();
  for (const row of rows) {
    for (const key of Object.keys(row)) {
      if (isNoiseColumnKey(key) || seen.has(key)) {
        continue;
      }
      seen.add(key);
      order.push(key);
    }
  }
  return order;
}

/**
 * Keep non-empty fields from the workbook row.
 * Field names come from the file (plus parser-added keys like Ferry_No), not a fixed schema.
 */
function compactRow(row: EventRecord, fieldOrder?: string[]): EventRecord {
  const keys = fieldOrder?.length
    ? fieldOrder
    : Object.keys(row).filter((key) => !isNoiseColumnKey(key));

  const out: EventRecord = {};
  for (const key of keys) {
    if (isNoiseColumnKey(key)) {
      continue;
    }
    const value = row[key]?.trim();
    if (!value) {
      continue;
    }
    out[key === "NAME" ? "Name" : key] = value;
  }

  // Sparse rows may have extra keys not in the sheet-wide order
  if (fieldOrder?.length) {
    for (const [key, value] of Object.entries(row)) {
      if (isNoiseColumnKey(key) || out[key] || out.Name === value || !value?.trim()) {
        continue;
      }
      out[key === "NAME" ? "Name" : key] = value.trim();
    }
  }

  return out;
}

function withSheetStats(sheet: SheetData, rows: EventRecord[], options?: { summaryOnly?: boolean }): SheetData {
  const fieldOrder = fieldOrderFromRows(rows.length ? rows : sheet.rows);
  const compactRows = options?.summaryOnly
    ? rows.slice(0, 5).map((row) => compactRow(row, fieldOrder))
    : rows.map((row) => compactRow(row, fieldOrder));

  return {
    sheet: sheet.sheet,
    numbered_item_count: sheet.numbered_item_count ?? countNumberedItems(sheet.rows),
    summary: sheet.summary,
    rows: compactRows,
  };
}

function resultLimit(maxResults?: number): number {
  const raw = maxResults ?? DEFAULT_MAX_RESULTS;
  return Math.max(1, Math.min(raw, 200));
}

function answerHintForQuery(query: string, sheets: SheetData[]): string | undefined {
  if (!sheets.length) {
    return undefined;
  }

  if (isCountOrSummaryQuery(query)) {
    const participants = sheets.find((sheet) => /participant/i.test(sheet.sheet));
    if (participants?.summary) {
      return `Use Participants.summary.total_people / participate / cannot_participate. Do not dump every row.`;
    }
    const agenda = sheets.find((sheet) => /agenda/i.test(sheet.sheet));
    if (agenda) {
      return `Use Agenda.numbered_item_count (or summary.agenda_items) for the timed program count.`;
    }
    return `Prefer each sheet's summary and numbered_item_count for totals.`;
  }

  if (isMenuQuery(query)) {
    const table = sheets.find((sheet) => /table layout/i.test(sheet.sheet));
    const menuRows = (table?.rows || []).filter((row) => row.Section === "Menu" && row.Dish);
    const summaryMenu = table?.summary?.menu_by_category;
    if (menuRows.length || summaryMenu) {
      const lines = menuRows.map((row) => `${row.Category}: ${row.Dish}`);
      return `MENU FROM EXCEL ONLY. Quote Dish text exactly (keep Burmese). Do not translate or invent English dish names. Rows: ${
        lines.join(" || ") || JSON.stringify(summaryMenu)
      }`;
    }
    return `For menu questions use Table Layout menu rows (Category + Dish) or summary.menu_by_category. Quote Dish exactly; never invent/translate dish names.`;
  }

  if (isFerryQuery(query)) {
    const ferrySheet = sheets.find((sheet) => /ferry/i.test(sheet.sheet) || /အပြန်/.test(sheet.sheet));
    if (ferrySheet?.rows?.length) {
      const q = normalizeText(query);
      const exact = ferrySheet.rows.find((row) => {
        const location = normalizeText(row.Location || "");
        return (
          location &&
          (location.includes(q) ||
            q
              .split(/[,\s၊]+/)
              .filter((t) => t.length > 1)
              .every((t) => location.includes(t)))
        );
      });
      if (exact?.Ferry_No) {
        return `Location match: answer with Ferry_No=${exact.Ferry_No}, Driver_Name=${exact.Driver_Name || exact.Driver || ""}, Car_Plate=${exact.Car_Plate || ""}, Location=${exact.Location || ""}. Do not invent other ferries/drivers.`;
      }
      return `For ferry/drop-off questions use Ferry_No, Driver_Name, Car_Plate, and Location from matching ferry rows. Ferry_No is the sheet order (1,2,3…). Never invent names or car plates.`;
    }
  }

  return undefined;
}

export async function searchEvents(
  query: string,
  maxResults = DEFAULT_MAX_RESULTS,
  conversationIdOrOptions?: string | WorkbookLookupOptions
): Promise<EventRecord[]> {
  const events = await loadEvents(conversationIdOrOptions);
  const matches = !query.trim()
    ? events
    : events.filter((event) =>
        matchesQuery([...Object.keys(event), ...Object.values(event)].join(" "), query)
      );

  return matches.slice(0, resultLimit(maxResults)).map(compactRow);
}

export async function searchWorkbook(
  query: string,
  maxResults = DEFAULT_MAX_RESULTS,
  conversationIdOrOptions?: string | WorkbookLookupOptions
): Promise<WorkbookSearchResult> {
  const options: WorkbookLookupOptions =
    typeof conversationIdOrOptions === "string"
      ? { conversationId: conversationIdOrOptions }
      : conversationIdOrOptions || {};

  const sheetsData = await loadWorkbook(false, options);
  const sourceType = workbookCache?.sourceType || getEventsSource();
  const source =
    workbookCache?.source ||
    (sourceType === "graph" ? describeGraphExcelConfig() : resolveExcelPath());
  const fileName = workbookCache?.fileName;
  const totalRows = sheetsData.reduce((sum, sheet) => sum + sheet.rows.length, 0);
  const limit = resultLimit(maxResults);
  const trimmedQuery = query.trim();

  if (!trimmedQuery) {
    const sheets = sheetsData.map((sheet) =>
      withSheetStats(sheet, [], { summaryOnly: true })
    );
    return {
      source,
      source_type: sourceType,
      file_name: fileName,
      total_rows: totalRows,
      match_count: totalRows,
      retrieval: "overview",
      answer_hint: "Workbook overview. Use each sheet.summary and numbered_item_count. Call again with a sheet keyword for details.",
      sheets,
    };
  }

  const summaryOnly = isCountOrSummaryQuery(trimmedQuery) && !isListQuery(trimmedQuery);
  const ragConfig = getRagConfig();

  // Menu questions: return structured Category/Dish rows from Table Layout only (no hallucination surface)
  if (isMenuQuery(trimmedQuery)) {
    const tableSheet = sheetsData.find((sheet) => /table layout/i.test(sheet.sheet));
    if (tableSheet) {
      const menuRows = filterMenuRows(tableSheet.rows, trimmedQuery);
      const sheet = withSheetStats(tableSheet, menuRows);
      return {
        source,
        source_type: sourceType,
        file_name: fileName,
        total_rows: totalRows,
        match_count: menuRows.length,
        retrieval: "lexical",
        answer_hint: answerHintForQuery(trimmedQuery, [sheet]),
        sheets: [sheet],
      };
    }
  }

  // Count/summary questions stay on structured sheet summaries (more reliable than vectors)
  if (!summaryOnly && ragConfig.enabled && workbookCache) {
    try {
      await ensureWorkbookIndexed(
        sheetsData,
        { source: workbookCache.source, loadedAt: workbookCache.loadedAt },
        ragConfig
      );
      const rag = await retrieveHybrid(trimmedQuery, sheetsData, {
        ...ragConfig,
        topK: Math.min(limit, ragConfig.topK),
      });

      if (rag.sheets.length > 0 || rag.hits.length > 0) {
        return {
          source,
          source_type: sourceType,
          file_name: fileName,
          total_rows: totalRows,
          match_count: rag.match_count,
          retrieval: "hybrid",
          rag: {
            enabled: true,
            provider: rag.provider,
            top_k: rag.top_k,
            hit_count: rag.hits.length,
          },
          answer_hint: answerHintForQuery(
            trimmedQuery,
            rag.sheets.map((sheet) => ({
              sheet: sheet.sheet,
              rows: sheet.rows,
              summary: sheet.summary,
              numbered_item_count: sheet.numbered_item_count,
            }))
          ),
          sheets: rag.sheets.map((sheet) => ({
            sheet: sheet.sheet,
            numbered_item_count: sheet.numbered_item_count,
            summary: sheet.summary,
            rows: sheet.rows.slice(0, limit),
          })),
        };
      }
    } catch (error) {
      console.warn(
        `RAG retrieve failed, falling back to lexical: ${
          error instanceof Error ? error.message : String(error)
        }`
      );
    }
  }

  const targeted = sheetsData.filter((sheet) => queryTargetsSheet(sheet.sheet, trimmedQuery));
  const candidateSheets = targeted.length > 0 ? targeted : sheetsData;
  const sheets: SheetData[] = [];
  let matchCount = 0;

  for (const sheet of candidateSheets) {
    const targetsSheet = targeted.length > 0;
    const rows = targetsSheet
      ? sheet.rows
      : sheet.rows.filter((row) =>
          matchesQuery([...Object.keys(row), ...Object.values(row), sheet.sheet].join(" "), trimmedQuery)
        );

    if (rows.length === 0 && !(targetsSheet && sheet.summary && Object.keys(sheet.summary).length)) {
      continue;
    }

    matchCount += rows.length || (targetsSheet ? 1 : 0);
    const rowLimit = summaryOnly ? 5 : limit;
    sheets.push(withSheetStats(sheet, rows.slice(0, rowLimit), { summaryOnly }));
  }

  return {
    source,
    source_type: sourceType,
    file_name: fileName,
    total_rows: totalRows,
    match_count: matchCount,
    retrieval: "lexical",
    rag: {
      enabled: ragConfig.enabled,
      provider: ragConfig.embedding.provider,
      top_k: ragConfig.topK,
      hit_count: matchCount,
    },
    answer_hint: answerHintForQuery(trimmedQuery, sheets),
    sheets,
  };
}

export function clearWorkbookCache(): void {
  workbookCache = null;
  clearRagIndex();
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
