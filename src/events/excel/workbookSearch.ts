import { ensureWorkbookIndexed, retrieveHybrid } from "../../rag";
import { getRagConfig } from "../../rag/config";
import { describeGraphExcelConfig } from "../graphExcelClient";
import {
  normalizeCacheQuery,
  workbookSearchCache,
} from "../../utils/queryCache";
import { isCheckedMark, withSheetStats } from "./sheetParse";
import { expandPersonNamesAgainstWorkbook, scoreEnglishNameSimilarity } from "./nameMatch";
import type {
  EventRecord,
  SheetData,
  WorkbookLookupOptions,
  WorkbookSearchResult,
} from "./types";
import {
  getEventsSource,
  getWorkbookCache,
  resolveExcelPath,
  resolveUploadedWorkbook,
} from "./workbookCache";
import { loadWorkbook, resolveWorkbookOptions } from "./workbookLoad";

const DEFAULT_MAX_RESULTS = 20;

const SEARCH_STOP_WORDS = new Set([
  "a",
  "an",
  "all",
  "are",
  "can",
  "could",
  "count",
  "did",
  "do",
  "does",
  "find",
  "for",
  "give",
  "goes",
  "going",
  "how",
  "i",
  "in",
  "is",
  "item",
  "items",
  "list",
  "many",
  "me",
  "my",
  "of",
  "please",
  "should",
  "show",
  "tell",
  "the",
  "them",
  "to",
  "total",
  "what",
  "when",
  "where",
  "which",
  "who",
  "whose",
  "will",
  "would",
  "about",
]);

function normalizeText(value: string): string {
  return value.toLocaleLowerCase("my").normalize("NFC");
}

function significantTerms(query: string): string[] {
  const terms = normalizeText(query)
    .split(/\s+/)
    .map((term) => term.trim())
    .filter((term) => term.length > 0 && !SEARCH_STOP_WORDS.has(term));

  const expanded: string[] = [];
  for (const term of terms) {
    expanded.push(term);
    if (term.length > 4 && term.endsWith("s") && !term.endsWith("ss")) {
      expanded.push(term.slice(0, -1));
    }
  }
  return [...new Set(expanded)];
}

function expandQueryForRetrieval(query: string): string {
  const q = normalizeText(query);
  const extras: string[] = [];
  if (/\boffice\b/.test(q) || /\bmtm\b/.test(q)) {
    extras.push("ရုံး");
  }
  if (/\bferry\b/.test(q) || /\bshuttle\b/.test(q)) {
    extras.push("ဖယ်ရီ", "ကား");
  }
  if (/\b(when|time|depart|departure|leave|leaving|take\s*off|start)\b/.test(q)) {
    extras.push("ထွက်", "ထွက်ခွာ", "စထွက်", "အချိန်");
  }
  if (/\bpresenter|emcee|\bmc\b/.test(q)) {
    extras.push("Presenter", "တင်ဆက်");
  }
  if (/\b(sing|singing|singer|karaoke)\b/.test(q)) {
    extras.push("Karaoke", "ကာရာအိုကေ", "သီချင်း");
  }
  if (
    /\b(seat|seating|sit|table)\b/.test(q) ||
    q.includes("ထိုင်") ||
    q.includes("စားပွဲ")
  ) {
    extras.push("Table", "Name", "စားပွဲ");
  }
  if (
    /\b(menu|dish|food|appetizer|salad|soup|dessert|main\s*course)\b/.test(q) ||
    q.includes("မီနူး") ||
    q.includes("ဟင်း")
  ) {
    extras.push("Dish", "Category", "Menu");
  }
  if (/\b(agenda|schedule|program|timeline)\b/.test(q) || q.includes("အစီအစဉ်")) {
    extras.push("Detail", "Time", "Start Time", "End Time");
  }
  if (
    /\b(drink|beverage|beer|juice|cocktail)\b/.test(q) ||
    q.includes("သောက်") ||
    q.includes("အဖျော်")
  ) {
    // Column name only — never append specific drink tokens (BEER/JUICE).
    // requestedDrinks() requires every token, so polluting the query breaks full-scan.
    extras.push("Beverage");
  }
  if (
    /\b(who|whom|whose|people|person|member|members|participant|participants)\b/.test(q) ||
    q.includes("ဘယ်သူ") ||
    q.includes("သူတွေ") ||
    q.includes("အဖွဲ့ဝင်") ||
    q.includes("ပါဝင်")
  ) {
    extras.push("Members", "Member name", "Name", "Group");
  }
  return extras.length ? `${query} ${extras.join(" ")}` : query;
}

function isWhoQuery(query: string): boolean {
  const q = normalizeText(query);
  if (isCountOrSummaryQuery(query) && !isListQuery(query)) {
    return false;
  }
  return (
    /\b(who|whom|whose|people|person|member|members|participant|participants)\b/.test(q) ||
    q.includes("ဘယ်သူ") ||
    q.includes("သူတွေ") ||
    q.includes("အဖွဲ့ဝင်") ||
    q.includes("ပါဝင်")
  );
}

function looksLikeInventoryValue(value: string): boolean {
  const text = value.trim();
  if (!text) {
    return false;
  }
  return (
    /=>/.test(text) ||
    /:\s*\d/.test(text) ||
    /\b(buy|dress|yukata|sakura|maypal)\b/i.test(text) ||
    /ပန်ဖို့|ခြေအိတ်|ဆံထုံး|အကျီ|ပုဆိုး/.test(text)
  );
}

function rowLooksLikePersonFact(row: EventRecord): boolean {
  if (row.Members || row["Member name"] || row.Name || row.NAME) {
    return true;
  }
  const value = row.Value || "";
  if (!value) {
    return Object.keys(row).some((key) => /name|member|leader/i.test(key));
  }
  return !looksLikeInventoryValue(value);
}

/** Keep only checkmark totals when dropping the rest of the summary (who-queries / row samples). */
function checkmarkOnlySummary(
  summary: SheetData["summary"]
): SheetData["summary"] {
  if (!summary?.checkmark_columns) {
    return undefined;
  }
  return { checkmark_columns: summary.checkmark_columns };
}

/** For "who is in X" questions, keep people rows and drop costume/rental line-items. */
function preferPeopleRows(sheets: SheetData[], query: string): SheetData[] {
  if (!isWhoQuery(query)) {
    return sheets;
  }
  return sheets
    .map((sheet) => {
      const people = sheet.rows.filter((row) => {
        if (row.__chunk_type === "summary") {
          return false;
        }
        if (looksLikeInventoryValue(row.Value || "")) {
          return false;
        }
        return rowLooksLikePersonFact(row);
      });
      if (!people.length) {
        return {
          ...sheet,
          rows: [],
          numbered_item_count: 0,
          summary: checkmarkOnlySummary(sheet.summary),
        };
      }
      return {
        ...sheet,
        rows: people,
        numbered_item_count: people.length,
        summary: checkmarkOnlySummary(sheet.summary),
      };
    })
    .filter((sheet) => sheet.rows.length > 0);
}

function isSelfReferenceQuery(query: string): boolean {
  const q = normalizeText(query);
  return (
    /\b(i|me|my|myself|mine)\b/.test(q) ||
    q.includes("ကျွန်တော်") ||
    q.includes("ကျွန်မ") ||
    q.includes("ကျွန်ုပ်") ||
    q.includes("ကိုယ်") ||
    q.includes("ငါ")
  );
}

function enrichQueryWithRequester(query: string, requesterName?: string): string {
  let enriched = query;
  const name = (requesterName || "").trim();
  if (name && isSelfReferenceQuery(query) && !normalizeText(query).includes(normalizeText(name))) {
    enriched = `${enriched} ${name}`;
  }
  return expandQueryForRetrieval(enriched);
}

function termInHaystack(haystack: string, term: string): boolean {
  const normalizedHaystack = normalizeText(haystack);
  const normalizedTerm = normalizeText(term).trim();
  if (!normalizedTerm) {
    return true;
  }
  if (normalizedHaystack === normalizedTerm) {
    return true;
  }
  const escaped = normalizedTerm.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  return new RegExp(`(?:^|[^\\p{L}\\p{N}_])${escaped}(?:[^\\p{L}\\p{N}_]|$)`, "u").test(
    normalizedHaystack
  );
}

function matchesQuery(haystack: string, query: string): boolean {
  const normalizedQuery = normalizeText(query).trim();
  if (!normalizedQuery) {
    return true;
  }
  if (termInHaystack(haystack, normalizedQuery)) {
    return true;
  }
  const terms = significantTerms(query);
  return terms.length > 0 && terms.every((term) => termInHaystack(haystack, term));
}

function rowHaystack(row: EventRecord): string {
  return [...Object.keys(row), ...Object.values(row)]
    .filter((value) => typeof value === "string" && value.trim())
    .join(" ");
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
  return significantTerms(query).some(
    (term) => name.includes(term) || term.includes(name)
  );
}

function contentQueryForSheet(sheetName: string, query: string): string {
  const nameTokens = new Set(normalizeText(sheetName).split(/\s+/).filter(Boolean));
  const terms = significantTerms(query).filter((term) => {
    if (nameTokens.has(term)) {
      return false;
    }
    for (const token of nameTokens) {
      if (token.includes(term) || term.includes(token)) {
        return false;
      }
    }
    return true;
  });
  return terms.join(" ");
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

/** Built-in aliases; workbook-discovered drinks are merged in at query time. */
const BUILTIN_DRINK_ALIASES: Array<{ token: string; pattern: RegExp }> = [
  { token: "JUICE", pattern: /\bjuice\b|ဖျော်ရည်/i },
  { token: "BEER", pattern: /\bbeer\b/i },
  { token: "COCKTAIL", pattern: /\bcocktail\b|\bcocktial\b/i },
  { token: "WINE", pattern: /\bwine\b/i },
  { token: "TEA", pattern: /\btea\b/i },
  { token: "WATER", pattern: /\bwater\b/i },
  { token: "SODA", pattern: /\bsoda\b|\bsoft\s*drink\b/i },
];

const MULTI_DRINK_SPLIT = /\s*(?:&|\/|,|\+|and|AND|နဲ့|နှင့်)\s*/;

function beverageCell(row: EventRecord): string {
  for (const [key, value] of Object.entries(row)) {
    if (/beverag|drink/i.test(key) && value?.trim()) {
      return value.trim();
    }
  }
  return "";
}

function personLabel(row: EventRecord): string {
  return (
    row.Name ||
    row.NAME ||
    row["Member name"] ||
    row.Members ||
    ""
  ).trim();
}

function isBeverageAllCell(value: string): boolean {
  const upper = value.trim().toUpperCase().replace(/\s+/g, " ");
  return upper === "ALL" || upper === "အားလုံး";
}

function normalizeDrinkToken(part: string): string {
  return part
    .trim()
    .toUpperCase()
    .replace(/COCKTIAL/g, "COCKTAIL")
    .replace(/[^A-Z0-9\u1000-\u109F]+/gi, " ")
    .trim()
    .split(/\s+/)[0];
}

/** Atomic drink labels present in this workbook (plus builtins). */
function discoverWorkbookDrinks(sheets: SheetData[]): string[] {
  const found = new Set<string>(BUILTIN_DRINK_ALIASES.map((drink) => drink.token));
  for (const sheet of sheets) {
    for (const row of sheet.rows) {
      const cell = beverageCell(row);
      if (!cell || isBeverageAllCell(cell) || /^both$/i.test(cell.trim())) {
        continue;
      }
      for (const part of cell.split(MULTI_DRINK_SPLIT)) {
        const token = normalizeDrinkToken(part);
        if (token && token.length >= 2 && token.length <= 24 && !/^\d+$/.test(token)) {
          found.add(token);
        }
      }
    }
  }
  return [...found];
}

/** Top atomic drinks that appear alone in cells — used to interpret BOTH. */
function mostCommonAtomicDrinks(sheets: SheetData[]): string[] {
  const counts = new Map<string, number>();
  for (const sheet of sheets) {
    for (const row of sheet.rows) {
      const cell = beverageCell(row);
      if (!cell || isBeverageAllCell(cell) || /^both$/i.test(cell.trim())) {
        continue;
      }
      const parts = cell
        .split(MULTI_DRINK_SPLIT)
        .map(normalizeDrinkToken)
        .filter((token) => token.length >= 2 && !/^\d+$/.test(token));
      if (parts.length !== 1) {
        continue;
      }
      counts.set(parts[0], (counts.get(parts[0]) || 0) + 1);
    }
  }
  return [...counts.entries()]
    .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]))
    .map(([token]) => token);
}

function drinkTokens(value: string, catalog: string[], bothPair: string[] = []): Set<string> {
  const upper = value
    .toUpperCase()
    .replace(/COCKTIAL/g, "COCKTAIL")
    .replace(/&/g, " AND ");
  const tokens = new Set<string>();
  if (isBeverageAllCell(value) || /\bALL\b/.test(upper)) {
    for (const drink of catalog) {
      if (!/^(ALL|BOTH)$/i.test(drink)) {
        tokens.add(drink);
      }
    }
    return tokens;
  }
  if (/\bBOTH\b/.test(upper)) {
    const pair =
      bothPair.length >= 2
        ? bothPair.slice(0, 2)
        : ["JUICE", "BEER"].filter((drink) => catalog.includes(drink));
    for (const drink of pair) {
      tokens.add(drink);
    }
  }
  for (const drink of catalog) {
    if (drink.length >= 2 && upper.includes(drink) && !/^(ALL|BOTH)$/i.test(drink)) {
      tokens.add(drink);
    }
  }
  return tokens;
}

function requestedDrinks(query: string, catalog: string[]): string[] {
  const q = normalizeText(query);
  const hits: string[] = [];
  for (const alias of BUILTIN_DRINK_ALIASES) {
    if (alias.pattern.test(query) && catalog.includes(alias.token) && !hits.includes(alias.token)) {
      hits.push(alias.token);
    }
  }
  for (const drink of catalog) {
    const needle = normalizeText(drink);
    if (needle.length >= 2 && (q.includes(needle) || termInHaystack(query, drink))) {
      if (!hits.includes(drink)) {
        hits.push(drink);
      }
    }
  }
  return hits;
}

/** True when the user asks about beverage value ALL (not "list all who drink beer"). */
function isBeverageAllQuery(query: string, catalog: string[]): boolean {
  const q = normalizeText(query);
  const drinkContext =
    /\b(drink|drinks|beverage|beverages)\b/.test(q) ||
    query.includes("အချိုရည်") ||
    query.includes("သောက်");
  if (!drinkContext) {
    return false;
  }
  if (requestedDrinks(query, catalog).length > 0) {
    return false;
  }
  return /\ball\b/.test(q);
}

function isExclusiveDrinkQuery(query: string): boolean {
  const q = normalizeText(query);
  return (
    /\b(only|just|exclusively|solely)\b/.test(q) ||
    q.includes("ပဲ") ||
    q.includes("သီးသန့်") ||
    q.includes("ချည်း")
  );
}

function rowMatchesDrinks(
  row: EventRecord,
  drinks: string[],
  exclusive: boolean,
  catalog: string[],
  bothPair: string[]
): boolean {
  if (!personLabel(row)) {
    return false;
  }
  const cell = beverageCell(row);
  if (!cell) {
    return false;
  }
  const tokens = drinkTokens(cell, catalog, bothPair);
  if (!tokens.size) {
    return false;
  }
  if (exclusive && drinks.length === 1) {
    if (isBeverageAllCell(cell)) {
      return false;
    }
    return tokens.size === 1 && tokens.has(drinks[0]);
  }
  return drinks.every((drink) => tokens.has(drink));
}

function rowMatchesBeverageAll(row: EventRecord): boolean {
  if (!personLabel(row)) {
    return false;
  }
  return isBeverageAllCell(beverageCell(row));
}

/**
 * Beverage questions must scan every participant row.
 * Drink tokens are discovered from the active workbook (not only JUICE/BEER/COCKTAIL).
 */
function scanBeverageMatches(sheets: SheetData[], query: string): SheetData[] | null {
  const catalog = discoverWorkbookDrinks(sheets);
  const bothPair = mostCommonAtomicDrinks(sheets).slice(0, 2);
  const wantsAll = isBeverageAllQuery(query, catalog);
  const drinks = requestedDrinks(query, catalog);
  if (!wantsAll && !drinks.length) {
    return null;
  }
  const exclusive = isExclusiveDrinkQuery(query);
  const matched: SheetData[] = [];

  for (const sheet of sheets) {
    const rows = sheet.rows.filter((row) =>
      wantsAll
        ? rowMatchesBeverageAll(row)
        : rowMatchesDrinks(row, drinks, exclusive, catalog, bothPair)
    );
    if (!rows.length) {
      continue;
    }
    matched.push(
      withSheetStats(sheet, rows, { omitSummary: true, keepCheckmarkColumns: true })
    );
  }

  return matched.length ? matched : null;
}

function rowNo(row: EventRecord): number {
  const raw = String(row["No."] || row.No || "").trim();
  const n = Number(raw);
  return Number.isFinite(n) ? n : Number.POSITIVE_INFINITY;
}

function ferryBlob(row: EventRecord): string {
  return [row.Ferry, row.Driver, row.ferry, row.driver]
    .map((value) => String(value || "").trim())
    .filter(Boolean)
    .join(" ");
}

function normalizeFerryKey(value: string): string {
  return value.toUpperCase().replace(/\s+/g, "");
}

/** Plate / ferry ids mentioned in the question (e.g. 4N-9173) plus free-text ferry titles. */
function extractFerryKeysFromQuery(query: string): string[] {
  const keys = new Set<string>();
  for (const match of query.match(/\b\d+[A-Za-z]-?\d+\b/g) || []) {
    keys.add(normalizeFerryKey(match));
  }
  // Broader plate-like tokens (ABC-12, 12-AB-34).
  for (const match of query.match(/\b[A-Za-z0-9]+-[A-Za-z0-9]+\b/g) || []) {
    if (!/^(table|seat)-/i.test(match)) {
      keys.add(normalizeFerryKey(match));
    }
  }
  for (const match of query.match(/\bferry\s*#?\s*([A-Za-z0-9-]+)\b/gi) || []) {
    const id = match.replace(/^ferry\s*#?\s*/i, "").trim();
    if (id) {
      keys.add(normalizeFerryKey(id));
    }
  }
  return [...keys];
}

function isFerryRosterQuery(query: string): boolean {
  const q = normalizeText(query);
  return (
    /\bferry\b/.test(q) ||
    /\bshuttle\b/.test(q) ||
    /\bdrop(?:ped)?\s*-?\s*off\b/.test(q) ||
    /\bafter me\b/.test(q) ||
    /\bsame ferry\b/.test(q) ||
    /\bdriver\b/.test(q) ||
    query.includes("ဖယ်ရီ") ||
    extractFerryKeysFromQuery(query).length > 0
  );
}

function rowMatchesFerryKey(row: EventRecord, key: string): boolean {
  const blob = normalizeFerryKey(ferryBlob(row));
  if (!blob || !key) {
    return false;
  }
  return blob.includes(key) || key.includes(blob);
}

function findRequesterFerryKeys(
  sheets: SheetData[],
  requesterName?: string
): string[] {
  const name = (requesterName || "").trim();
  if (!name) {
    return [];
  }
  let best: { score: number; keys: string[] } | null = null;
  for (const sheet of sheets) {
    for (const row of sheet.rows) {
      const label = personLabel(row);
      if (!label || !ferryBlob(row)) {
        continue;
      }
      const score = Math.max(
        scoreEnglishNameSimilarity(name, label),
        scoreEnglishNameSimilarity(label, name)
      );
      if (score < 0.72) {
        continue;
      }
      const keys = extractFerryKeysFromQuery(ferryBlob(row));
      if (!keys.length) {
        // Fall back to whole ferry title as a key fragment.
        const ferry = String(row.Ferry || "").trim();
        if (ferry) {
          keys.push(normalizeFerryKey(ferry));
        }
      }
      if (!keys.length) {
        continue;
      }
      if (!best || score > best.score) {
        best = { score, keys };
      }
    }
  }
  return best?.keys || [];
}

/**
 * Ferry / drop-off / "after me" questions need the full passenger list for that ferry.
 * Hybrid RAG only returns a sample (e.g. No 5 without No 6+), which breaks order questions.
 */
function scanFerryMatches(
  sheets: SheetData[],
  query: string,
  requesterName?: string
): SheetData[] | null {
  if (!isFerryRosterQuery(query)) {
    return null;
  }

  let keys = extractFerryKeysFromQuery(query);
  if (!keys.length && (isSelfReferenceQuery(query) || /\bafter me\b/i.test(query))) {
    keys = findRequesterFerryKeys(sheets, requesterName);
  }
  if (!keys.length) {
    // Still ferry-related but no plate — use requester ferry when a name is known.
    keys = findRequesterFerryKeys(sheets, requesterName);
  }
  if (!keys.length) {
    return null;
  }

  const matched: SheetData[] = [];
  for (const sheet of sheets) {
    const rows = sheet.rows
      .filter((row) => personLabel(row) && keys.some((key) => rowMatchesFerryKey(row, key)))
      .sort((a, b) => rowNo(a) - rowNo(b));
    if (!rows.length) {
      continue;
    }
    matched.push(
      withSheetStats(sheet, rows, { omitSummary: true, keepCheckmarkColumns: true })
    );
  }

  return matched.length ? matched : null;
}

const DUTY_FIELD_KEY =
  /description|role|duty|task|position|responsibility|incharge|တာဝန်/i;

/** Generic role fluff — never hardcode specific duties like Camera/Karaoke. */
const ROLE_QUERY_FLUFF = new Set([
  ...SEARCH_STOP_WORDS,
  "volunteer",
  "volunteers",
  "volunteering",
  "role",
  "roles",
  "duty",
  "duties",
  "person",
  "people",
  "member",
  "members",
  "participant",
  "participants",
  "team",
  "staff",
  "assigned",
  "assignment",
  "doing",
  "does",
  "name",
  "names",
]);

function dutyFieldText(row: EventRecord): string {
  const parts: string[] = [];
  for (const [key, value] of Object.entries(row)) {
    if (!value?.trim() || !DUTY_FIELD_KEY.test(key)) {
      continue;
    }
    parts.push(value.trim());
  }
  return parts.join(" ");
}

function expandRoleTerm(term: string): string[] {
  if (term === "man") {
    return ["man", "men"];
  }
  if (term === "men") {
    return ["men", "man"];
  }
  return [term];
}

/** Content-bearing tokens for matching Description/Role (not "who" / "volunteer"). */
function roleTermsFromQuery(query: string): string[] {
  return significantTerms(query).filter((term) => !ROLE_QUERY_FLUFF.has(term) && term.length > 1);
}

function termMatchesDuty(dutyNorm: string, dutyRaw: string, term: string): boolean {
  for (const variant of expandRoleTerm(term)) {
    if (termInHaystack(dutyRaw, variant) || dutyNorm.includes(normalizeText(variant))) {
      return true;
    }
  }
  return false;
}

/**
 * Score how well a row's Description/Role answers the question.
 * High scores = titled duties ("Camera Men", "Karaoke Incharge"); low/0 = fall through.
 */
function scoreDutyRoleMatch(row: EventRecord, query: string): number {
  const duty = dutyFieldText(row);
  if (!duty) {
    return 0;
  }
  const dutyNorm = normalizeText(duty);
  const queryNorm = normalizeText(query);
  const roleTerms = roleTermsFromQuery(query);
  if (!roleTerms.length) {
    return 0;
  }

  if (dutyNorm.length >= 3 && queryNorm.includes(dutyNorm)) {
    return 100;
  }

  const dutyTerms = significantTerms(duty).filter((term) => !ROLE_QUERY_FLUFF.has(term));
  if (dutyTerms.length > 0 && dutyTerms.every((term) => termInHaystack(query, term))) {
    return 95;
  }

  const hits = roleTerms.filter((term) => termMatchesDuty(dutyNorm, duty, term));
  if (hits.length === roleTerms.length && roleTerms.length >= 2) {
    return 80 + hits.length;
  }
  if (hits.length === roleTerms.length && roleTerms.length === 1 && roleTerms[0].length >= 5) {
    // Single long token only when it is the whole duty title (avoids "karaoke" → Karaoke 〇).
    if (dutyNorm === normalizeText(roleTerms[0]) || dutyTerms.length === 1) {
      return 75;
    }
  }
  if (hits.length >= 2 && hits.length >= Math.ceil(roleTerms.length * 0.75)) {
    return 60 + hits.length * 5;
  }
  return 0;
}

/**
 * Match any duty/role by Description (or Role/Duty/…) before checkmark columns.
 * Works for Camera Men, Karaoke Incharge, and future roles without per-question patches.
 */
function scanDutyRoleMatches(sheets: SheetData[], query: string): SheetData[] | null {
  type Scored = { sheet: SheetData; row: EventRecord; score: number };
  const scored: Scored[] = [];

  for (const sheet of sheets) {
    for (const row of sheet.rows) {
      const score = scoreDutyRoleMatch(row, query);
      if (score >= 60) {
        scored.push({ sheet, row, score });
      }
    }
  }

  if (!scored.length) {
    return null;
  }

  const best = Math.max(...scored.map((entry) => entry.score));
  // Prefer strong title matches; keep near-ties so co-incharges both appear.
  const threshold = best >= 90 ? best - 10 : Math.max(60, best - 15);
  const winners = scored.filter((entry) => entry.score >= threshold);

  const bySheet = new Map<string, { sheet: SheetData; rows: EventRecord[] }>();
  for (const entry of winners) {
    const key = entry.sheet.sheet;
    const bucket = bySheet.get(key) || { sheet: entry.sheet, rows: [] };
    if (!bucket.rows.includes(entry.row)) {
      bucket.rows.push(entry.row);
    }
    bySheet.set(key, bucket);
  }

  return [...bySheet.values()].map(({ sheet, rows }) =>
    withSheetStats(sheet, rows, { omitSummary: true, keepCheckmarkColumns: true })
  );
}

function personNameMatches(label: string, name: string): boolean {
  if (!label || !name) {
    return false;
  }
  return (
    scoreEnglishNameSimilarity(label, name) >= 0.72 ||
    termInHaystack(label, name) ||
    termInHaystack(name, label)
  );
}

function findRowsForPerson(sheets: SheetData[], name: string): SheetData[] {
  const matched: SheetData[] = [];
  for (const sheet of sheets) {
    const rows = sheet.rows.filter((row) => personNameMatches(personLabel(row), name));
    if (!rows.length) {
      continue;
    }
    matched.push(
      withSheetStats(sheet, rows, { omitSummary: true, keepCheckmarkColumns: true })
    );
  }
  return matched;
}

/** "What's my volunteer role?" → sender's row (Description), not the Volunteer 〇 list. */
function scanSelfDutyProfile(
  sheets: SheetData[],
  query: string,
  requesterName?: string
): SheetData[] | null {
  const name = (requesterName || "").trim();
  if (!name || !isSelfReferenceQuery(query)) {
    return null;
  }
  const asksRole =
    /\b(volunteer|role|duty|duties|incharge|in[\s-]?charge|assigned)\b|စေတနာ|တာဝန်/i.test(
      query
    );
  if (!asksRole) {
    return null;
  }

  const matched: SheetData[] = [];
  for (const sheet of sheets) {
    const rows = sheet.rows.filter((row) => {
      const label = personLabel(row);
      return Boolean(label && dutyFieldText(row) && personNameMatches(label, name));
    });
    if (!rows.length) {
      continue;
    }
    matched.push(
      withSheetStats(sheet, rows, { omitSummary: true, keepCheckmarkColumns: true })
    );
  }
  return matched.length ? matched : null;
}

/**
 * First-person / named-person facts across the whole workbook (table, drink, ferry seat,
 * participate flags, etc.). Full rows — never a RAG sample.
 */
function scanPersonProfileMatches(
  sheets: SheetData[],
  query: string,
  requesterName?: string
): SheetData[] | null {
  const q = normalizeText(query);
  const wantsSelf =
    isSelfReferenceQuery(query) &&
    /\b(table|seat|seating|sit|beverage|drink|ferry|shuttle|participate|participation|volunteer|role|duty|where|which)\b|စားပွဲ|ထိုင်|သောက်|ဖယ်ရီ|ပါဝင်|စေတနာ|တာဝန်/.test(
      q
    );
  if (!wantsSelf) {
    return null;
  }
  const name = (requesterName || "").trim();
  if (!name) {
    return null;
  }
  const roleAsk = /\b(volunteer|role|duty)\b|စေတနာ|တာဝန်/i.test(query);
  const matched: SheetData[] = [];
  for (const sheet of findRowsForPerson(sheets, name)) {
    const rows = roleAsk
      ? sheet.rows
      : sheet.rows.filter((row) => {
          // Skip multi-person duty Name cells ("A, B, C") for table/drink questions.
          const label = personLabel(row);
          return !/[,;/]/.test(label) || Boolean(row.Table || row.Beverage || ferryBlob(row));
        });
    if (!rows.length) {
      continue;
    }
    matched.push({ ...sheet, rows, numbered_item_count: rows.length });
  }
  return matched.length ? matched : null;
}

function extractTableIdsFromQuery(query: string): string[] {
  const ids = new Set<string>();
  for (const match of query.match(/\b(?:table|seat|卓)[\s_-]*\d+\b/gi) || []) {
    const num = match.replace(/\D+/g, "");
    if (num) {
      ids.add(`table-${num}`);
    }
  }
  if (/\bvip\b/i.test(query)) {
    ids.add("vip");
  }
  return [...ids];
}

function rowTableId(row: EventRecord): string {
  const raw = String(row.Table || row.table || row.Seat || "").trim();
  if (!raw) {
    return "";
  }
  if (/^vip$/i.test(raw)) {
    return "vip";
  }
  const num = raw.match(/\d+/);
  return num ? `table-${num[0]}` : normalizeText(raw);
}

function isSeatingQuery(query: string): boolean {
  const q = normalizeText(query);
  return (
    /\b(seat|seating|sit|sits|sitting|table|tables|mates?|neighbor|neighbours?)\b/.test(q) ||
    q.includes("ထိုင်") ||
    q.includes("စားပွဲ") ||
    extractTableIdsFromQuery(query).length > 0
  );
}

/**
 * Seating / table-mates: complete list for Table-N (or the sender's table).
 * Hybrid RAG previously returned wrong neighbors from adjacent tables.
 */
function scanSeatingMatches(
  sheets: SheetData[],
  query: string,
  requesterName?: string
): SheetData[] | null {
  if (!isSeatingQuery(query)) {
    return null;
  }

  let tableIds = extractTableIdsFromQuery(query);
  if (!tableIds.length && isSelfReferenceQuery(query) && requesterName) {
    for (const sheet of sheets) {
      for (const row of sheet.rows) {
        if (!personNameMatches(personLabel(row), requesterName)) {
          continue;
        }
        const id = rowTableId(row);
        if (id) {
          tableIds = [id];
          break;
        }
      }
      if (tableIds.length) {
        break;
      }
    }
  }
  if (!tableIds.length) {
    return null;
  }

  const idSet = new Set(tableIds);
  const matched: SheetData[] = [];
  for (const sheet of sheets) {
    const rows = sheet.rows.filter((row) => {
      const id = rowTableId(row);
      return Boolean(id && idSet.has(id) && personLabel(row));
    });
    if (!rows.length) {
      continue;
    }
    matched.push(
      withSheetStats(sheet, rows, { omitSummary: true, keepCheckmarkColumns: true })
    );
  }
  return matched.length ? matched : null;
}

function isMenuQuery(query: string): boolean {
  const q = normalizeText(query);
  return (
    /\b(menu|dish|dishes|food|meal|appetizer|appetisers?|salad|soup|dessert|starter|entree|entrée|main\s*course|dinner\s*menu)\b/.test(
      q
    ) ||
    q.includes("မီနူး") ||
    q.includes("ဟင်း") ||
    q.includes("အစားအသောက်")
  );
}

function rowLooksLikeMenu(row: EventRecord): boolean {
  if (row.Dish || row.Category) {
    return true;
  }
  return /menu/i.test(String(row.Section || ""));
}

/**
 * Dinner menu / course questions: return every Dish/Category row (optionally filtered).
 */
function scanMenuMatches(sheets: SheetData[], query: string): SheetData[] | null {
  if (!isMenuQuery(query)) {
    return null;
  }

  const categoryHints = significantTerms(query).filter(
    (term) =>
      !ROLE_QUERY_FLUFF.has(term) &&
      !/^(menu|dish|dishes|food|meal|dinner|list|show)$/i.test(term)
  );

  const matched: SheetData[] = [];
  for (const sheet of sheets) {
    let rows = sheet.rows.filter((row) => rowLooksLikeMenu(row));
    if (!rows.length) {
      continue;
    }
    if (categoryHints.length) {
      const filtered = rows.filter((row) => {
        const blob = normalizeText(
          [row.Category, row.Dish, row.Section].filter(Boolean).join(" ")
        );
        return categoryHints.some(
          (hint) => blob.includes(normalizeText(hint)) || termInHaystack(blob, hint)
        );
      });
      // Keep filter only when it still finds menu rows; otherwise return full menu.
      if (filtered.length) {
        rows = filtered;
      }
    }
    matched.push(
      withSheetStats(sheet, rows, { omitSummary: true, keepCheckmarkColumns: true })
    );
  }
  return matched.length ? matched : null;
}

/** Columns that define groups of rows (full roster needed, not a RAG sample). */
const GROUP_FIELD_KEY =
  /^(table|category|group|section|ferry|dish|location|notice|team)$/i;

const GENERIC_GROUP_VALUES = /^(event|menu|notice|group|team|list|other|n\/?a)$/i;

/**
 * When the question names a concrete group value that exists in the workbook
 * (e.g. a Group title, Category, Notice heading), return every row in that group.
 */
function scanGroupFieldMatches(sheets: SheetData[], query: string): SheetData[] | null {
  type Hit = { sheet: SheetData; row: EventRecord; score: number; value: string };
  const hits: Hit[] = [];
  const queryNorm = normalizeText(query);
  const queryTerms = significantTerms(query);

  for (const sheet of sheets) {
    for (const row of sheet.rows) {
      for (const [key, raw] of Object.entries(row)) {
        if (!GROUP_FIELD_KEY.test(key) || !raw?.trim()) {
          continue;
        }
        // Table/Ferry have dedicated scanners; skip here to avoid double-firing.
        if (/^(table|ferry)$/i.test(key)) {
          continue;
        }
        const value = raw.trim();
        const valueNorm = normalizeText(value);
        // Skip sheet-meta Section:"Event" and other one-word generics that match any "event" Q.
        if (valueNorm.length < 4 || GENERIC_GROUP_VALUES.test(valueNorm)) {
          continue;
        }
        let score = 0;
        if (queryNorm.includes(valueNorm)) {
          // Require a distinctive value (long or multi-token) for substring score 100.
          score =
            valueNorm.length >= 8 || significantTerms(value).length >= 2 ? 100 : 0;
        } else {
          const valueTerms = significantTerms(value).filter((term) => !ROLE_QUERY_FLUFF.has(term));
          if (
            valueTerms.length >= 1 &&
            valueTerms.every((term) => termInHaystack(query, term))
          ) {
            score = valueTerms.length >= 2 ? 90 : 70;
          } else if (queryTerms.length >= 2) {
            const overlap = queryTerms.filter((term) => termMatchesDuty(valueNorm, value, term));
            if (overlap.length >= 2 && overlap.length >= Math.ceil(queryTerms.length * 0.6)) {
              score = 65;
            }
          }
        }
        if (score >= 70) {
          hits.push({ sheet, row, score, value });
        }
      }
    }
  }

  if (!hits.length) {
    return null;
  }

  const best = Math.max(...hits.map((hit) => hit.score));
  const bestValues = new Set(
    hits.filter((hit) => hit.score >= best - 10).map((hit) => normalizeText(hit.value))
  );

  // Expand to every row sharing the winning group value(s) on the same field family.
  const matched: SheetData[] = [];
  for (const sheet of sheets) {
    const rows = sheet.rows.filter((row) =>
      Object.entries(row).some(([key, raw]) => {
        if (!GROUP_FIELD_KEY.test(key) || /^(table|ferry)$/i.test(key) || !raw?.trim()) {
          return false;
        }
        return bestValues.has(normalizeText(raw));
      })
    );
    if (!rows.length) {
      continue;
    }
    matched.push(
      withSheetStats(sheet, rows, { omitSummary: true, keepCheckmarkColumns: true })
    );
  }
  return matched.length ? matched : null;
}

function rowLooksLikeAgenda(row: EventRecord): boolean {
  const keys = Object.keys(row);
  const hasTime = keys.some((key) => /^(start\s*time|end\s*time|time|period|ကြာချိန်)$/i.test(key));
  const hasDetail = keys.some((key) => /^(detail|description|item|activity|program)$/i.test(key));
  return hasTime && (hasDetail || Boolean(row.Detail || row.Description));
}

function isEventDateQuery(query: string): boolean {
  const q = normalizeText(query);
  if (
    /\b(event\s*date|party\s*date|what day|which day|what date|which date|when is (the )?(event|party)|held on|開催)\b/.test(
      q
    )
  ) {
    return true;
  }
  // Burmese: "event က ဘယ်နေ့ လုပ်မှာလဲ", ရက်စွဲ, ကျင်းပမည့်နေ့, …
  return (
    /ဘယ်နေ့|ရက်စွဲ|ကျင်းပ|ပွဲရက်|ပွဲနေ့/.test(query) ||
    (q.includes("နေ့") && /\b(event|party|ပွဲ)\b/.test(q))
  );
}

const DATE_VALUE_RE =
  /(?:\d{4}[/\-.]\d{1,2}[/\-.]\d{1,2})|(?:\d{1,2}[/\-.]\d{1,2}[/\-.]\d{2,4})|(?:\b(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:t(?:ember)?)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)\.?\s+\d{1,2}(?:st|nd|rd|th)?(?:,?\s*\d{4})?)|(?:\d{1,2}(?:st|nd|rd|th)?\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*(?:\s+\d{4})?)|(?:ဇန်နဝါရီ|ဖေဖော်ဝါရီ|မတ်|ဧပြီ|မေ|ဇွန်|ဇူလိုင်|ဩဂုတ်|စက်တင်ဘာ|အောက်တိုဘာ|နိုဝင်ဘာ|ဒီဇင်ဘာ)[^\n]{0,12}\d{0,4}/i;

function cellLooksLikeCalendarDate(text: string): boolean {
  const value = (text || "").trim();
  if (!value || value.length > 160) {
    return false;
  }
  // Pure clock times are not calendar dates.
  if (/^\d{1,2}:\d{2}(\s*[-–~]\s*\d{1,2}:\d{2})?$/.test(value)) {
    return false;
  }
  return DATE_VALUE_RE.test(value);
}

function rowHasCalendarDate(row: EventRecord): string[] {
  const hits: string[] = [];
  for (const [key, value] of Object.entries(row)) {
    if (!value?.trim() || key.startsWith("__")) {
      continue;
    }
    if (cellLooksLikeCalendarDate(value)) {
      hits.push(`${key}: ${value.trim()}`);
      continue;
    }
    // Explicit Date/Day columns / Event meta Field even when the value is free text.
    if (/^(date|day|ရက်စွဲ)$/i.test(key) || (/^field$/i.test(key) && /date|day|ရက်/i.test(value))) {
      hits.push(`${key}: ${value.trim()}`);
    }
  }
  // Sheet-level Event meta rows from Participants (Date / Place / Estimate time).
  if (/^event$/i.test(String(row.Section || "")) && (row.Date || row.Place || row["Estimate time"])) {
    if (row.Date) {
      hits.push(`Date: ${row.Date}`);
    }
    if (row.Place) {
      hits.push(`Place: ${row.Place}`);
    }
    if (row["Estimate time"]) {
      hits.push(`Estimate time: ${row["Estimate time"]}`);
    }
  }
  return hits;
}

/**
 * Event Place / Estimate time / Date from Participants (or any sheet) Event meta rows.
 */
function scanEventMetaFieldMatches(sheets: SheetData[], query: string): SheetData[] | null {
  const q = normalizeText(query);
  const wantsPlace =
    /\b(place|venue|location|where)\b/.test(q) || /နေရာ|နေရာဘယ်|ဘယ်နေရာ|会場/.test(query);
  const wantsEstimate =
    /\b(estimate|duration|how long|length)\b/.test(q) ||
    /ကြာချိန်|ဘယ်လောက်ကြာ|ခန့်မှန်း/.test(query);
  const wantsDate = isEventDateQuery(query);
  if (!wantsPlace && !wantsEstimate && !wantsDate) {
    return null;
  }

  const matched: SheetData[] = [];
  for (const sheet of sheets) {
    const rows = sheet.rows.filter((row) => {
      if (!/^event$/i.test(String(row.Section || ""))) {
        // Also accept plain Date/Place/Estimate columns on any row.
        if (wantsDate && row.Date) {
          return true;
        }
        if (wantsPlace && (row.Place || row.Venue || row.Location)) {
          return true;
        }
        if (wantsEstimate && (row["Estimate time"] || row.Duration || row["ကြာချိန်"])) {
          return true;
        }
        return false;
      }
      if (wantsDate && row.Date) {
        return true;
      }
      if (wantsPlace && row.Place) {
        return true;
      }
      if (wantsEstimate && row["Estimate time"]) {
        return true;
      }
      // Field/Value shape from extractEventMetaRecords
      const field = String(row.Field || "").toLocaleLowerCase("en");
      if (wantsDate && /date|day/.test(field)) {
        return true;
      }
      if (wantsPlace && /place|venue|location/.test(field)) {
        return true;
      }
      if (wantsEstimate && /estimate|duration|time/.test(field)) {
        return true;
      }
      return false;
    });
    if (!rows.length) {
      continue;
    }
    matched.push(
      withSheetStats(sheet, rows, { omitSummary: true, keepCheckmarkColumns: true })
    );
  }
  return matched.length ? matched : null;
}

/**
 * "When is the event / ဘယ်နေ့" — collect Notice/Date/Detail cells with calendar dates
 * across the whole workbook (not a hybrid sample). Always include file_name meta.
 */
function scanEventDateMatches(
  sheets: SheetData[],
  query: string,
  fileName?: string
): SheetData[] | null {
  if (!isEventDateQuery(query)) {
    return null;
  }

  const datedRows: EventRecord[] = [];
  const seen = new Set<string>();
  for (const sheet of sheets) {
    for (const row of sheet.rows) {
      const hits = rowHasCalendarDate(row);
      if (!hits.length) {
        continue;
      }
      const key = hits.join("|");
      if (seen.has(key)) {
        continue;
      }
      seen.add(key);
      datedRows.push({
        ...row,
        __date_hits: hits.join(" · "),
        __sheet: sheet.sheet,
      });
    }
  }

  const meta: EventRecord = {
    Section: "Event",
    File: fileName || "",
    Question: "event date",
  };
  if (datedRows.length) {
    meta.Dates_found = String(datedRows.length);
  } else {
    meta.Note =
      "No calendar date cell found in the workbook rows. Do not invent a date. Tell the user the Excel does not list the event day clearly.";
  }

  return [
    {
      sheet: "Event meta",
      rows: [meta, ...datedRows.slice(0, 40)],
      numbered_item_count: datedRows.length,
      summary: {
        total_rows: datedRows.length + 1,
        file_name: fileName || "",
        date_rows: datedRows.length,
      },
    },
  ];
}

function isAgendaQuery(query: string): boolean {
  const q = normalizeText(query);
  // Event-day questions have their own scanner — don't dump the whole agenda.
  if (isEventDateQuery(query) && !/\b(agenda|schedule|program|timeline)\b|အစီအစဉ်/.test(q)) {
    return false;
  }
  if (
    /\b(agenda|schedule|program|timeline)\b/.test(q) ||
    q.includes("အစီအစဉ်")
  ) {
    return true;
  }
  // Time questions only when they look schedule-related (not ferry departure alone).
  const timeAsk =
    /\b(when|what time|starts?|ends?|begin)\b/.test(q) ||
    q.includes("ဘယ်အချိန်") ||
    q.includes("အချိန်");
  if (!timeAsk) {
    return false;
  }
  if (isFerryRosterQuery(query) || isSeatingQuery(query) || isMenuQuery(query)) {
    return false;
  }
  return true;
}

/**
 * Agenda / schedule: match Detail (or time fields) against the question across timed rows.
 */
function scanAgendaMatches(sheets: SheetData[], query: string): SheetData[] | null {
  if (!isAgendaQuery(query)) {
    return null;
  }

  const terms = significantTerms(query).filter(
    (term) =>
      !ROLE_QUERY_FLUFF.has(term) &&
      !/^(agenda|schedule|program|timeline|time|start|starts|end|ends|begin|when|what)$/i.test(
        term
      )
  );

  type Scored = { sheet: SheetData; row: EventRecord; score: number };
  const scored: Scored[] = [];

  for (const sheet of sheets) {
    for (const row of sheet.rows) {
      if (!rowLooksLikeAgenda(row)) {
        continue;
      }
      const detail = [
        row.Detail,
        row.Description,
        row.Item,
        row.Activity,
        row.Program,
        row.Time,
        row["Start Time"],
        row["End Time"],
      ]
        .filter(Boolean)
        .join(" ");
      const detailNorm = normalizeText(detail);
      if (!detailNorm) {
        continue;
      }

      let score = 0;
      if (!terms.length) {
        // Bare "agenda/schedule" → every timed row (not only numbered).
        score = 50;
      } else {
        const hits = terms.filter(
          (term) => termInHaystack(detail, term) || detailNorm.includes(normalizeText(term))
        );
        if (hits.length === terms.length) {
          score = 90 + hits.length;
        } else if (hits.length >= 2) {
          score = 60 + hits.length * 5;
        } else if (hits.length === 1 && terms.length === 1 && terms[0].length >= 4) {
          score = 70;
        }
      }
      if (score >= 50) {
        scored.push({ sheet, row, score });
      }
    }
  }

  if (!scored.length) {
    return null;
  }

  // Bare agenda list: return all timed rows, not a sample.
  if (!terms.length) {
    const bySheet = new Map<string, { sheet: SheetData; rows: EventRecord[] }>();
    for (const entry of scored) {
      const key = entry.sheet.sheet;
      const bucket = bySheet.get(key) || { sheet: entry.sheet, rows: [] };
      if (!bucket.rows.includes(entry.row)) {
        bucket.rows.push(entry.row);
      }
      bySheet.set(key, bucket);
    }
    return [...bySheet.values()].map(({ sheet, rows }) =>
      withSheetStats(sheet, rows, { omitSummary: true, keepCheckmarkColumns: true })
    );
  }

  const best = Math.max(...scored.map((entry) => entry.score));
  if (best < 70) {
    return null;
  }
  const winners = scored.filter((entry) => entry.score >= best - 15);
  const bySheet = new Map<string, { sheet: SheetData; rows: EventRecord[] }>();
  for (const entry of winners) {
    const key = entry.sheet.sheet;
    const bucket = bySheet.get(key) || { sheet: entry.sheet, rows: [] };
    if (!bucket.rows.includes(entry.row)) {
      bucket.rows.push(entry.row);
    }
    bySheet.set(key, bucket);
  }
  return [...bySheet.values()].map(({ sheet, rows }) =>
    withSheetStats(sheet, rows, { omitSummary: true, keepCheckmarkColumns: true })
  );
}

/** Map free-text questions (EN/Burmese) to checkmark columns that exist in the workbook. */
function checkmarkColumnTargets(sheets: SheetData[], query: string): string[] {
  // Role titles ("Karaoke Incharge") are handled by scanDutyRoleMatches — don't
  // treat them as Karaoke 〇 column hits.
  if (roleTermsFromQuery(query).length >= 2 && /incharge|in[\s-]?charge|men|man|lead|leader/i.test(query)) {
    return [];
  }

  const knownColumns = new Set<string>();
  for (const sheet of sheets) {
    const marks = sheet.summary?.checkmark_columns;
    if (!marks) {
      continue;
    }
    for (const column of Object.keys(marks)) {
      knownColumns.add(column);
    }
  }

  const targets: string[] = [];
  for (const column of knownColumns) {
    if (termInHaystack(query, column)) {
      targets.push(column);
      continue;
    }
    if (/karaoke/i.test(column) && /\b(karaoke|sing|singing|singer|song)\b|သီချင်း/.test(query)) {
      targets.push(column);
      continue;
    }
    if (/volunteer/i.test(column) && /\bvolunteer\b|စေတနာ/.test(query)) {
      targets.push(column);
      continue;
    }
    if (
      /ferry\s*info|early/i.test(column) &&
      /\b(early|ferry\s*info|go\s*early|leave\s*early)\b/i.test(query)
    ) {
      targets.push(column);
    }
  }
  return targets;
}

/**
 * Checkmark questions (Karaoke 〇, Volunteer 〇, …) must scan every row.
 * Hybrid RAG returns an arbitrary sample of × rows and drops the people who are actually checked.
 */
function scanCheckmarkMatches(sheets: SheetData[], query: string): SheetData[] | null {
  const targets = checkmarkColumnTargets(sheets, query);
  if (!targets.length) {
    return null;
  }

  const matched: SheetData[] = [];
  for (const sheet of sheets) {
    const marks = sheet.summary?.checkmark_columns;
    const markMap =
      marks && typeof marks === "object"
        ? (marks as Record<string, { checked?: number; total?: number }>)
        : undefined;
    const columns = targets.filter((column) => Boolean(markMap?.[column]));
    if (!columns.length) {
      continue;
    }
    const rows = sheet.rows.filter((row) => columns.some((column) => isCheckedMark(row[column])));
    if (!rows.length) {
      continue;
    }
    matched.push(
      withSheetStats(sheet, rows, { omitSummary: true, keepCheckmarkColumns: true })
    );
  }

  return matched.length ? matched : null;
}

function resultLimit(maxResults?: number): number {
  const raw = maxResults ?? DEFAULT_MAX_RESULTS;
  return Math.max(1, Math.min(raw, 200));
}

function answerHintForQuery(query: string, sheets: SheetData[]): string | undefined {
  if (!sheets.length) {
    return undefined;
  }
  if (isCountOrSummaryQuery(query) && !isListQuery(query)) {
    return `Prefer sheet.summary / numbered_item_count for totals. Do not dump every row.`;
  }
  const totalRows = sheets.reduce((sum, sheet) => sum + (sheet.rows?.length || 0), 0);
  if (totalRows > 0) {
    return `Answer only from the ${totalRows} matched row(s). Quote cell values exactly. Sheet and column names vary by workbook.`;
  }
  return `No detail rows matched. If a summary is present, use it; otherwise say you could not find that in the event file.`;
}

async function searchWorkbookUncached(
  trimmedQuery: string,
  limit: number,
  options: WorkbookLookupOptions
): Promise<WorkbookSearchResult> {
  const sheetsData = await loadWorkbook(false, options);
  const active =
    resolveUploadedWorkbook(options.conversationId, options.userId) || getWorkbookCache();
  const sourceType = active?.sourceType || getEventsSource();
  const source =
    active?.source ||
    (sourceType === "graph" ? describeGraphExcelConfig() : resolveExcelPath());
  const fileName = active?.fileName;
  const totalRows = sheetsData.reduce((sum, sheet) => sum + sheet.rows.length, 0);
  const retrievalQuery = expandPersonNamesAgainstWorkbook(
    enrichQueryWithRequester(trimmedQuery, options.requesterName),
    sheetsData
  );

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
      answer_hint:
        "Workbook overview. Use each sheet.summary (columns / totals). Call again with keywords for matching rows.",
      sheets,
    };
  }

  const summaryOnly = isCountOrSummaryQuery(trimmedQuery) && !isListQuery(trimmedQuery);
  const ragConfig = getRagConfig();

  // Drink detection must use the original query — retrievalQuery may append
  // synonym tokens that requestedDrinks() would then require together (AND).
  const beverageSheets = scanBeverageMatches(sheetsData, trimmedQuery);
  if (beverageSheets) {
    const catalog = discoverWorkbookDrinks(sheetsData);
    const matchCount = beverageSheets.reduce((sum, sheet) => sum + sheet.rows.length, 0);
    const exclusive = isExclusiveDrinkQuery(trimmedQuery);
    const wantsAll = isBeverageAllQuery(trimmedQuery, catalog);
    const drinks = wantsAll ? "ALL" : requestedDrinks(trimmedQuery, catalog).join(" + ");
    return {
      source,
      source_type: sourceType,
      file_name: fileName,
      total_rows: totalRows,
      match_count: matchCount,
      retrieval: "beverage",
      answer_hint: wantsAll
        ? `Complete list: ${matchCount} people whose beverage is ALL. List every Name. Do not treat this as a sample.`
        : exclusive
          ? `Complete list: ${matchCount} people whose beverage is only ${drinks}. List every Name. Do not treat this as a sample.`
          : `Complete list: ${matchCount} people whose beverage includes ${drinks}. List every Name. Do not treat this as a sample.`,
      sheets: beverageSheets,
    };
  }

  // Checkmark 〇 lists (who sings / volunteers) before duty titles, so "karaoke"
  // isn't answered as Karaoke Incharge. Role titles still win via checkmarkColumnTargets guard.
  const checkmarkSheets = scanCheckmarkMatches(sheetsData, trimmedQuery);
  if (checkmarkSheets) {
    const matchCount = checkmarkSheets.reduce((sum, sheet) => sum + sheet.rows.length, 0);
    const targets = checkmarkColumnTargets(sheetsData, trimmedQuery).join(", ");
    return {
      source,
      source_type: sourceType,
      file_name: fileName,
      total_rows: totalRows,
      match_count: matchCount,
      retrieval: "checkmark",
      answer_hint: `Complete list: ${matchCount} people checked 〇 in [${targets}]. List every Name from the returned rows. Do not treat this as a sample.`,
      sheets: checkmarkSheets,
    };
  }

  const dutySheets = scanDutyRoleMatches(sheetsData, trimmedQuery);
  if (dutySheets) {
    const matchCount = dutySheets.reduce((sum, sheet) => sum + sheet.rows.length, 0);
    return {
      source,
      source_type: sourceType,
      file_name: fileName,
      total_rows: totalRows,
      match_count: matchCount,
      retrieval: "duty",
      answer_hint:
        `Complete duty/role match: ${matchCount} row(s) whose Description/Role matches the question. ` +
        `Name every person and quote their Description/Role exactly. Do not treat this as a sample.`,
      sheets: dutySheets,
    };
  }

  const selfDutySheets = scanSelfDutyProfile(
    sheetsData,
    trimmedQuery,
    options.requesterName
  );
  if (selfDutySheets) {
    const matchCount = selfDutySheets.reduce((sum, sheet) => sum + sheet.rows.length, 0);
    return {
      source,
      source_type: sourceType,
      file_name: fileName,
      total_rows: totalRows,
      match_count: matchCount,
      retrieval: "duty",
      answer_hint:
        `Sender profile: ${matchCount} row(s) for this person. Answer from Description/Role (and related fields). ` +
        `Do not list every Volunteer 〇 unless that is what they asked.`,
      sheets: selfDutySheets,
    };
  }

  const ferrySheets = scanFerryMatches(
    sheetsData,
    trimmedQuery,
    options.requesterName
  );
  if (ferrySheets) {
    const matchCount = ferrySheets.reduce((sum, sheet) => sum + sheet.rows.length, 0);
    return {
      source,
      source_type: sourceType,
      file_name: fileName,
      total_rows: totalRows,
      match_count: matchCount,
      retrieval: "ferry",
      answer_hint:
        `Complete ferry roster: ${matchCount} passengers in No. order (Ferry/Driver fields included). ` +
        `Use every returned row — do not treat this as a sample. ` +
        `For "Nth after me", find the sender's No. then answer with the passenger at No. + N.`,
      sheets: ferrySheets,
    };
  }

  // Whole-workbook structured scans (any sheet) before hybrid RAG samples.
  const seatingSheets = scanSeatingMatches(
    sheetsData,
    trimmedQuery,
    options.requesterName
  );
  if (seatingSheets) {
    const matchCount = seatingSheets.reduce((sum, sheet) => sum + sheet.rows.length, 0);
    return {
      source,
      source_type: sourceType,
      file_name: fileName,
      total_rows: totalRows,
      match_count: matchCount,
      retrieval: "seating",
      answer_hint:
        `Complete seating list: ${matchCount} people at the matched table(s). ` +
        `Name every person. Do not pull names from other tables.`,
      sheets: seatingSheets,
    };
  }

  const menuSheets = scanMenuMatches(sheetsData, trimmedQuery);
  if (menuSheets) {
    const matchCount = menuSheets.reduce((sum, sheet) => sum + sheet.rows.length, 0);
    return {
      source,
      source_type: sourceType,
      file_name: fileName,
      total_rows: totalRows,
      match_count: matchCount,
      retrieval: "menu",
      answer_hint:
        `Complete menu match: ${matchCount} dish row(s). Quote Category and Dish exactly. ` +
        `If they asked for one course, only list that course's dishes.`,
      sheets: menuSheets,
    };
  }

  const personSheets = scanPersonProfileMatches(
    sheetsData,
    trimmedQuery,
    options.requesterName
  );
  if (personSheets) {
    const matchCount = personSheets.reduce((sum, sheet) => sum + sheet.rows.length, 0);
    return {
      source,
      source_type: sourceType,
      file_name: fileName,
      total_rows: totalRows,
      match_count: matchCount,
      retrieval: "person",
      answer_hint:
        `Complete person profile: ${matchCount} row(s) for this sender across the workbook. ` +
        `Answer from Table / Beverage / Ferry / Description / participation fields present.`,
      sheets: personSheets,
    };
  }

  const groupSheets = scanGroupFieldMatches(sheetsData, trimmedQuery);
  if (groupSheets) {
    const matchCount = groupSheets.reduce((sum, sheet) => sum + sheet.rows.length, 0);
    return {
      source,
      source_type: sourceType,
      file_name: fileName,
      total_rows: totalRows,
      match_count: matchCount,
      retrieval: "group",
      answer_hint:
        `Complete group match: ${matchCount} row(s) sharing the Category/Group/Section/… value in the question. ` +
        `Use every returned row.`,
      sheets: groupSheets,
    };
  }

  const eventMetaSheets = scanEventMetaFieldMatches(sheetsData, trimmedQuery);
  if (eventMetaSheets) {
    const matchCount = eventMetaSheets.reduce((sum, sheet) => sum + sheet.rows.length, 0);
    return {
      source,
      source_type: sourceType,
      file_name: fileName,
      total_rows: totalRows,
      match_count: matchCount,
      retrieval: "event_meta",
      answer_hint:
        `Event meta from the workbook (Date / Place / Estimate time). ` +
        `Quote those field values exactly. Use every returned row. ` +
        `This is not an empty date lookup — answer from Place/Estimate/Date as present.`,
      sheets: eventMetaSheets,
    };
  }

  const eventDateSheets = scanEventDateMatches(sheetsData, trimmedQuery, fileName);
  if (eventDateSheets) {
    const matchCount = eventDateSheets.reduce((sum, sheet) => sum + sheet.rows.length, 0);
    return {
      source,
      source_type: sourceType,
      file_name: fileName,
      total_rows: totalRows,
      match_count: matchCount,
      retrieval: "event_date",
      answer_hint:
        `Event-date lookup: quote calendar dates from Notice/Date/Detail / __date_hits exactly. ` +
        `If Note says no calendar date was found, say the Excel does not list the day — do NOT use the unrelated-topic refusal, and do NOT invent a date from the filename year alone.`,
      sheets: eventDateSheets,
    };
  }

  const agendaSheets = scanAgendaMatches(sheetsData, trimmedQuery);
  if (agendaSheets) {
    const matchCount = agendaSheets.reduce((sum, sheet) => sum + sheet.rows.length, 0);
    return {
      source,
      source_type: sourceType,
      file_name: fileName,
      total_rows: totalRows,
      match_count: matchCount,
      retrieval: "agenda",
      answer_hint:
        `Complete agenda/schedule match: ${matchCount} timed row(s). Quote Time/Detail exactly. ` +
        `Use every returned row — do not treat this as a sample.`,
      sheets: agendaSheets,
    };
  }

  if (!summaryOnly && ragConfig.enabled && active) {
    try {
      let documentId = options.documentId;
      if (!documentId && options.conversationId) {
        const { getEventSession } = await import("../eventSessionLazy").then((mod) =>
          mod.importEventSession()
        );
        const session = await getEventSession(options.conversationId);
        documentId = session.documentId || undefined;
      }

      await ensureWorkbookIndexed(
        sheetsData,
        {
          source: active.source,
          loadedAt: active.loadedAt,
          documentId,
        },
        ragConfig
      );
      const rag = await retrieveHybrid(
        retrievalQuery,
        sheetsData,
        {
          ...ragConfig,
          topK: Math.min(limit, ragConfig.topK),
        },
        documentId
      );

        if (rag.sheets.length > 0 || rag.hits.length > 0) {
          const detailSheets = preferPeopleRows(
            rag.sheets.map((sheet) => {
              const rows = sheet.rows.slice(0, limit);
              const hasDetailRows = rows.some((row) => row.__chunk_type !== "summary");
              return {
                sheet: sheet.sheet,
                numbered_item_count: hasDetailRows ? rows.length : sheet.numbered_item_count,
                summary: hasDetailRows
                  ? checkmarkOnlySummary(sheet.summary)
                  : sheet.summary,
                rows,
              };
            }),
            retrievalQuery
          );
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
          answer_hint:
            answerHintForQuery(retrievalQuery, detailSheets) ||
            "Answer only from the matched rows. Sheet/column names vary by workbook.",
          sheets: detailSheets,
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

  const targeted = sheetsData.filter((sheet) =>
    queryTargetsSheet(sheet.sheet, retrievalQuery)
  );
  const candidateSheets = targeted.length > 0 ? targeted : sheetsData;
  const sheets: SheetData[] = [];
  let matchCount = 0;

  for (const sheet of candidateSheets) {
    const contentQuery = contentQueryForSheet(sheet.sheet, retrievalQuery);

    if (summaryOnly) {
      if (!sheet.summary || !Object.keys(sheet.summary).length) {
        continue;
      }
      matchCount += 1;
      sheets.push(withSheetStats(sheet, [], { summaryOnly: true }));
      continue;
    }

    let rows: EventRecord[];
    if (contentQuery) {
      rows = sheet.rows.filter((row) => matchesQuery(rowHaystack(row), contentQuery));
    } else if (isListQuery(trimmedQuery) && targeted.length > 0) {
      rows = sheet.rows.slice(0, limit);
    } else if (targeted.length > 0) {
      if (sheet.summary && Object.keys(sheet.summary).length) {
        matchCount += 1;
        sheets.push(withSheetStats(sheet, [], { summaryOnly: true }));
      }
      continue;
    } else {
      rows = sheet.rows.filter((row) => matchesQuery(rowHaystack(row), retrievalQuery));
    }

    if (!rows.length) {
      continue;
    }

    matchCount += rows.length;
    sheets.push(
      withSheetStats(sheet, rows.slice(0, limit), {
        omitSummary: true,
        keepCheckmarkColumns: true,
      })
    );
  }

  const peopleSheets = preferPeopleRows(sheets, retrievalQuery);
  const finalSheets = isWhoQuery(retrievalQuery) ? peopleSheets : sheets;
  const finalMatchCount = finalSheets.reduce((sum, sheet) => sum + sheet.rows.length, 0);

  return {
    source,
    source_type: sourceType,
    file_name: fileName,
    total_rows: totalRows,
    match_count: isWhoQuery(retrievalQuery) ? finalMatchCount : matchCount,
    retrieval: "lexical",
    rag: {
      enabled: ragConfig.enabled,
      provider: ragConfig.embedding.provider,
      top_k: ragConfig.topK,
      hit_count: isWhoQuery(retrievalQuery) ? finalMatchCount : matchCount,
    },
    answer_hint:
      answerHintForQuery(retrievalQuery, finalSheets) ||
      "Answer only from the matched rows. Do not invent values not present.",
    sheets: finalSheets,
  };
}

export async function searchWorkbook(
  query: string,
  maxResults = DEFAULT_MAX_RESULTS,
  conversationIdOrOptions?: string | WorkbookLookupOptions
): Promise<WorkbookSearchResult> {
  const options = resolveWorkbookOptions(conversationIdOrOptions);
  const limit = resultLimit(maxResults);
  const trimmedQuery = query.trim();
  const cacheKey = [
    "search",
    "v8-multi-event",
    normalizeCacheQuery(trimmedQuery),
    String(limit),
    options.documentId || "",
    options.conversationId || "",
    options.userId || "",
    normalizeCacheQuery(options.requesterName || ""),
  ].join("::");

  const cached = workbookSearchCache.get(cacheKey) as WorkbookSearchResult | undefined;
  if (cached) {
    return {
      ...cached,
      answer_hint: cached.answer_hint
        ? `${cached.answer_hint} (cached)`
        : "Cached workbook lookup result.",
    };
  }

  const result = await searchWorkbookUncached(trimmedQuery, limit, options);
  workbookSearchCache.set(cacheKey, result);
  return result;
}
