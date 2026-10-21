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

const DRINK_ALIASES: Array<{ token: "JUICE" | "BEER" | "COCKTAIL"; pattern: RegExp }> = [
  { token: "JUICE", pattern: /\bjuice\b|ဖျော်ရည်/i },
  { token: "BEER", pattern: /\bbeer\b/i },
  { token: "COCKTAIL", pattern: /\bcocktail\b|\bcocktial\b/i },
];

function requestedDrinks(query: string): Array<"JUICE" | "BEER" | "COCKTAIL"> {
  return DRINK_ALIASES.filter((drink) => drink.pattern.test(query)).map((drink) => drink.token);
}

/** True when the user asks about beverage value ALL (not "list all who drink beer"). */
function isBeverageAllQuery(query: string): boolean {
  const q = normalizeText(query);
  const drinkContext =
    /\b(drink|drinks|beverage|beverages)\b/.test(q) ||
    query.includes("အချိုရည်") ||
    query.includes("သောက်");
  if (!drinkContext) {
    return false;
  }
  // Specific drinks win over ALL.
  if (requestedDrinks(query).length > 0) {
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

function drinkTokens(value: string): Set<"JUICE" | "BEER" | "COCKTAIL"> {
  const upper = value.toUpperCase().replace(/COCKTIAL/g, "COCKTAIL");
  const tokens = new Set<"JUICE" | "BEER" | "COCKTAIL">();
  // Literal "ALL" means every drink type in workbook scans for juice/beer/cocktail.
  if (/\bALL\b/.test(upper) || upper.trim() === "အားလုံး") {
    tokens.add("JUICE");
    tokens.add("BEER");
    tokens.add("COCKTAIL");
    return tokens;
  }
  if (upper.includes("JUICE")) {
    tokens.add("JUICE");
  }
  if (upper.includes("BEER")) {
    tokens.add("BEER");
  }
  if (upper.includes("COCKTAIL")) {
    tokens.add("COCKTAIL");
  }
  return tokens;
}

function isBeverageAllCell(value: string): boolean {
  const upper = value.trim().toUpperCase().replace(/\s+/g, " ");
  return upper === "ALL" || upper === "အားလုံး";
}

function rowMatchesDrinks(
  row: EventRecord,
  drinks: Array<"JUICE" | "BEER" | "COCKTAIL">,
  exclusive: boolean
): boolean {
  if (!personLabel(row)) {
    return false;
  }
  const cell = beverageCell(row);
  if (!cell) {
    return false;
  }
  const tokens = drinkTokens(cell);
  if (!tokens.size) {
    return false;
  }
  if (exclusive && drinks.length === 1) {
    // "only JUICE" should not include people whose cell is ALL.
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
 * Hybrid RAG only returns a sample, which drops people who only chose JUICE/BEER/ALL.
 */
function scanBeverageMatches(sheets: SheetData[], query: string): SheetData[] | null {
  const wantsAll = isBeverageAllQuery(query);
  const drinks = requestedDrinks(query);
  if (!wantsAll && !drinks.length) {
    return null;
  }
  const exclusive = isExclusiveDrinkQuery(query);
  const matched: SheetData[] = [];

  for (const sheet of sheets) {
    const rows = sheet.rows.filter((row) =>
      wantsAll ? rowMatchesBeverageAll(row) : rowMatchesDrinks(row, drinks, exclusive)
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

/** Plate / ferry ids mentioned in the question (e.g. 4N-9173). */
function extractFerryKeysFromQuery(query: string): string[] {
  const keys = new Set<string>();
  for (const match of query.match(/\b\d+[A-Za-z]-?\d+\b/g) || []) {
    keys.add(normalizeFerryKey(match));
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

/** Map free-text questions (EN/Burmese) to checkmark columns that exist in the workbook. */
function checkmarkColumnTargets(sheets: SheetData[], query: string): string[] {
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
    const columns = targets.filter((column) => sheet.summary?.checkmark_columns?.[column]);
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
    const matchCount = beverageSheets.reduce((sum, sheet) => sum + sheet.rows.length, 0);
    const exclusive = isExclusiveDrinkQuery(trimmedQuery);
    const wantsAll = isBeverageAllQuery(trimmedQuery);
    const drinks = wantsAll ? "ALL" : requestedDrinks(trimmedQuery).join(" + ");
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

  if (!summaryOnly && ragConfig.enabled && active) {
    try {
      let documentId = options.documentId;
      if (!documentId && options.conversationId) {
        const { getEventSession } = await import("../eventSession.js");
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
