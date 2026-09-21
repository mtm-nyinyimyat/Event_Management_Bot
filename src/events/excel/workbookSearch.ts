import { ensureWorkbookIndexed, retrieveHybrid } from "../../rag";
import { getRagConfig } from "../../rag/config";
import { describeGraphExcelConfig } from "../graphExcelClient";
import {
  normalizeCacheQuery,
  workbookSearchCache,
} from "../../utils/queryCache";
import { withSheetStats } from "./sheetParse";
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
  if (/\b(seat|seating|sit|table)\b/.test(q)) {
    extras.push("Table", "စားပွဲ");
  }
  if (/\b(drink|beverage|beer|juice|cocktail)\b/.test(q)) {
    extras.push("Beverage", "BEER", "JUICE");
  }
  return extras.length ? `${query} ${extras.join(" ")}` : query;
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
  const retrievalQuery = enrichQueryWithRequester(trimmedQuery, options.requesterName);

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
        const detailSheets = rag.sheets.map((sheet) => {
          const rows = sheet.rows.slice(0, limit);
          const hasDetailRows = rows.some((row) => row.__chunk_type !== "summary");
          return {
            sheet: sheet.sheet,
            numbered_item_count: hasDetailRows ? rows.length : sheet.numbered_item_count,
            summary: hasDetailRows ? undefined : sheet.summary,
            rows,
          };
        });
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
    sheets.push(withSheetStats(sheet, rows.slice(0, limit), { omitSummary: true }));
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
    answer_hint:
      answerHintForQuery(retrievalQuery, sheets) ||
      "Answer only from the matched rows. Do not invent values not present.",
    sheets,
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
