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
  retrieval?: "hybrid" | "lexical" | "overview" | "beverage";
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
  /** Bump when sheet parsing changes so stale JSON caches are re-downloaded. */
  parseVersion?: number;
}

export interface WorkbookLookupOptions {
  conversationId?: string;
  /** RAG scope — preferred over conversationId for vector index lookup. */
  documentId?: string;
  userId?: string;
  requesterName?: string;
}
