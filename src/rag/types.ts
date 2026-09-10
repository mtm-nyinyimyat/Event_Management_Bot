export type EmbeddingProviderName = "local" | "openai" | "azure";

export type RagRow = Record<string, string>;
export type RagSummary = Record<string, string | number | Record<string, number>>;

export interface RagChunk {
  id: string;
  sheet: string;
  text: string;
  row: RagRow;
}

export interface IndexedChunk extends RagChunk {
  embedding: Float32Array;
}

export interface RagHit {
  chunk: RagChunk;
  score: number;
  lexicalScore: number;
  vectorScore: number;
}

export interface RagSheetResult {
  sheet: string;
  numbered_item_count?: number;
  summary?: RagSummary;
  rows: RagRow[];
}

export interface RagSearchResult {
  enabled: boolean;
  provider: EmbeddingProviderName;
  match_count: number;
  top_k: number;
  retrieval: "hybrid" | "lexical" | "vector";
  hits: Array<{
    sheet: string;
    score: number;
    lexical_score: number;
    vector_score: number;
    row: RagRow;
  }>;
  sheets: RagSheetResult[];
}
