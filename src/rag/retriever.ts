import type { SheetData } from "../events/excelStore";
import { chunkWorkbookSheets, dynamicRow } from "./chunker";
import {
  createEmbeddingClient,
  embedInBatches,
  type EmbeddingClient,
  type EmbeddingConfig,
} from "./embeddings";
import type { EmbeddingProviderName, IndexedChunk, RagHit, RagSearchResult } from "./types";
import { clearAllWorkbookVectorStores, getWorkbookVectorStore } from "./vectorStore";

export interface RagRuntimeConfig {
  enabled: boolean;
  topK: number;
  hybridAlpha: number;
  minScore: number;
  embedding: EmbeddingConfig;
}

let embeddingClient: EmbeddingClient | null = null;
const indexingPromises = new Map<string, Promise<void>>();

function normalizeText(value: string): string {
  return value.toLocaleLowerCase("my").normalize("NFC");
}

function significantTerms(query: string): string[] {
  const stop = new Set([
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

  return normalizeText(query)
    .split(/\s+/)
    .map((term) => term.trim())
    .filter((term) => term.length > 0 && !stop.has(term));
}

function lexicalScore(text: string, query: string): number {
  const haystack = normalizeText(text);
  const normalizedQuery = normalizeText(query).trim();
  if (!normalizedQuery) {
    return 0;
  }

  if (haystack.includes(normalizedQuery)) {
    return 1;
  }

  const terms = significantTerms(query);
  if (terms.length === 0) {
    return 0;
  }

  const hits = terms.filter((term) => haystack.includes(term)).length;
  return hits / terms.length;
}

function hashText(value: string): number {
  let hash = 2166136261;
  for (let i = 0; i < value.length; i += 1) {
    hash ^= value.charCodeAt(i);
    hash = Math.imul(hash, 16777619);
  }
  return hash >>> 0;
}

/** Content-based fingerprint (stable across process restarts / workbook TTL reloads). */
function workbookFingerprint(
  sheets: SheetData[],
  source: string,
  embeddingModel: string
): string {
  const chunks = chunkWorkbookSheets(sheets);
  let contentHash = 2166136261;
  for (const chunk of chunks) {
    contentHash ^= hashText(`${chunk.id}\0${chunk.text}`);
    contentHash = Math.imul(contentHash, 16777619);
  }
  return `${source}::${embeddingModel}::${chunks.length}::${contentHash >>> 0}`;
}

function getClient(config: RagRuntimeConfig): EmbeddingClient {
  if (!embeddingClient || embeddingClient.provider !== config.embedding.provider) {
    embeddingClient = createEmbeddingClient(config.embedding);
  }
  return embeddingClient;
}

function scopeKey(conversationId?: string): string {
  return conversationId?.trim() || "__global__";
}

export function clearRagIndex(options?: { persist?: boolean; conversationId?: string }): void {
  void clearRagIndexAsync(options);
}

export async function clearRagIndexAsync(options?: {
  persist?: boolean;
  conversationId?: string;
}): Promise<void> {
  await clearAllWorkbookVectorStores({
    persist: options?.persist === true,
    conversationId: options?.conversationId,
  });
  if (options?.conversationId) {
    indexingPromises.delete(scopeKey(options.conversationId));
  } else {
    indexingPromises.clear();
  }
}

export async function ensureWorkbookIndexed(
  sheets: SheetData[],
  meta: { source: string; loadedAt: number; conversationId?: string },
  config: RagRuntimeConfig
): Promise<void> {
  const client = getClient(config);
  const fingerprint = workbookFingerprint(sheets, meta.source, client.model);
  const store = getWorkbookVectorStore(meta.conversationId);
  const backend = store.getBackend();
  const key = scopeKey(meta.conversationId);

  if (
    store.getFingerprint() === fingerprint &&
    store.getProvider() === client.provider &&
    store.getModel() === client.model &&
    store.size > 0
  ) {
    return;
  }

  // Prefer persisted DB index over re-embedding
  if (await store.loadFromDb(fingerprint, client.provider, client.model)) {
    console.debug(
      `🗃️ RAG index loaded from ${backend} (${store.size} chunks, model=${client.model}, scope=${key})`
    );
    return;
  }

  const existing = indexingPromises.get(key);
  if (existing) {
    await existing;
    if (store.getFingerprint() === fingerprint && store.size > 0) {
      return;
    }
  }

  const indexingPromise = (async () => {
    const chunks = chunkWorkbookSheets(sheets);
    const embeddings = await embedInBatches(
      client,
      chunks.map((chunk) => chunk.text)
    );

    const indexed: IndexedChunk[] = chunks.map((chunk, index) => ({
      ...chunk,
      embedding: embeddings[index],
    }));

    await store.replaceAll(indexed, fingerprint, {
      provider: client.provider,
      model: client.model,
      persist: true,
    });
    console.debug(
      `💾 RAG index embedded and saved to ${backend} (${indexed.length} chunks, model=${client.model}, scope=${key})`
    );
  })();

  indexingPromises.set(key, indexingPromise);

  try {
    await indexingPromise;
  } finally {
    indexingPromises.delete(key);
  }
}

export async function retrieveHybrid(
  query: string,
  sheets: SheetData[],
  config: RagRuntimeConfig,
  conversationId?: string
): Promise<RagSearchResult> {
  const topK = Math.max(1, Math.min(config.topK, 50));
  const client = getClient(config);
  const store = getWorkbookVectorStore(conversationId);
  const [queryEmbedding] = await client.embed([query]);
  const vectorHits = store.search(queryEmbedding, Math.max(topK * 3, topK));

  const merged = new Map<string, RagHit>();

  for (const hit of vectorHits) {
    const lex = lexicalScore(hit.chunk.text, query);
    const score = config.hybridAlpha * hit.score + (1 - config.hybridAlpha) * lex;
    merged.set(hit.chunk.id, {
      chunk: hit.chunk,
      score,
      lexicalScore: lex,
      vectorScore: hit.score,
    });
  }

  // Ensure strong lexical matches are not missed by vector-only ranking
  for (const sheet of sheets) {
    sheet.rows.forEach((row, index) => {
      const text = `Sheet: ${sheet.sheet} | ${Object.entries(row)
        .filter(([key, value]) => !key.startsWith("__") && !key.startsWith("Column_") && !!value?.trim())
        .map(([k, v]) => `${k}: ${v}`)
        .join(" | ")}`;
      const lex = lexicalScore(text, query);
      if (lex < 0.34) {
        return;
      }
      const id = `${sheet.sheet}::row::${index}`;
      const existing = merged.get(id);
      const vectorScore = existing?.vectorScore ?? 0;
      const score = config.hybridAlpha * vectorScore + (1 - config.hybridAlpha) * lex;
      if (!existing || score > existing.score) {
        merged.set(id, {
          chunk: { id, sheet: sheet.sheet, text, row: dynamicRow(row) },
          score,
          lexicalScore: lex,
          vectorScore,
        });
      }
    });
  }

  const ranked = [...merged.values()]
    .filter((hit) => hit.score >= config.minScore || hit.lexicalScore >= 0.5)
    .sort((a, b) => b.score - a.score)
    .slice(0, topK);

  const bySheet = new Map<
    string,
    {
      sheet: string;
      numbered_item_count?: number;
      summary?: SheetData["summary"];
      rows: Record<string, string>[];
    }
  >();

  for (const hit of ranked) {
    if (hit.chunk.row.__chunk_type === "summary") {
      const sheetMeta = sheets.find((sheet) => sheet.sheet === hit.chunk.sheet);
      if (!sheetMeta) {
        continue;
      }
      if (!bySheet.has(hit.chunk.sheet)) {
        bySheet.set(hit.chunk.sheet, {
          sheet: hit.chunk.sheet,
          numbered_item_count: sheetMeta.numbered_item_count,
          summary: sheetMeta.summary,
          rows: [],
        });
      }
      continue;
    }

    const sheetMeta = sheets.find((sheet) => sheet.sheet === hit.chunk.sheet);
    const bucket =
      bySheet.get(hit.chunk.sheet) ||
      ({
        sheet: hit.chunk.sheet,
        numbered_item_count: sheetMeta?.numbered_item_count,
        summary: sheetMeta?.summary,
        rows: [],
      } as {
        sheet: string;
        numbered_item_count?: number;
        summary?: SheetData["summary"];
        rows: Record<string, string>[];
      });

    bucket.rows.push(dynamicRow(hit.chunk.row));
    bySheet.set(hit.chunk.sheet, bucket);
  }

  return {
    enabled: true,
    provider: client.provider as EmbeddingProviderName,
    match_count: ranked.length,
    top_k: topK,
    retrieval: "hybrid",
    hits: ranked.map((hit) => ({
      sheet: hit.chunk.sheet,
      score: Number(hit.score.toFixed(4)),
      lexical_score: Number(hit.lexicalScore.toFixed(4)),
      vector_score: Number(hit.vectorScore.toFixed(4)),
      row: dynamicRow(hit.chunk.row),
    })),
    sheets: [...bySheet.values()],
  };
}
