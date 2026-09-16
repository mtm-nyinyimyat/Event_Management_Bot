import { cosineSimilarity } from "./embeddings";
import { getPostgresVectorStore } from "./postgresVectorStore";
import type { IndexedChunk, RagChunk } from "./types";

export type VectorBackendName = "postgres";

export function resolveVectorBackend(): VectorBackendName {
  return "postgres";
}

/**
 * Warm in-memory view of one document's RAG index for fast cosine search.
 * Source of truth is Postgres (pgvector), keyed by document_id.
 */
export class InMemoryVectorStore {
  private chunks: IndexedChunk[] = [];
  private fingerprint = "";
  private provider = "";
  private model = "";
  private documentId = "";

  constructor(documentId = "") {
    this.documentId = documentId;
  }

  get size(): number {
    return this.chunks.length;
  }

  getFingerprint(): string {
    return this.fingerprint;
  }

  getProvider(): string {
    return this.provider;
  }

  getModel(): string {
    return this.model;
  }

  getDocumentId(): string {
    return this.documentId;
  }

  getBackend(): VectorBackendName {
    return "postgres";
  }

  async clear(options?: { persist?: boolean }): Promise<void> {
    this.chunks = [];
    this.fingerprint = "";
    this.provider = "";
    this.model = "";
    if (options?.persist) {
      await getPostgresVectorStore().clear(this.documentId || undefined);
    }
  }

  /** Load persisted index when fingerprint/provider/model match. */
  async loadFromDb(fingerprint: string, provider: string, model: string): Promise<boolean> {
    if (!this.documentId) {
      return false;
    }
    const db = getPostgresVectorStore();
    if (!(await db.hasFingerprint(fingerprint, provider, model, this.documentId))) {
      return false;
    }
    const chunks = await db.loadAll(this.documentId);
    if (!chunks.length) {
      return false;
    }
    this.chunks = chunks;
    this.fingerprint = fingerprint;
    this.provider = provider;
    this.model = model;
    return true;
  }

  async replaceAll(
    chunks: IndexedChunk[],
    fingerprint: string,
    meta: { provider: string; model: string; persist?: boolean }
  ): Promise<void> {
    this.chunks = chunks;
    this.fingerprint = fingerprint;
    this.provider = meta.provider;
    this.model = meta.model;

    if (meta.persist === false || !this.documentId) {
      return;
    }

    await getPostgresVectorStore().replaceAll(chunks, {
      fingerprint,
      provider: meta.provider,
      model: meta.model,
      documentId: this.documentId,
    });
  }

  search(queryEmbedding: Float32Array, topK: number): Array<{ chunk: RagChunk; score: number }> {
    if (this.chunks.length === 0 || topK <= 0) {
      return [];
    }

    const scored = this.chunks.map((chunk) => ({
      chunk,
      score: cosineSimilarity(queryEmbedding, chunk.embedding),
    }));

    scored.sort((a, b) => b.score - a.score);
    return scored.slice(0, topK).map(({ chunk, score }) => ({
      chunk: {
        id: chunk.id,
        sheet: chunk.sheet,
        text: chunk.text,
        row: chunk.row,
      },
      score,
    }));
  }
}

const stores = new Map<string, InMemoryVectorStore>();

function scopeKey(documentId?: string): string {
  return documentId?.trim() || "__global__";
}

export function getWorkbookVectorStore(documentId?: string): InMemoryVectorStore {
  const key = scopeKey(documentId);
  let store = stores.get(key);
  if (!store) {
    store = new InMemoryVectorStore(key === "__global__" ? "" : key);
    stores.set(key, store);
  }
  return store;
}

/** @deprecated Prefer getWorkbookVectorStore(documentId) */
export const workbookVectorStore = getWorkbookVectorStore();

export async function clearAllWorkbookVectorStores(options?: {
  persist?: boolean;
  documentId?: string;
}): Promise<void> {
  if (options?.documentId) {
    const store = getWorkbookVectorStore(options.documentId);
    await store.clear({ persist: options.persist === true });
    stores.delete(scopeKey(options.documentId));
    return;
  }

  const keys = [...stores.keys()];
  for (const key of keys) {
    const store = stores.get(key);
    if (!store) {
      continue;
    }
    await store.clear({ persist: options?.persist === true });
  }
  stores.clear();
  if (options?.persist) {
    await getPostgresVectorStore().clear();
  }
}
