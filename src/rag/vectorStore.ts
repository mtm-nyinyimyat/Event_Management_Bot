import { cosineSimilarity } from "./embeddings";
import { getPostgresVectorStore } from "./postgresVectorStore";
import type { IndexedChunk, RagChunk } from "./types";

export type VectorBackendName = "postgres";

export function resolveVectorBackend(): VectorBackendName {
  return "postgres";
}

/**
 * Warm in-memory view of one conversation's RAG index for fast cosine search.
 * Source of truth is Postgres (pgvector).
 */
export class InMemoryVectorStore {
  private chunks: IndexedChunk[] = [];
  private fingerprint = "";
  private provider = "";
  private model = "";
  private conversationId = "";

  constructor(conversationId = "") {
    this.conversationId = conversationId;
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

  getConversationId(): string {
    return this.conversationId;
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
      await getPostgresVectorStore().clear(this.conversationId || undefined);
    }
  }

  /** Load persisted index when fingerprint/provider/model match. */
  async loadFromDb(fingerprint: string, provider: string, model: string): Promise<boolean> {
    const db = getPostgresVectorStore();
    if (!(await db.hasFingerprint(fingerprint, provider, model, this.conversationId || undefined))) {
      return false;
    }
    const chunks = await db.loadAll(this.conversationId || undefined);
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

    if (meta.persist === false) {
      return;
    }

    await getPostgresVectorStore().replaceAll(chunks, {
      fingerprint,
      provider: meta.provider,
      model: meta.model,
      conversationId: this.conversationId || undefined,
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

function scopeKey(conversationId?: string): string {
  return conversationId?.trim() || "__global__";
}

export function getWorkbookVectorStore(conversationId?: string): InMemoryVectorStore {
  const key = scopeKey(conversationId);
  let store = stores.get(key);
  if (!store) {
    store = new InMemoryVectorStore(key === "__global__" ? "" : key);
    stores.set(key, store);
  }
  return store;
}

/** @deprecated Prefer getWorkbookVectorStore(conversationId) */
export const workbookVectorStore = getWorkbookVectorStore();

export async function clearAllWorkbookVectorStores(options?: {
  persist?: boolean;
  conversationId?: string;
}): Promise<void> {
  if (options?.conversationId) {
    const store = getWorkbookVectorStore(options.conversationId);
    await store.clear({ persist: options.persist === true });
    stores.delete(scopeKey(options.conversationId));
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
