import { cosineSimilarity } from "./embeddings";
import { getPostgresVectorStore } from "./postgresVectorStore";
import type { IndexedChunk, RagChunk } from "./types";

export type VectorBackendName = "postgres";

export function resolveVectorBackend(): VectorBackendName {
  return "postgres";
}

/**
 * Warm in-memory view of the active RAG index for fast cosine search.
 * Source of truth is Postgres (pgvector).
 */
export class InMemoryVectorStore {
  private chunks: IndexedChunk[] = [];
  private fingerprint = "";
  private provider = "";
  private model = "";

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

  getBackend(): VectorBackendName {
    return "postgres";
  }

  async clear(options?: { persist?: boolean }): Promise<void> {
    this.chunks = [];
    this.fingerprint = "";
    this.provider = "";
    this.model = "";
    if (options?.persist) {
      await getPostgresVectorStore().clear();
    }
  }

  /** Load persisted index when fingerprint/provider/model match. */
  async loadFromDb(fingerprint: string, provider: string, model: string): Promise<boolean> {
    const db = getPostgresVectorStore();
    if (!(await db.hasFingerprint(fingerprint, provider, model))) {
      return false;
    }
    const chunks = await db.loadAll();
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

export const workbookVectorStore = new InMemoryVectorStore();
