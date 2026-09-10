import { cosineSimilarity } from "./embeddings";
import type { IndexedChunk, RagChunk } from "./types";

export class InMemoryVectorStore {
  private chunks: IndexedChunk[] = [];
  private fingerprint = "";

  get size(): number {
    return this.chunks.length;
  }

  getFingerprint(): string {
    return this.fingerprint;
  }

  clear(): void {
    this.chunks = [];
    this.fingerprint = "";
  }

  replaceAll(chunks: IndexedChunk[], fingerprint: string): void {
    this.chunks = chunks;
    this.fingerprint = fingerprint;
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
