import type { IndexedChunk } from "./types";
import { ensureDocumentSchema } from "./documents";
import {
  getPostgresPool,
  parsePgVector,
  toPgVectorLiteral,
} from "../storage/postgres";

const DEFAULT_DIMS = 384;

/**
 * Postgres + pgvector store for workbook RAG chunks (per document).
 */
export class PostgresVectorStore {
  private schemaReadyForDims = new Map<number, Promise<void>>();

  private ensureSchema(dims = DEFAULT_DIMS): Promise<void> {
    const existing = this.schemaReadyForDims.get(dims);
    if (existing) {
      return existing;
    }

    const ready = (async () => {
      await ensureDocumentSchema();
      const pool = getPostgresPool();

      const meta = await pool.query<{ dims: number }>(
        `SELECT dims FROM rag_meta LIMIT 1`
      );
      const existingDims = meta.rows[0]?.dims;
      if (existingDims && existingDims !== dims) {
        await pool.query(`DROP TABLE IF EXISTS rag_chunks`);
        await pool.query(`
          CREATE TABLE rag_chunks (
            document_id TEXT NOT NULL
              REFERENCES rag_documents(id) ON DELETE CASCADE ON UPDATE CASCADE,
            chunk_id TEXT NOT NULL,
            sheet TEXT NOT NULL,
            text TEXT NOT NULL,
            row_json JSONB NOT NULL,
            embedding vector(${dims}) NOT NULL,
            dims INT NOT NULL,
            PRIMARY KEY (document_id, chunk_id)
          );
        `);
        await pool.query(
          `CREATE INDEX IF NOT EXISTS idx_rag_chunks_sheet ON rag_chunks(document_id, sheet)`
        );
        await pool.query(`DELETE FROM rag_meta`);
      }
    })();

    this.schemaReadyForDims.set(dims, ready);
    return ready;
  }

  private requireDocumentId(documentId?: string): string {
    const id = documentId?.trim();
    if (!id) {
      throw new Error("documentId is required for RAG vector operations");
    }
    return id;
  }

  async getMeta(documentId?: string): Promise<{
    fingerprint: string;
    provider: string;
    model: string;
    dims: number;
    chunk_count: number;
  } | null> {
    await this.ensureSchema();
    const id = this.requireDocumentId(documentId);
    const result = await getPostgresPool().query<{
      fingerprint: string;
      provider: string;
      model: string;
      dims: number;
      chunk_count: number;
    }>(
      `SELECT fingerprint, provider, model, dims, chunk_count
       FROM rag_meta WHERE document_id = $1`,
      [id]
    );
    return result.rows[0] || null;
  }

  async hasFingerprint(
    fingerprint: string,
    provider: string,
    model: string,
    documentId?: string
  ): Promise<boolean> {
    const meta = await this.getMeta(documentId);
    return (
      !!meta &&
      meta.fingerprint === fingerprint &&
      meta.provider === provider &&
      meta.model === model &&
      meta.chunk_count > 0
    );
  }

  async loadAll(documentId?: string): Promise<IndexedChunk[]> {
    await this.ensureSchema();
    const id = this.requireDocumentId(documentId);
    const result = await getPostgresPool().query<{
      chunk_id: string;
      sheet: string;
      text: string;
      row_json: Record<string, string> | string;
      embedding: unknown;
      dims: number;
    }>(
      `SELECT chunk_id, sheet, text, row_json, embedding::text AS embedding, dims
       FROM rag_chunks WHERE document_id = $1`,
      [id]
    );

    return result.rows.map((row) => {
      const rowJson =
        typeof row.row_json === "string"
          ? (JSON.parse(row.row_json) as Record<string, string>)
          : row.row_json;
      return {
        id: row.chunk_id,
        sheet: row.sheet,
        text: row.text,
        row: rowJson,
        embedding: parsePgVector(row.embedding, row.dims),
      };
    });
  }

  async replaceAll(
    chunks: IndexedChunk[],
    meta: { fingerprint: string; provider: string; model: string; documentId: string }
  ): Promise<void> {
    const dims = chunks[0]?.embedding.length || DEFAULT_DIMS;
    await this.ensureSchema(dims);
    const documentId = this.requireDocumentId(meta.documentId);
    const client = await getPostgresPool().connect();

    try {
      await client.query("BEGIN");
      await client.query(`DELETE FROM rag_chunks WHERE document_id = $1`, [documentId]);

      for (const chunk of chunks) {
        await client.query(
          `INSERT INTO rag_chunks (document_id, chunk_id, sheet, text, row_json, embedding, dims)
           VALUES ($1, $2, $3, $4, $5::jsonb, $6::vector, $7)`,
          [
            documentId,
            chunk.id,
            chunk.sheet,
            chunk.text,
            JSON.stringify(chunk.row),
            toPgVectorLiteral(chunk.embedding),
            chunk.embedding.length,
          ]
        );
      }

      await client.query(
        `INSERT INTO rag_meta (
           document_id, fingerprint, provider, model, dims, chunk_count, updated_at
         ) VALUES ($1, $2, $3, $4, $5, $6, NOW())
         ON CONFLICT (document_id) DO UPDATE SET
           fingerprint = EXCLUDED.fingerprint,
           provider = EXCLUDED.provider,
           model = EXCLUDED.model,
           dims = EXCLUDED.dims,
           chunk_count = EXCLUDED.chunk_count,
           updated_at = EXCLUDED.updated_at`,
        [documentId, meta.fingerprint, meta.provider, meta.model, dims, chunks.length]
      );

      await client.query("COMMIT");
    } catch (error) {
      await client.query("ROLLBACK");
      throw error;
    } finally {
      client.release();
    }
  }

  async clear(documentId?: string): Promise<void> {
    await this.ensureSchema();
    const pool = getPostgresPool();
    if (documentId) {
      const id = this.requireDocumentId(documentId);
      await pool.query(`DELETE FROM rag_chunks WHERE document_id = $1`, [id]);
      await pool.query(`DELETE FROM rag_meta WHERE document_id = $1`, [id]);
      return;
    }
    await pool.query("DELETE FROM rag_chunks");
    await pool.query("DELETE FROM rag_meta");
  }
}

let postgresStore: PostgresVectorStore | null = null;

export function getPostgresVectorStore(): PostgresVectorStore {
  if (!postgresStore) {
    postgresStore = new PostgresVectorStore();
  }
  return postgresStore;
}
