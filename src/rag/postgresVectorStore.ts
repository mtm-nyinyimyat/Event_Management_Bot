import type { IndexedChunk } from "./types";
import {
  getPostgresPool,
  parsePgVector,
  toPgVectorLiteral,
} from "../storage/postgres";

const DEFAULT_DIMS = 384;
const DEFAULT_CONVERSATION = "__global__";

/**
 * Postgres + pgvector store for workbook RAG chunks (per conversation).
 */
export class PostgresVectorStore {
  private schemaReadyForDims = new Map<number, Promise<void>>();

  private ensureSchema(dims = DEFAULT_DIMS): Promise<void> {
    const existing = this.schemaReadyForDims.get(dims);
    if (existing) {
      return existing;
    }

    const ready = (async () => {
      const pool = getPostgresPool();
      await pool.query(`CREATE EXTENSION IF NOT EXISTS vector`);

      // Prefer conversation-scoped tables. Migrate away from legacy single-index schema.
      await pool.query(`
        CREATE TABLE IF NOT EXISTS rag_meta (
          conversation_id TEXT PRIMARY KEY,
          fingerprint TEXT NOT NULL,
          provider TEXT NOT NULL,
          model TEXT NOT NULL,
          dims INT NOT NULL,
          chunk_count INT NOT NULL,
          updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
        );
      `);

      // Legacy single-row meta (id = 1) → drop and recreate scoped tables if needed
      const legacyMeta = await pool.query(`
        SELECT column_name FROM information_schema.columns
        WHERE table_name = 'rag_meta' AND column_name = 'id'
      `);
      if (legacyMeta.rows.length > 0) {
        await pool.query(`DROP TABLE IF EXISTS rag_chunks`);
        await pool.query(`DROP TABLE IF EXISTS rag_meta`);
        await pool.query(`
          CREATE TABLE rag_meta (
            conversation_id TEXT PRIMARY KEY,
            fingerprint TEXT NOT NULL,
            provider TEXT NOT NULL,
            model TEXT NOT NULL,
            dims INT NOT NULL,
            chunk_count INT NOT NULL,
            updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
          );
        `);
      }

      await pool.query(`
        CREATE TABLE IF NOT EXISTS rag_chunks (
          conversation_id TEXT NOT NULL,
          chunk_id TEXT NOT NULL,
          sheet TEXT NOT NULL,
          text TEXT NOT NULL,
          row_json JSONB NOT NULL,
          embedding vector(${dims}) NOT NULL,
          dims INT NOT NULL,
          PRIMARY KEY (conversation_id, chunk_id)
        );
      `);

      const legacyChunks = await pool.query(`
        SELECT column_name FROM information_schema.columns
        WHERE table_name = 'rag_chunks' AND column_name = 'conversation_id'
      `);
      if (legacyChunks.rows.length === 0) {
        await pool.query(`DROP TABLE IF EXISTS rag_chunks`);
        await pool.query(`
          CREATE TABLE rag_chunks (
            conversation_id TEXT NOT NULL,
            chunk_id TEXT NOT NULL,
            sheet TEXT NOT NULL,
            text TEXT NOT NULL,
            row_json JSONB NOT NULL,
            embedding vector(${dims}) NOT NULL,
            dims INT NOT NULL,
            PRIMARY KEY (conversation_id, chunk_id)
          );
        `);
      }

      await pool.query(`
        CREATE INDEX IF NOT EXISTS idx_rag_chunks_sheet
        ON rag_chunks(conversation_id, sheet);
      `);

      const meta = await pool.query<{ dims: number; conversation_id: string }>(
        `SELECT dims, conversation_id FROM rag_meta LIMIT 1`
      );
      const existingDims = meta.rows[0]?.dims;
      if (existingDims && existingDims !== dims) {
        await pool.query(`DROP TABLE IF EXISTS rag_chunks`);
        await pool.query(`
          CREATE TABLE rag_chunks (
            conversation_id TEXT NOT NULL,
            chunk_id TEXT NOT NULL,
            sheet TEXT NOT NULL,
            text TEXT NOT NULL,
            row_json JSONB NOT NULL,
            embedding vector(${dims}) NOT NULL,
            dims INT NOT NULL,
            PRIMARY KEY (conversation_id, chunk_id)
          );
        `);
        await pool.query(
          `CREATE INDEX IF NOT EXISTS idx_rag_chunks_sheet ON rag_chunks(conversation_id, sheet)`
        );
        await pool.query(`DELETE FROM rag_meta`);
      }
    })();

    this.schemaReadyForDims.set(dims, ready);
    return ready;
  }

  private scope(conversationId?: string): string {
    return conversationId?.trim() || DEFAULT_CONVERSATION;
  }

  async getMeta(conversationId?: string): Promise<{
    fingerprint: string;
    provider: string;
    model: string;
    dims: number;
    chunk_count: number;
  } | null> {
    await this.ensureSchema();
    const result = await getPostgresPool().query<{
      fingerprint: string;
      provider: string;
      model: string;
      dims: number;
      chunk_count: number;
    }>(
      `SELECT fingerprint, provider, model, dims, chunk_count
       FROM rag_meta WHERE conversation_id = $1`,
      [this.scope(conversationId)]
    );
    return result.rows[0] || null;
  }

  async hasFingerprint(
    fingerprint: string,
    provider: string,
    model: string,
    conversationId?: string
  ): Promise<boolean> {
    const meta = await this.getMeta(conversationId);
    return (
      !!meta &&
      meta.fingerprint === fingerprint &&
      meta.provider === provider &&
      meta.model === model &&
      meta.chunk_count > 0
    );
  }

  async loadAll(conversationId?: string): Promise<IndexedChunk[]> {
    await this.ensureSchema();
    const result = await getPostgresPool().query<{
      chunk_id: string;
      sheet: string;
      text: string;
      row_json: Record<string, string> | string;
      embedding: unknown;
      dims: number;
    }>(
      `SELECT chunk_id, sheet, text, row_json, embedding::text AS embedding, dims
       FROM rag_chunks WHERE conversation_id = $1`,
      [this.scope(conversationId)]
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
    meta: { fingerprint: string; provider: string; model: string; conversationId?: string }
  ): Promise<void> {
    const dims = chunks[0]?.embedding.length || DEFAULT_DIMS;
    await this.ensureSchema(dims);
    const conversationId = this.scope(meta.conversationId);
    const client = await getPostgresPool().connect();

    try {
      await client.query("BEGIN");
      await client.query(`DELETE FROM rag_chunks WHERE conversation_id = $1`, [conversationId]);

      for (const chunk of chunks) {
        await client.query(
          `INSERT INTO rag_chunks (conversation_id, chunk_id, sheet, text, row_json, embedding, dims)
           VALUES ($1, $2, $3, $4, $5::jsonb, $6::vector, $7)`,
          [
            conversationId,
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
           conversation_id, fingerprint, provider, model, dims, chunk_count, updated_at
         ) VALUES ($1, $2, $3, $4, $5, $6, NOW())
         ON CONFLICT (conversation_id) DO UPDATE SET
           fingerprint = EXCLUDED.fingerprint,
           provider = EXCLUDED.provider,
           model = EXCLUDED.model,
           dims = EXCLUDED.dims,
           chunk_count = EXCLUDED.chunk_count,
           updated_at = EXCLUDED.updated_at`,
        [conversationId, meta.fingerprint, meta.provider, meta.model, dims, chunks.length]
      );

      await client.query("COMMIT");
    } catch (error) {
      await client.query("ROLLBACK");
      throw error;
    } finally {
      client.release();
    }
  }

  async clear(conversationId?: string): Promise<void> {
    await this.ensureSchema();
    const pool = getPostgresPool();
    if (conversationId) {
      const scope = this.scope(conversationId);
      await pool.query(`DELETE FROM rag_chunks WHERE conversation_id = $1`, [scope]);
      await pool.query(`DELETE FROM rag_meta WHERE conversation_id = $1`, [scope]);
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
