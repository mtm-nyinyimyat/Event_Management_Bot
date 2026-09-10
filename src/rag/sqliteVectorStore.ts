import fs from "node:fs";
import path from "node:path";
import Database from "better-sqlite3";
import type { IndexedChunk } from "./types";

function resolveRagDbPath(): string {
  if (process.env.RAG_DB_PATH) {
    return path.resolve(process.env.RAG_DB_PATH);
  }
  if (process.env.CONVERSATIONS_DB_PATH) {
    const conversationsPath = path.resolve(process.env.CONVERSATIONS_DB_PATH);
    return path.join(path.dirname(conversationsPath), "rag_vectors.db");
  }
  return path.resolve(__dirname, "../../src/storage/rag_vectors.db");
}

function embeddingToBuffer(embedding: Float32Array): Buffer {
  return Buffer.from(embedding.buffer, embedding.byteOffset, embedding.byteLength);
}

function bufferToEmbedding(buffer: Buffer, dims: number): Float32Array {
  const copy = Buffer.from(buffer);
  const aligned = new Float32Array(dims);
  aligned.set(new Float32Array(copy.buffer, copy.byteOffset, dims));
  return aligned;
}

/**
 * SQLite-backed vector store for workbook RAG chunks.
 * Source of truth on disk; callers may keep a warm in-memory copy for search.
 */
export class SqliteVectorStore {
  private db: Database.Database;
  private readonly dbPath: string;

  constructor(dbPath = resolveRagDbPath()) {
    this.dbPath = dbPath;
    fs.mkdirSync(path.dirname(dbPath), { recursive: true });
    this.db = new Database(dbPath);
    this.db.pragma("journal_mode = WAL");
    this.initialize();
  }

  private initialize(): void {
    this.db.exec(`
      CREATE TABLE IF NOT EXISTS rag_meta (
        id INTEGER PRIMARY KEY CHECK (id = 1),
        fingerprint TEXT NOT NULL,
        provider TEXT NOT NULL,
        model TEXT NOT NULL,
        dims INTEGER NOT NULL,
        chunk_count INTEGER NOT NULL,
        updated_at TEXT NOT NULL
      );

      CREATE TABLE IF NOT EXISTS rag_chunks (
        chunk_id TEXT PRIMARY KEY,
        sheet TEXT NOT NULL,
        text TEXT NOT NULL,
        row_json TEXT NOT NULL,
        embedding BLOB NOT NULL,
        dims INTEGER NOT NULL
      );

      CREATE INDEX IF NOT EXISTS idx_rag_chunks_sheet ON rag_chunks(sheet);
    `);
  }

  get path(): string {
    return this.dbPath;
  }

  getMeta(): {
    fingerprint: string;
    provider: string;
    model: string;
    dims: number;
    chunk_count: number;
  } | null {
    const row = this.db
      .prepare(
        `SELECT fingerprint, provider, model, dims, chunk_count
         FROM rag_meta WHERE id = 1`
      )
      .get() as
      | {
          fingerprint: string;
          provider: string;
          model: string;
          dims: number;
          chunk_count: number;
        }
      | undefined;
    return row || null;
  }

  hasFingerprint(fingerprint: string, provider: string, model: string): boolean {
    const meta = this.getMeta();
    return (
      !!meta &&
      meta.fingerprint === fingerprint &&
      meta.provider === provider &&
      meta.model === model &&
      meta.chunk_count > 0
    );
  }

  loadAll(): IndexedChunk[] {
    const rows = this.db
      .prepare(
        `SELECT chunk_id, sheet, text, row_json, embedding, dims
         FROM rag_chunks`
      )
      .all() as Array<{
      chunk_id: string;
      sheet: string;
      text: string;
      row_json: string;
      embedding: Buffer;
      dims: number;
    }>;

    return rows.map((row) => ({
      id: row.chunk_id,
      sheet: row.sheet,
      text: row.text,
      row: JSON.parse(row.row_json) as Record<string, string>,
      embedding: bufferToEmbedding(row.embedding, row.dims),
    }));
  }

  replaceAll(
    chunks: IndexedChunk[],
    meta: { fingerprint: string; provider: string; model: string }
  ): void {
    const dims = chunks[0]?.embedding.length || 0;
    const insert = this.db.prepare(
      `INSERT INTO rag_chunks (chunk_id, sheet, text, row_json, embedding, dims)
       VALUES (@chunk_id, @sheet, @text, @row_json, @embedding, @dims)`
    );
    const upsertMeta = this.db.prepare(
      `INSERT INTO rag_meta (id, fingerprint, provider, model, dims, chunk_count, updated_at)
       VALUES (1, @fingerprint, @provider, @model, @dims, @chunk_count, @updated_at)
       ON CONFLICT(id) DO UPDATE SET
         fingerprint = excluded.fingerprint,
         provider = excluded.provider,
         model = excluded.model,
         dims = excluded.dims,
         chunk_count = excluded.chunk_count,
         updated_at = excluded.updated_at`
    );

    const tx = this.db.transaction(() => {
      this.db.exec("DELETE FROM rag_chunks");
      for (const chunk of chunks) {
        insert.run({
          chunk_id: chunk.id,
          sheet: chunk.sheet,
          text: chunk.text,
          row_json: JSON.stringify(chunk.row),
          embedding: embeddingToBuffer(chunk.embedding),
          dims: chunk.embedding.length,
        });
      }
      upsertMeta.run({
        fingerprint: meta.fingerprint,
        provider: meta.provider,
        model: meta.model,
        dims,
        chunk_count: chunks.length,
        updated_at: new Date().toISOString(),
      });
    });

    tx();
  }

  clear(): void {
    const tx = this.db.transaction(() => {
      this.db.exec("DELETE FROM rag_chunks");
      this.db.exec("DELETE FROM rag_meta");
    });
    tx();
  }
}

let sqliteStore: SqliteVectorStore | null = null;

export function getSqliteVectorStore(): SqliteVectorStore {
  if (!sqliteStore) {
    sqliteStore = new SqliteVectorStore();
  }
  return sqliteStore;
}
