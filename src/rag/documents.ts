import { createHash, randomUUID } from "crypto";
import { getPostgresPool } from "../storage/postgres";

export type DocumentSourceType = "upload" | "sharepoint" | "graph" | "local";
export type DocumentStatus = "pending" | "ready" | "syncing" | "error" | "archived";

export interface RagDocumentRecord {
  id: string;
  sourceType: DocumentSourceType;
  sourceUri: string;
  fileName: string | null;
  contentHash: string | null;
  status: DocumentStatus;
  createdAt: string;
  updatedAt: string;
}

export interface UpsertSharepointSourceInput {
  documentId: string;
  driveId?: string | null;
  itemId?: string | null;
  siteId?: string | null;
  webUrl?: string | null;
  subscriptionId?: string | null;
  subscriptionExpiresAt?: Date | string | null;
  deltaLink?: string | null;
  etag?: string | null;
  lastModifiedAt?: Date | string | null;
  lastDeltaSyncAt?: Date | string | null;
}

export interface RagSharepointSourceRecord {
  documentId: string;
  driveId: string | null;
  itemId: string | null;
  siteId: string | null;
  webUrl: string | null;
  subscriptionId: string | null;
  subscriptionExpiresAt: string | null;
  deltaLink: string | null;
  etag: string | null;
  lastModifiedAt: string | null;
  lastDeltaSyncAt: string | null;
}

let schemaReady: Promise<void> | null = null;

/**
 * Ensure document-centric tables exist (also handled by Prisma migrate).
 * Migrates away from legacy conversation-scoped rag_* if still present.
 */
export async function ensureDocumentSchema(): Promise<void> {
  if (!schemaReady) {
    schemaReady = (async () => {
      const pool = getPostgresPool();
      await pool.query(`CREATE EXTENSION IF NOT EXISTS vector`);

      await pool.query(`
        CREATE TABLE IF NOT EXISTS rag_documents (
          id TEXT PRIMARY KEY,
          source_type TEXT NOT NULL CHECK (source_type IN ('upload', 'sharepoint', 'graph', 'local')),
          source_uri TEXT NOT NULL,
          file_name TEXT,
          content_hash TEXT,
          status TEXT NOT NULL DEFAULT 'pending'
            CHECK (status IN ('pending', 'ready', 'syncing', 'error', 'archived')),
          created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
          updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
        );
      `);

      await pool.query(`
        CREATE INDEX IF NOT EXISTS idx_rag_documents_status ON rag_documents(status);
      `);
      await pool.query(`
        CREATE INDEX IF NOT EXISTS idx_rag_documents_content_hash ON rag_documents(content_hash);
      `);

      await pool.query(`
        CREATE TABLE IF NOT EXISTS rag_sharepoint_sources (
          document_id TEXT PRIMARY KEY
            REFERENCES rag_documents(id) ON DELETE CASCADE ON UPDATE CASCADE,
          drive_id TEXT,
          item_id TEXT,
          site_id TEXT,
          web_url TEXT,
          subscription_id TEXT,
          subscription_expires_at TIMESTAMPTZ,
          delta_link TEXT,
          etag TEXT,
          last_modified_at TIMESTAMPTZ,
          last_delta_sync_at TIMESTAMPTZ,
          updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
        );
      `);

      const deltaSyncCol = await pool.query(`
        SELECT 1 FROM information_schema.columns
        WHERE table_name = 'rag_sharepoint_sources' AND column_name = 'last_delta_sync_at'
        LIMIT 1
      `);
      if (deltaSyncCol.rows.length === 0) {
        await pool.query(`
          ALTER TABLE rag_sharepoint_sources
          ADD COLUMN last_delta_sync_at TIMESTAMPTZ
        `);
      }

      await pool.query(`
        CREATE UNIQUE INDEX IF NOT EXISTS uq_rag_sharepoint_drive_item
        ON rag_sharepoint_sources(drive_id, item_id)
        WHERE drive_id IS NOT NULL AND item_id IS NOT NULL;
      `);

      // Drop legacy conversation-scoped RAG if still present
      const legacyChunks = await pool.query(`
        SELECT 1 FROM information_schema.columns
        WHERE table_name = 'rag_chunks' AND column_name = 'conversation_id'
        LIMIT 1
      `);
      if (legacyChunks.rows.length > 0) {
        await pool.query(`DROP TABLE IF EXISTS rag_chunks`);
        await pool.query(`DROP TABLE IF EXISTS rag_meta`);
      }

      const metaHasDoc = await pool.query(`
        SELECT 1 FROM information_schema.columns
        WHERE table_name = 'rag_meta' AND column_name = 'document_id'
        LIMIT 1
      `);
      if (metaHasDoc.rows.length === 0) {
        await pool.query(`DROP TABLE IF EXISTS rag_meta`);
      }

      await pool.query(`
        CREATE TABLE IF NOT EXISTS rag_meta (
          document_id TEXT PRIMARY KEY
            REFERENCES rag_documents(id) ON DELETE CASCADE ON UPDATE CASCADE,
          fingerprint TEXT NOT NULL,
          provider TEXT NOT NULL,
          model TEXT NOT NULL,
          dims INT NOT NULL,
          chunk_count INT NOT NULL,
          updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
        );
      `);

      await pool.query(`
        CREATE TABLE IF NOT EXISTS rag_chunks (
          document_id TEXT NOT NULL
            REFERENCES rag_documents(id) ON DELETE CASCADE ON UPDATE CASCADE,
          chunk_id TEXT NOT NULL,
          sheet TEXT NOT NULL,
          text TEXT NOT NULL,
          row_json JSONB NOT NULL,
          embedding vector(384) NOT NULL,
          dims INT NOT NULL,
          PRIMARY KEY (document_id, chunk_id)
        );
      `);

      await pool.query(`
        CREATE INDEX IF NOT EXISTS idx_rag_chunks_sheet
        ON rag_chunks(document_id, sheet);
      `);

      // event_sessions: ensure document_id column exists
      await pool.query(`
        CREATE TABLE IF NOT EXISTS event_sessions (
          conversation_id TEXT PRIMARY KEY,
          document_id TEXT REFERENCES rag_documents(id) ON DELETE SET NULL ON UPDATE CASCADE,
          status TEXT NOT NULL CHECK (status IN ('idle', 'pending', 'active', 'ended')),
          pending_urls JSONB NOT NULL DEFAULT '[]'::jsonb,
          active_source TEXT,
          active_file_name TEXT,
          started_by TEXT,
          started_at TIMESTAMPTZ,
          updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
        );
      `);

      const sessionDocCol = await pool.query(`
        SELECT 1 FROM information_schema.columns
        WHERE table_name = 'event_sessions' AND column_name = 'document_id'
        LIMIT 1
      `);
      if (sessionDocCol.rows.length === 0) {
        await pool.query(`
          ALTER TABLE event_sessions
          ADD COLUMN document_id TEXT
            REFERENCES rag_documents(id) ON DELETE SET NULL ON UPDATE CASCADE
        `);
      }

      await pool.query(`
        CREATE INDEX IF NOT EXISTS idx_event_sessions_document_id
        ON event_sessions(document_id);
      `);
      await pool.query(`
        CREATE INDEX IF NOT EXISTS idx_event_sessions_status
        ON event_sessions(status);
      `);
    })();
  }
  await schemaReady;
}

export function hashContent(bufferOrText: Buffer | string): string {
  return createHash("sha256").update(bufferOrText).digest("hex");
}

function mapDocument(row: {
  id: string;
  source_type: DocumentSourceType;
  source_uri: string;
  file_name: string | null;
  content_hash: string | null;
  status: DocumentStatus;
  created_at: Date | string;
  updated_at: Date | string;
}): RagDocumentRecord {
  return {
    id: row.id,
    sourceType: row.source_type,
    sourceUri: row.source_uri,
    fileName: row.file_name,
    contentHash: row.content_hash,
    status: row.status,
    createdAt: String(row.created_at),
    updatedAt: String(row.updated_at),
  };
}

export async function createRagDocument(input: {
  sourceType: DocumentSourceType;
  sourceUri: string;
  fileName?: string | null;
  contentHash?: string | null;
  status?: DocumentStatus;
  id?: string;
}): Promise<RagDocumentRecord> {
  await ensureDocumentSchema();
  const id = input.id || randomUUID();
  const status = input.status || "pending";
  const result = await getPostgresPool().query(
    `INSERT INTO rag_documents (
       id, source_type, source_uri, file_name, content_hash, status, created_at, updated_at
     ) VALUES ($1, $2, $3, $4, $5, $6, NOW(), NOW())
     RETURNING id, source_type, source_uri, file_name, content_hash, status, created_at, updated_at`,
    [
      id,
      input.sourceType,
      input.sourceUri,
      input.fileName || null,
      input.contentHash || null,
      status,
    ]
  );
  return mapDocument(result.rows[0]);
}

export async function getRagDocument(documentId: string): Promise<RagDocumentRecord | null> {
  await ensureDocumentSchema();
  const result = await getPostgresPool().query(
    `SELECT id, source_type, source_uri, file_name, content_hash, status, created_at, updated_at
     FROM rag_documents WHERE id = $1`,
    [documentId]
  );
  return result.rows[0] ? mapDocument(result.rows[0]) : null;
}

export async function updateRagDocumentStatus(
  documentId: string,
  status: DocumentStatus,
  extras?: { contentHash?: string | null; fileName?: string | null }
): Promise<void> {
  await ensureDocumentSchema();
  await getPostgresPool().query(
    `UPDATE rag_documents SET
       status = $2,
       content_hash = COALESCE($3, content_hash),
       file_name = COALESCE($4, file_name),
       updated_at = NOW()
     WHERE id = $1`,
    [documentId, status, extras?.contentHash ?? null, extras?.fileName ?? null]
  );
}

export async function archiveRagDocument(documentId: string): Promise<void> {
  await updateRagDocumentStatus(documentId, "archived");
}

export async function deleteRagDocument(documentId: string): Promise<void> {
  await ensureDocumentSchema();
  // Cascades to rag_meta, rag_chunks, rag_sharepoint_sources
  await getPostgresPool().query(`DELETE FROM rag_documents WHERE id = $1`, [documentId]);
}

export async function upsertSharepointSource(
  input: UpsertSharepointSourceInput
): Promise<void> {
  await ensureDocumentSchema();
  await getPostgresPool().query(
    `INSERT INTO rag_sharepoint_sources (
       document_id, drive_id, item_id, site_id, web_url,
       subscription_id, subscription_expires_at, delta_link, etag, last_modified_at,
       last_delta_sync_at, updated_at
     ) VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, NOW())
     ON CONFLICT (document_id) DO UPDATE SET
       drive_id = COALESCE(EXCLUDED.drive_id, rag_sharepoint_sources.drive_id),
       item_id = COALESCE(EXCLUDED.item_id, rag_sharepoint_sources.item_id),
       site_id = COALESCE(EXCLUDED.site_id, rag_sharepoint_sources.site_id),
       web_url = COALESCE(EXCLUDED.web_url, rag_sharepoint_sources.web_url),
       subscription_id = COALESCE(EXCLUDED.subscription_id, rag_sharepoint_sources.subscription_id),
       subscription_expires_at = COALESCE(
         EXCLUDED.subscription_expires_at, rag_sharepoint_sources.subscription_expires_at
       ),
       delta_link = COALESCE(EXCLUDED.delta_link, rag_sharepoint_sources.delta_link),
       etag = COALESCE(EXCLUDED.etag, rag_sharepoint_sources.etag),
       last_modified_at = COALESCE(
         EXCLUDED.last_modified_at, rag_sharepoint_sources.last_modified_at
       ),
       last_delta_sync_at = COALESCE(
         EXCLUDED.last_delta_sync_at, rag_sharepoint_sources.last_delta_sync_at
       ),
       updated_at = NOW()`,
    [
      input.documentId,
      input.driveId ?? null,
      input.itemId ?? null,
      input.siteId ?? null,
      input.webUrl ?? null,
      input.subscriptionId ?? null,
      input.subscriptionExpiresAt ?? null,
      input.deltaLink ?? null,
      input.etag ?? null,
      input.lastModifiedAt ?? null,
      input.lastDeltaSyncAt ?? null,
    ]
  );
}

function mapSharepointSource(row: {
  document_id: string;
  drive_id: string | null;
  item_id: string | null;
  site_id: string | null;
  web_url: string | null;
  subscription_id: string | null;
  subscription_expires_at: Date | string | null;
  delta_link: string | null;
  etag: string | null;
  last_modified_at: Date | string | null;
  last_delta_sync_at: Date | string | null;
}): RagSharepointSourceRecord {
  return {
    documentId: row.document_id,
    driveId: row.drive_id,
    itemId: row.item_id,
    siteId: row.site_id,
    webUrl: row.web_url,
    subscriptionId: row.subscription_id,
    subscriptionExpiresAt: row.subscription_expires_at
      ? String(row.subscription_expires_at)
      : null,
    deltaLink: row.delta_link,
    etag: row.etag,
    lastModifiedAt: row.last_modified_at ? String(row.last_modified_at) : null,
    lastDeltaSyncAt: row.last_delta_sync_at ? String(row.last_delta_sync_at) : null,
  };
}

/** Active SharePoint-backed documents linked to an active event session. */
export async function listActiveSharepointSources(): Promise<RagSharepointSourceRecord[]> {
  await ensureDocumentSchema();
  const result = await getPostgresPool().query(
    `SELECT s.document_id, s.drive_id, s.item_id, s.site_id, s.web_url,
            s.subscription_id, s.subscription_expires_at, s.delta_link, s.etag,
            s.last_modified_at, s.last_delta_sync_at
     FROM rag_sharepoint_sources s
     INNER JOIN rag_documents d ON d.id = s.document_id
     INNER JOIN event_sessions e ON e.document_id = s.document_id
     WHERE e.status = 'active'
       AND d.status IN ('ready', 'syncing', 'error')
       AND s.drive_id IS NOT NULL
       AND s.item_id IS NOT NULL`
  );
  return result.rows.map(mapSharepointSource);
}

export async function getSharepointSource(
  documentId: string
): Promise<RagSharepointSourceRecord | null> {
  await ensureDocumentSchema();
  const result = await getPostgresPool().query(
    `SELECT document_id, drive_id, item_id, site_id, web_url,
            subscription_id, subscription_expires_at, delta_link, etag,
            last_modified_at, last_delta_sync_at
     FROM rag_sharepoint_sources WHERE document_id = $1`,
    [documentId]
  );
  return result.rows[0] ? mapSharepointSource(result.rows[0]) : null;
}
