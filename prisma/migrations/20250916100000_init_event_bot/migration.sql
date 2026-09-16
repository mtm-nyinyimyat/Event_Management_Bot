-- Event Management bot — document-centric schema (Prisma Postgres + pgvector)

CREATE EXTENSION IF NOT EXISTS vector;

CREATE TABLE "conversations" (
    "id" BIGSERIAL NOT NULL,
    "conversation_id" TEXT NOT NULL,
    "role" TEXT NOT NULL,
    "name" TEXT NOT NULL,
    "content" TEXT NOT NULL,
    "activity_id" TEXT NOT NULL,
    "timestamp" TEXT NOT NULL,
    "blob" JSONB NOT NULL,

    CONSTRAINT "conversations_pkey" PRIMARY KEY ("id")
);

CREATE INDEX "idx_conversations_conversation_id" ON "conversations"("conversation_id");
CREATE INDEX "idx_conversations_timestamp" ON "conversations"("conversation_id", "timestamp");

CREATE TABLE "feedback" (
    "id" BIGSERIAL NOT NULL,
    "reply_to_id" TEXT NOT NULL,
    "reaction" TEXT NOT NULL,
    "feedback" TEXT,
    "created_at" TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "feedback_pkey" PRIMARY KEY ("id")
);

CREATE INDEX "idx_feedback_reply_to_id" ON "feedback"("reply_to_id");

CREATE TABLE "rag_documents" (
    "id" TEXT NOT NULL,
    "source_type" TEXT NOT NULL,
    "source_uri" TEXT NOT NULL,
    "file_name" TEXT,
    "content_hash" TEXT,
    "status" TEXT NOT NULL DEFAULT 'pending',
    "created_at" TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updated_at" TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "rag_documents_pkey" PRIMARY KEY ("id"),
    CONSTRAINT "rag_documents_status_check" CHECK (
      "status" IN ('pending', 'ready', 'syncing', 'error', 'archived')
    ),
    CONSTRAINT "rag_documents_source_type_check" CHECK (
      "source_type" IN ('upload', 'sharepoint', 'graph', 'local')
    )
);

CREATE INDEX "idx_rag_documents_status" ON "rag_documents"("status");
CREATE INDEX "idx_rag_documents_content_hash" ON "rag_documents"("content_hash");

CREATE TABLE "rag_sharepoint_sources" (
    "document_id" TEXT NOT NULL,
    "drive_id" TEXT,
    "item_id" TEXT,
    "site_id" TEXT,
    "web_url" TEXT,
    "subscription_id" TEXT,
    "subscription_expires_at" TIMESTAMPTZ,
    "delta_link" TEXT,
    "etag" TEXT,
    "last_modified_at" TIMESTAMPTZ,
    "updated_at" TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "rag_sharepoint_sources_pkey" PRIMARY KEY ("document_id"),
    CONSTRAINT "rag_sharepoint_sources_document_id_fkey"
      FOREIGN KEY ("document_id") REFERENCES "rag_documents"("id") ON DELETE CASCADE ON UPDATE CASCADE
);

CREATE UNIQUE INDEX "uq_rag_sharepoint_drive_item"
  ON "rag_sharepoint_sources"("drive_id", "item_id")
  WHERE "drive_id" IS NOT NULL AND "item_id" IS NOT NULL;

CREATE TABLE "event_sessions" (
    "conversation_id" TEXT NOT NULL,
    "document_id" TEXT,
    "status" TEXT NOT NULL,
    "pending_urls" JSONB NOT NULL DEFAULT '[]'::jsonb,
    "active_source" TEXT,
    "active_file_name" TEXT,
    "started_by" TEXT,
    "started_at" TIMESTAMPTZ,
    "updated_at" TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "event_sessions_pkey" PRIMARY KEY ("conversation_id"),
    CONSTRAINT "event_sessions_status_check" CHECK (
      "status" IN ('idle', 'pending', 'active', 'ended')
    ),
    CONSTRAINT "event_sessions_document_id_fkey"
      FOREIGN KEY ("document_id") REFERENCES "rag_documents"("id") ON DELETE SET NULL ON UPDATE CASCADE
);

CREATE INDEX "idx_event_sessions_document_id" ON "event_sessions"("document_id");
CREATE INDEX "idx_event_sessions_status" ON "event_sessions"("status");

CREATE TABLE "rag_meta" (
    "document_id" TEXT NOT NULL,
    "fingerprint" TEXT NOT NULL,
    "provider" TEXT NOT NULL,
    "model" TEXT NOT NULL,
    "dims" INTEGER NOT NULL,
    "chunk_count" INTEGER NOT NULL,
    "updated_at" TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "rag_meta_pkey" PRIMARY KEY ("document_id"),
    CONSTRAINT "rag_meta_document_id_fkey"
      FOREIGN KEY ("document_id") REFERENCES "rag_documents"("id") ON DELETE CASCADE ON UPDATE CASCADE
);

CREATE TABLE "rag_chunks" (
    "document_id" TEXT NOT NULL,
    "chunk_id" TEXT NOT NULL,
    "sheet" TEXT NOT NULL,
    "text" TEXT NOT NULL,
    "row_json" JSONB NOT NULL,
    "embedding" vector(384) NOT NULL,
    "dims" INTEGER NOT NULL,

    CONSTRAINT "rag_chunks_pkey" PRIMARY KEY ("document_id","chunk_id"),
    CONSTRAINT "rag_chunks_document_id_fkey"
      FOREIGN KEY ("document_id") REFERENCES "rag_documents"("id") ON DELETE CASCADE ON UPDATE CASCADE
);

CREATE INDEX "idx_rag_chunks_sheet" ON "rag_chunks"("document_id", "sheet");
