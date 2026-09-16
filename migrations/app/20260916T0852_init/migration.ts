#!/usr/bin/env -S node
import type { Contract as End } from '../../snapshots/45171294cfb86c7e16d6205724b8f2cec224bc5cdceb3feefad1eb4178cb9ba7/contract';
import endContract from '../../snapshots/45171294cfb86c7e16d6205724b8f2cec224bc5cdceb3feefad1eb4178cb9ba7/contract.json' with { type: 'json' };
import { Migration, MigrationCLI, col, fn, lit, primaryKey } from '@prisma/orm-postgres/migration';

export default class M extends Migration<never, End> {
  override readonly endContractJson = endContract;

  override get operations() {
    return [
      this.createSchema({ schema: 'public' }),
      this.createTable({
        schema: 'public',
        table: 'conversations',
        columns: [
          col('activity_id', 'text', { notNull: true, codecRef: { codecId: 'pg/text@1' } }),
          col('blob', 'jsonb', { notNull: true, codecRef: { codecId: 'pg/jsonb@1' } }),
          col('content', 'text', { notNull: true, codecRef: { codecId: 'pg/text@1' } }),
          col('conversation_id', 'text', { notNull: true, codecRef: { codecId: 'pg/text@1' } }),
          col('id', 'BIGSERIAL', { notNull: true, codecRef: { codecId: 'pg/int8@1' } }),
          col('name', 'text', { notNull: true, codecRef: { codecId: 'pg/text@1' } }),
          col('role', 'text', { notNull: true, codecRef: { codecId: 'pg/text@1' } }),
          col('timestamp', 'text', { notNull: true, codecRef: { codecId: 'pg/text@1' } }),
        ],
        constraints: [primaryKey(['id'])],
      }),
      this.createTable({
        schema: 'public',
        table: 'event_sessions',
        columns: [
          col('active_file_name', 'text', { codecRef: { codecId: 'pg/text@1' } }),
          col('active_source', 'text', { codecRef: { codecId: 'pg/text@1' } }),
          col('conversation_id', 'text', { notNull: true, codecRef: { codecId: 'pg/text@1' } }),
          col('document_id', 'text', { codecRef: { codecId: 'pg/text@1' } }),
          col('pending_urls', 'jsonb', {
            notNull: true,
            default: lit('[]'),
            codecRef: { codecId: 'pg/jsonb@1' },
          }),
          col('started_at', 'timestamptz', { codecRef: { codecId: 'pg/timestamptz-string@1' } }),
          col('started_by', 'text', { codecRef: { codecId: 'pg/text@1' } }),
          col('status', 'text', { notNull: true, codecRef: { codecId: 'pg/text@1' } }),
          col('updated_at', 'timestamptz', {
            notNull: true,
            default: fn('now()'),
            codecRef: { codecId: 'pg/timestamptz-string@1' },
          }),
        ],
        constraints: [primaryKey(['conversation_id'])],
      }),
      this.createTable({
        schema: 'public',
        table: 'feedback',
        columns: [
          col('created_at', 'timestamptz', {
            notNull: true,
            default: fn('now()'),
            codecRef: { codecId: 'pg/timestamptz-string@1' },
          }),
          col('feedback', 'text', { codecRef: { codecId: 'pg/text@1' } }),
          col('id', 'BIGSERIAL', { notNull: true, codecRef: { codecId: 'pg/int8@1' } }),
          col('reaction', 'text', { notNull: true, codecRef: { codecId: 'pg/text@1' } }),
          col('reply_to_id', 'text', { notNull: true, codecRef: { codecId: 'pg/text@1' } }),
        ],
        constraints: [primaryKey(['id'])],
      }),
      this.createTable({
        schema: 'public',
        table: 'rag_chunks',
        columns: [
          col('chunk_id', 'text', { notNull: true, codecRef: { codecId: 'pg/text@1' } }),
          col('dims', 'int4', { notNull: true, codecRef: { codecId: 'pg/int4@1' } }),
          col('document_id', 'text', { notNull: true, codecRef: { codecId: 'pg/text@1' } }),
          col('embedding', 'vector(384)', {
            notNull: true,
            codecRef: { codecId: 'pg/vector@1', typeParams: { length: 384 } },
          }),
          col('row_json', 'jsonb', { notNull: true, codecRef: { codecId: 'pg/jsonb@1' } }),
          col('sheet', 'text', { notNull: true, codecRef: { codecId: 'pg/text@1' } }),
          col('text', 'text', { notNull: true, codecRef: { codecId: 'pg/text@1' } }),
        ],
        constraints: [primaryKey(['document_id', 'chunk_id'])],
      }),
      this.createTable({
        schema: 'public',
        table: 'rag_documents',
        columns: [
          col('content_hash', 'text', { codecRef: { codecId: 'pg/text@1' } }),
          col('created_at', 'timestamptz', {
            notNull: true,
            default: fn('now()'),
            codecRef: { codecId: 'pg/timestamptz-string@1' },
          }),
          col('file_name', 'text', { codecRef: { codecId: 'pg/text@1' } }),
          col('id', 'text', { notNull: true, codecRef: { codecId: 'pg/text@1' } }),
          col('source_type', 'text', { notNull: true, codecRef: { codecId: 'pg/text@1' } }),
          col('source_uri', 'text', { notNull: true, codecRef: { codecId: 'pg/text@1' } }),
          col('status', 'text', {
            notNull: true,
            default: lit('pending'),
            codecRef: { codecId: 'pg/text@1' },
          }),
          col('updated_at', 'timestamptz', {
            notNull: true,
            default: fn('now()'),
            codecRef: { codecId: 'pg/timestamptz-string@1' },
          }),
        ],
        constraints: [primaryKey(['id'])],
      }),
      this.createTable({
        schema: 'public',
        table: 'rag_meta',
        columns: [
          col('chunk_count', 'int4', { notNull: true, codecRef: { codecId: 'pg/int4@1' } }),
          col('dims', 'int4', { notNull: true, codecRef: { codecId: 'pg/int4@1' } }),
          col('document_id', 'text', { notNull: true, codecRef: { codecId: 'pg/text@1' } }),
          col('fingerprint', 'text', { notNull: true, codecRef: { codecId: 'pg/text@1' } }),
          col('model', 'text', { notNull: true, codecRef: { codecId: 'pg/text@1' } }),
          col('provider', 'text', { notNull: true, codecRef: { codecId: 'pg/text@1' } }),
          col('updated_at', 'timestamptz', {
            notNull: true,
            default: fn('now()'),
            codecRef: { codecId: 'pg/timestamptz-string@1' },
          }),
        ],
        constraints: [primaryKey(['document_id'])],
      }),
      this.createTable({
        schema: 'public',
        table: 'rag_sharepoint_sources',
        columns: [
          col('delta_link', 'text', { codecRef: { codecId: 'pg/text@1' } }),
          col('document_id', 'text', { notNull: true, codecRef: { codecId: 'pg/text@1' } }),
          col('drive_id', 'text', { codecRef: { codecId: 'pg/text@1' } }),
          col('etag', 'text', { codecRef: { codecId: 'pg/text@1' } }),
          col('item_id', 'text', { codecRef: { codecId: 'pg/text@1' } }),
          col('last_modified_at', 'timestamptz', {
            codecRef: { codecId: 'pg/timestamptz-string@1' },
          }),
          col('site_id', 'text', { codecRef: { codecId: 'pg/text@1' } }),
          col('subscription_expires_at', 'timestamptz', {
            codecRef: { codecId: 'pg/timestamptz-string@1' },
          }),
          col('subscription_id', 'text', { codecRef: { codecId: 'pg/text@1' } }),
          col('updated_at', 'timestamptz', {
            notNull: true,
            default: fn('now()'),
            codecRef: { codecId: 'pg/timestamptz-string@1' },
          }),
          col('web_url', 'text', { codecRef: { codecId: 'pg/text@1' } }),
        ],
        constraints: [primaryKey(['document_id'])],
      }),
      this.addUnique({
        schema: 'public',
        table: 'rag_sharepoint_sources',
        constraint: 'uq_rag_sharepoint_drive_item',
        columns: ['drive_id', 'item_id'],
      }),
      this.createIndex({
        schema: 'public',
        table: 'conversations',
        index: 'idx_conversations_conversation_id',
        columns: ['conversation_id'],
      }),
      this.createIndex({
        schema: 'public',
        table: 'conversations',
        index: 'idx_conversations_timestamp',
        columns: ['conversation_id', 'timestamp'],
      }),
      this.createIndex({
        schema: 'public',
        table: 'event_sessions',
        index: 'idx_event_sessions_document_id',
        columns: ['document_id'],
      }),
      this.createIndex({
        schema: 'public',
        table: 'event_sessions',
        index: 'idx_event_sessions_status',
        columns: ['status'],
      }),
      this.createIndex({
        schema: 'public',
        table: 'feedback',
        index: 'idx_feedback_reply_to_id',
        columns: ['reply_to_id'],
      }),
      this.createIndex({
        schema: 'public',
        table: 'rag_chunks',
        index: 'idx_rag_chunks_sheet',
        columns: ['document_id', 'sheet'],
      }),
      this.createIndex({
        schema: 'public',
        table: 'rag_chunks',
        index: 'rag_chunks_document_id_idx_d3d0944e',
        columns: ['document_id'],
      }),
      this.createIndex({
        schema: 'public',
        table: 'rag_documents',
        index: 'idx_rag_documents_content_hash',
        columns: ['content_hash'],
      }),
      this.createIndex({
        schema: 'public',
        table: 'rag_documents',
        index: 'idx_rag_documents_status',
        columns: ['status'],
      }),
      this.addForeignKey({
        schema: 'public',
        table: 'event_sessions',
        foreignKey: {
          name: 'event_sessions_document_id_fkey',
          columns: ['document_id'],
          references: { schema: 'public', table: 'rag_documents', columns: ['id'] },
          onDelete: 'setNull',
        },
      }),
      this.addForeignKey({
        schema: 'public',
        table: 'rag_chunks',
        foreignKey: {
          name: 'rag_chunks_document_id_fkey',
          columns: ['document_id'],
          references: { schema: 'public', table: 'rag_documents', columns: ['id'] },
          onDelete: 'cascade',
        },
      }),
      this.addForeignKey({
        schema: 'public',
        table: 'rag_meta',
        foreignKey: {
          name: 'rag_meta_document_id_fkey',
          columns: ['document_id'],
          references: { schema: 'public', table: 'rag_documents', columns: ['id'] },
          onDelete: 'cascade',
        },
      }),
      this.addForeignKey({
        schema: 'public',
        table: 'rag_sharepoint_sources',
        foreignKey: {
          name: 'rag_sharepoint_sources_document_id_fkey',
          columns: ['document_id'],
          references: { schema: 'public', table: 'rag_documents', columns: ['id'] },
          onDelete: 'cascade',
        },
      }),
    ];
  }
}

MigrationCLI.run(import.meta.url, M);
