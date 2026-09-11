import { ILogger } from "@microsoft/teams.common";
import type { Pool } from "pg";
import { IDatabase } from "./database";
import { getPostgresPool } from "./postgres";
import { MessageRecord } from "./types";

/**
 * Postgres-backed conversation / feedback store.
 */
export class PostgresKVStore implements IDatabase {
  private pool: Pool;

  constructor(private logger: ILogger) {
    this.pool = getPostgresPool();
  }

  async initialize(): Promise<void> {
    await this.pool.query(`
      CREATE TABLE IF NOT EXISTS conversations (
        id BIGSERIAL PRIMARY KEY,
        conversation_id TEXT NOT NULL,
        role TEXT NOT NULL,
        name TEXT NOT NULL,
        content TEXT NOT NULL,
        activity_id TEXT NOT NULL,
        timestamp TEXT NOT NULL,
        blob JSONB NOT NULL
      );

      CREATE INDEX IF NOT EXISTS idx_conversations_conversation_id
        ON conversations(conversation_id);
      CREATE INDEX IF NOT EXISTS idx_conversations_timestamp
        ON conversations(conversation_id, timestamp);

      CREATE TABLE IF NOT EXISTS feedback (
        id BIGSERIAL PRIMARY KEY,
        reply_to_id TEXT NOT NULL,
        reaction TEXT NOT NULL,
        feedback TEXT,
        created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
      );

      CREATE INDEX IF NOT EXISTS idx_feedback_reply_to_id ON feedback(reply_to_id);
    `);
    this.logger.debug("✅ Postgres conversation schema ready");
  }

  async clearAll(): Promise<void> {
    await this.pool.query("DELETE FROM conversations");
    this.logger.debug("🧹 Cleared all conversations from Postgres store.");
  }

  async get(conversationId: string): Promise<MessageRecord[]> {
    const result = await this.pool.query<{ blob: MessageRecord }>(
      `SELECT blob FROM conversations
       WHERE conversation_id = $1
       ORDER BY timestamp ASC`,
      [conversationId]
    );
    return result.rows.map((row) => row.blob);
  }

  async getMessagesByTimeRange(
    conversationId: string,
    startTime: string,
    endTime: string
  ): Promise<MessageRecord[]> {
    const result = await this.pool.query<{ blob: MessageRecord }>(
      `SELECT blob FROM conversations
       WHERE conversation_id = $1
         AND timestamp >= $2
         AND timestamp <= $3
       ORDER BY timestamp ASC`,
      [conversationId, startTime, endTime]
    );
    return result.rows.map((row) => row.blob);
  }

  async getRecentMessages(conversationId: string, limit = 10): Promise<MessageRecord[]> {
    const messages = await this.get(conversationId);
    return messages.slice(-limit);
  }

  async clearConversation(conversationId: string): Promise<void> {
    await this.pool.query("DELETE FROM conversations WHERE conversation_id = $1", [
      conversationId,
    ]);
  }

  async addMessages(messages: MessageRecord[]): Promise<void> {
    if (!messages.length) {
      return;
    }

    const client = await this.pool.connect();
    try {
      await client.query("BEGIN");
      for (const message of messages) {
        await client.query(
          `INSERT INTO conversations
            (conversation_id, role, name, content, activity_id, timestamp, blob)
           VALUES ($1, $2, $3, $4, $5, $6, $7::jsonb)`,
          [
            message.conversation_id,
            message.role,
            message.name,
            message.content,
            message.activity_id,
            message.timestamp,
            JSON.stringify(message),
          ]
        );
      }
      await client.query("COMMIT");
    } catch (error) {
      await client.query("ROLLBACK");
      throw error;
    } finally {
      client.release();
    }
  }

  async countMessages(conversationId: string): Promise<number> {
    const result = await this.pool.query<{ count: string }>(
      `SELECT COUNT(*)::text AS count FROM conversations WHERE conversation_id = $1`,
      [conversationId]
    );
    return Number(result.rows[0]?.count || 0);
  }

  async clearAllMessages(): Promise<void> {
    const result = await this.pool.query("DELETE FROM conversations");
    this.logger.debug(
      `🧹 Cleared all conversations from Postgres. Deleted ${result.rowCount ?? 0} records.`
    );
  }

  async getFilteredMessages(
    conversationId: string,
    keywords: string[],
    startTime: string,
    endTime: string,
    participants?: string[],
    maxResults?: number
  ): Promise<MessageRecord[]> {
    const values: unknown[] = [conversationId, startTime, endTime];
    const where: string[] = [
      "conversation_id = $1",
      "timestamp >= $2",
      "timestamp <= $3",
    ];

    if (keywords.length) {
      const keywordClauses = keywords.map((_, index) => {
        values.push(`%${keywords[index].toLowerCase()}%`);
        return `LOWER(content) LIKE $${values.length}`;
      });
      where.push(`(${keywordClauses.join(" OR ")})`);
    }

    if (participants?.length) {
      const participantClauses = participants.map((_, index) => {
        values.push(`%${participants[index].toLowerCase()}%`);
        return `LOWER(name) LIKE $${values.length}`;
      });
      where.push(`(${participantClauses.join(" OR ")})`);
    }

    const limit = maxResults && typeof maxResults === "number" ? maxResults : 5;
    values.push(limit);

    const result = await this.pool.query<{ blob: MessageRecord }>(
      `SELECT blob FROM conversations
       WHERE ${where.join(" AND ")}
       ORDER BY timestamp DESC
       LIMIT $${values.length}`,
      values
    );
    return result.rows.map((row) => row.blob);
  }

  async recordFeedback(
    replyToId: string,
    reaction: "like" | "dislike" | string,
    feedbackJson?: unknown
  ): Promise<boolean> {
    try {
      const result = await this.pool.query(
        `INSERT INTO feedback (reply_to_id, reaction, feedback)
         VALUES ($1, $2, $3)`,
        [replyToId, reaction, feedbackJson ? JSON.stringify(feedbackJson) : null]
      );
      return (result.rowCount ?? 0) > 0;
    } catch (error) {
      this.logger.error(`❌ recordFeedback error:`, error);
      return false;
    }
  }

  async close(): Promise<void> {
    // Shared pool is closed separately via closePostgresPool()
    this.logger.debug("🔌 Postgres conversation store closed (shared pool retained)");
  }
}
