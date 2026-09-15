import { MessageRecord } from "./types";

/**
 * Abstract database interface for conversation history (Postgres).
 */
export interface IDatabase {
  initialize(): Promise<void>;
  get(conversationId: string): MessageRecord[] | Promise<MessageRecord[]>;
  clearConversation(conversationId: string): void | Promise<void>;
  addMessages(messages: MessageRecord[]): void | Promise<void>;
  countMessages(conversationId: string): number | Promise<number>;
  recordFeedback(
    replyToId: string,
    reaction: "like" | "dislike" | string,
    feedbackJson?: unknown
  ): boolean | Promise<boolean>;
  close(): void | Promise<void>;
}
