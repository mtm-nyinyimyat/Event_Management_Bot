import { IDatabase } from "./database";
import { MessageRecord } from "./types";

export class ConversationMemory {
  constructor(
    private store: IDatabase,
    private conversationId: string
  ) {}

  async addMessages(messages: MessageRecord[]): Promise<void> {
    await this.store.addMessages(messages);
  }

  /** Recent turns for this conversation (oldest → newest), for follow-up context. */
  async getRecentMessages(limit = 8): Promise<MessageRecord[]> {
    const capped = Math.max(1, Math.min(limit, 20));
    const store = this.store as IDatabase & {
      getRecentMessages?: (conversationId: string, limit?: number) => MessageRecord[] | Promise<MessageRecord[]>;
    };
    if (typeof store.getRecentMessages === "function") {
      return store.getRecentMessages(this.conversationId, capped);
    }
    const all = await store.get(this.conversationId);
    return all.slice(-capped);
  }
}
