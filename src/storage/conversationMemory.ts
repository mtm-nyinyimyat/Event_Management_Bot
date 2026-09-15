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
}
