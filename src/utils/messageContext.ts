import { Client, IMessageActivity } from "@microsoft/teams.api";
import { ConversationMemory } from "../storage/conversationMemory";
import { IDatabase } from "../storage/database";

/**
 * Context object that stores all important information for processing a message
 */
export interface MessageContext {
  text: string;
  conversationId: string;
  /** When set, workbook/RAG lookups use this conversation's active event (shared across DM/group). */
  activeEventConversationId?: string;
  /** Display name of the active event Excel file (for user-facing refusals). */
  activeEventFileName?: string;
  userId?: string;
  userName: string;
  timestamp: string;
  isPersonalChat: boolean;
  activityId: string;
  members: Array<{ name: string; id: string }>;
  memory: ConversationMemory;
}

async function getConversationParticipantsFromAPI(
  api: Client,
  conversationId: string
): Promise<Array<{ name: string; id: string }>> {
  try {
    const members = await api.conversations.members(conversationId).get();
    if (!Array.isArray(members)) {
      return [];
    }
    return members.map((member) => ({
      name: member.name || "Unknown",
      id: member.aadObjectId || member.id,
    }));
  } catch {
    return [];
  }
}

/**
 * Factory function to create a MessageContext from a Teams activity
 */
export async function createMessageContext(
  storage: IDatabase,
  activity: IMessageActivity,
  api?: Client
): Promise<MessageContext> {
  const conversationId = `${activity.conversation.id}`;
  const userId = activity.from.aadObjectId || activity.from.id;
  let members: Array<{ name: string; id: string }> = [];
  if (api) {
    members = await getConversationParticipantsFromAPI(api, conversationId);
  }

  return {
    text: activity.text || "",
    conversationId,
    userId,
    userName: activity.from.name || "User",
    timestamp: activity.timestamp?.toString() || "Unknown",
    isPersonalChat: activity.conversation.conversationType === "personal",
    activityId: activity.id,
    members,
    memory: new ConversationMemory(storage, conversationId),
  };
}
