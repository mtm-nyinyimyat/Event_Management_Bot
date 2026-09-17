import { Account, IMessageActivity, MessageActivity } from "@microsoft/teams.api";
import { MessageRecord } from "../storage/types";

/**
 * User-facing event title from an Excel file name.
 * "2025 MTM Annual Staff Party 3.xlsx" → "2025 MTM Annual Staff Party"
 */
export function formatEventDisplayName(
  fileName?: string | null,
  fallback = "the active event"
): string {
  let name = String(fileName || "").trim();
  if (!name) {
    return fallback;
  }

  name = name.replace(/\.(xlsx|xls)$/i, "").trim();
  name = name
    .replace(/\s*[\(\[]?\s*v?\d+\s*[\)\]]?\s*$/i, "")
    .replace(/\s*[-_]\s*(copy|final|new)\s*$/i, "")
    .trim();

  return name || fallback;
}

export interface BotReplyOptions {
  /** @mention this Teams account in the reply. */
  mentionAccount?: Account;
  /** Reply in the thread under this activity id (channel/group). */
  replyToId?: string;
  /** Attach AI-generated + feedback entities (default true). */
  withFeedback?: boolean;
}

/**
 * Build an outgoing message with optional @mention and thread replyToId.
 */
export function finalizePromptResponse(
  text: string,
  options?: BotReplyOptions
): MessageActivity {
  const body = String(text || "").trim();
  const activity = new MessageActivity(body);

  if (options?.withFeedback !== false) {
    activity.addAiGenerated().addFeedback();
  }

  if (options?.mentionAccount?.id && options.mentionAccount.name) {
    activity.addMention(options.mentionAccount, { addText: false });
    activity.text = `<at>${options.mentionAccount.name}</at> ${body}`;
  }

  if (options?.replyToId) {
    activity.replyToId = options.replyToId;
  }

  return activity;
}

export function createMessageRecords(activities: IMessageActivity[]): MessageRecord[] {
  const conversation_id = activities[0].conversation.id;
  return activities.map((activity) => ({
    conversation_id,
    role: activity.entities?.some((e: { additionalType?: string[] }) =>
      e.additionalType?.includes("AIGeneratedContent")
    )
      ? "model"
      : "user",
    content: activity.text?.replace(/<\/?at>/g, "") || "",
    timestamp: activity.timestamp?.toString() || new Date().toISOString(),
    activity_id: activity.id,
    name: activity.from?.name || "Event Bot",
  }));
}
