import { IMessageActivity, MessageActivity } from "@microsoft/teams.api";
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

export function finalizePromptResponse(text: string): MessageActivity {
  return new MessageActivity(text).addAiGenerated().addFeedback();
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
