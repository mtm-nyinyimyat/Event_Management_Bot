import { ChatPrompt } from "@microsoft/teams.ai";
import { ILogger } from "@microsoft/teams.common";
import { ServiceUnavailableError } from "../../events/activeEventFile";
import {
  getLastEventLookup,
  isEventFollowUpAsk,
  rememberEventLookup,
  resolveFollowUpLookupQuery,
} from "../../events/excel/eventLookupMemory";
import { searchWorkbook } from "../../events/excelStore";
import { createChatModel } from "../../utils/config";
import { MessageContext } from "../../utils/messageContext";
import { BaseCapability, CapabilityDefinition, CapabilityResult } from "../capability";
import { buildEventsPrompt } from "./prompt";
import { LOOKUP_EVENTS_SCHEMA, LookupEventsArgs } from "./schema";

function formatRecentTurns(
  turns: Array<{ role?: string; content?: string; name?: string }>
): string {
  if (!turns.length) {
    return "(none)";
  }
  return turns
    .map((turn) => {
      const role = turn.role === "model" ? "Bot" : "User";
      const who = turn.name ? ` (${turn.name})` : "";
      const body = String(turn.content || "")
        .replace(/\s+/g, " ")
        .trim()
        .slice(0, 400);
      return `${role}${who}: ${body}`;
    })
    .join("\n");
}

export class EventsCapability extends BaseCapability {
  readonly name = "events";

  createPrompt(context: MessageContext): ChatPrompt {
    const modelConfig = this.getModelConfig("events");
    const senderName = context.userName || "Unknown";
    const conversationKey =
      context.activeEventConversationId || context.conversationId;

    const prompt = new ChatPrompt({
      instructions: buildEventsPrompt(senderName, context.activeEventFileName),
      model: createChatModel(modelConfig),
    }).function(
      "lookup_events",
      "Search the event Excel workbook with hybrid RAG (semantic + keyword) and return summaries plus matching rows. For short follow-ups (list/send/yes), pass the prior topic keywords.",
      LOOKUP_EVENTS_SCHEMA,
      async ({ query, max_results }: LookupEventsArgs) => {
        const lookupQuery = resolveFollowUpLookupQuery(
          conversationKey,
          query || "",
          context.text || ""
        );
        const followUp = isEventFollowUpAsk(context.text || "");
        const limit = Math.max(max_results ?? 20, followUp ? 80 : 20);
        this.logger.debug(
          `📊 lookup_events query="${lookupQuery}" sender="${senderName}" followUp=${followUp}`
        );
        try {
          const result = await searchWorkbook(lookupQuery, limit, {
            conversationId: conversationKey,
            documentId: context.activeEventDocumentId,
            userId: context.userId,
            requesterName: senderName,
          });
          if ((result.match_count || 0) > 0) {
            rememberEventLookup(conversationKey, {
              query: lookupQuery,
              retrieval: result.retrieval,
              matchCount: result.match_count,
            });
          }
          const fromCache = /\(cached\)|Cached workbook/i.test(result.answer_hint || "");
          if (fromCache) {
            this.logger.debug(`🗃️ lookup_events cache hit query="${lookupQuery}"`);
          }
          return JSON.stringify(result);
        } catch (error) {
          const message = error instanceof Error ? error.message : "Unknown Excel error";
          this.logger.error(`❌ Excel lookup failed: ${message}`);
          return JSON.stringify({ error: message, sheets: [], match_count: 0 });
        }
      }
    );

    this.logger.debug(`Initialized Events Capability for sender="${senderName}"`);
    return prompt;
  }

  async processRequest(context: MessageContext): Promise<CapabilityResult> {
    try {
      const prompt = this.createPrompt(context);
      const senderName = context.userName || "Unknown";
      const conversationKey =
        context.activeEventConversationId || context.conversationId;
      const recent = await context.memory.getRecentMessages(8);
      const lastLookup = getLastEventLookup(conversationKey);
      const followUp = isEventFollowUpAsk(context.text || "");

      const message = [
        `Teams sender: ${senderName}`,
        `User question: ${context.text}`,
        `Recent conversation turns:`,
        formatRecentTurns(recent),
        lastLookup
          ? `Last workbook lookup: query="${lastLookup.query}" retrieval=${lastLookup.retrieval || "?"} matches=${lastLookup.matchCount ?? "?"}`
          : `Last workbook lookup: (none)`,
        followUp
          ? `This looks like a follow-up (list/send/yes). Continue the last workbook topic — call lookup_events with those prior keywords. Do NOT look up the sender's personal profile unless they asked about themselves.`
          : `If this question refers to the speaker (I/me/my), look up "${senderName}" in the workbook.`,
      ].join("\n");

      const response = await prompt.send(message);
      return {
        response:
          response.content ||
          "Couldn't find that one — try asking another way?",
      };
    } catch (error) {
      throw new ServiceUnavailableError(error);
    }
  }
}

export const EVENTS_CAPABILITY_DEFINITION: CapabilityDefinition = {
  name: "events",
  manager_desc: `**events**: ONLY capability for the active event Excel Q&A (any sheet/column layout). Use for questions about that file in English or Burmese. Do NOT use for general knowledge. If unrelated to the active event file, do not delegate — refuse with the event-file-only message.`,
  handler: async (context: MessageContext, logger: ILogger) => {
    const capability = new EventsCapability(logger);
    const result = await capability.processRequest(context);
    if (result.error) {
      logger.error(`❌ Error in Events Capability: ${result.error}`);
      throw new ServiceUnavailableError(result.error);
    }
    return result.response || "Couldn't find that one — try asking another way?";
  },
};
