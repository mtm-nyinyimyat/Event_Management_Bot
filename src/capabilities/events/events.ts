import { ChatPrompt } from "@microsoft/teams.ai";
import { ILogger } from "@microsoft/teams.common";
import { ServiceUnavailableError } from "../../events/activeEventFile";
import { searchWorkbook } from "../../events/excelStore";
import { createChatModel } from "../../utils/config";
import { MessageContext } from "../../utils/messageContext";
import { BaseCapability, CapabilityDefinition, CapabilityResult } from "../capability";
import { buildEventsPrompt } from "./prompt";
import { LOOKUP_EVENTS_SCHEMA, LookupEventsArgs } from "./schema";

export class EventsCapability extends BaseCapability {
  readonly name = "events";

  createPrompt(context: MessageContext): ChatPrompt {
    const modelConfig = this.getModelConfig("events");
    const senderName = context.userName || "Unknown";

    const prompt = new ChatPrompt({
      instructions: buildEventsPrompt(senderName, context.activeEventFileName),
      model: createChatModel(modelConfig),
    }).function(
      "lookup_events",
      "Search the event Excel workbook with hybrid RAG (semantic + keyword) and return summaries plus matching rows",
      LOOKUP_EVENTS_SCHEMA,
      async ({ query, max_results }: LookupEventsArgs) => {
        this.logger.debug(
          `📊 lookup_events query="${query || ""}" sender="${senderName}"`
        );
        try {
          const result = await searchWorkbook(query || "", max_results ?? 20, {
            conversationId: context.activeEventConversationId || context.conversationId,
            userId: context.userId,
            requesterName: senderName,
          });
          const fromCache = /\(cached\)|Cached workbook/i.test(result.answer_hint || "");
          if (fromCache) {
            this.logger.debug(`🗃️ lookup_events cache hit query="${query || ""}"`);
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
      const message = [
        `Teams sender: ${senderName}`,
        `User question: ${context.text}`,
        `If this question refers to the speaker (I/me/my), look up "${senderName}" in the workbook.`,
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
  manager_desc: `**events**: ONLY capability for event Excel Q&A. Use for questions about the active event file (English or Burmese): participants, agenda, volunteers, ferry, seating, menu, beverages, karaoke. Do NOT use for general knowledge. If the user question is unrelated to that file, do not delegate — refuse with the event-file-only message.`,
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
