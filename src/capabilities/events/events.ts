import { ChatPrompt } from "@microsoft/teams.ai";
import { ILogger } from "@microsoft/teams.common";
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
      instructions: buildEventsPrompt(senderName),
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
            conversationId: context.conversationId,
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
        response: response.content || "No response generated",
      };
    } catch (error) {
      return {
        response: "",
        error: error instanceof Error ? error.message : "Unknown error",
      };
    }
  }
}

export const EVENTS_CAPABILITY_DEFINITION: CapabilityDefinition = {
  name: "events",
  manager_desc: `**events**: Use for any question about the event Excel workbook (English or Burmese) — participants counts, agenda/program, volunteers, ferry routes/drivers, table seating (including "where will I sit"), beverages, karaoke. Examples: "how many participants", "list the agenda", "who is on ferry", "volunteer list", "where does X sit", "where will I sit".`,
  handler: async (context: MessageContext, logger: ILogger) => {
    const capability = new EventsCapability(logger);
    const result = await capability.processRequest(context);
    if (result.error) {
      logger.error(`❌ Error in Events Capability: ${result.error}`);
      return `Error looking up events: ${result.error}`;
    }
    return result.response || "No matching rows were found in the Excel file.";
  },
};
