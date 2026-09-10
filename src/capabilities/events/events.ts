import { ChatPrompt } from "@microsoft/teams.ai";
import { ILogger } from "@microsoft/teams.common";
import { searchWorkbook } from "../../events/excelStore";
import { createChatModel } from "../../utils/config";
import { MessageContext } from "../../utils/messageContext";
import { BaseCapability, CapabilityDefinition } from "../capability";
import { EVENTS_PROMPT } from "./prompt";
import { LOOKUP_EVENTS_SCHEMA, LookupEventsArgs } from "./schema";

export class EventsCapability extends BaseCapability {
  readonly name = "events";

  createPrompt(_context: MessageContext): ChatPrompt {
    const modelConfig = this.getModelConfig("events");

    const prompt = new ChatPrompt({
      instructions: EVENTS_PROMPT,
      model: createChatModel(modelConfig),
    }).function(
      "lookup_events",
      "Search the event Excel workbook with hybrid RAG (semantic + keyword) and return summaries plus matching rows",
      LOOKUP_EVENTS_SCHEMA,
      async ({ query, max_results }: LookupEventsArgs) => {
        this.logger.debug(`📊 lookup_events query="${query || ""}"`);
        try {
          return JSON.stringify(await searchWorkbook(query || "", max_results ?? 80));
        } catch (error) {
          const message = error instanceof Error ? error.message : "Unknown Excel error";
          this.logger.error(`❌ Excel lookup failed: ${message}`);
          return JSON.stringify({ error: message, sheets: [], match_count: 0 });
        }
      }
    );

    this.logger.debug("Initialized Events Capability");
    return prompt;
  }
}

export const EVENTS_CAPABILITY_DEFINITION: CapabilityDefinition = {
  name: "events",
  manager_desc: `**events**: Use for any question about the event Excel workbook (English or Burmese) — participants counts, agenda/program, volunteers, ferry routes/drivers, table seating, beverages, karaoke. Examples: "how many participants", "list the agenda", "who is on ferry", "volunteer list", "where does X sit".`,
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
