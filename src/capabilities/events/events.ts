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

  createPrompt(context: MessageContext): ChatPrompt {
    const modelConfig = this.getModelConfig("events");
    const conversationId = context.conversationId;
    const userId = context.userId;

    const prompt = new ChatPrompt({
      instructions: EVENTS_PROMPT,
      model: createChatModel(modelConfig),
    }).function(
      "lookup_events",
      "Search all sheets in the Excel workbook (chat upload preferred; remembered across chats) and return matching rows",
      LOOKUP_EVENTS_SCHEMA,
      async ({ query, max_results }: LookupEventsArgs) => {
        this.logger.debug(`📊 lookup_events query="${query || ""}"`);
        try {
          return JSON.stringify(
            await searchWorkbook(query || "", max_results ?? 20, { conversationId, userId })
          );
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
  manager_desc: `**events**: Use for questions about the party/event Excel workbook uploaded in any chat (English or Burmese) — dance members, team leaders, practice schedules, rehearsal locations, costumes, buy/rent lists, costs, and related remarks. The last uploaded workbook is remembered across private/group/channel chats. Examples: "who is on ရှမ်းအက", "Practice_Schedule for 09/13", "ဝယ်ငှားစာရင်း", "list dance members". If no file is loaded yet, ask the user to attach an .xlsx.`,
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
