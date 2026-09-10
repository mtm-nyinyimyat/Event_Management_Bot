import { ChatPrompt } from "@microsoft/teams.ai";
import { ILogger } from "@microsoft/teams.common";
import { CAPABILITY_DEFINITIONS } from "../capabilities/registry";
import { createChatModel, getModelConfig } from "../utils/config";
import { MessageContext } from "../utils/messageContext";
import { extractTimeRange } from "../utils/utils";
import { generateManagerPrompt } from "./prompt";

export interface ManagerResult {
  response: string;
}

// Manager prompt that coordinates all sub-tasks
// ChatPrompt with a couple of utility functions and delegation functions for each capability
export class ManagerPrompt {
  private prompt: ChatPrompt;

  private isInitialized = false;

  constructor(
    private context: MessageContext,
    private logger: ILogger
  ) {}

  private async createManagerPrompt(): Promise<ChatPrompt> {
    const managerModelConfig = getModelConfig("manager");
    this.logger.debug(
      `🤖 Manager model=${managerModelConfig.model} baseUrl=${managerModelConfig.baseUrl}`
    );

    // Keep recent turns only — full history can re-send old Groq errors and blow token budgets
    const history = await this.context.memory.values();
    const recentHistory = history.slice(-20);

    const prompt = new ChatPrompt({
      instructions: generateManagerPrompt(CAPABILITY_DEFINITIONS),
      model: createChatModel(managerModelConfig),
      messages: recentHistory,
    })
      .function(
        "calculate_time_range",
        "Parse natural language time expressions and calculate exact start/end times for time-based queries",
        {
          type: "object",
          properties: {
            time_phrase: {
              type: "string",
              description:
                'Natural language time expression extracted from the user request (e.g., "yesterday", "last week", "2 days ago", "past 3 hours")',
            },
          },
          required: ["time_phrase"],
        },
        async (time_phrase: string) => {
          this.logger.debug(`🕒 FUNCTION CALL: calculate_time_range - parsing "${time_phrase}"`);

          const timeRange = extractTimeRange(time_phrase);

          this.context.startTime = timeRange ? timeRange?.from.toISOString() : this.context.endTime;
          this.context.endTime = timeRange ? timeRange?.to.toISOString() : this.context.endTime;

          this.logger.debug(this.context.startTime);
          this.logger.debug(this.context.endTime);
        }
      )
      .function(
        "clear_conversation_history",
        "Clear conversation history in the database for the current conversation",
        async () => {
          await this.context.memory.clear();
          this.logger.debug("The conversation history has been cleared!");
        }
      );

    return prompt;
  }

  private addCapabilities() {
    for (const capability of CAPABILITY_DEFINITIONS) {
      this.prompt.function(
        `delegate_to_${capability.name}`,
        `Delegate to ${capability.name} capability`,
        async () => {
          return capability.handler(this.context, this.logger.child(capability.name));
        }
      );
    }
  }

  private async initialize(): Promise<void> {
    if (!this.isInitialized) {
      this.prompt = await this.createManagerPrompt();
      this.addCapabilities();
      this.isInitialized = true;
    }
  }

  async processRequest(): Promise<ManagerResult> {
    try {
      await this.initialize();
      const response = await this.prompt.send(this.context.text);
      return {
        response: response.content || "No response generated",
      };
    } catch (error) {
      const status =
        typeof error === "object" && error !== null && "status" in error
          ? (error as { status?: number }).status
          : undefined;
      const message = error instanceof Error ? error.message : "Unknown error";
      this.logger.error(`❌ Error in Manager: ${status ? `${status} ` : ""}${message}`);

      if (status === 503) {
        return {
          response:
            "The AI provider is busy right now (503). Wait a few seconds and send the message again.",
        };
      }

      return {
        response: `Sorry, I encountered an error processing your request: ${message}`,
      };
    }
  }
}
