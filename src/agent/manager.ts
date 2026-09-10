import { ChatPrompt } from "@microsoft/teams.ai";
import { ILogger } from "@microsoft/teams.common";
import { CAPABILITY_DEFINITIONS } from "../capabilities/registry";
import { createChatModel, getModelConfig } from "../utils/config";
import { MessageContext } from "../utils/messageContext";
import { answerCache, normalizeCacheQuery } from "../utils/queryCache";
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

    // Keep a short window only — Groq free gpt-oss-120b is ~8K TPM; history balloons fast
    const historyLimit = Math.max(2, Number(process.env.MANAGER_HISTORY_LIMIT || 6) || 6);
    const history = await this.context.memory.values();
    const recentHistory = history.slice(-historyLimit).map((msg) => ({
      role: msg.role,
      content: msg.content,
    }));

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
    const normalizedQuestion = normalizeCacheQuery(this.context.text);
    const answerKey = `answer::${normalizedQuestion}`;

    if (normalizedQuestion) {
      const cachedAnswer = answerCache.get(answerKey);
      if (cachedAnswer) {
        this.logger.debug(`🗃️ Answer cache hit for "${normalizedQuestion}"`);
        return { response: cachedAnswer };
      }
    }

    try {
      await this.initialize();
      const response = await this.prompt.send(this.context.text);
      const text = response.content || "No response generated";

      if (
        normalizedQuestion &&
        text &&
        !/^sorry,/i.test(text) &&
        !/rate limit|error processing|provider is busy/i.test(text)
      ) {
        answerCache.set(answerKey, text);
      }

      return { response: text };
    } catch (error) {
      const status =
        typeof error === "object" && error !== null && "status" in error
          ? (error as { status?: number }).status
          : undefined;
      const message = error instanceof Error ? error.message : "Unknown error";
      const cause =
        typeof error === "object" && error !== null && "cause" in error
          ? (error as { cause?: unknown }).cause
          : undefined;
      const causeText =
        cause instanceof Error
          ? cause.message
          : cause
            ? JSON.stringify(cause)
            : "";
      this.logger.error(
        `❌ Error in Manager: ${status ? `${status} ` : ""}${message}${
          causeText ? ` | cause: ${causeText}` : ""
        }`
      );

      if (status === 503) {
        return {
          response:
            "The AI provider is busy right now (503). Wait a few seconds and send the message again.",
        };
      }

      if (status === 429 || /rate limit|tokens per minute|TPM/i.test(message)) {
        return {
          response:
            "AI rate limit hit. Wait a bit and try again, or switch models.",
        };
      }

      return {
        response: `Sorry, I encountered an error processing your request: ${message}`,
      };
    }
  }
}
