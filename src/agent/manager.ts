import { ChatPrompt } from "@microsoft/teams.ai";
import { ILogger } from "@microsoft/teams.common";
import { EVENTS_CAPABILITY_DEFINITION } from "../capabilities/events/events";
import { CAPABILITY_DEFINITIONS } from "../capabilities/registry";
import { getEventSession, findActiveEventSession } from "../events/eventSession";
import { createChatModel, getModelConfig } from "../utils/config";
import { MessageContext } from "../utils/messageContext";
import { answerCache, normalizeCacheQuery } from "../utils/queryCache";
import { extractTimeRange, formatEventDisplayName } from "../utils/utils";
import { ServiceUnavailableError } from "../events/activeEventFile";
import { generateManagerPrompt } from "./prompt";

export interface ManagerResult {
  response: string;
}

function eventOnlyRefusal(fileName?: string | null): string {
  return `I only answer questions from "${formatEventDisplayName(fileName)}". Please ask about that event.`;
}

function noActiveEventMsg(): string {
  return "I only answer questions from the active event. Upload an .xlsx and send /start first.";
}

function activeEventGreeting(fileName?: string | null): string {
  return `Hi! Ask me anything about "${formatEventDisplayName(fileName)}"`;
}

function activeEventPrompt(fileName?: string | null): string {
  return `Ask me anything about "${formatEventDisplayName(fileName)}"`;
}

function isGreeting(text: string): boolean {
  const t = text.replace(/<\/?at>/gi, " ").replace(/\s+/g, " ").trim().toLowerCase();
  return /^(hi|hello|hey|good morning|good afternoon|good evening|မင်္ဂလာပါ|ဟယ်လို)[\s!.?]*$/i.test(
    t
  );
}

function looksLikeOffTopic(text: string): boolean {
  const t = text.toLowerCase();
  const offTopic =
    /\b(how to host|how do i host|host a project|deploy|docker|kubernetes|python|javascript|programming|write (a |an )?(code|essay|email)|weather|news|joke|recipe)\b/i.test(
      t
    );
  const eventHints =
    /\b(participant|agenda|volunteer|ferry|seat|table|menu|beverage|beer|juice|cocktail|karaoke|event|party|driver|dish|appetizer|schedule|who|how many|where)\b|ပါဝင်|အစီအစဉ်|ဖယ်ရီ|စားပွဲ|မီနူး/i.test(
      t
    );
  return offTopic && !eventHints;
}

// Manager prompt that coordinates all sub-tasks
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

    const historyLimit = Math.max(2, Number(process.env.MANAGER_HISTORY_LIMIT || 6) || 6);
    const history = await this.context.memory.values();
    const recentHistory = history.slice(-historyLimit).map((msg) => ({
      role: msg.role,
      content: msg.content,
    }));

    const prompt = new ChatPrompt({
      instructions: generateManagerPrompt(
        CAPABILITY_DEFINITIONS,
        this.context.userName,
        this.context.activeEventFileName
      ),
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
      const session = (await findActiveEventSession()) || (await getEventSession(this.context.conversationId));
      const eventFileName = session.activeFileName;

      if (isGreeting(this.context.text)) {
        const greeting =
          session.status === "active"
            ? activeEventGreeting(eventFileName)
            : noActiveEventMsg();
        return { response: greeting };
      }

      const trimmed = (this.context.text || "").replace(/<\/?at>/gi, " ").trim();
      if (!trimmed) {
        return {
          response:
            session.status === "active"
              ? activeEventPrompt(eventFileName)
              : noActiveEventMsg(),
        };
      }

      if (looksLikeOffTopic(this.context.text)) {
        this.logger.debug(`🚫 Off-topic refused: "${this.context.text}"`);
        return {
          response:
            session.status === "active"
              ? eventOnlyRefusal(eventFileName)
              : noActiveEventMsg(),
        };
      }

      if (session.status !== "active") {
        return { response: noActiveEventMsg() };
      }

      // Active event (from this chat or another): always use workbook capability
      this.logger.debug(
        `📊 Routing to events capability (workbook-only; activeSession=${session.conversationId}; file=${eventFileName || "unknown"})`
      );
      // Stash active workbook conversation so events lookup uses the indexed RAG scope
      (this.context as MessageContext).activeEventConversationId = session.conversationId;
      (this.context as MessageContext).activeEventFileName = eventFileName || undefined;
      const eventsResponse = await EVENTS_CAPABILITY_DEFINITION.handler(
        this.context,
        this.logger.child("events")
      );
      const text = eventsResponse || eventOnlyRefusal(eventFileName);

      if (
        normalizedQuestion &&
        text &&
        !/^sorry,/i.test(text) &&
        !/rate limit|error processing|provider is busy|error looking up events/i.test(text)
      ) {
        answerCache.set(answerKey, text);
      }

      if (/error looking up events|sorry, i encountered an error/i.test(text)) {
        throw new ServiceUnavailableError(text);
      }

      return { response: text };
    } catch (error) {
      if (error instanceof ServiceUnavailableError) {
        throw error;
      }
      const message = error instanceof Error ? error.message : "Unknown error";
      this.logger.error(`❌ Error in Manager: ${message}`);
      throw new ServiceUnavailableError(error);
    }
  }
}
