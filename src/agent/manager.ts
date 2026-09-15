import { ILogger } from "@microsoft/teams.common";
import { EVENTS_CAPABILITY_DEFINITION } from "../capabilities/events/events";
import { ServiceUnavailableError } from "../events/activeEventFile";
import { findActiveEventSession, getEventSession } from "../events/eventSession";
import { MessageContext } from "../utils/messageContext";
import { answerCache, normalizeCacheQuery } from "../utils/queryCache";
import { formatEventDisplayName } from "../utils/utils";

export interface ManagerResult {
  response: string;
}

function eventOnlyRefusal(fileName?: string | null): string {
  return `I mostly just know about "${formatEventDisplayName(fileName)}" — what were you wondering about for that?`;
}

function noActiveEventMsg(): string {
  return "We're not live yet — upload an .xlsx and send /start whenever you're ready.";
}

function activeEventGreeting(fileName?: string | null): string {
  return `Hey — what's up? Anything you need for "${formatEventDisplayName(fileName)}"?`;
}

function activeEventPrompt(fileName?: string | null): string {
  return `What's on your mind for "${formatEventDisplayName(fileName)}"?`;
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

/** Routes event workbook questions to the events capability (no general LLM manager). */
export class ManagerPrompt {
  constructor(
    private context: MessageContext,
    private logger: ILogger
  ) {}

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
      const session =
        (await findActiveEventSession()) || (await getEventSession(this.context.conversationId));
      const eventFileName = session.activeFileName;

      if (isGreeting(this.context.text)) {
        return {
          response:
            session.status === "active"
              ? activeEventGreeting(eventFileName)
              : noActiveEventMsg(),
        };
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

      this.logger.debug(
        `📊 Routing to events capability (activeSession=${session.conversationId}; file=${eventFileName || "unknown"})`
      );
      this.context.activeEventConversationId = session.conversationId;
      this.context.activeEventFileName = eventFileName || undefined;

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
