import "dotenv/config";
import "./proxy";
import { ManagedIdentityCredential } from "@azure/identity";
import { TokenCredentials } from "@microsoft/teams.api";
import { App } from "@microsoft/teams.apps";
import { ConsoleLogger } from "@microsoft/teams.common";
import { DevtoolsPlugin } from "@microsoft/teams.dev";
import { ManagerPrompt } from "./agent/manager";
import {
  addPendingShareUrls,
  endEventSession,
  extractSharePointUrls,
  isEndEventCommand,
  isStartEventCommand,
  startEventSession,
} from "./events/eventSession";
import { IDatabase } from "./storage/database";
import { StorageFactory } from "./storage/storageFactory";
import { logModelConfigs, validateEnvironment } from "./utils/config";
import { createMessageContext } from "./utils/messageContext";
import { createMessageRecords, finalizePromptResponse } from "./utils/utils";

const logger = new ConsoleLogger("collaborator", { level: "debug" });

const createTokenFactory = () => {
  return async (scope: string | string[], tenantId?: string): Promise<string> => {
    const managedIdentityCredential = new ManagedIdentityCredential({
      clientId: process.env.CLIENT_ID,
    });
    const scopes = Array.isArray(scope) ? scope : [scope];
    const tokenResponse = await managedIdentityCredential.getToken(scopes, {
      tenantId: tenantId,
    });

    return tokenResponse.token;
  };
};

// Configure authentication using TokenCredentials
const tokenCredentials: TokenCredentials = {
  clientId: process.env.CLIENT_ID || "",
  token: createTokenFactory(),
};

// Use managed identity in cloud environment, otherwise use devtools plugin for local development
const options =
  process.env.BOT_TYPE === "UserAssignedMsi"
    ? { ...tokenCredentials }
    : { plugins: [new DevtoolsPlugin()] };

const app = new App({
  ...options,
  logger,
  // JWT check fetches login.botframework.com. If that host is blocked, set SKIP_BOT_AUTH=1 locally.
  skipAuth: !process.env.CLIENT_ID || process.env.SKIP_BOT_AUTH === "1",
});

// Initialize storage
let storage: IDatabase;
let feedbackStorage: IDatabase;

app.on("message.submit.feedback", async ({ activity }) => {
  try {
    const { reaction, feedback: feedbackJson } = activity.value.actionValue;

    if (!activity.replyToId) {
      logger.warn(`No replyToId found for messageId ${activity.id}`);
      return;
    }

    const success = await feedbackStorage.recordFeedback(
      activity.replyToId,
      reaction,
      feedbackJson
    );

    if (success) {
      logger.debug(`✅ Successfully recorded feedback for message ${activity.replyToId}`);
    } else {
      logger.warn(`Failed to record feedback for message ${activity.replyToId}`);
    }
  } catch (error) {
    logger.error(
      `Error processing feedback: ${error instanceof Error ? error.message : "Unknown error"}`
    );
  }
});

function describeNetworkError(error: unknown): string {
  const err = error as { message?: string; code?: string; cause?: { message?: string; code?: string } };
  const parts = [err?.message, err?.code, err?.cause?.message, err?.cause?.code].filter(Boolean);
  return parts.join(" | ") || String(error);
}

function isOutboundBlocked(error: unknown): boolean {
  const text = describeNetworkError(error);
  return /ECONNREFUSED|ETIMEDOUT|ENOTFOUND|certificate|unable to verify/i.test(text);
}

function stripMentions(text?: string): string {
  return String(text || "")
    .replace(/<\/?at>/gi, " ")
    .replace(/<at>[^<]*<\/at>/gi, " ")
    .replace(/\s+/g, " ")
    .trim();
}

app.on("message", async ({ send, activity, api }) => {
  const botMentioned = activity.entities?.some((e) => e.type === "mention");
  const context = botMentioned
    ? await createMessageContext(storage, activity, api)
    : await createMessageContext(storage, activity);

  let trackedMessages;

  if (!activity.conversation.isGroup || botMentioned) {
    try {
      await send({ type: "typing" });
    } catch (error) {
      logger.warn(`Typing indicator failed (continuing): ${describeNetworkError(error)}`);
    }

    try {
      const text = stripMentions(activity.text);
      const conversationId = context.conversationId;
      const userName = context.userName || activity.from?.name || "unknown";

      // Lifecycle commands take priority over LLM routing
      if (isStartEventCommand(text)) {
        const confirmation = await startEventSession({
          conversationId,
          startedBy: userName,
          logger: logger.child("event-session"),
        });
        const sent = await send(confirmation);
        trackedMessages = createMessageRecords([activity]);
        logger.debug(`Event start replied id=${sent.id}`);
      } else if (isEndEventCommand(text)) {
        const confirmation = await endEventSession({
          conversationId,
          logger: logger.child("event-session"),
        });
        const sent = await send(confirmation);
        trackedMessages = createMessageRecords([activity]);
        logger.debug(`Event end replied id=${sent.id}`);
      } else {
        const shareUrls = extractSharePointUrls(activity.text);
        // Stage SharePoint URLs only when the message is primarily a link paste
        const textWithoutUrls = shareUrls
          .reduce((acc, url) => acc.replace(url, " "), text)
          .replace(/\s+/g, " ")
          .trim();
        if (shareUrls.length > 0 && textWithoutUrls.length < 8) {
          const session = await addPendingShareUrls(conversationId, shareUrls);
          const confirmation =
            `Saved ${shareUrls.length} SharePoint link(s) for this chat.\n` +
            `Pending URL(s): ${session.pendingUrls.length}\n\n` +
            `I will not process the workbook until you send /start or startevent.\n` +
            `When the event is over, send /end or endevent to clear all event data.`;
          const sent = await send(confirmation);
          trackedMessages = createMessageRecords([activity]);
          logger.debug(`Pending SharePoint URL(s) saved; replied id=${sent.id}`);
        } else {
          const manager = new ManagerPrompt(context, logger.child("manager"));
          const result = await manager.processRequest();
          const formattedResult = finalizePromptResponse(result.response, context, logger);

          const sent = await send(formattedResult);
          formattedResult.id = sent.id;
          trackedMessages = createMessageRecords([activity, formattedResult]);
        }
      }
    } catch (error) {
      const detail = describeNetworkError(error);
      logger.error(`❌ Failed to handle/reply to message: ${detail}`);
      if (isOutboundBlocked(error)) {
        logger.error(
          "Outbound HTTPS to Microsoft is blocked (VPN/proxy/firewall). " +
            "The bot received the Teams message but cannot call login.botframework.com or smba.trafficmanager.net:443. " +
            "Connect the corporate VPN, or set HTTPS_PROXY in .env if you use a proxy."
        );
      }
      try {
        const failMsg =
          error instanceof Error &&
          /SharePoint|download|No active event|No SharePoint|Could not download/i.test(error.message)
            ? error.message
            : "Sorry — I hit an error handling that message. Please try again.";
        await send(failMsg);
      } catch {
        // ignore secondary send failure
      }
      trackedMessages = createMessageRecords([activity]);
    }
  } else {
    trackedMessages = createMessageRecords([activity]);
  }

  logger.debug(trackedMessages);
  await context.memory.addMessages(trackedMessages);
});

app.on("install.add", async ({ send }) => {
  try {
    await send(
      "👋 Hi! I'm the Event Management bot.\n\n" +
        "1) Paste a SharePoint Excel URL\n" +
        "2) Send /start (or startevent) to load it\n" +
        "3) Ask questions about the event\n" +
        "4) Send /end (or endevent) when finished to clear the data"
    );
  } catch (error) {
    logger.error(`Welcome message failed: ${describeNetworkError(error)}`);
  }
});

(async () => {
  const port = process.env.PORT || process.env.port || 3978;
  try {
    validateEnvironment(logger);
    logModelConfigs(logger);

    // Initialize storage
    storage = await StorageFactory.createStorage(logger.child("storage"));
    feedbackStorage = storage;

    logger.debug("✅ Storage initialized successfully");
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    logger.error(`❌ Configuration error: ${message}`);
    process.exit(1);
  }

  await app.start(port);

  logger.debug(`🚀 Collab Agent started on port ${port}`);
})();
