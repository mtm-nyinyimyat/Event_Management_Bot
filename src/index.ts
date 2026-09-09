import "dotenv/config";
import "./proxy";
import { ManagedIdentityCredential } from "@azure/identity";
import { TokenCredentials } from "@microsoft/teams.api";
import { App } from "@microsoft/teams.apps";
import { ConsoleLogger } from "@microsoft/teams.common";
import { DevtoolsPlugin } from "@microsoft/teams.dev";
import { ManagerPrompt } from "./agent/manager";
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
      const manager = new ManagerPrompt(context, logger.child("manager"));
      const result = await manager.processRequest();
      const formattedResult = finalizePromptResponse(result.response, context, logger);

      const sent = await send(formattedResult);
      formattedResult.id = sent.id;
      trackedMessages = createMessageRecords([activity, formattedResult]);
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
      "👋 Hi! I'm the Event Management bot. Ask me about the event Excel workbook and I will look up the details."
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
