import "dotenv/config";
import "./proxy";
import { ManagedIdentityCredential } from "@azure/identity";
import { TokenCredentials } from "@microsoft/teams.api";
import { App } from "@microsoft/teams.apps";
import { ConsoleLogger } from "@microsoft/teams.common";
import { DevtoolsPlugin } from "@microsoft/teams.dev";
import { ManagerPrompt } from "./agent/manager";
import {
  FILE_UPLOAD_HELP,
  ingestExcelAttachmentsFromActivity,
  looksLikeUndeliveredFileShare,
} from "./events/chatExcelAttachment";
import { getEventsSource, resolveUploadedWorkbook } from "./events/excelStore";
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
  skipAuth: !process.env.CLIENT_ID,
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

app.on("message", async ({ send, activity, api }) => {
  const botMentioned = activity.entities?.some((e) => e.type === "mention");
  const context = botMentioned
    ? await createMessageContext(storage, activity, api)
    : await createMessageContext(storage, activity);

  let trackedMessages;

  if (!activity.conversation.isGroup || botMentioned) {
    // process request if One-on-One chat or if @mentioned in Groupchat
    await send({ type: "typing" });

    let uploadNote: string | undefined;
    try {
      const ingested = await ingestExcelAttachmentsFromActivity(
        activity,
        context.conversationId,
        logger.child("excel-upload"),
        context.userId
      );
      if (ingested) {
        uploadNote = ingested.confirmation;
        logger.debug(`✅ Loaded chat Excel: ${ingested.files.map((f) => f.fileName).join(", ")}`);
      }
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      logger.error(`❌ Chat Excel ingest failed: ${message}`);
      const formattedFailure = finalizePromptResponse(
        `I couldn't read the attached Excel file.\n${message}`,
        context,
        logger
      );
      const sent = await send(formattedFailure);
      formattedFailure.id = sent.id;
      trackedMessages = createMessageRecords([activity, formattedFailure]);
      await context.memory.addMessages(trackedMessages);
      return;
    }

    const strippedText = (activity.text || "").replace(/<at>[^<]*<\/at>/gi, "").trim();
    const hasQuestion = Boolean(strippedText);
    const looksLikeGreetingOnly =
      hasQuestion && /^(hi|hello|hey|yo|thanks|thank you|ok|okay|hola)[\s!.?]*$/i.test(strippedText);

    // After a successful upload, confirm the real filename. Don't let the model invent names.
    if (uploadNote && (!hasQuestion || looksLikeGreetingOnly)) {
      const formattedUpload = finalizePromptResponse(uploadNote, context, logger);
      const sent = await send(formattedUpload);
      formattedUpload.id = sent.id;
      trackedMessages = createMessageRecords([activity, formattedUpload]);
      await context.memory.addMessages(trackedMessages);
      return;
    }

    const activeWorkbook = resolveUploadedWorkbook(context.conversationId, context.userId);

    // Teams showed a file card, but no downloadable Excel arrived for the bot.
    if (!uploadNote && looksLikeUndeliveredFileShare(activity)) {
      logger.warn(
        `⚠️ File share visible in Teams UI but no Excel bytes delivered. ${JSON.stringify(activity.attachments || [])}`
      );
      const formatted = finalizePromptResponse(
        `I can see you shared a file in Teams, but I still could not read the Excel bytes.\n\n${FILE_UPLOAD_HELP}`,
        context,
        logger
      );
      const sent = await send(formatted);
      formatted.id = sent.id;
      trackedMessages = createMessageRecords([activity, formatted]);
      await context.memory.addMessages(trackedMessages);
      return;
    }

    // Chat mode with no loaded workbook: never let the LLM invent 202611_* from history.
    if (getEventsSource() === "chat" && !activeWorkbook && !uploadNote) {
      const formatted = finalizePromptResponse(
        looksLikeGreetingOnly || !hasQuestion
          ? `Hi! I don't have an Excel workbook loaded yet.\n\n${FILE_UPLOAD_HELP}`
          : `I don't have an Excel workbook loaded, so I can't answer from a file yet.\n\n${FILE_UPLOAD_HELP}`,
        context,
        logger
      );
      const sent = await send(formatted);
      formatted.id = sent.id;
      trackedMessages = createMessageRecords([activity, formatted]);
      await context.memory.addMessages(trackedMessages);
      return;
    }

    // Tell the model which file is actually loaded (prevents stale 202611 replies from chat history).
    if (activeWorkbook?.fileName) {
      context.text =
        `[Active workbook: ${activeWorkbook.fileName} (source=${activeWorkbook.sourceType})]\n` +
        context.text;
    }

    const manager = new ManagerPrompt(context, logger.child("manager"));
    const result = await manager.processRequest();
    let formattedResult = finalizePromptResponse(result.response, context, logger);
    if (uploadNote) {
      formattedResult = finalizePromptResponse(
        `${uploadNote}\n\n${result.response || ""}`.trim(),
        context,
        logger
      );
    }

    const sent = await send(formattedResult);
    formattedResult.id = sent.id;

    trackedMessages = createMessageRecords([activity, formattedResult]);
  } else {
    trackedMessages = createMessageRecords([activity]);
  }

  logger.debug(trackedMessages);
  await context.memory.addMessages(trackedMessages);
});

app.on("install.add", async ({ send }) => {
  await send(
    "👋 Hi! I'm the Event Management bot.\n\n" +
      "Attach an .xlsx with the paperclip in a 1:1 chat (Upload from this device). " +
      "I only use a file after I confirm its exact name — I will not silently use a local 202611_* workbook."
  );
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
