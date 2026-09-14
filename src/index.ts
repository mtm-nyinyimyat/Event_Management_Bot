import "dotenv/config";
import "./proxy";
import { ManagedIdentityCredential } from "@azure/identity";
import { ActivityLike, MessageActivity, TokenCredentials } from "@microsoft/teams.api";
import { App } from "@microsoft/teams.apps";
import { ConsoleLogger, ILogger } from "@microsoft/teams.common";
import { DevtoolsPlugin } from "@microsoft/teams.dev";
import { ManagerPrompt } from "./agent/manager";
import {
  activityLikelyHasExcelUpload,
  downloadExcelFilesFromActivity,
  FILE_UPLOAD_HELP,
  findExcelAttachments,
  isEmoticonOnlyActivity,
} from "./events/chatExcelAttachment";
import {
  buildFileInfoAttachment,
  createServiceUnavailableReply,
  uploadAcceptedEventWorkbook,
} from "./events/activeEventFile";
import {
  addPendingShareUrls,
  addPendingUploads,
  endEventSession,
  extractSharePointUrls,
  findActiveEventSession,
  getEventIngestMode,
  isAnyEventActive,
  isEndEventCommand,
  isStartEventCommand,
  isUploadIngestMode,
  startEventSession,
} from "./events/eventSession";
import { IDatabase } from "./storage/database";
import { StorageFactory } from "./storage/storageFactory";
import { logModelConfigs, validateEnvironment } from "./utils/config";
import { createMessageContext } from "./utils/messageContext";
import { createMessageRecords, finalizePromptResponse, formatEventDisplayName } from "./utils/utils";

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

async function sendServiceUnavailable(
  send: (activity: ActivityLike) => Promise<unknown>,
  error: unknown,
  log: ILogger
): Promise<void> {
  const detail = describeNetworkError(error);
  log.error(`❌ User-facing service unavailable: ${detail}`);
  try {
    const reply = await createServiceUnavailableReply();
    await send(reply.message);
    if (reply.hasFileConsent) {
      log.debug(`📎 Sent file consent card for ${reply.fileName}`);
    } else {
      log.warn("Service unavailable reply sent without Excel consent card (no active workbook in memory)");
    }
  } catch (sendError) {
    log.error(`Failed to send service-unavailable reply: ${describeNetworkError(sendError)}`);
  }
}

async function replyFromManager(
  send: (activity: ActivityLike) => Promise<{ id?: string }>,
  context: Awaited<ReturnType<typeof createMessageContext>>,
  activity: Parameters<typeof createMessageRecords>[0][0],
  log: ILogger
): Promise<ReturnType<typeof createMessageRecords>> {
  try {
    const manager = new ManagerPrompt(context, log.child("manager"));
    const result = await manager.processRequest();
    const formattedResult = finalizePromptResponse(result.response, context, log);
    const sent = await send(formattedResult);
    if (sent.id) {
      formattedResult.id = sent.id;
    }
    return createMessageRecords([activity, formattedResult]);
  } catch (error) {
    await sendServiceUnavailable(send, error, log);
    return createMessageRecords([activity]);
  }
}

app.on("message", async ({ send, activity, api }) => {
  const botMentioned = activity.entities?.some((e) => e.type === "mention");
  // In groups, also handle Excel uploads without @mention (test /start flow).
  const excelUploadLikely = isUploadIngestMode() && activityLikelyHasExcelUpload(activity);
  const shouldHandle = !activity.conversation.isGroup || botMentioned || excelUploadLikely;

  const context = botMentioned
    ? await createMessageContext(storage, activity, api)
    : await createMessageContext(storage, activity);

  let trackedMessages;

  if (shouldHandle) {
    try {
      await send({ type: "typing" });
    } catch (error) {
      logger.warn(`Typing indicator failed (continuing): ${describeNetworkError(error)}`);
    }

    try {
      const text = stripMentions(activity.text);
      const conversationId = context.conversationId;
      const userName = context.userName || activity.from?.name || "unknown";
      const ingestMode = getEventIngestMode();

      // Lifecycle commands take priority over LLM routing
      if (isEmoticonOnlyActivity(activity)) {
        // Stickers/animated emoji are not questions and must not hit LLM or error fallback.
        const active = await findActiveEventSession();
        const eventLabel = formatEventDisplayName(active?.activeFileName, "");
        const ack = active
          ? eventLabel
            ? `Got it! Ask me anything about "${eventLabel}" (seating, ferry, menu, beverage, etc.).`
            : "Got it! Ask me anything about the event (seating, ferry, menu, beverage, etc.)."
          : "Got it! Upload an .xlsx and send /start when you are ready.";
        const sent = await send(ack);
        trackedMessages = createMessageRecords([activity]);
        logger.debug(`Ignored emoticon-only message; replied id=${sent.id}`);
      } else if (isStartEventCommand(text)) {
        if (await isAnyEventActive()) {
          const confirmation =
            "A previous event is still in progress and has not been ended yet.\n" +
            "Send /end (or endevent) to finish it before starting another event.";
          const sent = await send(confirmation);
          trackedMessages = createMessageRecords([activity]);
          logger.debug(`Blocked /start while event active; replied id=${sent.id}`);
        } else {
          const confirmation = await startEventSession({
            conversationId,
            startedBy: userName,
            logger: logger.child("event-session"),
          });
          const sent = await send(confirmation);
          trackedMessages = createMessageRecords([activity]);
          logger.debug(`Event start replied id=${sent.id}`);
        }
      } else if (isEndEventCommand(text)) {
        const confirmation = await endEventSession({
          conversationId,
          logger: logger.child("event-session"),
        });
        const sent = await send(confirmation);
        trackedMessages = createMessageRecords([activity]);
        logger.debug(`Event end replied id=${sent.id}`);
      } else if (ingestMode === "upload") {
        // --- Test mode: stage chat Excel uploads; SharePoint URLs ignored ---
        const maybeExcel =
          excelUploadLikely || findExcelAttachments(activity).length > 0;
        try {
          const downloaded = maybeExcel
            ? await downloadExcelFilesFromActivity(
                activity,
                conversationId,
                logger.child("excel-upload"),
                { allowGraphFallback: false }
              )
            : null;

          if (downloaded && downloaded.length > 0) {
            if (await isAnyEventActive()) {
              const confirmation =
                "A previous event is still in progress and has not been ended yet.\n" +
                "Send /end (or endevent) to finish it before uploading a new Excel file.";
              const sent = await send(confirmation);
              trackedMessages = createMessageRecords([activity]);
              logger.debug(`Blocked upload while event active; replied id=${sent.id}`);
            } else {
              const staged = await addPendingUploads(conversationId, downloaded);
              const list = staged.fileNames.map((name) => `• ${name}`).join("\n");
              const confirmation =
                `Saved ${staged.fileNames.length} Excel upload(s) for this chat:\n${list}\n\n` +
                `Pending file(s): ${staged.pendingCount}\n` +
                `I will not process them until you send /start\n` +
                `When the event is over, send /end to clear all event data.`;
              const sent = await send(confirmation);
              trackedMessages = createMessageRecords([activity]);
              logger.debug(`Pending Excel upload(s) saved; replied id=${sent.id}`);
            }
          } else if (excelUploadLikely) {
            await send(FILE_UPLOAD_HELP);
            trackedMessages = createMessageRecords([activity]);
          } else {
            trackedMessages = await replyFromManager(send, context, activity, logger);
          }
        } catch (uploadError) {
          await sendServiceUnavailable(send, uploadError, logger);
          trackedMessages = createMessageRecords([activity]);
        }
      } else {
        // --- SharePoint URL mode (EVENTS_INGEST_MODE=sharepoint) ---
        const shareUrls = extractSharePointUrls(activity.text);
        const textWithoutUrls = shareUrls
          .reduce((acc, url) => acc.replace(url, " "), text)
          .replace(/\s+/g, " ")
          .trim();
        if (shareUrls.length > 0 && textWithoutUrls.length < 8) {
          if (await isAnyEventActive()) {
            const confirmation =
              "A previous event is still in progress and has not been ended yet.\n" +
              "Send /end (or endevent) to finish it before pasting a new SharePoint URL.";
            const sent = await send(confirmation);
            trackedMessages = createMessageRecords([activity]);
            logger.debug(`Blocked SharePoint URL while event active; replied id=${sent.id}`);
          } else {
            const session = await addPendingShareUrls(conversationId, shareUrls);
            const confirmation =
              `Saved ${shareUrls.length} SharePoint link(s) for this chat.\n` +
              `Pending URL(s): ${session.pendingUrls.length}\n\n` +
              `I will not process the workbook until you send /start.\n` +
              `When the event is over, send /end or endevent to clear all event data.`;
            const sent = await send(confirmation);
            trackedMessages = createMessageRecords([activity]);
            logger.debug(`Pending SharePoint URL(s) saved; replied id=${sent.id}`);
          }
        } else {
          trackedMessages = await replyFromManager(send, context, activity, logger);
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
      await sendServiceUnavailable(send, error, logger);
      trackedMessages = createMessageRecords([activity]);
    }
  } else {
    trackedMessages = createMessageRecords([activity]);
  }

  logger.debug(trackedMessages);
  await context.memory.addMessages(trackedMessages);
});

app.on("file.consent.accept", async ({ activity, send }) => {
  try {
    const value = activity.value;
    const uploadInfo = value?.uploadInfo;
    const token =
      typeof value?.context?.token === "string" ? value.context.token : undefined;

    if (!uploadInfo?.uploadUrl) {
      logger.warn("file.consent.accept missing uploadUrl");
      return { status: 200 };
    }

    const uploaded = await uploadAcceptedEventWorkbook({
      uploadUrl: uploadInfo.uploadUrl,
      token,
      logger: logger.child("event-file"),
    });

    if (!uploaded) {
      await send("Service unavailable this time. The event workbook is no longer available.");
      return { status: 200 };
    }

    await send(
      new MessageActivity("Here is the active event Excel file.").addAttachments(
        buildFileInfoAttachment({
          name: uploadInfo.name || uploaded.fileName,
          contentUrl: uploadInfo.contentUrl,
          uniqueId: uploadInfo.uniqueId,
          fileType: uploadInfo.fileType || "xlsx",
        })
      )
    );
    return { status: 200 };
  } catch (error) {
    logger.error(`file.consent.accept failed: ${describeNetworkError(error)}`);
    try {
      await send("Service unavailable this time. Please try again shortly.");
    } catch {
      // ignore
    }
    return { status: 200 };
  }
});

app.on("file.consent.decline", async ({ send }) => {
  try {
    await send("Okay — I won’t send the Excel file.");
  } catch (error) {
    logger.warn(`file.consent.decline reply failed: ${describeNetworkError(error)}`);
  }
  return { status: 200 };
});

app.on("install.add", async ({ send }) => {
  try {
    if (isUploadIngestMode()) {
      await send(
        "👋 Hi! I'm the Event Management bot (upload test mode).\n\n" +
          "1) Upload an .xlsx (paperclip → Upload from this device)\n" +
          "2) Send /start to load it\n" +
          "3) Ask questions about the event (@mention me in groups)\n" +
          "4) Send /end when finished to clear the data"
      );
    } else {
      await send(
        "👋 Hi! I'm the Event Management bot.\n\n" +
          "1) Paste a SharePoint Excel URL\n" +
          "2) Send /start to load it\n" +
          "3) Ask questions about the event\n" +
          "4) Send /end when finished to clear the data"
      );
    }
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

    logger.debug(
      `📎 Event ingest mode: ${getEventIngestMode()} (set EVENTS_INGEST_MODE=sharepoint to use SharePoint URLs)`
    );
    logger.debug("✅ Storage initialized successfully");
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    logger.error(`❌ Configuration error: ${message}`);
    process.exit(1);
  }

  await app.start(port);

  logger.debug(`🚀 Collab Agent started on port ${port}`);
})();
