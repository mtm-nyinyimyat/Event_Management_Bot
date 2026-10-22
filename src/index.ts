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
  extractSharePointUrlsFromActivity,
  extractSharePointUrlsFromGraphMessage,
  findActiveEventSession,
  getEventIngestMode,
  isAnyEventActive,
  isEndEventCommand,
  isStartEventCommand,
  isUploadIngestMode,
  mergeSharePointUrls,
  startEventSession,
} from "./events/eventSession";
import { fetchTeamsMessageFromGraph } from "./events/graphExcelClient";
import {
  ensureGraphChangeSubscription,
  runDeltaSyncTick,
  startDeltaPolling,
} from "./events/deltaSync";
import {
  completeSsoTokenExchange,
  getAuthStartUrl,
  getGraphAuthMode,
  getGraphOperatorUserId,
  GraphAuthRequiredError,
  saveUserGraphTokens,
  userHasGraphAuth,
} from "./events/graphAuth";
import {
  resolveTeamsAuthUserIds,
  runWithGraphAuthUserAsync,
} from "./events/graphAuthContext";
import {
  ensureGraphTokenSchema,
  getStoredGraphTokens,
} from "./events/graphTokenStore";
import { registerEntraAuthRoutes } from "./events/entraAuthRoutes";
import {
  buildGraphSsoOAuthCard,
  extractTokenExchangeAssertion,
} from "./events/teamsSso";
import {
  registerGraphWebhookRoutes,
  setGraphWebhookChangeHandler,
} from "./events/graphWebhook";
import { IDatabase } from "./storage/database";
import { StorageFactory } from "./storage/storageFactory";
import { logModelConfigs, validateEnvironment } from "./utils/config";
import { createMessageContext, MessageContext } from "./utils/messageContext";
import {
  BotReplyOptions,
  createMessageRecords,
  finalizePromptResponse,
  formatEventDisplayName,
} from "./utils/utils";

const logger = new ConsoleLogger("collaborator", { level: "debug" });

/** @mention sender in groups/channels only; reply in the thread of the triggering post. */
function replyOptionsFromContext(
  context: MessageContext,
  extras?: Pick<BotReplyOptions, "withFeedback">
): BotReplyOptions {
  return {
    mentionAccount: context.isPersonalChat ? undefined : context.senderAccount,
    replyToId: context.replyToId,
    withFeedback: extras?.withFeedback,
  };
}

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
  applicationIdUri:
    process.env.GRAPH_SSO_RESOURCE_URI?.trim() ||
    (process.env.BOT_ID || process.env.CLIENT_ID
      ? `api://botid-${process.env.BOT_ID || process.env.CLIENT_ID}`
      : undefined),
  // JWT check fetches login.botframework.com. If that host is blocked, set SKIP_BOT_AUTH=1 locally.
  skipAuth: !process.env.CLIENT_ID || process.env.SKIP_BOT_AUTH === "1",
});

const webhookLogger = logger.child("graph-webhook");
registerGraphWebhookRoutes(app, webhookLogger);
registerEntraAuthRoutes(app, logger.child("entra-auth"));
setGraphWebhookChangeHandler((log) => {
  void runDeltaSyncTick(log || webhookLogger).catch((error) => {
    const message = error instanceof Error ? error.message : String(error);
    webhookLogger.warn(`Webhook-triggered delta sync failed: ${message}`);
  });
});

// Initialize storage
let storage: IDatabase;
let feedbackStorage: IDatabase;

function resolveActivityUserId(activity: {
  from?: { aadObjectId?: string; id?: string };
}): string {
  return resolveTeamsAuthUserIds(activity.from).canonical;
}

async function promptGraphSso(
  send: (activity: ActivityLike) => Promise<unknown>,
  userId: string,
  log: ILogger,
  options?: { isPersonalChat?: boolean }
): Promise<void> {
  // Teams silent SSO / OAuth cards are unreliable outside personal scope.
  // Prefer a browser sign-in link in group/channel so one personal consent is enough.
  if (options?.isPersonalChat === false) {
    try {
      const startUrl = getAuthStartUrl(userId);
      await send(
        "Please sign in once so I can access Microsoft 365 on your behalf. " +
          "After that, the same sign-in works in this group/channel and in 1:1 chat.\n\n" +
          `Sign in: ${startUrl}`
      );
      log.debug("Sent Graph sign-in link for group/channel (no OAuth card)");
      return;
    } catch (error) {
      log.warn(
        `Group sign-in link failed, falling back to OAuth card: ${
          error instanceof Error ? error.message : String(error)
        }`
      );
    }
  }

  try {
    await send(buildGraphSsoOAuthCard({ userId }));
    log.debug("Sent Teams SSO card (tokenExchange + /auth/start fallback)");
  } catch (error) {
    log.error(
      `Failed to send SSO card: ${error instanceof Error ? error.message : String(error)}`
    );
    try {
      const startUrl = getAuthStartUrl(userId);
      await send(
        `Please sign in so I can access Microsoft 365 on your behalf:\n${startUrl}`
      );
    } catch {
      await send(
        "Please sign in so I can access Microsoft 365 on your behalf. " +
          "Set PUBLIC_BASE_URL (HTTPS tunnel) for the sign-in link."
      );
    }
  }
}

/**
 * Ensure Graph SSO tokens exist for this user when GRAPH_AUTH_MODE=sso.
 * Returns false when an OAuth card was sent and the caller should stop.
 */
async function ensureMessageGraphAuth(
  send: (activity: ActivityLike) => Promise<unknown>,
  userId: string,
  log: ILogger,
  options?: { alternateUserIds?: string[]; isPersonalChat?: boolean }
): Promise<boolean> {
  if (getGraphAuthMode() !== "sso") {
    return true;
  }
  if (!userId) {
    log.warn("Cannot start Graph SSO: missing activity.from.aadObjectId");
    return true;
  }
  // Validates / refreshes the stored token — not merely that a DB row exists.
  // Also checks alternate Teams ids so DM sign-in is reused in group/channel.
  if (await userHasGraphAuth(userId, options?.alternateUserIds || [])) {
    return true;
  }
  log.debug(`No usable Graph SSO token for ${userId}; sending sign-in prompt`);
  await promptGraphSso(send, userId, log, {
    isPersonalChat: options?.isPersonalChat,
  });
  return false;
}

app.on("signin.token-exchange", async ({ activity, log }) => {
  const authIds = resolveTeamsAuthUserIds(activity.from);
  const userId = authIds.canonical;
  const assertion = extractTokenExchangeAssertion(activity);
  const connectionName =
    (activity.value as { connectionName?: string } | undefined)?.connectionName ||
    "GraphSSO";
  const exchangeId =
    (activity.value as { id?: string } | undefined)?.id || "token-exchange";

  if (!userId || !assertion) {
    log.warn("signin.token-exchange missing userId or token");
    return {
      status: 412,
      body: {
        id: exchangeId,
        connectionName,
        failureDetail: "missing user or SSO token",
      },
    };
  }

  try {
    await completeSsoTokenExchange(userId, assertion);
    // Mirror under Teams id when both are present so DM / group / channel share one sign-in.
    const stored = await getStoredGraphTokens(userId);
    if (stored) {
      for (const alt of authIds.alternates) {
        await saveUserGraphTokens(alt, {
          accessToken: stored.accessToken,
          refreshToken: stored.refreshToken ?? undefined,
          expiresAt: stored.expiresAt,
          scope: stored.scope,
        });
      }
    }
    log.debug(`Graph SSO OBO complete for user ${userId}`);
    return { status: 200 };
  } catch (error) {
    log.error(
      `Graph SSO OBO failed: ${error instanceof Error ? error.message : String(error)}`
    );
    return {
      status: 412,
      body: {
        id: exchangeId,
        connectionName,
        failureDetail: "unable to exchange token for Graph",
      },
    };
  }
});

/** Consent is handled by /auth/start → /auth/end (Entra auth code), not Bot Framework verify-state. */
app.on("signin.verify-state", async ({ log }) => {
  log.debug(
    "signin.verify-state ignored — use Entra /auth/start consent fallback (no Azure Bot OAuth)"
  );
  return { status: 200 };
});

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
  return (
    String(text || "")
      .replace(/<at\b[^>]*>[\s\S]*?<\/at>/gi, " ")
      .replace(/<\/?at\b[^>]*>/gi, " ")
      .replace(/&nbsp;/gi, " ")
      .replace(/\u200b|\ufeff|\u00a0/g, " ")
      .replace(/\s+/g, " ")
      .trim()
  );
}

/** Teams mention entities vary; also treat explicit <at>bot</at> body as a mention. */
function isBotMentioned(activity: {
  entities?: unknown[];
  text?: string;
}): boolean {
  const entities = activity.entities || [];
  const hasMentionEntity = entities.some((entity) => {
    const e = entity as { type?: string; mentioned?: { id?: string; name?: string } };
    if (e?.type === "mention") {
      return true;
    }
    return Boolean(e?.mentioned);
  });
  if (hasMentionEntity) {
    return true;
  }
  return /<at\b[^>]*>[\s\S]*?<\/at>/i.test(String(activity.text || ""));
}

/** Best-effort Graph re-read: BF often omits channel post previews/files. */
async function fetchGraphMessageSafe(
  activity: {
    id?: string;
    conversation?: { id?: string };
    channelData?: unknown;
  }
): Promise<Awaited<ReturnType<typeof fetchTeamsMessageFromGraph>>> {
  const messageId = activity.id;
  const conversationId = activity.conversation?.id;
  if (!messageId || !conversationId) {
    return undefined;
  }
  const channelData = (activity.channelData || {}) as {
    team?: { id?: string };
    channel?: { id?: string };
  };
  try {
    return await fetchTeamsMessageFromGraph(conversationId, messageId, {
      team: channelData.team,
      channel: channelData.channel,
    });
  } catch (error) {
    logger.debug(
      `Graph message enrich skipped: ${error instanceof Error ? error.message : String(error)}`
    );
    return undefined;
  }
}

async function sendServiceUnavailable(
  send: (activity: ActivityLike) => Promise<unknown>,
  error: unknown,
  log: ILogger,
  context?: MessageContext
): Promise<void> {
  const detail = describeNetworkError(error);
  log.error(`❌ User-facing service unavailable: ${detail}`);
  try {
    const reply = await createServiceUnavailableReply();
    if (context?.replyToId) {
      reply.message.replyToId = context.replyToId;
    }
    if (
      context &&
      !context.isPersonalChat &&
      context.senderAccount?.id &&
      context.senderAccount.name
    ) {
      reply.message.addMention(context.senderAccount, { addText: false });
      const body = String(reply.message.text || "").trim();
      reply.message.text = `<at>${context.senderAccount.name}</at> ${body}`;
    }
    await send(reply.message);
  } catch (sendError) {
    log.error(`Failed to send service-unavailable reply: ${describeNetworkError(sendError)}`);
  }
}

async function replyFromManager(
  send: (activity: ActivityLike) => Promise<{ id?: string }>,
  context: MessageContext,
  activity: Parameters<typeof createMessageRecords>[0][0],
  log: ILogger
): Promise<ReturnType<typeof createMessageRecords>> {
  try {
    const manager = new ManagerPrompt(context, log.child("manager"));
    const result = await manager.processRequest();
    const formattedResult = finalizePromptResponse(
      result.response,
      replyOptionsFromContext(context)
    );
    const sent = await send(formattedResult);
    if (sent.id) {
      formattedResult.id = sent.id;
    }
    return createMessageRecords([activity, formattedResult]);
  } catch (error) {
    await sendServiceUnavailable(send, error, log, context);
    return createMessageRecords([activity]);
  }
}

app.on("message", async ({ send, activity, api }) => {
  const botMentioned = isBotMentioned(activity);
  // In groups, also handle Excel uploads without @mention (test /start flow).
  const excelUploadLikely = isUploadIngestMode() && activityLikelyHasExcelUpload(activity);
  const shouldHandle = !activity.conversation.isGroup || botMentioned || excelUploadLikely;

  const context = botMentioned
    ? await createMessageContext(storage, activity, api)
    : await createMessageContext(storage, activity);

  const authIds = resolveTeamsAuthUserIds(activity.from);
  const graphUserId = authIds.canonical || context.userId || "";

  await runWithGraphAuthUserAsync(
    graphUserId || "anonymous",
    async () => {
  let trackedMessages;

  if (shouldHandle) {
    try {
      await send({ type: "typing" });
    } catch (error) {
      logger.warn(`Typing indicator failed (continuing): ${describeNetworkError(error)}`);
    }

    try {
      // Channel posts: BF payload often lacks preview/file links — re-read via Graph.
      // Best-effort only — do not require every group member to be signed in for Q&A.
      const graphMessage = await fetchGraphMessageSafe(activity);
      const activityText = graphMessage?.body?.content || activity.text;
      const text = stripMentions(activityText);
      // Prefer Graph body so the manager sees full channel-post content after @mention.
      if (graphMessage?.body?.content && !String(activity.text || "").includes(graphMessage.body.content)) {
        context.text = activityText;
      }
      if (isStartEventCommand(activityText) || isEndEventCommand(activityText)) {
        context.text = stripMentions(activityText);
        logger.debug(`Command from Graph/BF body; context.text="${context.text}"`);
      }
      if (graphMessage) {
        logger.debug(
          `Graph enrich ok: attachments=${graphMessage.attachments?.length || 0} bodyLen=${graphMessage.body?.content?.length || 0}`
        );
      }
      const conversationId = context.conversationId;
      const userName = context.userName || activity.from?.name || "unknown";
      const ingestMode = getEventIngestMode();

      // Extract SharePoint links from this post FIRST so "/start" in the same
      // message (or thread) can load them. BF often omits preview/file URLs;
      // Graph re-read fills them in.
      let shareUrlsThisMessage: string[] = [];
      if (ingestMode === "sharepoint" && !isEmoticonOnlyActivity(activity)) {
        shareUrlsThisMessage = mergeSharePointUrls(
          extractSharePointUrlsFromActivity(activity),
          extractSharePointUrlsFromGraphMessage(graphMessage)
        );
        logger.debug(
          `Share URLs in message: count=${shareUrlsThisMessage.length} ` +
            `bfAttachments=${activity.attachments?.length || 0} ` +
            `graphAttachments=${graphMessage?.attachments?.length || 0} ` +
            `urls=${JSON.stringify(shareUrlsThisMessage)}`
        );
        if (shareUrlsThisMessage.length > 0 && !(await isAnyEventActive())) {
          const staged = await addPendingShareUrls(conversationId, shareUrlsThisMessage);
          logger.debug(
            `Pre-staged SharePoint URL(s) before command routing (pending=${staged.pendingUrls.length})`
          );
        }
      }

      // Lifecycle commands take priority over LLM routing
      const commandTexts = [text, stripMentions(activity.text), String(activity.text || "")];
      const isStartCmd = commandTexts.some((value) => isStartEventCommand(value));
      const isEndCmd = commandTexts.some((value) => isEndEventCommand(value));
      if (isStartCmd || isEndCmd) {
        logger.debug(
          `Command match start=${isStartCmd} end=${isEndCmd} raw="${String(activity.text || "").slice(0, 80)}" stripped="${text.slice(0, 80)}"`
        );
      }
      if (isEmoticonOnlyActivity(activity)) {
        // Stickers/animated emoji are not questions and must not hit LLM or error fallback.
        const active = await findActiveEventSession();
        const eventLabel = formatEventDisplayName(active?.activeFileName, "");
        const ack = active
          ? eventLabel
            ? `Nice 🙂 What do you need for "${eventLabel}"?`
            : "Nice 🙂 What do you need for the event?"
          : "Nice 🙂 Upload an .xlsx and hit /start when you're ready.";
        const sent = await send(
          finalizePromptResponse(ack, replyOptionsFromContext(context, { withFeedback: false }))
        );
        trackedMessages = createMessageRecords([activity]);
        logger.debug(`Ignored emoticon-only message; replied id=${sent.id}`);
      } else if (isStartCmd) {
        if (await isAnyEventActive()) {
          const confirmation =
            "There's already an event running.\n" +
            "Send /end first if you want to start a new one.";
          const sent = await send(
            finalizePromptResponse(
              confirmation,
              replyOptionsFromContext(context, { withFeedback: false })
            )
          );
          trackedMessages = createMessageRecords([activity]);
          logger.debug(`Blocked /start while event active; replied id=${sent.id}`);
        } else {
          // SharePoint /start downloads via Graph — only the starter must be signed in.
          // Other group members can @mention and ask about the loaded event without SSO.
          if (getGraphAuthMode() === "sso" && getEventIngestMode() === "sharepoint") {
            const ok = await ensureMessageGraphAuth(send, graphUserId, logger, {
              alternateUserIds: authIds.alternates,
              isPersonalChat: context.isPersonalChat,
            });
            if (!ok) {
              trackedMessages = createMessageRecords([activity]);
              return;
            }
          }
          // Ensure links from this same post are saved before start reads pendingUrls.
          if (shareUrlsThisMessage.length > 0) {
            await addPendingShareUrls(conversationId, shareUrlsThisMessage);
          }
          const confirmation = await startEventSession({
            conversationId,
            startedBy: userName,
            logger: logger.child("event-session"),
          });
          const sent = await send(
            finalizePromptResponse(
              confirmation,
              replyOptionsFromContext(context, { withFeedback: false })
            )
          );
          trackedMessages = createMessageRecords([activity]);
          logger.debug(`Event start replied id=${sent.id}`);
        }
      } else if (isEndCmd) {
        const confirmation = await endEventSession({
          conversationId,
          logger: logger.child("event-session"),
        });
        const sent = await send(
          finalizePromptResponse(
            confirmation,
            replyOptionsFromContext(context, { withFeedback: false })
          )
        );
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
                logger.child("excel-upload")
              )
            : null;

          if (downloaded && downloaded.length > 0) {
            if (await isAnyEventActive()) {
              const confirmation =
                "There's already an event running.\n" +
                "Send /end first if you want to upload a new Excel.";
              const sent = await send(
                finalizePromptResponse(
                  confirmation,
                  replyOptionsFromContext(context, { withFeedback: false })
                )
              );
              trackedMessages = createMessageRecords([activity]);
              logger.debug(`Blocked upload while event active; replied id=${sent.id}`);
            } else {
              const staged = await addPendingUploads(conversationId, downloaded);
              const list = staged.fileNames.map((name) => `• ${name}`).join("\n");
              const confirmation =
                `Got your Excel 🙂\n${list}\n\n` +
                `Send /start when you want me to load it.\n` +
                `Send /end later to clear everything.`;
              const sent = await send(
                finalizePromptResponse(
                  confirmation,
                  replyOptionsFromContext(context, { withFeedback: false })
                )
              );
              trackedMessages = createMessageRecords([activity]);
              logger.debug(`Pending Excel upload(s) saved; replied id=${sent.id}`);
            }
          } else if (excelUploadLikely) {
            await send(
              finalizePromptResponse(
                FILE_UPLOAD_HELP,
                replyOptionsFromContext(context, { withFeedback: false })
              )
            );
            trackedMessages = createMessageRecords([activity]);
          } else {
            trackedMessages = await replyFromManager(send, context, activity, logger);
          }
        } catch (uploadError) {
          await sendServiceUnavailable(send, uploadError, logger, context);
          trackedMessages = createMessageRecords([activity]);
        }
      } else {
        // --- SharePoint URL mode (EVENTS_INGEST_MODE=sharepoint) ---
        // Prefer URLs already extracted (BF + Graph) above.
        const shareUrls = shareUrlsThisMessage;
        const textUrls = mergeSharePointUrls(
          extractSharePointUrls(activity.text),
          extractSharePointUrls(activityText)
        );
        const textWithoutUrls = textUrls
          .reduce((acc, url) => acc.split(url).join(" "), text)
          .replace(/<[^>]+>/g, " ")
          .replace(/\s+/g, " ")
          .trim();
        const remainingWords = textWithoutUrls ? textWithoutUrls.split(/\s+/).length : 0;
        const hasLinkContent = shareUrls.length > 0;
        // Bare paste, short ask after @mention, or preview/file-only card.
        const looksLikeLinkShare =
          hasLinkContent &&
          (remainingWords <= 8 ||
            (botMentioned && remainingWords <= 20) ||
            (remainingWords === 0 && (activity.attachments?.length || 0) > 0));

        if (hasLinkContent && !(await isAnyEventActive())) {
          const session = await addPendingShareUrls(conversationId, shareUrls);
          const confirmation =
            `Got your SharePoint link 🙂 (${session.pendingUrls.length} pending)\n\n` +
            `Send /start when you want me to load it.\n` +
            `Send /end later to clear everything.`;
          const sent = await send(
            finalizePromptResponse(
              confirmation,
              replyOptionsFromContext(context, { withFeedback: false })
            )
          );
          trackedMessages = createMessageRecords([activity]);
          logger.debug(
            `Pending SharePoint URL(s) saved from post (urls=${shareUrls.length}, pending=${session.pendingUrls.length}); replied id=${sent.id}`
          );
        } else if (looksLikeLinkShare) {
          const confirmation =
            "There's already an event running.\n" +
            "Send /end first if you want to paste a new link.";
          const sent = await send(
            finalizePromptResponse(
              confirmation,
              replyOptionsFromContext(context, { withFeedback: false })
            )
          );
          trackedMessages = createMessageRecords([activity]);
          logger.debug(`Blocked SharePoint URL while event active; replied id=${sent.id}`);
        } else {
          trackedMessages = await replyFromManager(send, context, activity, logger);
        }
      }
    } catch (error) {
      if (error instanceof GraphAuthRequiredError) {
        const authUser = graphUserId || error.userId;
        // Token may already exist under the other Teams id (DM vs group).
        if (authUser && (await userHasGraphAuth(authUser, authIds.alternates))) {
          await send(
            finalizePromptResponse(
              "You're already signed in. Please send your question again.",
              replyOptionsFromContext(context, { withFeedback: false })
            )
          );
        } else if (!context.isPersonalChat && (await isAnyEventActive())) {
          // Group Q&A should not require every member to SSO — only /start does.
          await send(
            finalizePromptResponse(
              "I couldn't refresh the Excel from SharePoint with your account. " +
                "Ask the person who started the event (or anyone already signed in) to send /start again, " +
                "or @mention me with your question once the file is loaded.",
              replyOptionsFromContext(context, { withFeedback: false })
            )
          );
        } else {
          await promptGraphSso(send, authUser, logger, {
            isPersonalChat: context.isPersonalChat,
          });
        }
        trackedMessages = createMessageRecords([activity]);
      } else {
      const detail = describeNetworkError(error);
      logger.error(`❌ Failed to handle/reply to message: ${detail}`);
      if (isOutboundBlocked(error)) {
        logger.error(
          "Outbound HTTPS to Microsoft is blocked (VPN/proxy/firewall). " +
            "The bot received the Teams message but cannot call login.botframework.com or smba.trafficmanager.net:443. " +
            "Connect the corporate VPN, or set HTTPS_PROXY in .env if you use a proxy."
        );
      }
      await sendServiceUnavailable(send, error, logger, context);
      trackedMessages = createMessageRecords([activity]);
      }
    }
  } else {
    trackedMessages = createMessageRecords([activity]);
  }

  logger.debug(trackedMessages);
  await context.memory.addMessages(trackedMessages);
  },
    authIds.alternates
  );
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
      await send("Looks like that Excel isn't available anymore — sorry about that.");
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
      await send("Sorry, I'm a bit stuck right now. Try again in a moment?");
    } catch {
      // ignore
    }
    return { status: 200 };
  }
});

app.on("file.consent.decline", async ({ send }) => {
  try {
    await send("No worries — I won't send the file.");
  } catch (error) {
    logger.warn(`file.consent.decline reply failed: ${describeNetworkError(error)}`);
  }
  return { status: 200 };
});

app.on("install.add", async ({ send }) => {
  try {
    if (isUploadIngestMode()) {
      await send(
        "Hey! Happy to help with the event once we're set up.\n\n" +
          "1) Upload an .xlsx (paperclip → Upload from this device)\n" +
          "2) Send /start\n" +
          "3) Just chat with me about seating, ferry, menu, drinks… (@mention in groups)\n" +
          "4) Send /end when you're done"
      );
    } else {
      await send(
        "Hey! Happy to help with the event once we're set up.\n\n" +
          "1) Paste a SharePoint Excel URL\n" +
          "2) Send /start\n" +
          "3) Just chat with me about the event\n" +
          "4) Send /end when you're done"
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
    if (getGraphAuthMode() === "sso") {
      await ensureGraphTokenSchema();
      logger.debug(
        "Graph auth mode: SSO (Teams Native SSO → OBO). Device login is disabled."
      );
      if (
        getEventIngestMode() === "sharepoint" ||
        Boolean(process.env.EVENTS_GRAPH_DRIVE_ID?.trim()) ||
        Boolean(process.env.GRAPH_WEBHOOK_URL?.trim())
      ) {
        try {
          const operatorId = getGraphOperatorUserId();
          const ready = await userHasGraphAuth(operatorId);
          if (!ready) {
            logger.warn(
              `Graph operator ${operatorId} has no stored SSO tokens yet. ` +
                "Have the operator message the bot in Teams once to complete SSO before delta/webhook Graph calls work."
            );
          } else {
            logger.debug(`Graph operator ${operatorId} has stored tokens for background jobs`);
          }
        } catch (error) {
          logger.warn(
            `GRAPH_OPERATOR_USER_ID not ready: ${error instanceof Error ? error.message : String(error)}`
          );
        }
      }
    }
    logger.debug("✅ Storage initialized successfully");
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    logger.error(`❌ Configuration error: ${message}`);
    process.exit(1);
  }

  await app.start(port);

  const deltaLogger = logger.child("delta");
  startDeltaPolling(deltaLogger);
  void ensureGraphChangeSubscription(deltaLogger).catch((error) => {
    const message = error instanceof Error ? error.message : String(error);
    deltaLogger.warn(`Graph change subscription setup failed: ${message}`);
  });
  logger.debug(`🚀 Collab Agent started on port ${port}`);
})();
