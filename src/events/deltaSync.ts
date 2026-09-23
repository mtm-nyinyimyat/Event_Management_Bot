import { existsSync, mkdirSync, readFileSync, writeFileSync } from "fs";
import path from "path";
import { ILogger } from "@microsoft/teams.common";
import {
  ensureWorkbookIndexed,
  getRagConfig,
  hashContent,
  listActiveSharepointSources,
  updateRagDocumentStatus,
  upsertSharepointSource,
  type RagSharepointSourceRecord,
} from "../rag";
import { clearQueryCaches } from "../utils/queryCache";
import {
  findActiveEventSession,
  getEventIngestMode,
} from "./eventSession";
import {
  clearConversationWorkbook,
  loadWorkbookFromBuffer,
  setConversationWorkbook,
  type SheetData,
} from "./excelStore";
import {
  createGraphSubscription,
  downloadExcelBinaryFromDriveItem,
  getDeltaChanges,
  renewGraphSubscription,
  resolveConfiguredDriveId,
  type GraphDeltaItem,
} from "./graphExcelClient";
import { resolveGraphNotificationUrl } from "./graphWebhook";

const DEFAULT_POLL_INTERVAL_MS = 15 * 60 * 1000;
const SUBSCRIPTION_STORE = path.join(process.cwd(), "data", "graph-webhook-subscription.json");
/** Renew when fewer than this many ms remain before expiry. */
const RENEW_WITHIN_MS = 6 * 60 * 60 * 1000;

let pollTimer: ReturnType<typeof setInterval> | null = null;
let pollInFlight = false;

function resolvePollIntervalMs(): number {
  const raw = Number(process.env.EVENTS_DELTA_POLL_INTERVAL_MS || DEFAULT_POLL_INTERVAL_MS);
  if (!Number.isFinite(raw) || raw < 60_000) {
    return DEFAULT_POLL_INTERVAL_MS;
  }
  return Math.floor(raw);
}

export async function executeDeltaQuery(source: RagSharepointSourceRecord) {
  if (!source.driveId || !source.itemId) {
    throw new Error(`SharePoint source ${source.documentId} is missing driveId/itemId`);
  }

  return getDeltaChanges({
    deltaLink: source.deltaLink,
    driveId: source.driveId,
    itemId: source.itemId,
    knownEtag: source.etag,
  });
}

/**
 * Apply delta/eTag changes: re-download + reindex when the tracked Excel changed.
 */
export async function processChanges(
  source: RagSharepointSourceRecord,
  changes: GraphDeltaItem[],
  meta: { deltaLink: string | null; logger?: ILogger }
): Promise<{ reindexed: boolean }> {
  const logger = meta.logger;
  const relevant = changes.filter((item) => itemAffectsTrackedFile(item, source));

  const deleted = relevant.find((item) => item.deleted);
  if (deleted) {
    logger?.warn(
      `SharePoint item deleted for document ${source.documentId}; marking document error`
    );
    await updateRagDocumentStatus(source.documentId, "error");
    await upsertSharepointSource({
      documentId: source.documentId,
      deltaLink: meta.deltaLink,
      lastDeltaSyncAt: new Date().toISOString(),
    });
    return { reindexed: false };
  }

  const latest = relevant[relevant.length - 1];
  const etagChanged = Boolean(latest?.eTag && latest.eTag !== source.etag);
  const shouldReindex =
    etagChanged ||
    // First etag-poll with no prior etag still returns the item as "changed"
    (Boolean(latest) && !source.etag);

  if (!shouldReindex || !source.driveId || !source.itemId) {
    await upsertSharepointSource({
      documentId: source.documentId,
      deltaLink: meta.deltaLink ?? source.deltaLink,
      etag: latest?.eTag || source.etag,
      lastModifiedAt: latest?.lastModifiedDateTime || source.lastModifiedAt,
      lastDeltaSyncAt: new Date().toISOString(),
    });
    return { reindexed: false };
  }

  logger?.debug(
    `♻️ Delta sync: reindexing document ${source.documentId} (etag ${source.etag || "none"} → ${latest.eTag})`
  );

  await updateRagDocumentStatus(source.documentId, "syncing");

  try {
    const { buffer, fileName } = await downloadExcelBinaryFromDriveItem(
      source.driveId,
      source.itemId,
      latest.name || undefined
    );
    const sheets: SheetData[] = loadWorkbookFromBuffer(buffer, fileName);
    const contentHash = hashContent(buffer);
    const sourceUri = source.webUrl || `graph://drives/${source.driveId}/items/${source.itemId}`;

    const active = await findActiveEventSession();
    if (active?.documentId === source.documentId) {
      clearConversationWorkbook(active.conversationId);
      setConversationWorkbook(active.conversationId, {
        source: sourceUri,
        sourceType: "graph",
        fileName,
        sheets,
        loadedAt: Date.now(),
      });
    }

    await ensureWorkbookIndexed(
      sheets,
      { source: sourceUri, loadedAt: Date.now(), documentId: source.documentId },
      getRagConfig()
    );

    await updateRagDocumentStatus(source.documentId, "ready", {
      contentHash,
      fileName,
    });

    await upsertSharepointSource({
      documentId: source.documentId,
      deltaLink: meta.deltaLink ?? source.deltaLink,
      etag: latest.eTag || source.etag,
      lastModifiedAt: latest.lastModifiedDateTime || source.lastModifiedAt,
      lastDeltaSyncAt: new Date().toISOString(),
    });

    clearQueryCaches();
    return { reindexed: true };
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    logger?.error(`Delta reindex failed for ${source.documentId}: ${message}`);
    await updateRagDocumentStatus(source.documentId, "error");
    await upsertSharepointSource({
      documentId: source.documentId,
      lastDeltaSyncAt: new Date().toISOString(),
    });
    return { reindexed: false };
  }
}

async function syncOneSource(
  source: RagSharepointSourceRecord,
  logger?: ILogger
): Promise<void> {
  const delta = await executeDeltaQuery(source);
  await processChanges(source, delta.items, {
    deltaLink: delta.deltaLink,
    logger,
  });
}

export async function runDeltaSyncTick(logger?: ILogger): Promise<void> {
  if (pollInFlight) {
    return;
  }
  pollInFlight = true;
  try {
    const sources = await listActiveSharepointSources();
    if (!sources.length) {
      return;
    }
    logger?.debug(`⏱️ Delta poll: checking ${sources.length} SharePoint document(s)`);
    for (const source of sources) {
      try {
        await syncOneSource(source, logger);
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        logger?.warn(`Delta poll failed for ${source.documentId}: ${message}`);
      }
    }
  } finally {
    pollInFlight = false;
  }
}

type StoredSubscription = {
  id: string;
  resource: string;
  driveId: string;
  /** Filter-only: EventBot folder id (ODB webhooks subscribe to drive root). */
  folderId: string;
  notificationUrl: string;
  expirationDateTime: string;
  clientState: string;
};

function readStoredSubscription(): StoredSubscription | null {
  try {
    if (!existsSync(SUBSCRIPTION_STORE)) {
      return null;
    }
    const raw = JSON.parse(readFileSync(SUBSCRIPTION_STORE, "utf8")) as StoredSubscription;
    if (!raw?.id || !raw.expirationDateTime) {
      return null;
    }
    return raw;
  } catch {
    return null;
  }
}

function writeStoredSubscription(sub: StoredSubscription): void {
  mkdirSync(path.dirname(SUBSCRIPTION_STORE), { recursive: true });
  writeFileSync(SUBSCRIPTION_STORE, JSON.stringify(sub, null, 2), "utf8");
}

function resolveWatchFolderId(): string {
  return (process.env.EVENTS_GRAPH_FOLDER_ID || "").trim();
}

/** True when a delta item is under the configured EventBot folder (or is that folder). */
export function itemIsUnderWatchFolder(
  item: GraphDeltaItem,
  folderId: string
): boolean {
  if (!folderId) {
    return true;
  }
  if (item.id === folderId) {
    return true;
  }
  if (item.parentReference?.id === folderId) {
    return true;
  }
  const path = item.parentReference?.path || "";
  // Graph paths look like: /drives/{driveId}/root:/EventBot or .../items/{folderId}:
  if (path.includes("/items/" + folderId)) {
    return true;
  }
  return false;
}

function itemAffectsTrackedFile(
  item: GraphDeltaItem,
  source: RagSharepointSourceRecord
): boolean {
  if (!source.itemId) {
    return false;
  }
  // Exact match on the tracked Excel — do not require parentReference (etag poll omits it).
  if (item.id === source.itemId) {
    return true;
  }
  return false;
}

async function persistSubscriptionOnSources(
  subscriptionId: string,
  expirationDateTime: string
): Promise<void> {
  const sources = await listActiveSharepointSources();
  for (const source of sources) {
    await upsertSharepointSource({
      documentId: source.documentId,
      subscriptionId,
      subscriptionExpiresAt: expirationDateTime,
    });
  }
}

/**
 * Create (or reuse) a Graph change-notification subscription on the drive root.
 * OneDrive for Business only allows root; folderId is stored for delta filtering only.
 * Requires public HTTPS GRAPH_WEBHOOK_URL (e.g. ngrok).
 */
export async function createDeltaSubscription(
  driveId: string,
  folderId: string,
  logger?: ILogger
): Promise<{ id: string; expirationDateTime?: string } | null> {
  const notificationUrl = resolveGraphNotificationUrl();
  if (!notificationUrl) {
    logger?.debug(
      "Graph subscription skipped: set GRAPH_WEBHOOK_URL to your public HTTPS webhook (ngrok)"
    );
    return null;
  }
  if (!driveId.trim()) {
    logger?.warn("Graph subscription skipped: missing driveId");
    return null;
  }

  // ODB: only /root is supported. Personal OneDrive can use subfolders, but root
  // works everywhere and we filter to EVENTS_GRAPH_FOLDER_ID in delta processing.
  const resource = "/drives/" + driveId.trim() + "/root";
  const existing = readStoredSubscription();
  if (
    existing &&
    existing.resource === resource &&
    existing.notificationUrl === notificationUrl &&
    Date.parse(existing.expirationDateTime) > Date.now() + RENEW_WITHIN_MS
  ) {
    logger?.debug(
      "Reusing Graph subscription " + existing.id + " until " + existing.expirationDateTime
    );
    await persistSubscriptionOnSources(existing.id, existing.expirationDateTime);
    return { id: existing.id, expirationDateTime: existing.expirationDateTime };
  }

  if (existing?.id && existing.resource === resource) {
    try {
      const renewed = await renewGraphSubscription(existing.id);
      if (renewed?.id && renewed.expirationDateTime) {
        const stored: StoredSubscription = {
          ...existing,
          id: renewed.id,
          expirationDateTime: renewed.expirationDateTime,
          notificationUrl,
          folderId: folderId.trim() || existing.folderId,
        };
        writeStoredSubscription(stored);
        await persistSubscriptionOnSources(renewed.id, renewed.expirationDateTime);
        logger?.debug(
          "Renewed Graph subscription " + renewed.id + " -> " + renewed.expirationDateTime
        );
        return renewed;
      }
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      logger?.warn(
        "Graph subscription renew failed (" + existing.id + "): " + message + "; creating new"
      );
    }
  }

  const created = await createGraphSubscription({
    notificationUrl,
    resource,
    changeType: "updated",
  });
  if (!created?.id) {
    return null;
  }

  const stored: StoredSubscription = {
    id: created.id,
    resource,
    driveId: driveId.trim(),
    folderId: folderId.trim(),
    notificationUrl,
    expirationDateTime:
      created.expirationDateTime ||
      new Date(Date.now() + 2 * 24 * 60 * 60 * 1000).toISOString(),
    clientState: process.env.GRAPH_WEBHOOK_CLIENT_STATE || "event-bot",
  };
  writeStoredSubscription(stored);
  await persistSubscriptionOnSources(stored.id, stored.expirationDateTime);
  logger?.debug(
    "Created Graph subscription " +
      stored.id +
      " on " +
      resource +
      " (filter folder=" +
      (stored.folderId || "none") +
      "; expires " +
      stored.expirationDateTime +
      ")"
  );
  return { id: stored.id, expirationDateTime: stored.expirationDateTime };
}

/** Renew drive-root subscription when within RENEW_WITHIN_MS of expiry. */
export async function renewExpiringSubscriptions(logger?: ILogger): Promise<void> {
  const stored = readStoredSubscription();
  if (!stored?.id) {
    return;
  }
  const expiresAt = Date.parse(stored.expirationDateTime);
  if (!Number.isFinite(expiresAt) || expiresAt > Date.now() + RENEW_WITHIN_MS) {
    return;
  }
  try {
    const renewed = await renewGraphSubscription(stored.id);
    if (!renewed?.id || !renewed.expirationDateTime) {
      return;
    }
    writeStoredSubscription({
      ...stored,
      id: renewed.id,
      expirationDateTime: renewed.expirationDateTime,
    });
    await persistSubscriptionOnSources(renewed.id, renewed.expirationDateTime);
    logger?.debug(
      "Renewed Graph subscription " + renewed.id + " -> " + renewed.expirationDateTime
    );
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    logger?.warn("Failed to renew Graph subscription " + stored.id + ": " + message);
  }
}

/**
 * Ensure a drive-root change subscription exists when GRAPH_WEBHOOK_URL is set.
 * EVENTS_GRAPH_FOLDER_ID is optional filter scope for delta (not the subscription resource).
 */
export async function ensureGraphChangeSubscription(logger?: ILogger): Promise<void> {
  if (!resolveGraphNotificationUrl()) {
    logger?.debug(
      "Graph change subscription idle: set GRAPH_WEBHOOK_URL (https://.../api/sharepoint/webhook)"
    );
    return;
  }

  const folderId = resolveWatchFolderId();
  if (!folderId) {
    logger?.debug(
      "Graph change subscription: EVENTS_GRAPH_FOLDER_ID unset — will subscribe to drive root with no folder filter"
    );
  }

  try {
    const driveId = await resolveConfiguredDriveId();
    await createDeltaSubscription(driveId, folderId, logger);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    logger?.warn("ensureGraphChangeSubscription failed: " + message);
  }
}

/**
 * Start periodic Delta / eTag polling for active SharePoint event documents.
 * Interval: EVENTS_DELTA_POLL_INTERVAL_MS (default 900000 = 15 minutes).
 * Webhook notifications also trigger the same sync path (faster).
 */
export function startDeltaPolling(logger?: ILogger): void {
  if (pollTimer) {
    return;
  }

  // Upload-only local testing: still allow poller; it no-ops when no SP sources.
  const intervalMs = resolvePollIntervalMs();
  logger?.debug(
    "⏱️ Delta polling enabled every " +
      intervalMs +
      "ms (EVENTS_DELTA_POLL_INTERVAL_MS; ingest=" +
      getEventIngestMode() +
      ")"
  );

  // Kick once shortly after startup, then on interval.
  void runDeltaSyncTick(logger).catch((error) => {
    const message = error instanceof Error ? error.message : String(error);
    logger?.warn("Initial delta poll failed: " + message);
  });

  pollTimer = setInterval(() => {
    void runDeltaSyncTick(logger).catch((error) => {
      const message = error instanceof Error ? error.message : String(error);
      logger?.warn("Delta poll tick failed: " + message);
    });
    void renewExpiringSubscriptions(logger);
  }, intervalMs);

  // Do not keep the process alive solely for the timer in some runtimes.
  if (typeof pollTimer === "object" && "unref" in pollTimer) {
    pollTimer.unref();
  }
}

export function stopDeltaPolling(): void {
  if (pollTimer) {
    clearInterval(pollTimer);
    pollTimer = null;
  }
}
