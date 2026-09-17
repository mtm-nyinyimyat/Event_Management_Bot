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
  downloadExcelBinaryFromDriveItem,
  getDeltaChanges,
  type GraphDeltaItem,
} from "./graphExcelClient";

const DEFAULT_POLL_INTERVAL_MS = 15 * 60 * 1000;

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

function itemAffectsTrackedFile(
  item: GraphDeltaItem,
  source: RagSharepointSourceRecord
): boolean {
  if (!source.itemId) {
    return false;
  }
  if (item.id === source.itemId) {
    return true;
  }
  // Parent-folder deltas may include the file under the same id only.
  return false;
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

/** Phase 2 stub — subscriptions need a public webhook URL. */
export async function createDeltaSubscription(
  _driveId: string,
  _itemId: string
): Promise<null> {
  return null;
}

/** Phase 2 stub — renew when Graph Change Notifications are enabled. */
export async function renewExpiringSubscriptions(_logger?: ILogger): Promise<void> {
  return;
}

/**
 * Start periodic Delta / eTag polling for active SharePoint event documents.
 * Interval: EVENTS_DELTA_POLL_INTERVAL_MS (default 900000 = 15 minutes).
 */
export function startDeltaPolling(logger?: ILogger): void {
  if (pollTimer) {
    return;
  }

  // Upload-only local testing: still allow poller; it no-ops when no SP sources.
  const intervalMs = resolvePollIntervalMs();
  logger?.debug(
    `⏱️ Delta polling enabled every ${intervalMs}ms (EVENTS_DELTA_POLL_INTERVAL_MS; ingest=${getEventIngestMode()})`
  );

  // Kick once shortly after startup, then on interval.
  void runDeltaSyncTick(logger).catch((error) => {
    const message = error instanceof Error ? error.message : String(error);
    logger?.warn(`Initial delta poll failed: ${message}`);
  });

  pollTimer = setInterval(() => {
    void runDeltaSyncTick(logger).catch((error) => {
      const message = error instanceof Error ? error.message : String(error);
      logger?.warn(`Delta poll tick failed: ${message}`);
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
