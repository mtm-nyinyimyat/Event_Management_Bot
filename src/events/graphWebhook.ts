import { App, ExpressAdapter } from "@microsoft/teams.apps";
import { ILogger } from "@microsoft/teams.common";

const WEBHOOK_PATH = "/api/sharepoint/webhook";

type ExpressReq = {
  query: Record<string, string | string[] | undefined>;
  body?: unknown;
};

type ExpressRes = {
  status: (code: number) => ExpressRes;
  type: (t: string) => ExpressRes;
  send: (body?: unknown) => void;
  sendStatus: (code: number) => void;
  headersSent: boolean;
};

let onChangeNotify: ((logger?: ILogger) => void) | null = null;
let syncTimer: ReturnType<typeof setTimeout> | null = null;
let lastNotifyAt = 0;

/** Register the work to run after a change notification (typically delta sync). */
export function setGraphWebhookChangeHandler(
  handler: (logger?: ILogger) => void
): void {
  onChangeNotify = handler;
}

function expectedClientState(): string {
  return (process.env.GRAPH_WEBHOOK_CLIENT_STATE || "event-bot").trim();
}

function scheduleChangeHandler(logger?: ILogger): void {
  // Graph expects a fast 2xx; do the real work after responding.
  const debounceMs = Number(process.env.GRAPH_WEBHOOK_DEBOUNCE_MS || 2_000);
  if (syncTimer) {
    clearTimeout(syncTimer);
  }
  syncTimer = setTimeout(() => {
    syncTimer = null;
    try {
      onChangeNotify?.(logger);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      logger?.warn("Webhook change handler failed: " + message);
    }
  }, Number.isFinite(debounceMs) && debounceMs >= 0 ? debounceMs : 2_000);
}

function validationTokenFromRequest(req: ExpressReq): string | undefined {
  const raw = req.query.validationToken ?? req.query.validationtoken;
  if (Array.isArray(raw)) {
    const first = raw[0];
    return typeof first === "string" ? first : undefined;
  }
  return typeof raw === "string" ? raw : undefined;
}

type GraphNotification = {
  subscriptionId?: string;
  clientState?: string;
  changeType?: string;
  resource?: string;
};

type GraphNotificationPayload = {
  value?: GraphNotification[];
};

function handleWebhookRequest(req: ExpressReq, res: ExpressRes, logger?: ILogger): void {
  const validationToken = validationTokenFromRequest(req);
  if (validationToken) {
    logger?.debug("Graph webhook validation request received");
    res.status(200).type("text/plain").send(validationToken);
    return;
  }

  const expected = expectedClientState();
  const body = (req.body || {}) as GraphNotificationPayload;
  const notifications = Array.isArray(body.value) ? body.value : [];

  if (!notifications.length) {
    logger?.debug("Graph webhook POST with empty body (ack 202)");
    res.sendStatus(202);
    return;
  }

  const accepted = notifications.filter((n) => {
    if (!n.clientState) {
      return true;
    }
    return n.clientState === expected;
  });

  if (!accepted.length) {
    logger?.warn("Graph webhook notification rejected: clientState mismatch");
    res.sendStatus(202);
    return;
  }

  lastNotifyAt = Date.now();
  logger?.debug(
    "Graph change notification: " +
      accepted.length +
      " item(s) — " +
      accepted.map((n) => (n.changeType || "?") + " " + (n.resource || "")).join("; ")
  );

  res.sendStatus(202);
  scheduleChangeHandler(logger);
}

function parseJsonBody(
  req: ExpressReq & { on?: (event: string, cb: (...args: unknown[]) => void) => void },
  _res: ExpressRes,
  next: (err?: unknown) => void
): void {
  if (req.body !== undefined && req.body !== null) {
    next();
    return;
  }
  const chunks: Buffer[] = [];
  req.on?.("data", (chunk: unknown) => {
    chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(String(chunk)));
  });
  req.on?.("end", () => {
    try {
      const raw = Buffer.concat(chunks).toString("utf8").trim();
      (req as ExpressReq).body = raw ? JSON.parse(raw) : {};
      next();
    } catch (err) {
      next(err);
    }
  });
  req.on?.("error", (err: unknown) => next(err));
}

/**
 * Mount POST/GET /api/sharepoint/webhook on the Teams app Express adapter.
 * GET is included because some Graph validation probes use it.
 */
export function registerGraphWebhookRoutes(app: App, logger?: ILogger): void {
  const adapter = app.server.adapter;
  if (!(adapter instanceof ExpressAdapter)) {
    logger?.warn(
      "Graph webhook not registered: HttpServer adapter is not ExpressAdapter"
    );
    return;
  }

  const handler = (req: ExpressReq, res: ExpressRes) => {
    try {
      handleWebhookRequest(req, res, logger);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      logger?.error("Graph webhook handler error: " + message);
      if (!res.headersSent) {
        res.sendStatus(500);
      }
    }
  };

  adapter.post(WEBHOOK_PATH, parseJsonBody as never, handler as never);
  adapter.get(WEBHOOK_PATH, handler as never);
  logger?.debug("Graph webhook listening at " + WEBHOOK_PATH);
}

export function getGraphWebhookPath(): string {
  return WEBHOOK_PATH;
}

export function getLastWebhookNotifyAt(): number {
  return lastNotifyAt;
}

/** Public HTTPS URL Graph will call (ngrok / deployed host). */
export function resolveGraphNotificationUrl(): string | null {
  const explicit = process.env.GRAPH_WEBHOOK_URL?.trim();
  if (explicit) {
    return explicit.replace(/\/+$/, "");
  }
  const base = process.env.BOT_ENDPOINT?.trim() || process.env.WEBSITE_HOSTNAME?.trim();
  if (!base) {
    return null;
  }
  const origin = /^https?:\/\//i.test(base)
    ? base.replace(/\/+$/, "")
    : "https://" + base.replace(/\/+$/, "");
  return origin + WEBHOOK_PATH;
}
