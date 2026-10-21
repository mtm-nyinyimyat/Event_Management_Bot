import { randomBytes } from "node:crypto";
import { App, ExpressAdapter } from "@microsoft/teams.apps";
import { ILogger } from "@microsoft/teams.common";
import {
  buildEntraAuthorizeUrl,
  exchangeAuthorizationCodeForGraph,
  getAuthRedirectUri,
  saveUserGraphTokens,
} from "./graphAuth";

const AUTH_START_PATH = "/auth/start";
const AUTH_END_PATH = "/auth/end";
const STATE_TTL_MS = 15 * 60 * 1000;

type ExpressReq = {
  query: Record<string, string | string[] | undefined>;
  protocol?: string;
  get?: (name: string) => string | undefined;
  headers?: Record<string, string | string[] | undefined>;
};

type ExpressRes = {
  status: (code: number) => ExpressRes;
  type: (t: string) => ExpressRes;
  send: (body?: unknown) => void;
  redirect: (url: string) => void;
  headersSent: boolean;
};

interface PendingAuthState {
  userId: string;
  expiresAt: number;
}

const pendingStates = new Map<string, PendingAuthState>();

function queryParam(
  query: Record<string, string | string[] | undefined>,
  key: string
): string {
  const raw = query[key];
  if (Array.isArray(raw)) {
    return String(raw[0] || "").trim();
  }
  return String(raw || "").trim();
}

function pruneStates(): void {
  const now = Date.now();
  for (const [key, value] of pendingStates) {
    if (value.expiresAt <= now) {
      pendingStates.delete(key);
    }
  }
}

function createState(userId: string): string {
  pruneStates();
  const nonce = randomBytes(16).toString("hex");
  pendingStates.set(nonce, {
    userId,
    expiresAt: Date.now() + STATE_TTL_MS,
  });
  return nonce;
}

function consumeState(state: string): string | null {
  pruneStates();
  const pending = pendingStates.get(state);
  pendingStates.delete(state);
  if (!pending || pending.expiresAt <= Date.now()) {
    return null;
  }
  return pending.userId;
}

function htmlPage(title: string, body: string): string {
  return `<!DOCTYPE html>
<html><head><meta charset="utf-8"/><title>${title}</title>
<style>
  body{font-family:Segoe UI,system-ui,sans-serif;max-width:32rem;margin:3rem auto;padding:0 1rem;color:#242424}
  h1{font-size:1.25rem} p{line-height:1.5}
</style></head><body>
<h1>${title}</h1>
<p>${body}</p>
<script>try{window.close();}catch(e){}</script>
</body></html>`;
}

function handleAuthStart(req: ExpressReq, res: ExpressRes, logger?: ILogger): void {
  const userId = queryParam(req.query, "userId");
  if (!userId) {
    res.status(400).type("html").send(
      htmlPage("Sign-in error", "Missing userId. Open the sign-in link from the Teams bot again.")
    );
    return;
  }
  try {
    const state = createState(userId);
    const authorizeUrl = buildEntraAuthorizeUrl(state);
    logger?.debug(`Entra auth start for user ${userId}`);
    res.redirect(authorizeUrl);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    logger?.error(`auth/start failed: ${message}`);
    res.status(500).type("html").send(htmlPage("Sign-in error", message));
  }
}

async function handleAuthEnd(
  req: ExpressReq,
  res: ExpressRes,
  logger?: ILogger
): Promise<void> {
  const error = queryParam(req.query, "error");
  const errorDescription = queryParam(req.query, "error_description");
  if (error) {
    logger?.warn(`auth/end Entra error: ${error} ${errorDescription}`);
    res.status(400).type("html").send(
      htmlPage(
        "Sign-in cancelled",
        errorDescription || error || "Authorization was denied."
      )
    );
    return;
  }

  const code = queryParam(req.query, "code");
  const state = queryParam(req.query, "state");
  if (!code || !state) {
    res.status(400).type("html").send(
      htmlPage("Sign-in error", "Missing authorization code or state.")
    );
    return;
  }

  const userId = consumeState(state);
  if (!userId) {
    res.status(400).type("html").send(
      htmlPage(
        "Sign-in expired",
        "This sign-in link expired. Ask the bot for a new sign-in card and try again."
      )
    );
    return;
  }

  try {
    const redirectUri = getAuthRedirectUri();
    const tokens = await exchangeAuthorizationCodeForGraph(code, redirectUri);
    await saveUserGraphTokens(userId, tokens);
    logger?.debug(`Entra auth-code tokens saved for user ${userId}`);
    res.status(200).type("html").send(
      htmlPage(
        "Signed in",
        "Microsoft 365 access is connected. You can close this window and return to Teams."
      )
    );
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    logger?.error(`auth/end failed for ${userId}: ${message}`);
    res.status(500).type("html").send(htmlPage("Sign-in error", message));
  }
}

/**
 * Mount GET /auth/start and GET /auth/end for Entra auth-code consent (no Azure Bot OAuth).
 */
export function registerEntraAuthRoutes(app: App, logger?: ILogger): void {
  const adapter = app.server.adapter;
  if (!(adapter instanceof ExpressAdapter)) {
    logger?.warn(
      "Entra auth routes not registered: HttpServer adapter is not ExpressAdapter"
    );
    return;
  }

  adapter.get(AUTH_START_PATH, ((req: ExpressReq, res: ExpressRes) => {
    try {
      handleAuthStart(req, res, logger);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      logger?.error("auth/start handler error: " + message);
      if (!res.headersSent) {
        res.status(500).type("html").send(htmlPage("Sign-in error", message));
      }
    }
  }) as never);

  adapter.get(AUTH_END_PATH, ((req: ExpressReq, res: ExpressRes) => {
    void handleAuthEnd(req, res, logger).catch((error) => {
      const message = error instanceof Error ? error.message : String(error);
      logger?.error("auth/end handler error: " + message);
      if (!res.headersSent) {
        res.status(500).type("html").send(htmlPage("Sign-in error", message));
      }
    });
  }) as never);

  logger?.debug(`Entra auth routes listening at ${AUTH_START_PATH} and ${AUTH_END_PATH}`);
}

export function getEntraAuthStartPath(): string {
  return AUTH_START_PATH;
}

export function getEntraAuthEndPath(): string {
  return AUTH_END_PATH;
}
