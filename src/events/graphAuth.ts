import { ClientSecretCredential } from "@azure/identity";
import {
  ConfidentialClientApplication,
  type AuthenticationResult,
  type Configuration,
} from "@azure/msal-node";
import {
  getGraphAuthAlternateUserIds,
  getGraphAuthUserId,
} from "./graphAuthContext";
import {
  clearStoredGraphTokens,
  getStoredGraphTokens,
  saveStoredGraphTokens,
} from "./graphTokenStore";

export type GraphAuthMode = "app" | "sso";

const APP_SCOPE = "https://graph.microsoft.com/.default";
const SKEW_MS = 120_000;

export class GraphAuthRequiredError extends Error {
  readonly userId: string;

  constructor(userId: string, message?: string) {
    super(message || `Graph SSO required for user ${userId}`);
    this.name = "GraphAuthRequiredError";
    this.userId = userId;
  }
}

interface TokenBundle {
  accessToken: string;
  refreshToken?: string;
  expiresAt: number;
  scope: string;
}

interface RawTokenResponse {
  access_token?: string;
  refresh_token?: string;
  expires_in?: number;
  scope?: string;
  error?: string;
  error_description?: string;
}

const inflightByUser = new Map<string, Promise<string>>();
let msalApp: ConfidentialClientApplication | null = null;

export function getGraphAuthMode(): GraphAuthMode {
  const raw = (process.env.GRAPH_AUTH_MODE || "app").trim().toLowerCase();
  // Legacy aliases from device-login era; all map to Entra SSO + OBO.
  if (raw === "sso" || raw === "delegated" || raw === "device" || raw === "user") {
    return "sso";
  }
  return "app";
}

function tenantId(): string {
  const value = process.env.TENANT_ID || process.env.GRAPH_TENANT_ID;
  if (!value?.trim()) {
    throw new Error("Graph SSO needs TENANT_ID.");
  }
  return value.trim();
}

function clientId(): string {
  const value = process.env.CLIENT_ID || process.env.GRAPH_CLIENT_ID;
  if (!value?.trim()) {
    throw new Error("Graph SSO needs CLIENT_ID.");
  }
  return value.trim();
}

function clientSecret(): string {
  const value = process.env.CLIENT_SECRET || process.env.GRAPH_CLIENT_SECRET;
  if (!value?.trim()) {
    throw new Error("Graph SSO/OBO needs CLIENT_SECRET (confidential client).");
  }
  return value.trim();
}

/**
 * Delegated Graph scopes for OBO / refresh.
 * Override with GRAPH_DELEGATED_SCOPES (space-separated).
 */
export function delegatedScopeString(): string {
  const raw = process.env.GRAPH_DELEGATED_SCOPES?.trim();
  if (raw) {
    return raw;
  }
  return [
    "https://graph.microsoft.com/Files.Read.All",
    "https://graph.microsoft.com/Sites.Read.All",
    "https://graph.microsoft.com/User.Read",
    "https://graph.microsoft.com/Files.Read",
    "offline_access",
  ].join(" ");
}

function delegatedScopesForMsal(): string[] {
  return delegatedScopeString()
    .split(/\s+/)
    .map((s) => s.trim())
    .filter((s) => s && s !== "offline_access");
}

function tokenEndpoint(): string {
  return `https://login.microsoftonline.com/${tenantId()}/oauth2/v2.0/token`;
}

export function getGraphOperatorUserId(): string {
  const value = process.env.GRAPH_OPERATOR_USER_ID?.trim();
  if (!value) {
    throw new Error(
      "Set GRAPH_OPERATOR_USER_ID to the operator Entra object id used for background Graph jobs."
    );
  }
  return value;
}

export function getGraphTokenExchangeResourceUri(): string {
  const botId = (process.env.BOT_ID || process.env.CLIENT_ID || clientId()).trim();
  const configured = process.env.GRAPH_SSO_RESOURCE_URI?.trim();
  return configured || `api://botid-${botId}`;
}

/** Public HTTPS origin for /auth/start and /auth/end (tunnel or deployed host). */
export function getPublicBaseUrl(): string {
  const explicit =
    process.env.PUBLIC_BASE_URL?.trim() ||
    process.env.BOT_ENDPOINT?.trim() ||
    process.env.WEBSITE_HOSTNAME?.trim();
  if (explicit) {
    return /^https?:\/\//i.test(explicit)
      ? explicit.replace(/\/+$/, "")
      : `https://${explicit.replace(/\/+$/, "")}`;
  }
  const webhook = process.env.GRAPH_WEBHOOK_URL?.trim();
  if (webhook) {
    try {
      return new URL(webhook).origin;
    } catch {
      // fall through
    }
  }
  throw new Error(
    "Set PUBLIC_BASE_URL (or BOT_ENDPOINT) to your public HTTPS origin for Entra auth redirects."
  );
}

export function getAuthRedirectUri(): string {
  return `${getPublicBaseUrl()}/auth/end`;
}

export function getAuthStartUrl(userId: string): string {
  const url = new URL(`${getPublicBaseUrl()}/auth/start`);
  url.searchParams.set("userId", userId);
  return url.toString();
}

function getMsalApp(): ConfidentialClientApplication {
  if (!msalApp) {
    const config: Configuration = {
      auth: {
        clientId: clientId(),
        authority: `https://login.microsoftonline.com/${tenantId()}`,
        clientSecret: clientSecret(),
      },
    };
    msalApp = new ConfidentialClientApplication(config);
  }
  return msalApp;
}

function bundleFromMsal(result: AuthenticationResult, fallbackScope: string): TokenBundle {
  if (!result.accessToken) {
    throw new Error("MSAL OBO response missing accessToken");
  }
  const expiresAt = result.expiresOn
    ? result.expiresOn.getTime() - SKEW_MS
    : Date.now() + 3600 * 1000 - SKEW_MS;
  return {
    accessToken: result.accessToken,
    expiresAt,
    scope: result.scopes?.join(" ") || fallbackScope,
  };
}

function bundleFromRaw(body: RawTokenResponse, fallbackScope: string): TokenBundle {
  if (!body.access_token) {
    throw new Error(body.error_description || body.error || "Token response missing access_token");
  }
  return {
    accessToken: body.access_token,
    refreshToken: body.refresh_token,
    expiresAt: Date.now() + Math.max(60, (body.expires_in || 3600) - 120) * 1000,
    scope: body.scope || fallbackScope,
  };
}

async function postForm(url: string, fields: Record<string, string>): Promise<RawTokenResponse> {
  const response = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams(fields),
  });
  return (await response.json()) as RawTokenResponse;
}

/**
 * OBO: raw token endpoint (captures refresh_token with offline_access).
 * Falls back to MSAL OBO if the raw call fails.
 */
export async function exchangeSsoTokenForGraph(ssoAssertion: string): Promise<TokenBundle> {
  const scope = delegatedScopeString();

  const raw = await postForm(tokenEndpoint(), {
    client_id: clientId(),
    client_secret: clientSecret(),
    grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
    assertion: ssoAssertion,
    requested_token_use: "on_behalf_of",
    scope,
  });
  if (raw.access_token && !raw.error) {
    return bundleFromRaw(raw, scope);
  }

  const msalResult = await getMsalApp().acquireTokenOnBehalfOf({
    oboAssertion: ssoAssertion,
    scopes: delegatedScopesForMsal(),
  });
  if (!msalResult?.accessToken) {
    throw new Error(
      raw.error_description ||
        raw.error ||
        "On-Behalf-Of failed to return an access token"
    );
  }
  return bundleFromMsal(msalResult, scope);
}

/**
 * Authorization-code grant (consent fallback). Returns Graph access + refresh directly.
 */
export async function exchangeAuthorizationCodeForGraph(
  code: string,
  redirectUri: string
): Promise<TokenBundle> {
  const scope = delegatedScopeString();
  const body = await postForm(tokenEndpoint(), {
    client_id: clientId(),
    client_secret: clientSecret(),
    grant_type: "authorization_code",
    code,
    redirect_uri: redirectUri,
    scope,
  });
  if (body.error || !body.access_token) {
    throw new Error(body.error_description || body.error || "Authorization code exchange failed");
  }
  return bundleFromRaw(body, scope);
}

/** Build Entra authorize URL for interactive consent. */
export function buildEntraAuthorizeUrl(state: string): string {
  const url = new URL(
    `https://login.microsoftonline.com/${tenantId()}/oauth2/v2.0/authorize`
  );
  url.searchParams.set("client_id", clientId());
  url.searchParams.set("response_type", "code");
  url.searchParams.set("redirect_uri", getAuthRedirectUri());
  url.searchParams.set("response_mode", "query");
  url.searchParams.set("scope", delegatedScopeString());
  url.searchParams.set("state", state);
  url.searchParams.set("prompt", "select_account");
  return url.toString();
}

export async function refreshGraphTokens(refreshToken: string): Promise<TokenBundle> {
  const scope = delegatedScopeString();

  try {
    const result = await getMsalApp().acquireTokenByRefreshToken({
      refreshToken,
      scopes: delegatedScopesForMsal(),
    });
    if (result?.accessToken) {
      const bundle = bundleFromMsal(result, scope);
      bundle.refreshToken = refreshToken;
      return bundle;
    }
  } catch {
    // fall through to raw refresh
  }

  const body = await postForm(tokenEndpoint(), {
    client_id: clientId(),
    client_secret: clientSecret(),
    grant_type: "refresh_token",
    refresh_token: refreshToken,
    scope,
  });
  if (body.error) {
    throw new Error(body.error_description || body.error);
  }
  const tokens = bundleFromRaw(body, scope);
  if (!tokens.refreshToken) {
    tokens.refreshToken = refreshToken;
  }
  return tokens;
}

export async function saveUserGraphTokens(
  userId: string,
  tokens: TokenBundle
): Promise<void> {
  await saveStoredGraphTokens({
    userId,
    accessToken: tokens.accessToken,
    refreshToken: tokens.refreshToken ?? null,
    expiresAt: tokens.expiresAt,
    scope: tokens.scope,
  });
}

export async function clearUserGraphTokens(userId: string): Promise<void> {
  await clearStoredGraphTokens(userId);
}

async function acquireAppToken(): Promise<string> {
  const secret = process.env.CLIENT_SECRET || process.env.GRAPH_CLIENT_SECRET;
  if (!secret?.trim()) {
    throw new Error("App-only Graph auth needs CLIENT_SECRET.");
  }
  const tid = process.env.TENANT_ID || process.env.GRAPH_TENANT_ID;
  const cid = process.env.CLIENT_ID || process.env.GRAPH_CLIENT_ID;
  if (!tid?.trim() || !cid?.trim()) {
    throw new Error("App-only Graph auth needs TENANT_ID and CLIENT_ID.");
  }
  const credential = new ClientSecretCredential(tid.trim(), cid.trim(), secret.trim());
  const token = await credential.getToken(APP_SCOPE);
  if (!token?.token) {
    throw new Error("Failed to acquire Microsoft Graph app token");
  }
  return token.token;
}

/** Load a usable access token for one exact store key (refreshing when needed). */
async function tryAcquireSsoTokenForExactId(userId: string): Promise<string | null> {
  const cached = await getStoredGraphTokens(userId);
  if (cached && cached.expiresAt > Date.now() && cached.accessToken) {
    return cached.accessToken;
  }
  if (cached?.refreshToken) {
    try {
      const refreshed = await refreshGraphTokens(cached.refreshToken);
      await saveUserGraphTokens(userId, refreshed);
      return refreshed.accessToken;
    } catch {
      await clearStoredGraphTokens(userId);
    }
  } else if (cached) {
    // Expired access with no refresh (e.g. offline_access never granted) — drop the row
    // so SSO checks don't treat the user as already signed in.
    await clearStoredGraphTokens(userId);
  }
  return null;
}

/**
 * Acquire a Graph token for this user, also checking alternate Teams ids.
 * When found under an alternate key, copies the row onto the canonical id so
 * later DM/group/channel lookups share one sign-in.
 */
async function acquireSsoTokenForUser(
  userId: string,
  alternateUserIds: string[] = []
): Promise<string> {
  const candidates = [
    ...new Set(
      [userId, ...alternateUserIds]
        .map((id) => id.trim())
        .filter(Boolean)
    ),
  ];

  for (const id of candidates) {
    const token = await tryAcquireSsoTokenForExactId(id);
    if (!token) {
      continue;
    }
    if (id !== userId) {
      const row = await getStoredGraphTokens(id);
      if (row) {
        await saveUserGraphTokens(userId, {
          accessToken: row.accessToken,
          refreshToken: row.refreshToken ?? undefined,
          expiresAt: row.expiresAt,
          scope: row.scope,
        });
      }
    }
    return token;
  }

  throw new GraphAuthRequiredError(
    userId,
    `No valid Graph token for user ${userId}. Complete Teams SSO first.`
  );
}

/**
 * Resolve a Graph access token.
 * - app mode: client credentials
 * - sso mode: per-user store (userId arg, AsyncLocalStorage, or operator)
 */
export async function getGraphAccessToken(userId?: string): Promise<string> {
  if (getGraphAuthMode() !== "sso") {
    return acquireAppToken();
  }

  const resolved =
    userId?.trim() ||
    getGraphAuthUserId()?.trim() ||
    process.env.GRAPH_OPERATOR_USER_ID?.trim();

  if (!resolved) {
    throw new GraphAuthRequiredError(
      "(unknown)",
      "SSO Graph calls need a userId (set Graph auth context or GRAPH_OPERATOR_USER_ID)."
    );
  }

  const alternates =
    userId?.trim() && userId.trim() !== getGraphAuthUserId()?.trim()
      ? []
      : getGraphAuthAlternateUserIds();

  const existing = inflightByUser.get(resolved);
  if (existing) {
    return existing;
  }
  const pending = acquireSsoTokenForUser(resolved, alternates).finally(() => {
    inflightByUser.delete(resolved);
  });
  inflightByUser.set(resolved, pending);
  return pending;
}

export async function getOperatorGraphAccessToken(): Promise<string> {
  return getGraphAccessToken(getGraphOperatorUserId());
}

/**
 * True when SSO mode is on and the user can obtain a Graph token now
 * (unexpired access, or a refresh token that still works).
 * Stale rows with expired access and no refresh are cleared.
 */
export async function userHasGraphAuth(
  userId: string,
  alternateUserIds: string[] = []
): Promise<boolean> {
  if (getGraphAuthMode() !== "sso") {
    return true;
  }
  try {
    // Bypass ALS so callers can pass explicit alternates (DM vs group id keys).
    await acquireSsoTokenForUser(userId.trim(), alternateUserIds);
    return true;
  } catch (error) {
    if (error instanceof GraphAuthRequiredError) {
      return false;
    }
    throw error;
  }
}

/**
 * Ensure the user can call Graph. Throws GraphAuthRequiredError when SSO sign-in is needed.
 */
export async function ensureUserGraphAuth(userId: string): Promise<string> {
  if (getGraphAuthMode() !== "sso") {
    return getGraphAccessToken();
  }
  if (!userId?.trim()) {
    throw new GraphAuthRequiredError("(unknown)", "Missing user id for Graph SSO");
  }
  return getGraphAccessToken(userId.trim());
}

/**
 * Exchange a Teams SSO assertion, persist tokens for the user, return access token.
 */
export async function completeSsoTokenExchange(
  userId: string,
  ssoAssertion: string
): Promise<string> {
  const tokens = await exchangeSsoTokenForGraph(ssoAssertion);
  await saveUserGraphTokens(userId, tokens);
  return tokens.accessToken;
}
