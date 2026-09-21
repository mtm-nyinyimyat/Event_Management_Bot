import { ClientSecretCredential } from "@azure/identity";
import fs from "node:fs";
import path from "node:path";

export type GraphAuthMode = "app" | "delegated";

const APP_SCOPE = "https://graph.microsoft.com/.default";
const CACHE_FILE = path.join(process.cwd(), ".graph-token-cache.json");

interface CachedDelegatedTokens {
  accessToken: string;
  refreshToken?: string;
  expiresAt: number;
  scope: string;
}

interface DeviceCodeResponse {
  device_code: string;
  user_code: string;
  verification_uri: string;
  expires_in: number;
  interval?: number;
  message?: string;
}

interface TokenResponse {
  access_token?: string;
  refresh_token?: string;
  expires_in?: number;
  scope?: string;
  error?: string;
  error_description?: string;
}

let memoryCache: CachedDelegatedTokens | null = null;
let inflight: Promise<string> | null = null;
let devicePrompt: ((message: string) => void) | null = null;

export function getGraphAuthMode(): GraphAuthMode {
  const raw = (process.env.GRAPH_AUTH_MODE || "app").trim().toLowerCase();
  return raw === "delegated" || raw === "device" || raw === "user" ? "delegated" : "app";
}

function tenantId(): string {
  const value = process.env.TENANT_ID || process.env.GRAPH_TENANT_ID;
  if (!value?.trim()) {
    throw new Error("Device login needs TENANT_ID (Entra tenant of the OneDrive account).");
  }
  return value.trim();
}

function clientId(): string {
  const value = process.env.CLIENT_ID || process.env.GRAPH_CLIENT_ID;
  if (!value?.trim()) {
    throw new Error("Device login needs CLIENT_ID.");
  }
  return value.trim();
}

function clientSecret(): string | undefined {
  const value = process.env.CLIENT_SECRET || process.env.GRAPH_CLIENT_SECRET;
  return value?.trim() || undefined;
}

/**
 * Delegated scopes for the signed-in user's OneDrive.
 * Files.Read covers personal /me files. Override with GRAPH_DELEGATED_SCOPES if needed.
 */
function delegatedScopeString(): string {
  const raw = process.env.GRAPH_DELEGATED_SCOPES?.trim();
  if (raw) {
    return raw;
  }
  return [
    "https://graph.microsoft.com/Files.Read",
    "https://graph.microsoft.com/User.Read",
    "offline_access",
  ].join(" ");
}

function tokenEndpoint(): string {
  return `https://login.microsoftonline.com/${tenantId()}/oauth2/v2.0/token`;
}

function loadCache(): CachedDelegatedTokens | null {
  if (memoryCache) {
    return memoryCache;
  }
  try {
    if (!fs.existsSync(CACHE_FILE)) {
      return null;
    }
    const parsed = JSON.parse(fs.readFileSync(CACHE_FILE, "utf8")) as CachedDelegatedTokens;
    if (parsed?.accessToken && parsed.expiresAt) {
      memoryCache = parsed;
      return parsed;
    }
  } catch {
    // ignore corrupt cache
  }
  return null;
}

function saveCache(tokens: CachedDelegatedTokens): void {
  memoryCache = tokens;
  try {
    fs.writeFileSync(CACHE_FILE, JSON.stringify(tokens), { mode: 0o600 });
  } catch {
    // non-fatal
  }
}

function tokensFromResponse(body: TokenResponse, fallbackScope: string): CachedDelegatedTokens {
  if (!body.access_token) {
    throw new Error(body.error_description || body.error || "Graph token response missing access_token");
  }
  return {
    accessToken: body.access_token,
    refreshToken: body.refresh_token,
    expiresAt: Date.now() + Math.max(60, (body.expires_in || 3600) - 120) * 1000,
    scope: body.scope || fallbackScope,
  };
}

async function postForm(url: string, fields: Record<string, string>): Promise<TokenResponse> {
  const response = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams(fields),
  });
  return (await response.json()) as TokenResponse;
}

async function refreshDelegated(refreshToken: string, scope: string): Promise<CachedDelegatedTokens> {
  const fields: Record<string, string> = {
    client_id: clientId(),
    grant_type: "refresh_token",
    refresh_token: refreshToken,
    scope,
  };
  const secret = clientSecret();
  if (secret) {
    fields.client_secret = secret;
  }
  let body = await postForm(tokenEndpoint(), fields);
  if (body.error === "invalid_client" || /public client/i.test(body.error_description || "")) {
    delete fields.client_secret;
    body = await postForm(tokenEndpoint(), fields);
  }
  if (body.error) {
    throw new Error(body.error_description || body.error);
  }
  const tokens = tokensFromResponse(body, scope);
  if (!tokens.refreshToken) {
    tokens.refreshToken = refreshToken;
  }
  saveCache(tokens);
  return tokens;
}

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function deviceCodeLogin(onPrompt: (message: string) => void): Promise<CachedDelegatedTokens> {
  const scope = delegatedScopeString();
  const start = await fetch(
    `https://login.microsoftonline.com/${tenantId()}/oauth2/v2.0/devicecode`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        client_id: clientId(),
        scope,
      }),
    }
  );
  const device = (await start.json()) as DeviceCodeResponse & TokenResponse;
  if (!start.ok || !device.device_code || device.error) {
    throw new Error(
      device.error_description ||
        device.error ||
        "Failed to start device login. In Entra, enable Authentication → Allow public client flows."
    );
  }

  const message =
    device.message ||
    `To sign in to personal OneDrive, open ${device.verification_uri} and enter code ${device.user_code}`;
  onPrompt(message);

  const deadline = Date.now() + (device.expires_in || 900) * 1000;
  let waitMs = Math.max(5, device.interval || 5) * 1000;
  const secret = clientSecret();

  while (Date.now() < deadline) {
    await sleep(waitMs);
    const fields: Record<string, string> = {
      grant_type: "urn:ietf:params:oauth:grant-type:device_code",
      client_id: clientId(),
      device_code: device.device_code,
    };
    if (secret) {
      fields.client_secret = secret;
    }
    let body = await postForm(tokenEndpoint(), fields);
    if (
      secret &&
      (body.error === "invalid_client" || /public client|AADSTS700025/i.test(body.error_description || ""))
    ) {
      delete fields.client_secret;
      body = await postForm(tokenEndpoint(), fields);
    }

    if (body.access_token) {
      const tokens = tokensFromResponse(body, scope);
      saveCache(tokens);
      return tokens;
    }

    const error = body.error || "";
    if (error === "authorization_pending") {
      continue;
    }
    if (error === "slow_down") {
      waitMs += 5000;
      continue;
    }
    throw new Error(body.error_description || error || "Device login failed");
  }

  throw new Error("Device login timed out before the code was entered.");
}

async function acquireAppToken(): Promise<string> {
  const secret = clientSecret();
  if (!secret) {
    throw new Error("App-only Graph auth needs CLIENT_SECRET. Set GRAPH_AUTH_MODE=delegated for device login.");
  }
  const credential = new ClientSecretCredential(tenantId(), clientId(), secret);
  const token = await credential.getToken(APP_SCOPE);
  if (!token?.token) {
    throw new Error("Failed to acquire Microsoft Graph app token");
  }
  return token.token;
}

async function acquireDelegatedToken(): Promise<string> {
  const scope = delegatedScopeString();
  const cached = loadCache();
  if (cached && cached.expiresAt > Date.now() && (!cached.scope || cached.scope === scope)) {
    return cached.accessToken;
  }
  if (cached?.refreshToken) {
    try {
      const refreshed = await refreshDelegated(cached.refreshToken, scope);
      return refreshed.accessToken;
    } catch {
      // fall through to a new device login
    }
  }
  const tokens = await deviceCodeLogin(
    devicePrompt ||
      ((message) => {
        console.log(`\n${message}\n`);
      })
  );
  return tokens.accessToken;
}

export async function getGraphAccessToken(): Promise<string> {
  if (getGraphAuthMode() !== "delegated") {
    return acquireAppToken();
  }
  if (!inflight) {
    inflight = acquireDelegatedToken().finally(() => {
      inflight = null;
    });
  }
  return inflight;
}

/** Prompt for device login at startup when delegated mode is on. */
export async function ensureGraphAuthReady(
  log?: (message: string) => void
): Promise<void> {
  if (getGraphAuthMode() !== "delegated") {
    return;
  }
  devicePrompt = (message) => {
    console.log(`\n${message}\n`);
    log?.(message);
  };
  try {
    await getGraphAccessToken();
    log?.("Personal OneDrive device login ready");
  } finally {
    devicePrompt = null;
  }
}
