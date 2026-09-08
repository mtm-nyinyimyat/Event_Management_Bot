import { ClientSecretCredential } from "@azure/identity";

export interface GraphDriveItem {
  id: string;
  name: string;
  webUrl?: string;
  lastModifiedDateTime?: string;
  size?: number;
  file?: { mimeType?: string };
  folder?: Record<string, unknown>;
}

export interface AllowedExcelTarget {
  path?: string;
  itemId?: string;
  /** Full Teams/SharePoint open link or Doc.aspx URL (resolved via Graph shares API). */
  shareUrl?: string;
  /** Optional expected file name extracted from a browser link. */
  expectedFileName?: string;
  label?: string;
}

export interface GraphSheetData {
  sheet: string;
  rows: Record<string, string>[];
}

export interface GraphWorkbookData {
  source: string;
  fileName: string;
  webUrl?: string;
  allowedAs: string;
  sheets: GraphSheetData[];
}

function requireEnv(name: string): string {
  const value = process.env[name]?.trim();
  if (!value) {
    throw new Error(`Missing required environment variable: ${name}`);
  }
  return value;
}

function getCredential(): ClientSecretCredential {
  const tenantId = process.env.TENANT_ID || process.env.GRAPH_TENANT_ID;
  const clientId = process.env.CLIENT_ID || process.env.GRAPH_CLIENT_ID;
  const clientSecret = process.env.CLIENT_SECRET || process.env.GRAPH_CLIENT_SECRET;

  if (!tenantId || !clientId || !clientSecret) {
    throw new Error(
      "Graph Excel access needs TENANT_ID, CLIENT_ID, and CLIENT_SECRET (bot app registration)."
    );
  }

  return new ClientSecretCredential(tenantId, clientId, clientSecret);
}

async function getGraphToken(): Promise<string> {
  const credential = getCredential();
  const token = await credential.getToken("https://graph.microsoft.com/.default");
  if (!token?.token) {
    throw new Error("Failed to acquire Microsoft Graph access token");
  }
  return token.token;
}

async function graphFetch<T>(url: string, init?: RequestInit): Promise<T> {
  const accessToken = await getGraphToken();
  const response = await fetch(url, {
    ...init,
    headers: {
      Authorization: `Bearer ${accessToken}`,
      Accept: "application/json",
      ...(init?.headers || {}),
    },
    redirect: "follow",
  });

  if (!response.ok) {
    const body = await response.text();
    throw new Error(`Graph request failed (${response.status}) ${url}: ${body}`);
  }

  if (response.status === 204) {
    return undefined as T;
  }

  const contentType = response.headers.get("content-type") || "";
  if (contentType.includes("application/json")) {
    return (await response.json()) as T;
  }

  return (await response.arrayBuffer()) as T;
}

function encodeDrivePath(filePath: string): string {
  const normalized = filePath.replace(/^\/+/, "").replace(/\\/g, "/");
  return normalized
    .split("/")
    .filter(Boolean)
    .map((segment) => encodeURIComponent(segment))
    .join("/");
}

function normalizePathKey(value: string): string {
  return value.replace(/^\/+/, "").replace(/\\/g, "/").trim().toLocaleLowerCase("my");
}

function cellToString(value: unknown): string {
  if (value === null || value === undefined) {
    return "";
  }
  return String(value).replace(/\r\n/g, "\n").trim();
}

/** Turn Teams/SharePoint open links into a Graph site path (/sites/Name). */
export function normalizeGraphSiteUrl(siteUrl: string): { hostname: string; pathname: string } {
  const parsed = new URL(siteUrl);
  let pathname = decodeURIComponent(parsed.pathname).replace(/\/$/, "");

  const openLinkMatch = pathname.match(/\/:[a-z]:\/[rw](\/(?:sites|teams|personal)\/[^/]+)/i);
  if (openLinkMatch) {
    pathname = openLinkMatch[1];
  }

  if (!/^\/(sites|teams|personal)\//i.test(pathname)) {
    throw new Error(
      `EVENTS_GRAPH_SITE_URL must point at a SharePoint site (e.g. https://contoso.sharepoint.com/sites/TeamName). Got path: ${pathname || "/"}`
    );
  }

  return { hostname: parsed.hostname, pathname };
}

function encodeSharingUrl(url: string): string {
  const base64 = Buffer.from(url, "utf8").toString("base64");
  return `u!${base64.replace(/=+$/g, "").replace(/\//g, "_").replace(/\+/g, "-")}`;
}

function extractFileNameFromUrl(urlOrPath: string): string | undefined {
  try {
    const parsed = new URL(urlOrPath, "https://sharepoint.local");
    const fileParam = parsed.searchParams.get("file");
    if (fileParam) {
      return decodeURIComponent(fileParam);
    }
  } catch {
    // ignore
  }
  const bare = urlOrPath.split("?")[0].split("/").filter(Boolean).pop();
  if (bare && /\.xlsx?$/i.test(bare)) {
    return decodeURIComponent(bare);
  }
  return undefined;
}

function looksLikeShareOrDocLink(value: string): boolean {
  const lower = value.toLowerCase();
  return (
    /^https?:\/\//i.test(value) ||
    lower.includes("_layouts/15/doc.aspx") ||
    lower.includes("sourcedoc=") ||
    lower.includes("/:x:/") ||
    lower.includes("/:f:/")
  );
}

function toAbsoluteShareUrl(entry: string): string {
  if (/^https?:\/\//i.test(entry)) {
    return entry;
  }

  const siteUrl = process.env.EVENTS_GRAPH_SITE_URL?.trim();
  if (!siteUrl) {
    throw new Error(
      `Allow-list entry looks like a Doc.aspx/share link but is not absolute. Set EVENTS_GRAPH_SITE_URL or paste the full file URL: ${entry}`
    );
  }

  const { hostname, pathname } = normalizeGraphSiteUrl(siteUrl);
  const relative = entry.replace(/^\/+/, "");
  return `https://${hostname}${pathname}/${relative}`;
}

function parseAllowListEntry(entry: string): AllowedExcelTarget {
  const value = entry.trim();
  if (!value) {
    throw new Error("Empty allow-list entry");
  }

  if (looksLikeShareOrDocLink(value)) {
    const shareUrl = toAbsoluteShareUrl(value);
    const expectedFileName = extractFileNameFromUrl(shareUrl) || extractFileNameFromUrl(value);
    return {
      shareUrl,
      expectedFileName,
      label: expectedFileName || shareUrl,
    };
  }

  return {
    path: value,
    expectedFileName: extractFileNameFromUrl(value),
    label: value,
  };
}

/**
 * Explicit allow-list only. The bot never browses a whole library.
 */
export function getAllowedExcelTargets(): AllowedExcelTarget[] {
  const targets: AllowedExcelTarget[] = [];

  const allowedFiles = process.env.EVENTS_GRAPH_ALLOWED_FILES?.trim();
  if (allowedFiles) {
    for (const entry of allowedFiles.split(",")) {
      const trimmed = entry.trim();
      if (trimmed) {
        targets.push(parseAllowListEntry(trimmed));
      }
    }
  }

  const singlePath = process.env.EVENTS_GRAPH_FILE_PATH?.trim();
  if (singlePath) {
    targets.push(parseAllowListEntry(singlePath));
  }

  const allowedItemIds = process.env.EVENTS_GRAPH_ALLOWED_ITEM_IDS?.trim();
  if (allowedItemIds) {
    for (const entry of allowedItemIds.split(",")) {
      const itemId = entry.trim();
      if (itemId) {
        targets.push({ itemId, label: itemId });
      }
    }
  }

  const singleItemId = process.env.EVENTS_GRAPH_ITEM_ID?.trim();
  if (singleItemId) {
    targets.push({ itemId: singleItemId, label: singleItemId });
  }

  const seen = new Set<string>();
  return targets.filter((target) => {
    const key = target.itemId
      ? `id:${target.itemId}`
      : target.shareUrl
        ? `share:${normalizePathKey(target.shareUrl)}`
        : `path:${normalizePathKey(target.path || "")}`;
    if (!key || key.endsWith(":") || seen.has(key)) {
      return false;
    }
    seen.add(key);
    return true;
  });
}

function assertTargetAllowed(item: GraphDriveItem, requested: AllowedExcelTarget): void {
  if (requested.itemId) {
    if (item.id !== requested.itemId) {
      throw new Error(`Refusing to read Graph item ${item.id}; it is not in the allow-list.`);
    }
    return;
  }

  const expectedName =
    requested.expectedFileName ||
    (requested.path ? requested.path.split("/").filter(Boolean).pop() : undefined);

  if (expectedName && normalizePathKey(item.name) !== normalizePathKey(expectedName)) {
    throw new Error(
      `Refusing to read "${item.name}"; only allow-listed Excel files can be accessed.`
    );
  }

  if (requested.path && !requested.shareUrl && !expectedName) {
    const allowedName = requested.path.split("/").filter(Boolean).pop() || requested.path;
    if (normalizePathKey(item.name) !== normalizePathKey(allowedName)) {
      throw new Error(
        `Refusing to read "${item.name}"; only allow-listed Excel files can be accessed.`
      );
    }
  }
}

async function resolveSiteId(): Promise<string> {
  if (process.env.EVENTS_GRAPH_SITE_ID?.trim()) {
    return process.env.EVENTS_GRAPH_SITE_ID.trim();
  }

  const siteUrl = process.env.EVENTS_GRAPH_SITE_URL?.trim();
  if (!siteUrl) {
    throw new Error(
      "Set EVENTS_GRAPH_SITE_ID or EVENTS_GRAPH_SITE_URL to the Team/SharePoint site that stores the Excel file."
    );
  }

  const { hostname, pathname } = normalizeGraphSiteUrl(siteUrl);
  const lookupUrl = `https://graph.microsoft.com/v1.0/sites/${hostname}:${pathname}`;

  try {
    const site = await graphFetch<{ id: string }>(lookupUrl);
    return site.id;
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    if (message.includes("(401)") || message.includes("(403)")) {
      throw new Error(
        `${message}\n\nSites.Selected cannot resolve this site until an admin grants the app access.\n` +
          `Do this once (as SharePoint/Global admin in Graph Explorer with Sites.FullControl.All):\n` +
          `1) GET ${lookupUrl}\n` +
          `2) Copy the returned "id" into EVENTS_GRAPH_SITE_ID in .env\n` +
          `3) POST https://graph.microsoft.com/v1.0/sites/{that-id}/permissions with body:\n` +
          `   {"roles":["read"],"grantedToIdentities":[{"application":{"id":"${process.env.CLIENT_ID || "YOUR_CLIENT_ID"}","displayName":"event_management_bot"}}]}\n` +
          `4) Restart the bot.\n` +
          `Faster alternative: temporarily add Application permission Sites.Read.All (admin consent), then switch back to Sites.Selected + site grant.`
      );
    }
    throw error;
  }
}

async function resolveDriveId(siteId: string): Promise<string> {
  if (process.env.EVENTS_GRAPH_DRIVE_ID?.trim()) {
    return process.env.EVENTS_GRAPH_DRIVE_ID.trim();
  }

  const drive = await graphFetch<{ id: string }>(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/drive`
  );
  return drive.id;
}

async function getItemByPath(driveId: string, itemPath: string): Promise<GraphDriveItem> {
  const encoded = encodeDrivePath(itemPath);
  return graphFetch<GraphDriveItem>(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${encoded}?$select=id,name,webUrl,lastModifiedDateTime,size,file,folder`
  );
}

async function getItemById(driveId: string, itemId: string): Promise<GraphDriveItem> {
  return graphFetch<GraphDriveItem>(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}?$select=id,name,webUrl,lastModifiedDateTime,size,file,folder`
  );
}

interface SharedDriveItem extends GraphDriveItem {
  parentReference?: { driveId?: string };
}

async function getItemByShareUrl(shareUrl: string): Promise<SharedDriveItem> {
  const encoded = encodeSharingUrl(shareUrl);
  return graphFetch<SharedDriveItem>(
    `https://graph.microsoft.com/v1.0/shares/${encoded}/driveItem?$select=id,name,webUrl,lastModifiedDateTime,size,file,folder,parentReference`
  );
}

function isExcelItem(item: GraphDriveItem): boolean {
  const name = item.name?.toLowerCase() || "";
  return Boolean(item.file) && (name.endsWith(".xlsx") || name.endsWith(".xls"));
}

function isBareFileName(value: string): boolean {
  return !/[\\/]/.test(value) && /\.xlsx?$/i.test(value.trim());
}

/**
 * Resolve an allow-listed file by exact name inside the site drive.
 * Prefer this over /shares when using Sites.Selected (shares often needs Files.Read.All).
 */
async function findExcelByFileName(driveId: string, fileName: string): Promise<GraphDriveItem> {
  const escaped = fileName.replace(/'/g, "''");
  const result = await graphFetch<{ value: GraphDriveItem[] }>(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/root/search(q='${encodeURIComponent(escaped)}')?$select=id,name,webUrl,lastModifiedDateTime,size,file,folder`
  );

  const matches = (result.value || []).filter(
    (item) => isExcelItem(item) && normalizePathKey(item.name) === normalizePathKey(fileName)
  );

  if (matches.length === 0) {
    throw new Error(
      `Allow-listed Excel file "${fileName}" was not found in the configured SharePoint/Teams site drive.`
    );
  }

  return matches[0];
}

function explainGraphAuthError(status: number, body: string): string {
  if (status !== 401 && status !== 403) {
    return "";
  }

  return (
    " Hint: Sites.Selected alone is not enough. After admin consent, an admin must also grant this app " +
    "`read` (or `write`) on the specific site via POST /sites/{site-id}/permissions. " +
    "Doc.aspx /shares links often need Files.Read.All; prefer file name or library path with Sites.Selected. " +
    `Graph body: ${body}`
  );
}

async function resolveAllowedItem(
  driveId: string | undefined,
  target: AllowedExcelTarget
): Promise<{ item: GraphDriveItem; driveId: string }> {
  let item: GraphDriveItem;
  let resolvedDriveId = driveId;

  try {
    const fileName = target.expectedFileName || (target.path && isBareFileName(target.path) ? target.path : undefined);

    // Sites.Selected-friendly path: look up exact allow-listed file name in the site drive.
    if (fileName && resolvedDriveId) {
      item = await findExcelByFileName(resolvedDriveId, fileName);
    } else if (target.shareUrl) {
      const shared = await getItemByShareUrl(target.shareUrl);
      item = shared;
      resolvedDriveId = shared.parentReference?.driveId;
      if (!resolvedDriveId) {
        throw new Error(
          `Could not resolve driveId for allow-listed share URL: ${target.label || target.shareUrl}`
        );
      }
    } else if (target.itemId) {
      if (!resolvedDriveId) {
        throw new Error("EVENTS_GRAPH_SITE_URL or EVENTS_GRAPH_SITE_ID is required when using item IDs");
      }
      item = await getItemById(resolvedDriveId, target.itemId);
    } else if (target.path) {
      if (!resolvedDriveId) {
        throw new Error("EVENTS_GRAPH_SITE_URL or EVENTS_GRAPH_SITE_ID is required when using file paths");
      }
      item = await getItemByPath(resolvedDriveId, target.path);
    } else {
      throw new Error("Allow-list entry is missing path, shareUrl, and itemId");
    }
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    const statusMatch = message.match(/Graph request failed \((\d+)\)/);
    const status = statusMatch ? Number(statusMatch[1]) : 0;
    const bodyMatch = message.match(/Graph request failed \(\d+\) [^:]+: ([\s\S]*)$/);
    const body = bodyMatch?.[1] || "";
    const hint = explainGraphAuthError(status, body);
    throw new Error(hint ? `${message}${hint}` : message);
  }

  assertTargetAllowed(item, target);

  if (!isExcelItem(item)) {
    throw new Error(`Allow-listed item "${item.name}" is not an Excel file`);
  }

  return { item, driveId: resolvedDriveId! };
}

function matrixToRecords(values: unknown[][]): Record<string, string>[] {
  if (!values?.length) {
    return [];
  }

  let headerIndex = 0;
  while (headerIndex < values.length) {
    const row = values[headerIndex] || [];
    if (row.some((cell) => cellToString(cell))) {
      break;
    }
    headerIndex += 1;
  }

  if (headerIndex >= values.length) {
    return [];
  }

  const headerRow = values[headerIndex] || [];
  const headers = headerRow.map((cell, index) => {
    const text = cellToString(cell);
    return text || `Column${index + 1}`;
  });

  const records: Record<string, string>[] = [];
  for (let r = headerIndex + 1; r < values.length; r += 1) {
    const row = values[r] || [];
    const record: Record<string, string> = {};
    let hasValue = false;

    headers.forEach((header, index) => {
      if (!header || header.startsWith("__EMPTY")) {
        return;
      }
      const text = cellToString(row[index]);
      if (text) {
        record[header] = text;
        hasValue = true;
      }
    });

    if (hasValue) {
      records.push(record);
    }
  }

  return records;
}

/**
 * Read worksheet data in place through Microsoft Graph Excel APIs.
 * This does not download the .xlsx binary.
 */
async function readWorkbookSheets(driveId: string, itemId: string): Promise<GraphSheetData[]> {
  const worksheets = await graphFetch<{ value: Array<{ id: string; name: string }> }>(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/workbook/worksheets?$select=id,name,position`
  );

  const sheets: GraphSheetData[] = [];

  for (const worksheet of worksheets.value || []) {
    try {
      const usedRange = await graphFetch<{ values?: unknown[][] }>(
        `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/workbook/worksheets/${worksheet.id}/usedRange(valuesOnly=true)`
      );
      sheets.push({
        sheet: worksheet.name,
        rows: matrixToRecords(usedRange.values || []),
      });
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      // Skip empty/hidden sheets that Graph cannot open for usedRange.
      if (message.includes("(404)") || message.includes("(400)")) {
        continue;
      }
      throw error;
    }
  }

  return sheets;
}

/** Read allow-listed Excel files via Graph workbook APIs (no file download). */
export async function readAllowedExcelWorkbooksFromGraph(): Promise<GraphWorkbookData[]> {
  const allowed = getAllowedExcelTargets();
  if (allowed.length === 0) {
    throw new Error(
      "No allow-listed Excel files configured. Set EVENTS_GRAPH_ALLOWED_FILES or EVENTS_GRAPH_FILE_PATH (or ITEM_ID)."
    );
  }

  // Always resolve the site drive when site is configured so Doc.aspx links can
  // be mapped by file name (Sites.Selected) instead of /shares (often Files.Read.All).
  const needsSiteDrive =
    Boolean(process.env.EVENTS_GRAPH_SITE_URL?.trim() || process.env.EVENTS_GRAPH_SITE_ID?.trim()) ||
    allowed.some((target) => !target.shareUrl || target.expectedFileName);
  let siteDriveId: string | undefined;
  if (needsSiteDrive) {
    const siteId = await resolveSiteId();
    siteDriveId = await resolveDriveId(siteId);
  }

  const workbooks: GraphWorkbookData[] = [];

  for (const target of allowed) {
    const { item, driveId } = await resolveAllowedItem(siteDriveId, target);
    const sheets = await readWorkbookSheets(driveId, item.id);
    workbooks.push({
      source: item.webUrl || `graph://drives/${driveId}/items/${item.id}/workbook`,
      fileName: item.name,
      webUrl: item.webUrl,
      allowedAs: target.label || target.path || target.shareUrl || target.itemId || item.name,
      sheets,
    });
  }

  return workbooks;
}

export function describeGraphExcelConfig(): string {
  const allowed = getAllowedExcelTargets();
  const site = process.env.EVENTS_GRAPH_SITE_URL || process.env.EVENTS_GRAPH_SITE_ID || "(site unset)";
  const files =
    allowed.length > 0
      ? allowed.map((t) => t.expectedFileName || t.path || t.shareUrl || t.itemId || "?").join(", ")
      : "(no allow-listed files)";
  return `${site} → [${files}] (Graph Excel API)`;
}

export function assertGraphExcelConfig(): void {
  requireEnv("CLIENT_ID");
  requireEnv("CLIENT_SECRET");
  requireEnv("TENANT_ID");

  const allowed = getAllowedExcelTargets();
  if (allowed.length === 0) {
    throw new Error(
      "Set EVENTS_GRAPH_ALLOWED_FILES or EVENTS_GRAPH_FILE_PATH / EVENTS_GRAPH_ITEM_ID. Broad folder scanning is disabled."
    );
  }

  const needsSite = allowed.some((target) => !target.shareUrl);
  if (needsSite && !process.env.EVENTS_GRAPH_SITE_ID?.trim() && !process.env.EVENTS_GRAPH_SITE_URL?.trim()) {
    throw new Error("Set EVENTS_GRAPH_SITE_URL or EVENTS_GRAPH_SITE_ID for Teams/SharePoint Excel access");
  }

  // Validate site URL shape early when provided.
  if (process.env.EVENTS_GRAPH_SITE_URL?.trim()) {
    normalizeGraphSiteUrl(process.env.EVENTS_GRAPH_SITE_URL.trim());
  }
}

export interface GraphChatAttachment {
  id?: string;
  name?: string;
  contentType?: string;
  contentUrl?: string;
  content?: string;
}

export interface GraphChatMessage {
  id?: string;
  body?: { content?: string; contentType?: string };
  attachments?: GraphChatAttachment[];
}

async function graphFetchOptional<T>(url: string): Promise<T | undefined> {
  try {
    return await graphFetch<T>(url);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    if (message.includes("(404)") || message.includes("(400)")) {
      return undefined;
    }
    throw error;
  }
}

/** Read a Teams chat/channel message via Graph (includes file attachments the Bot Framework payload often omits). */
export async function fetchTeamsMessageFromGraph(
  conversationId: string,
  messageId: string,
  channelData?: { team?: { id?: string }; channel?: { id?: string } }
): Promise<GraphChatMessage | undefined> {
  const encodedMessageId = encodeURIComponent(messageId);
  const teamId = channelData?.team?.id;
  const channelId = channelData?.channel?.id;

  if (teamId && channelId) {
    const encodedTeam = encodeURIComponent(teamId);
    const encodedChannel = encodeURIComponent(channelId);
    const fromChannel = await graphFetchOptional<GraphChatMessage>(
      `https://graph.microsoft.com/v1.0/teams/${encodedTeam}/channels/${encodedChannel}/messages/${encodedMessageId}?$expand=attachments`
    );
    if (fromChannel) {
      return fromChannel;
    }
  }

  const encodedChatId = encodeURIComponent(conversationId);
  return graphFetchOptional<GraphChatMessage>(
    `https://graph.microsoft.com/v1.0/chats/${encodedChatId}/messages/${encodedMessageId}`
  );
}

/** Download an Excel file from a SharePoint/OneDrive/Teams sharing or content URL. */
export async function downloadExcelBinaryFromShareUrl(
  shareUrl: string
): Promise<{ buffer: Buffer; fileName: string }> {
  const item = await getItemByShareUrl(shareUrl);
  const driveId = item.parentReference?.driveId;
  if (!driveId || !item.id) {
    throw new Error(`Could not resolve drive item for shared Excel URL: ${shareUrl}`);
  }

  const meta = await graphFetch<{
    name?: string;
    "@microsoft.graph.downloadUrl"?: string;
  }>(`https://graph.microsoft.com/v1.0/drives/${driveId}/items/${item.id}?$select=id,name,@microsoft.graph.downloadUrl`);

  const fileName = item.name || meta.name || extractFileNameFromUrl(shareUrl) || "workbook.xlsx";
  const downloadUrl = meta["@microsoft.graph.downloadUrl"];
  if (downloadUrl) {
    const response = await fetch(downloadUrl, { redirect: "follow" });
    if (!response.ok) {
      throw new Error(`Failed to download shared Excel (${response.status})`);
    }
    return { buffer: Buffer.from(await response.arrayBuffer()), fileName };
  }

  const content = await graphFetch<ArrayBuffer>(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${item.id}/content`
  );
  return { buffer: Buffer.from(content), fileName };
}
