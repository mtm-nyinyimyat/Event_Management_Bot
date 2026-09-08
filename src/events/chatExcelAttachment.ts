import { Attachment, IMessageActivity } from "@microsoft/teams.api";
import { ClientSecretCredential } from "@azure/identity";
import { ILogger } from "@microsoft/teams.common";
import {
  downloadExcelBinaryFromShareUrl,
  fetchTeamsMessageFromGraph,
} from "./graphExcelClient";
import {
  loadWorkbookFromBuffer,
  setUploadedWorkbook,
  SheetData,
} from "./excelStore";

const TEAMS_FILE_DOWNLOAD_INFO = "application/vnd.microsoft.teams.file.download.info";
const EXCEL_MIME =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

interface TeamsFileDownloadContent {
  downloadUrl?: string;
  uniqueId?: string;
  fileType?: string;
  etag?: string;
}

export interface IngestedExcelAttachment {
  fileName: string;
  sheetCount: number;
  rowCount: number;
}

function isExcelFileName(name?: string): boolean {
  if (!name) {
    return false;
  }
  const lower = name.toLowerCase();
  return lower.endsWith(".xlsx") || lower.endsWith(".xls") || lower.includes(".xlsx");
}

function normalizeAttachmentContent(
  attachment: Attachment
): TeamsFileDownloadContent | undefined {
  const raw = attachment.content;
  if (!raw) {
    return undefined;
  }
  if (typeof raw === "string") {
    try {
      return JSON.parse(raw) as TeamsFileDownloadContent;
    } catch {
      return undefined;
    }
  }
  if (typeof raw === "object") {
    return raw as TeamsFileDownloadContent;
  }
  return undefined;
}

function isHtmlOrPlainBodyAttachment(attachment: Attachment): boolean {
  const contentType = (attachment.contentType || "").toLowerCase();
  // Teams always includes an HTML rendering of the message text — ignore it for file detection.
  return contentType === "text/html" || contentType === "text/plain";
}

function isNonFileBodyAttachment(attachment: Attachment): boolean {
  const contentType = (attachment.contentType || "").toLowerCase();
  if (isHtmlOrPlainBodyAttachment(attachment)) {
    return true;
  }
  // Adaptive/hero cards are not Excel uploads.
  if (contentType.startsWith("application/vnd.microsoft.card.")) {
    return true;
  }
  return false;
}

function attachmentLooksLikeExcel(attachment: Attachment): boolean {
  if (isHtmlOrPlainBodyAttachment(attachment)) {
    return false;
  }

  if (isExcelFileName(attachment.name)) {
    return true;
  }

  const contentType = (attachment.contentType || "").toLowerCase();
  const content = normalizeAttachmentContent(attachment);
  const fileType = (content?.fileType || "").toLowerCase();

  if (
    contentType === EXCEL_MIME ||
    contentType.includes("spreadsheet") ||
    contentType.includes("excel")
  ) {
    return true;
  }

  // Teams personal-chat file download card.
  if (contentType === TEAMS_FILE_DOWNLOAD_INFO) {
    if (fileType === "xlsx" || fileType === "xls" || isExcelFileName(attachment.name)) {
      return true;
    }
    if (content?.downloadUrl) {
      return true;
    }
  }

  // File info card (sometimes used when a file is shared into chat).
  if (contentType === "application/vnd.microsoft.teams.card.file.info") {
    return fileType === "xlsx" || fileType === "xls" || isExcelFileName(attachment.name);
  }

  if (
    (contentType === "application/octet-stream" || contentType === "application/haansoftxlsx") &&
    (isExcelFileName(attachment.name) || isExcelFileName(attachment.contentUrl))
  ) {
    return true;
  }

  if (content?.downloadUrl && isExcelFileName(attachment.name)) {
    return true;
  }

  return false;
}

/**
 * True when Teams likely showed a file in the UI but did not give the bot downloadable Excel bytes.
 * Classic symptom with supportsFiles=false: empty text + only text/html attachment.
 */
export function looksLikeUndeliveredFileShare(activity: IMessageActivity): boolean {
  const text = (activity.text || "").replace(/<at>[^<]*<\/at>/gi, "").trim();
  if (text) {
    return false;
  }
  if (findExcelAttachments(activity).length > 0) {
    return false;
  }
  const attachments = activity.attachments || [];
  if (attachments.length === 0) {
    return true; // empty message — often a file card the bot cannot see
  }
  return attachments.every((attachment) => isHtmlOrPlainBodyAttachment(attachment));
}

export const FILE_UPLOAD_HELP =
  "Teams showed a file card, but this bot only received the chat HTML — not the Excel bytes.\n\n" +
  "That usually means the file was **shared from Teams/SharePoint Files** (preview card), not uploaded as a bot file.\n\n" +
  "Try this:\n" +
  "1. In a **1:1 chat** with the bot, click the paperclip\n" +
  "2. Choose **Upload from this device** (not OneDrive / Teams files)\n" +
  "3. Pick the .xlsx from your computer\n\n" +
  "If you keep sharing from Teams Files, an admin must grant this app Graph access to that chat/site so I can read the file from Microsoft Graph.";

export function describeActivityAttachments(activity: IMessageActivity): string {
  const attachments = activity.attachments || [];
  if (attachments.length === 0) {
    return "(no attachments on this message)";
  }
  return attachments
    .map((attachment, index) => {
      const content = normalizeAttachmentContent(attachment);
      return (
        `#${index + 1} name="${attachment.name || ""}" ` +
        `contentType="${attachment.contentType || ""}" ` +
        `hasContentUrl=${Boolean(attachment.contentUrl)} ` +
        `hasDownloadUrl=${Boolean(content?.downloadUrl)} ` +
        `fileType="${content?.fileType || ""}"`
      );
    })
    .join("; ");
}

/** Excel attachments Teams (or DevTools) attached to this message. */
export function findExcelAttachments(activity: IMessageActivity): Attachment[] {
  const attachments = activity.attachments || [];
  return attachments.filter(attachmentLooksLikeExcel);
}

async function getBotFrameworkToken(): Promise<string | undefined> {
  const tenantId = process.env.TENANT_ID || process.env.GRAPH_TENANT_ID;
  const clientId = process.env.CLIENT_ID || process.env.GRAPH_CLIENT_ID;
  const clientSecret = process.env.CLIENT_SECRET || process.env.GRAPH_CLIENT_SECRET;
  if (!tenantId || !clientId || !clientSecret) {
    return undefined;
  }

  try {
    const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
    const token = await credential.getToken("https://api.botframework.com/.default");
    return token?.token;
  } catch {
    return undefined;
  }
}

async function fetchBinary(url: string, token?: string): Promise<Buffer> {
  const headers: Record<string, string> = {
    Accept: "*/*",
  };
  if (token) {
    headers.Authorization = `Bearer ${token}`;
  }

  const response = await fetch(url, { headers, redirect: "follow" });
  if (!response.ok) {
    const body = await response.text();
    throw new Error(`Failed to download chat attachment (${response.status}): ${body.slice(0, 300)}`);
  }

  return Buffer.from(await response.arrayBuffer());
}

function looksLikeZipOrXlsx(buffer: Buffer): boolean {
  // XLSX is a ZIP (PK…). Legacy .xls starts with D0 CF 11 E0.
  if (buffer.length >= 2 && buffer[0] === 0x50 && buffer[1] === 0x4b) {
    return true;
  }
  if (
    buffer.length >= 4 &&
    buffer[0] === 0xd0 &&
    buffer[1] === 0xcf &&
    buffer[2] === 0x11 &&
    buffer[3] === 0xe0
  ) {
    return true;
  }
  return false;
}

function extractUrls(text?: string): string[] {
  if (!text) {
    return [];
  }
  const matches = text.match(/https?:\/\/[^\s"'<>]+/gi) || [];
  return [...new Set(matches.map((url) => url.replace(/&amp;/g, "&")))];
}

function isLikelyExcelShareUrl(url: string): boolean {
  const lower = url.toLowerCase();
  return (
    lower.includes("sharepoint.com") ||
    lower.includes("onedrive") ||
    lower.includes("/:x:/") ||
    isExcelFileName(url)
  );
}

function storeLoadedWorkbook(
  files: IngestedExcelAttachment[],
  sheets: SheetData[],
  conversationId: string,
  userId?: string
): { confirmation: string; files: IngestedExcelAttachment[] } {
  const primaryName = files.map((file) => file.fileName).join(", ");
  setUploadedWorkbook(
    {
      source: `chat-upload://${primaryName}`,
      sourceType: "upload",
      fileName: primaryName,
      sheets,
      loadedAt: Date.now(),
    },
    { conversationId, userId }
  );

  const details = files
    .map((file) => `• ${file.fileName} (${file.sheetCount} sheet(s), ${file.rowCount} row(s))`)
    .join("\n");

  return {
    files,
    confirmation:
      `Loaded Excel from chat:\n${details}\n\n` +
      `I will use this exact file in every chat until you upload a different .xlsx.\n` +
      `Ask me questions about it (in groups, @mention me).`,
  };
}

async function loadExcelBuffer(
  buffer: Buffer,
  fileName: string
): Promise<{ sheets: SheetData[]; file: IngestedExcelAttachment }> {
  if (!looksLikeZipOrXlsx(buffer)) {
    throw new Error(
      `Downloaded "${fileName}" but it does not look like a valid Excel file (unexpected binary format).`
    );
  }
  const sheets = loadWorkbookFromBuffer(buffer, undefined);
  return {
    sheets,
    file: {
      fileName,
      sheetCount: sheets.length,
      rowCount: sheets.reduce((sum, sheet) => sum + sheet.rows.length, 0),
    },
  };
}

/**
 * Teams often omits file bytes from the bot activity and only shows a SharePoint preview card.
 * Fetch the same chat message from Graph and download Excel attachments from there.
 */
async function ingestExcelFromGraphMessage(
  activity: IMessageActivity,
  conversationId: string,
  logger?: ILogger
): Promise<{ files: IngestedExcelAttachment[]; sheets: SheetData[] } | null> {
  const channelData = activity.channelData as
    | { team?: { id?: string }; channel?: { id?: string } }
    | undefined;

  let graphMessage;
  try {
    graphMessage = await fetchTeamsMessageFromGraph(conversationId, activity.id, channelData);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    logger?.warn(`Graph chat-message lookup failed: ${message}`);
    return null;
  }

  if (!graphMessage) {
    logger?.debug("Graph returned no chat message for this activity id");
    return null;
  }

  const candidates: Array<{ url: string; name?: string }> = [];
  for (const attachment of graphMessage.attachments || []) {
    if (!attachment.contentUrl) {
      continue;
    }
    if (
      isExcelFileName(attachment.name) ||
      isExcelFileName(attachment.contentUrl) ||
      isLikelyExcelShareUrl(attachment.contentUrl)
    ) {
      candidates.push({ url: attachment.contentUrl, name: attachment.name });
    }
  }

  for (const url of extractUrls(graphMessage.body?.content)) {
    if (isLikelyExcelShareUrl(url)) {
      candidates.push({ url, name: undefined });
    }
  }

  const unique = candidates.filter(
    (candidate, index) => candidates.findIndex((other) => other.url === candidate.url) === index
  );
  if (unique.length === 0) {
    logger?.debug(
      `Graph message ${graphMessage.id} has no Excel attachments (${graphMessage.attachments?.length || 0} attachment(s))`
    );
    return null;
  }

  const files: IngestedExcelAttachment[] = [];
  const allSheets: SheetData[] = [];
  const errors: string[] = [];

  for (const candidate of unique) {
    try {
      logger?.debug(`📎 Downloading Graph chat Excel: ${candidate.name || candidate.url}`);
      const { buffer, fileName } = await downloadExcelBinaryFromShareUrl(candidate.url);
      const loaded = await loadExcelBuffer(buffer, candidate.name || fileName);
      files.push(loaded.file);
      allSheets.push(...loaded.sheets);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      errors.push(message);
      logger?.warn(`Graph Excel download failed for ${candidate.url}: ${message}`);
    }
  }

  if (files.length === 0) {
    throw new Error(
      `Teams shared an Excel file, but Graph could not download it.\n${errors.join("\n")}`
    );
  }

  return { files, sheets: allSheets };
}

async function downloadExcelAttachment(attachment: Attachment): Promise<{
  buffer: Buffer;
  fileName: string;
}> {
  const content = normalizeAttachmentContent(attachment);
  const fileName =
    attachment.name ||
    (content?.fileType ? `workbook.${content.fileType}` : "workbook.xlsx");

  if (content?.downloadUrl) {
    const buffer = await fetchBinary(content.downloadUrl);
    return { buffer, fileName };
  }

  if (!attachment.contentUrl) {
    throw new Error(
      `Attachment "${fileName}" has no download URL. Attach an .xlsx file directly to the message in a 1:1 chat with the bot.`
    );
  }

  const token = await getBotFrameworkToken();
  try {
    const buffer = await fetchBinary(attachment.contentUrl, token);
    return { buffer, fileName };
  } catch (error) {
    if (!token) {
      throw error;
    }
    const buffer = await fetchBinary(attachment.contentUrl);
    return { buffer, fileName };
  }
}

/**
 * If the message includes Excel attachments, download/parse them and remember
 * globally so any chat can query the workbook.
 */
export async function ingestExcelAttachmentsFromActivity(
  activity: IMessageActivity,
  conversationId: string,
  logger?: ILogger,
  userId?: string
): Promise<{ confirmation: string; files: IngestedExcelAttachment[] } | null> {
  const allAttachments = activity.attachments || [];
  logger?.debug(`📎 Activity attachments: ${describeActivityAttachments(activity)}`);

  const fileLikeAttachments = allAttachments.filter((attachment) => !isNonFileBodyAttachment(attachment));
  const excelAttachments = findExcelAttachments(activity);

  if (excelAttachments.length > 0) {
    const files: IngestedExcelAttachment[] = [];
    const allSheets: SheetData[] = [];

    for (const attachment of excelAttachments) {
      logger?.debug(`📎 Downloading chat Excel attachment: ${attachment.name || "(unnamed)"}`);
      const { buffer, fileName } = await downloadExcelAttachment(attachment);
      const loaded = await loadExcelBuffer(buffer, fileName);
      files.push(loaded.file);
      allSheets.push(...loaded.sheets);
    }

    return storeLoadedWorkbook(files, allSheets, conversationId, userId);
  }

  if (fileLikeAttachments.length > 0) {
    throw new Error(
      `I received ${fileLikeAttachments.length} file attachment(s), but none looked like Excel (.xlsx).\n` +
        `Details: ${describeActivityAttachments(activity)}`
    );
  }

  // Bot Framework payload had no Excel. Try Graph (SharePoint/Teams file cards).
  if (looksLikeUndeliveredFileShare(activity) || !activity.text?.trim()) {
    const fromGraph = await ingestExcelFromGraphMessage(activity, conversationId, logger);
    if (fromGraph) {
      return storeLoadedWorkbook(fromGraph.files, fromGraph.sheets, conversationId, userId);
    }
  }

  return null;
}
