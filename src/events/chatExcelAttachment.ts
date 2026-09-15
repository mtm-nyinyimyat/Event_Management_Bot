import { Attachment, IMessageActivity } from "@microsoft/teams.api";
import { ClientSecretCredential } from "@azure/identity";
import { ILogger } from "@microsoft/teams.common";

const TEAMS_FILE_DOWNLOAD_INFO = "application/vnd.microsoft.teams.file.download.info";
const EXCEL_MIME =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

interface TeamsFileDownloadContent {
  downloadUrl?: string;
  uniqueId?: string;
  fileType?: string;
  etag?: string;
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
  return contentType === "text/html" || contentType === "text/plain";
}

function isNonFileBodyAttachment(attachment: Attachment): boolean {
  const contentType = (attachment.contentType || "").toLowerCase();
  if (isHtmlOrPlainBodyAttachment(attachment)) {
    return true;
  }
  if (contentType.startsWith("application/vnd.microsoft.card.")) {
    return true;
  }
  if (contentType.startsWith("image/") && !isExcelFileName(attachment.name)) {
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

  if (contentType === TEAMS_FILE_DOWNLOAD_INFO) {
    if (fileType === "xlsx" || fileType === "xls" || isExcelFileName(attachment.name)) {
      return true;
    }
    if (content?.downloadUrl) {
      return true;
    }
  }

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

/** Teams animated emoji / sticker with no real user text. */
export function isEmoticonOnlyActivity(activity: IMessageActivity): boolean {
  const text = (activity.text || "").replace(/<at>[^<]*<\/at>/gi, "").trim();
  if (text) {
    return false;
  }
  const attachments = activity.attachments || [];
  if (!attachments.length) {
    return false;
  }
  let hasImage = false;
  for (const attachment of attachments) {
    const contentType = (attachment.contentType || "").toLowerCase();
    if (contentType.startsWith("image/")) {
      hasImage = true;
      continue;
    }
    if (isHtmlOrPlainBodyAttachment(attachment)) {
      continue;
    }
    return false;
  }
  return hasImage;
}

/**
 * True when Teams likely showed a file in the UI but did not give the bot downloadable Excel bytes.
 */
export function looksLikeUndeliveredFileShare(activity: IMessageActivity): boolean {
  if (isEmoticonOnlyActivity(activity)) {
    return false;
  }
  const text = (activity.text || "").replace(/<at>[^<]*<\/at>/gi, "").trim();
  if (text) {
    return false;
  }
  if (findExcelAttachments(activity).length > 0) {
    return false;
  }
  const attachments = activity.attachments || [];
  if (attachments.length === 0) {
    return true;
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

export function findExcelAttachments(activity: IMessageActivity): Attachment[] {
  return (activity.attachments || []).filter(attachmentLooksLikeExcel);
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
  const headers: Record<string, string> = { Accept: "*/*" };
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

async function downloadExcelAttachment(attachment: Attachment): Promise<{
  buffer: Buffer;
  fileName: string;
}> {
  const content = normalizeAttachmentContent(attachment);
  const fileName =
    attachment.name ||
    (content?.fileType ? `workbook.${content.fileType}` : "workbook.xlsx");

  if (content?.downloadUrl) {
    return { buffer: await fetchBinary(content.downloadUrl), fileName };
  }

  if (!attachment.contentUrl) {
    throw new Error(
      `Attachment "${fileName}" has no download URL. Attach an .xlsx file directly to the message in a 1:1 chat with the bot.`
    );
  }

  const token = await getBotFrameworkToken();
  try {
    return { buffer: await fetchBinary(attachment.contentUrl, token), fileName };
  } catch (error) {
    if (!token) {
      throw error;
    }
    return { buffer: await fetchBinary(attachment.contentUrl), fileName };
  }
}

export interface DownloadedExcelFile {
  fileName: string;
  buffer: Buffer;
}

export function activityLikelyHasExcelUpload(activity: IMessageActivity): boolean {
  if (findExcelAttachments(activity).length > 0) {
    return true;
  }
  return looksLikeUndeliveredFileShare(activity);
}

/**
 * Download Excel bytes from chat attachments. Does not activate a workbook — callers stage via /start.
 */
export async function downloadExcelFilesFromActivity(
  activity: IMessageActivity,
  _conversationId: string,
  logger?: ILogger
): Promise<DownloadedExcelFile[] | null> {
  logger?.debug(`📎 Activity attachments: ${describeActivityAttachments(activity)}`);

  const allAttachments = activity.attachments || [];
  const fileLikeAttachments = allAttachments.filter((attachment) => !isNonFileBodyAttachment(attachment));
  const excelAttachments = findExcelAttachments(activity);

  if (excelAttachments.length > 0) {
    const files: DownloadedExcelFile[] = [];
    for (const attachment of excelAttachments) {
      logger?.debug(`📎 Downloading chat Excel attachment: ${attachment.name || "(unnamed)"}`);
      const downloaded = await downloadExcelAttachment(attachment);
      if (!looksLikeZipOrXlsx(downloaded.buffer)) {
        throw new Error(
          `Downloaded "${downloaded.fileName}" but it does not look like a valid Excel file.`
        );
      }
      files.push(downloaded);
    }
    return files;
  }

  if (fileLikeAttachments.length > 0) {
    if (looksLikeUndeliveredFileShare(activity) || isExcelFileName(activity.text)) {
      throw new Error(
        `I received ${fileLikeAttachments.length} file attachment(s), but none looked like Excel (.xlsx).\n` +
          `Details: ${describeActivityAttachments(activity)}`
      );
    }
    logger?.debug(
      `Ignoring ${fileLikeAttachments.length} non-Excel attachment(s): ${describeActivityAttachments(activity)}`
    );
    return null;
  }

  if (looksLikeUndeliveredFileShare(activity)) {
    throw new Error(FILE_UPLOAD_HELP);
  }

  return null;
}
