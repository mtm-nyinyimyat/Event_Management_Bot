import crypto from "crypto";
import { Attachment, MessageActivity } from "@microsoft/teams.api";
import { ILogger } from "@microsoft/teams.common";
import { findActiveEventSession } from "./eventSession";
import {
  getConversationWorkbook,
  sheetsToXlsxBuffer,
  type CachedWorkbook,
} from "./excelStore";

const FILE_CONSENT_CARD = "application/vnd.microsoft.teams.card.file.consent";
const FILE_INFO_CARD = "application/vnd.microsoft.teams.card.file.info";

export const SERVICE_UNAVAILABLE_TEXT =
  "Sorry, I'm a bit stuck right now — tap Accept below if you want the Excel file.";

const SERVICE_UNAVAILABLE_NO_FILE_TEXT =
  "Sorry, I'm a bit stuck right now. Try again in a moment?";

interface ActiveEventExport {
  buffer: Buffer;
  fileName: string;
  conversationId: string;
  downloadToken: string;
}

let activeExport: ActiveEventExport | null = null;

function sanitizeFileName(name: string): string {
  const base = (name || "event-workbook.xlsx").replace(/[^\w.\- ()]/g, "_");
  return base.toLowerCase().endsWith(".xlsx") ? base : `${base}.xlsx`;
}

export function clearActiveEventExport(): void {
  activeExport = null;
}

export function setActiveEventExport(options: {
  workbook: CachedWorkbook;
  conversationId: string;
  originalBuffer?: Buffer;
}): void {
  const fileName = sanitizeFileName(options.workbook.fileName || "event-workbook.xlsx");
  const buffer =
    options.originalBuffer && options.originalBuffer.length > 0
      ? options.originalBuffer
      : sheetsToXlsxBuffer(options.workbook.sheets);

  activeExport = {
    buffer,
    fileName,
    conversationId: options.conversationId,
    downloadToken: crypto.randomBytes(24).toString("hex"),
  };
}

export function getActiveEventExport(): ActiveEventExport | null {
  return activeExport;
}

export function getActiveEventExportForDownload(token: string): ActiveEventExport | null {
  if (!activeExport || activeExport.downloadToken !== token) {
    return null;
  }
  return activeExport;
}

async function resolveWorkbookBytes(): Promise<ActiveEventExport | null> {
  if (activeExport?.buffer?.length) {
    return activeExport;
  }

  const session = await findActiveEventSession();
  if (!session) {
    return null;
  }

  const workbook = getConversationWorkbook(session.conversationId);
  if (!workbook?.sheets?.length) {
    return null;
  }

  try {
    setActiveEventExport({ workbook, conversationId: session.conversationId });
    return activeExport;
  } catch {
    return null;
  }
}

function buildFileConsentAttachment(file: ActiveEventExport): Attachment {
  return {
    contentType: FILE_CONSENT_CARD,
    name: file.fileName,
    content: {
      description: "Active event workbook (Excel)",
      sizeInBytes: file.buffer.length,
      acceptContext: { token: file.downloadToken, kind: "active-event-workbook" },
      declineContext: { token: file.downloadToken, kind: "active-event-workbook" },
    },
  };
}

export interface ServiceUnavailableReply {
  message: MessageActivity;
  hasFileConsent: boolean;
  fileName?: string;
}

/** User-facing reply when an internal error occurred (no raw error text). */
export async function createServiceUnavailableReply(): Promise<ServiceUnavailableReply> {
  const file = await resolveWorkbookBytes();
  if (!file) {
    return {
      message: new MessageActivity(SERVICE_UNAVAILABLE_NO_FILE_TEXT),
      hasFileConsent: false,
    };
  }

  return {
    message: new MessageActivity(SERVICE_UNAVAILABLE_TEXT).addAttachments(
      buildFileConsentAttachment(file)
    ),
    hasFileConsent: true,
    fileName: file.fileName,
  };
}

export function buildFileInfoAttachment(uploadInfo: {
  name?: string;
  contentUrl?: string;
  uniqueId?: string;
  fileType?: string;
}): Attachment {
  return {
    contentType: FILE_INFO_CARD,
    name: uploadInfo.name || "event-workbook.xlsx",
    contentUrl: uploadInfo.contentUrl,
    content: {
      uniqueId: uploadInfo.uniqueId,
      fileType: uploadInfo.fileType || "xlsx",
    },
  };
}

export async function uploadAcceptedEventWorkbook(options: {
  uploadUrl: string;
  token?: string;
  logger?: ILogger;
}): Promise<{ buffer: Buffer; fileName: string } | null> {
  const file =
    (options.token ? getActiveEventExportForDownload(options.token) : null) || getActiveEventExport();
  if (!file?.buffer?.length) {
    options.logger?.warn("File consent accepted but no active event workbook is available");
    return null;
  }

  const response = await fetch(options.uploadUrl, {
    method: "PUT",
    headers: {
      "Content-Length": String(file.buffer.length),
      "Content-Range": `bytes 0-${file.buffer.length - 1}/${file.buffer.length}`,
    },
    body: new Uint8Array(file.buffer),
  });

  if (!response.ok) {
    const body = await response.text().catch(() => "");
    throw new Error(
      `Failed to upload event workbook (${response.status}): ${body.slice(0, 300)}`
    );
  }

  return { buffer: file.buffer, fileName: file.fileName };
}

export class ServiceUnavailableError extends Error {
  constructor(cause?: unknown) {
    super(SERVICE_UNAVAILABLE_TEXT);
    this.name = "ServiceUnavailableError";
    if (cause instanceof Error) {
      this.cause = cause;
    }
  }
}
