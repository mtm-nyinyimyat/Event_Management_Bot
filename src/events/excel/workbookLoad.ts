import fs from "fs";
import path from "path";
import * as XLSX from "xlsx";
import { ensureWorkbookIndexed } from "../../rag";
import { getRagConfig } from "../../rag/config";
import { clearQueryCaches } from "../../utils/queryCache";
import {
  assertGraphExcelConfig,
  describeGraphExcelConfig,
  readAllowedExcelWorkbooksFromGraph,
} from "../graphExcelClient";
import { buildSheetSummary, parseSheet } from "./sheetParse";
import type { CachedWorkbook, SheetData, WorkbookLookupOptions } from "./types";
import {
  cacheTtlMs,
  getEventsSource,
  getWorkbookCache,
  isEventSessionMode,
  resolveExcelPath,
  resolveUploadedWorkbook,
  setWorkbookCache,
} from "./workbookCache";

export function resolveWorkbookOptions(
  conversationIdOrOptions?: string | WorkbookLookupOptions
): WorkbookLookupOptions {
  if (typeof conversationIdOrOptions === "string") {
    return { conversationId: conversationIdOrOptions };
  }
  return conversationIdOrOptions || {};
}

function parseWorkbookBuffer(buffer: Buffer, sheetPrefix?: string): SheetData[] {
  const workbook = XLSX.read(buffer, { type: "buffer" });
  if (!workbook.SheetNames.length) {
    throw new Error("No worksheets found in Excel workbook");
  }

  return workbook.SheetNames.map((sheetName) => {
    const values = XLSX.utils.sheet_to_json<unknown[]>(workbook.Sheets[sheetName], {
      header: 1,
      defval: "",
      raw: false,
      blankrows: false,
    });

    const parsed = parseSheet(sheetName, values);
    return {
      ...parsed,
      sheet: sheetPrefix ? `${sheetPrefix} / ${sheetName}` : sheetName,
    };
  });
}

export function loadWorkbookFromBuffer(buffer: Buffer, fileName?: string): SheetData[] {
  return parseWorkbookBuffer(buffer, fileName);
}

/** Rebuild an .xlsx from parsed sheet rows (for re-sending the active event file). */
export function sheetsToXlsxBuffer(sheets: SheetData[]): Buffer {
  const workbook = XLSX.utils.book_new();
  const usedTitles = new Set<string>();

  for (const sheet of sheets) {
    let title = (sheet.sheet || "Sheet").slice(0, 31) || "Sheet";
    let suffix = 1;
    while (usedTitles.has(title)) {
      const base = (sheet.sheet || "Sheet").slice(0, Math.max(1, 31 - String(suffix).length - 1));
      title = `${base}_${suffix}`;
      suffix += 1;
    }
    usedTitles.add(title);
    const ws =
      sheet.rows.length > 0
        ? XLSX.utils.json_to_sheet(sheet.rows)
        : XLSX.utils.aoa_to_sheet([]);
    XLSX.utils.book_append_sheet(workbook, ws, title);
  }
  return Buffer.from(XLSX.write(workbook, { type: "buffer", bookType: "xlsx" }));
}

function loadLocalWorkbook(): CachedWorkbook {
  const filePath = resolveExcelPath();
  if (!fs.existsSync(filePath)) {
    throw new Error(
      `Excel file not found at ${filePath}. Place an .xlsx in data/ or set EVENTS_EXCEL_PATH.`
    );
  }

  const buffer = fs.readFileSync(filePath);
  return {
    source: filePath,
    sourceType: "local",
    fileName: path.basename(filePath),
    sheets: parseWorkbookBuffer(buffer),
    loadedAt: Date.now(),
  };
}

async function loadGraphWorkbook(): Promise<CachedWorkbook> {
  assertGraphExcelConfig();
  const workbooks = await readAllowedExcelWorkbooksFromGraph();
  if (!workbooks.length) {
    throw new Error(`No Excel workbooks returned from Graph (${describeGraphExcelConfig()})`);
  }

  const multi = workbooks.length > 1;
  const sheets: SheetData[] = workbooks.flatMap((workbook) =>
    workbook.sheets.map((sheet) => {
      const sheetName = multi ? `${workbook.fileName} / ${sheet.sheet}` : sheet.sheet;
      const rows = sheet.rows as SheetData["rows"];
      return {
        sheet: sheetName,
        rows,
        numbered_item_count: rows.filter((row) =>
          /^\d+$/.test(String(row["No."] || row.No || "").trim())
        ).length,
        summary: buildSheetSummary(sheetName, rows),
      };
    })
  );

  return {
    source: workbooks.map((workbook) => workbook.webUrl || workbook.source).join(" | "),
    sourceType: "graph",
    fileName: workbooks.map((workbook) => workbook.fileName).join(", "),
    sheets,
    loadedAt: Date.now(),
  };
}

export async function loadWorkbook(
  forceRefresh = false,
  conversationIdOrOptions?: string | WorkbookLookupOptions
): Promise<SheetData[]> {
  const options = resolveWorkbookOptions(conversationIdOrOptions);
  const conversationId = options.conversationId;
  const ttl = cacheTtlMs();

  const uploaded = resolveUploadedWorkbook(conversationId, options.userId);
  if (!uploaded && isEventSessionMode() && conversationId) {
    try {
      const { ensureActiveWorkbookInMemory, getEventSession } = await import(
        "../eventSessionLazy"
      ).then((mod) => mod.importEventSession());
      await ensureActiveWorkbookInMemory(conversationId);
      if (!options.documentId) {
        const session = await getEventSession(conversationId);
        options.documentId = session.documentId || undefined;
      }
    } catch (error) {
      console.warn(
        `Failed to rehydrate active event workbook: ${
          error instanceof Error ? error.message : String(error)
        }`
      );
    }
  }

  const resolvedUpload = resolveUploadedWorkbook(conversationId, options.userId);
  if (resolvedUpload) {
    const isChatUpload =
      resolvedUpload.sourceType === "upload" ||
      resolvedUpload.source.startsWith("chat-upload://");
    const stale = Date.now() - resolvedUpload.loadedAt >= ttl;

    // Graph/SharePoint session workbooks must refresh when TTL expires —
    // otherwise edits in Excel (e.g. flipping Event To Home ×) never appear.
    if (!isChatUpload && stale) {
      try {
        const { ensureActiveWorkbookInMemory } = await import(
          "../eventSessionLazy"
        ).then((mod) => mod.importEventSession());
        const refreshed = await ensureActiveWorkbookInMemory(
          conversationId || "",
          undefined,
          { force: true }
        );
        if (refreshed?.sheets?.length) {
          setWorkbookCache(refreshed);
          clearQueryCaches();
          const ragCfg = getRagConfig();
          try {
            await ensureWorkbookIndexed(
              refreshed.sheets,
              {
                source: refreshed.source,
                loadedAt: refreshed.loadedAt,
                documentId: options.documentId,
              },
              ragCfg
            );
          } catch (indexError) {
            console.warn(
              `RAG index failed after refresh, falling back to lexical: ${
                indexError instanceof Error ? indexError.message : String(indexError)
              }`
            );
          }
          return refreshed.sheets;
        }
      } catch (refreshError) {
        console.warn(
          `Stale conversation workbook refresh failed, using cached sheets: ${
            refreshError instanceof Error ? refreshError.message : String(refreshError)
          }`
        );
      }
    }

    if (!forceRefresh || isChatUpload) {
      setWorkbookCache(resolvedUpload);
      const ragConfig = getRagConfig();
      try {
        await ensureWorkbookIndexed(
          resolvedUpload.sheets,
          {
            source: resolvedUpload.source,
            loadedAt: resolvedUpload.loadedAt,
            documentId: options.documentId,
          },
          ragConfig
        );
      } catch (error) {
        console.warn(
          `RAG index failed, falling back to lexical: ${
            error instanceof Error ? error.message : String(error)
          }`
        );
      }
      return resolvedUpload.sheets;
    }
  }

  if (isEventSessionMode()) {
    const ingest = (process.env.EVENTS_INGEST_MODE || "upload").trim().toLowerCase();
    if (ingest === "sharepoint" || ingest === "url") {
      throw new Error(
        "No active event for this chat. Paste a SharePoint Excel URL, then send /start. " +
          "After the event, send /end (or endevent) to clear the data."
      );
    }
    throw new Error(
      "No active event for this chat. Upload an .xlsx (paperclip → Upload from this device), then send /start. " +
        "After the event, send /end (or endevent) to clear the data."
    );
  }

  const source = getEventsSource();
  const cached = getWorkbookCache();

  if (
    !forceRefresh &&
    cached &&
    cached.sourceType === source &&
    Date.now() - cached.loadedAt < ttl
  ) {
    return cached.sheets;
  }

  const loaded = source === "graph" ? await loadGraphWorkbook() : loadLocalWorkbook();
  setWorkbookCache(loaded);

  const ragConfig = getRagConfig();
  try {
    await ensureWorkbookIndexed(
      loaded.sheets,
      {
        source: loaded.source,
        loadedAt: loaded.loadedAt,
        documentId: options.documentId,
      },
      ragConfig
    );
  } catch (error) {
    console.warn(
      `RAG index failed, falling back to lexical: ${
        error instanceof Error ? error.message : String(error)
      }`
    );
  }

  return loaded.sheets;
}
