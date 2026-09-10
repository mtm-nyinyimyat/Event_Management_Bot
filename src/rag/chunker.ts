import type { SheetData } from "../events/excelStore";
import type { RagChunk } from "./types";

/** Skip parser/internal noise keys; keep every real column from the workbook. */
function isNoiseKey(key: string): boolean {
  if (!key || key.startsWith("__")) {
    return true;
  }
  if (key.startsWith("Column_") || key.startsWith("__EMPTY")) {
    return true;
  }
  return false;
}

/**
 * Discover column order from the sheet rows as they appear in the file.
 * First-seen key order is preserved so chunk text stays stable across rows.
 */
export function fieldOrderFromRows(rows: Array<Record<string, string>>): string[] {
  const order: string[] = [];
  const seen = new Set<string>();

  for (const row of rows) {
    for (const key of Object.keys(row)) {
      if (isNoiseKey(key) || seen.has(key)) {
        continue;
      }
      seen.add(key);
      order.push(key);
    }
  }

  return order;
}

/** Build chunk body from only the columns present on this row (dynamic headers). */
export function fieldsFromRow(
  row: Record<string, string>,
  fieldOrder?: string[]
): string {
  const keys = fieldOrder?.length
    ? fieldOrder
    : Object.keys(row).filter((key) => !isNoiseKey(key));

  const parts: string[] = [];
  for (const key of keys) {
    if (isNoiseKey(key)) {
      continue;
    }
    const value = row[key]?.trim();
    if (!value) {
      continue;
    }
    parts.push(`${key}: ${value}`);
  }

  // Include any extra keys on this row not yet covered (sparse / irregular sheets)
  if (fieldOrder?.length) {
    for (const [key, value] of Object.entries(row)) {
      if (isNoiseKey(key) || fieldOrder.includes(key) || !value?.trim()) {
        continue;
      }
      parts.push(`${key}: ${value.trim()}`);
    }
  }

  return parts.join(" | ");
}

/** Keep all non-empty workbook fields for returned RAG hits. */
export function dynamicRow(row: Record<string, string>): Record<string, string> {
  if (row.__chunk_type === "summary") {
    return { Content: row.Content || "" };
  }

  const out: Record<string, string> = {};
  for (const [key, value] of Object.entries(row)) {
    if (isNoiseKey(key) || !value?.trim()) {
      continue;
    }
    out[key] = value.trim();
  }
  return out;
}

function summaryToText(sheet: string, summary: SheetData["summary"]): string | null {
  if (!summary || Object.keys(summary).length === 0) {
    return null;
  }

  const parts = Object.entries(summary).map(([key, value]) => {
    if (typeof value === "object" && value !== null) {
      const nested = Object.entries(value)
        .map(([k, v]) => `${k}=${v}`)
        .join(", ");
      return `${key}: ${nested}`;
    }
    return `${key}: ${value}`;
  });

  return `Sheet: ${sheet} | Summary: ${parts.join(" | ")}`;
}

/**
 * Turn workbook sheets into retrieval chunks (one per row + optional summary chunk).
 * Column names come from the file rows, not a fixed schema.
 */
export function chunkWorkbookSheets(sheets: SheetData[]): RagChunk[] {
  const chunks: RagChunk[] = [];

  for (const sheet of sheets) {
    const summaryText = summaryToText(sheet.sheet, sheet.summary);
    if (summaryText) {
      chunks.push({
        id: `${sheet.sheet}::summary`,
        sheet: sheet.sheet,
        text: summaryText,
        row: { __chunk_type: "summary", Content: summaryText },
      });
    }

    const fieldOrder = fieldOrderFromRows(sheet.rows);

    sheet.rows.forEach((row, index) => {
      const body = fieldsFromRow(row, fieldOrder);
      if (!body) {
        return;
      }
      chunks.push({
        id: `${sheet.sheet}::row::${index}`,
        sheet: sheet.sheet,
        text: `Sheet: ${sheet.sheet} | ${body}`,
        row: dynamicRow(row),
      });
    });
  }

  return chunks;
}
