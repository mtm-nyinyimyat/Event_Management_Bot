import type { EventRecord, SheetData, SheetSummary } from "./types";

function cellToString(value: unknown): string {
  if (value === null || value === undefined) {
    return "";
  }
  return String(value).replace(/\r\n/g, "\n").trim();
}

function uniqueHeaders(headers: string[]): string[] {
  const used = new Map<string, number>();
  return headers.map((header, index) => {
    const base = header || `Column_${index + 1}`;
    const count = used.get(base) || 0;
    used.set(base, count + 1);
    return count === 0 ? base : `${base}_${count + 1}`;
  });
}

function rowHasValue(row: unknown[]): boolean {
  return (row || []).some((cell) => cellToString(cell));
}

function looksLikeHeaderRow(row: unknown[]): boolean {
  const cells = (row || []).map((cell) => cellToString(cell).toLocaleLowerCase("my"));
  const nonEmpty = cells.filter(Boolean);
  if (nonEmpty.length < 2) {
    return false;
  }
  // Prefer rows that look like labels (short tokens, shared header words)
  const headerish = nonEmpty.filter(
    (cell) =>
      cell.length <= 40 &&
      !/^\d+(\.\d+)?$/.test(cell) &&
      (/\b(no\.?|name|time|detail|description|location|table|ferry|beverage|remark)\b/i.test(
        cell
      ) ||
        cell.split(/\s+/).length <= 4)
  ).length;
  return headerish >= 2 || cells[0] === "no." || cells[0] === "no";
}

function isNumbered(value: string | undefined): boolean {
  return /^\d+$/.test((value || "").trim());
}

function primaryTableWidth(headerRow: unknown[]): number {
  let lastNamed = -1;
  (headerRow || []).forEach((cell, index) => {
    if (cellToString(cell)) {
      lastNamed = index;
    }
  });
  return Math.max(lastNamed + 1, 1);
}

function isNoiseColumnKey(key: string): boolean {
  if (!key || key.startsWith("__")) {
    return true;
  }
  if (key.startsWith("Column_") || key.startsWith("__EMPTY")) {
    return true;
  }
  return false;
}

function fieldOrderFromRows(rows: EventRecord[]): string[] {
  const order: string[] = [];
  const seen = new Set<string>();
  for (const row of rows) {
    for (const key of Object.keys(row)) {
      if (isNoiseColumnKey(key) || seen.has(key)) {
        continue;
      }
      seen.add(key);
      order.push(key);
    }
  }
  return order;
}

/** Generic matrix → records for any workbook layout. */
export function matrixToRecords(values: unknown[][]): EventRecord[] {
  if (!values?.length) {
    return [];
  }

  let headerIndex = values.findIndex((row) => looksLikeHeaderRow(row));
  if (headerIndex < 0) {
    headerIndex = values.findIndex((row) => rowHasValue(row));
  }
  if (headerIndex < 0) {
    return [];
  }

  const records: EventRecord[] = [];

  for (let i = 0; i < headerIndex; i += 1) {
    const text = (values[i] || []).map(cellToString).filter(Boolean).join(" ").trim();
    if (text) {
      records.push({ Section: "Notice", Content: text });
    }
  }

  const width = primaryTableWidth(values[headerIndex] || []);
  const headers = uniqueHeaders(
    (values[headerIndex] || []).slice(0, width).map((cell) => cellToString(cell))
  );
  let emptyStreak = 0;

  for (let i = headerIndex + 1; i < values.length; i += 1) {
    const row = values[i] || [];
    if (!row.slice(0, width).some((cell) => cellToString(cell))) {
      emptyStreak += 1;
      if (emptyStreak >= 8) {
        break;
      }
      continue;
    }
    emptyStreak = 0;

    if (looksLikeHeaderRow(row)) {
      continue;
    }

    const record: EventRecord = {};
    headers.forEach((header, column) => {
      if (!header || header.startsWith("__EMPTY") || header.startsWith("Column_")) {
        return;
      }
      const text = cellToString(row[column]);
      if (text) {
        record[header] = text;
      }
    });

    if (Object.keys(record).length > 0) {
      records.push(record);
    }
  }

  return records;
}

function defaultSummaryForRows(sheetName: string, rows: EventRecord[]): SheetSummary {
  const numbered = rows.filter((row) => isNumbered(row["No."] || row.No));
  const columns = fieldOrderFromRows(rows).slice(0, 24);
  const summary: SheetSummary = {
    total_rows: rows.length,
    numbered_item_count: numbered.length,
  };
  if (columns.length) {
    summary.columns = columns.join(", ");
  }
  // Soft hints only when sheet title suggests it — not required for retrieval
  if (/agenda|schedule|program/i.test(sheetName)) {
    summary.agenda_items = numbered.length;
  }
  if (/volunteer/i.test(sheetName)) {
    summary.volunteer_roles = numbered.length;
  }
  return summary;
}

export function parseSheet(sheetName: string, values: unknown[][]): SheetData {
  const rows = matrixToRecords(values);
  return {
    sheet: sheetName,
    rows,
    numbered_item_count: rows.filter((row) => isNumbered(row["No."] || row.No)).length,
    summary: defaultSummaryForRows(sheetName, rows),
  };
}

export function compactRow(row: EventRecord, fieldOrder?: string[]): EventRecord {
  const keys = fieldOrder?.length
    ? fieldOrder
    : Object.keys(row).filter((key) => !isNoiseColumnKey(key));

  const out: EventRecord = {};
  for (const key of keys) {
    if (isNoiseColumnKey(key)) {
      continue;
    }
    const value = row[key]?.trim();
    if (!value) {
      continue;
    }
    out[key === "NAME" ? "Name" : key] = value;
  }

  if (fieldOrder?.length) {
    for (const [key, value] of Object.entries(row)) {
      if (isNoiseColumnKey(key) || out[key] || out.Name === value || !value?.trim()) {
        continue;
      }
      out[key === "NAME" ? "Name" : key] = value.trim();
    }
  }

  return out;
}

export function withSheetStats(
  sheet: SheetData,
  rows: EventRecord[],
  options?: { summaryOnly?: boolean; omitSummary?: boolean }
): SheetData {
  const fieldOrder = fieldOrderFromRows(rows.length ? rows : sheet.rows);
  const compactRows = options?.summaryOnly
    ? rows.slice(0, 5).map((row) => compactRow(row, fieldOrder))
    : rows.map((row) => compactRow(row, fieldOrder));

  return {
    sheet: sheet.sheet,
    numbered_item_count: options?.omitSummary
      ? compactRows.length
      : sheet.numbered_item_count ?? rows.filter((row) => isNumbered(row["No."] || row.No)).length,
    summary: options?.omitSummary ? undefined : sheet.summary,
    rows: compactRows,
  };
}
