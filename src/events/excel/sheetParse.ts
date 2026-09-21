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

function isHeaderLabel(cell: string): boolean {
  const value = cell.trim();
  if (!value || value.length > 48) {
    return false;
  }
  // Times, durations, and plain numbers are data, not column titles.
  if (/^\d/.test(value) || /mins?\b/i.test(value) || value.includes("~") || value.includes(":")) {
    return false;
  }
  return (
    /^(no\.?|name|member name|members|team leader name|date|start time|end time|detail|description|location|remark|period|time|table|beverage|role|gps point|dev)$/i.test(
      value
    ) ||
    value === "ကြာချိန်" ||
    /volunteer|reference|participate/i.test(value)
  );
}

function looksLikeHeaderRow(row: unknown[]): boolean {
  const cells = (row || []).map((cell) => cellToString(cell)).filter(Boolean);
  if (cells.length < 2) {
    return false;
  }
  const labels = cells.filter(isHeaderLabel);
  const first = cellToString((row || [])[0]).toLocaleLowerCase("my");
  if ((first === "no." || first === "no") && labels.length >= 2) {
    return true;
  }
  return labels.length >= 3;
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

function nonEmptyCells(row: unknown[]): Array<{ index: number; text: string }> {
  const out: Array<{ index: number; text: string }> = [];
  (row || []).forEach((cell, index) => {
    const text = cellToString(cell);
    if (text) {
      out.push({ index, text });
    }
  });
  return out;
}

/** True when a header row repeats the same field labels in side-by-side blocks. */
function findRepeatedFieldBlocks(
  headerRow: unknown[]
): Array<{ start: number; end: number; fields: string[] }> | null {
  const width = Math.max((headerRow || []).length, 1);
  const blocks: Array<{ start: number; end: number; fields: string[] }> = [];
  let cursor = 0;

  while (cursor < width) {
    while (cursor < width && !cellToString(headerRow[cursor])) {
      cursor += 1;
    }
    if (cursor >= width) {
      break;
    }
    const start = cursor;
    const fields: string[] = [];
    while (cursor < width) {
      const text = cellToString(headerRow[cursor]);
      if (!text) {
        break;
      }
      fields.push(text);
      cursor += 1;
    }
    if (fields.length) {
      blocks.push({ start, end: cursor, fields });
    }
  }

  if (blocks.length < 2) {
    return null;
  }

  const signatures = blocks.map((block) =>
    block.fields.map((field) => field.toLocaleLowerCase("my")).join("|")
  );
  const first = signatures[0];
  const matching = signatures.filter((signature) => signature === first).length;
  if (matching < 2) {
    return null;
  }
  if (!blocks[0].fields.some((field) => isHeaderLabel(field))) {
    return null;
  }
  return blocks;
}

function sectionTitleForBlock(
  titleRow: unknown[] | undefined,
  blockStart: number,
  blockEnd: number,
  previousEnd: number
): string {
  if (!titleRow) {
    return "";
  }
  const titled = nonEmptyCells(titleRow).filter(
    (cell) => cell.index >= previousEnd && cell.index < blockEnd
  );
  if (titled.length) {
    return titled[0].text;
  }
  const anyInBlock = nonEmptyCells(titleRow).find(
    (cell) => cell.index >= blockStart && cell.index < blockEnd
  );
  return anyInBlock?.text || "";
}

/**
 * Side-by-side subtables: a sparse group-title row, then a repeated Date/Members/... header row.
 * Emits one record per group block per data row.
 */
function matrixToSectionedRecords(values: unknown[][]): EventRecord[] | null {
  if (values.length < 3) {
    return null;
  }

  for (let headerIndex = 1; headerIndex < Math.min(values.length, 6); headerIndex += 1) {
    const blocks = findRepeatedFieldBlocks(values[headerIndex] || []);
    if (!blocks) {
      continue;
    }

    const titleRow = values[headerIndex - 1];
    const titleCells = nonEmptyCells(titleRow || []);
    if (titleCells.length < 2) {
      continue;
    }

    const records: EventRecord[] = [];
    let emptyStreak = 0;

    for (let r = headerIndex + 1; r < values.length; r += 1) {
      const row = values[r] || [];
      if (!rowHasValue(row)) {
        emptyStreak += 1;
        if (emptyStreak >= 8) {
          break;
        }
        continue;
      }
      emptyStreak = 0;

      let previousEnd = 0;
      for (const block of blocks) {
        const group = sectionTitleForBlock(titleRow, block.start, block.end, previousEnd);
        previousEnd = block.end;
        const record: EventRecord = {};
        if (group) {
          record.Group = group;
        }
        let hasValue = false;
        block.fields.forEach((field, offset) => {
          const text = cellToString(row[block.start + offset]);
          if (text) {
            record[field] = text;
            hasValue = true;
          }
        });
        if (hasValue) {
          records.push(record);
        }
      }
    }

    return records.length ? records : null;
  }

  return null;
}

/**
 * Parallel independent columns (dance groups, purchase lists): one header cell per group,
 * values underneath are separate facts — never mash sibling columns into one row.
 */
function matrixToColumnListRecords(values: unknown[][]): EventRecord[] | null {
  if (values.length < 2) {
    return null;
  }

  let headerIndex = values.findIndex((row) => rowHasValue(row));
  if (headerIndex < 0) {
    return null;
  }

  // Prefer a later classic header row when present (normal tables).
  const classicHeader = values.findIndex((row) => looksLikeHeaderRow(row));
  if (classicHeader >= 0 && findRepeatedFieldBlocks(values[classicHeader] || [])) {
    return null;
  }
  if (classicHeader === headerIndex) {
    return null;
  }
  if (classicHeader > headerIndex) {
    // Title row + classic single table — not a parallel list.
    const titleCount = nonEmptyCells(values[headerIndex] || []).length;
    if (titleCount <= 1) {
      return null;
    }
  }

  const headerRow = values[headerIndex] || [];
  const headers = nonEmptyCells(headerRow);
  if (headers.length < 2) {
    return null;
  }

  // Classic single-entity tables keep No./Name style headers in one row.
  const classicLabels = headers.filter((cell) => isHeaderLabel(cell.text)).length;
  if (classicLabels >= 2) {
    return null;
  }

  // Build column ranges: each named header owns cells until the next named header.
  const columns = headers.map((header, index) => {
    const next = headers[index + 1];
    return {
      group: header.text,
      start: header.index,
      end: next ? next.index : Math.max(headerRow.length, header.index + 1),
    };
  });

  const records: EventRecord[] = [];
  let emptyStreak = 0;

  for (let r = headerIndex + 1; r < values.length; r += 1) {
    const row = values[r] || [];
    if (!rowHasValue(row)) {
      emptyStreak += 1;
      if (emptyStreak >= 8) {
        break;
      }
      continue;
    }
    emptyStreak = 0;

    for (const column of columns) {
      const valuesInColumn: string[] = [];
      for (let c = column.start; c < column.end; c += 1) {
        const text = cellToString(row[c]);
        if (text) {
          valuesInColumn.push(text);
        }
      }
      if (!valuesInColumn.length) {
        continue;
      }

      // First cell is the primary value; extra adjacent cells become Detail/Days/etc.
      const record: EventRecord = {
        Group: column.group,
        Value: valuesInColumn[0],
      };
      if (valuesInColumn.length > 1) {
        record.Detail = valuesInColumn.slice(1).join(" | ");
      }
      records.push(record);
    }
  }

  return records.length >= headers.length ? records : null;
}

function matrixToFlatRecords(values: unknown[][]): EventRecord[] {
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

/** Generic matrix → records for any workbook layout. */
export function matrixToRecords(values: unknown[][]): EventRecord[] {
  if (!values?.length) {
    return [];
  }

  const sectioned = matrixToSectionedRecords(values);
  if (sectioned) {
    return sectioned;
  }

  const columnLists = matrixToColumnListRecords(values);
  if (columnLists) {
    return columnLists;
  }

  return matrixToFlatRecords(values);
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
