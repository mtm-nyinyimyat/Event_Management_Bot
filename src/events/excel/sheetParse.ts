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
  // Allow "Estimate time" / "Start Time" style headers (words + time) before the ":" check.
  const normalized = value.replace(/:$/, "").trim();
  if (
    /^(date|day|place|venue|location|estimate\s*time|estimated?\s*time|duration|start\s*time|end\s*time|time|period|ကြာချိန်|ရက်စွဲ|နေရာ)$/i.test(
      normalized
    )
  ) {
    return true;
  }
  if (/^\d/.test(value) || /mins?\b/i.test(value) || value.includes("~") || value.includes(":")) {
    return false;
  }
  return (
    /^(no\.?|name|member name|members|team leader name|detail|description|remark|table|beverage|drink|role|gps point|dev|event participate|mtm to event|event to home|incharge contact number|status|email|phone|department|team|group|song|food|wine|tea|water|seat|section)$/i.test(
      value
    ) ||
    value === "ကြာချိန်" ||
    /volunteer|reference|participate/i.test(value)
  );
}

const EVENT_META_LABEL =
  /^(date|day|place|venue|location|estimate\s*time|estimated?\s*time|duration|ကြာချိန်|ရက်စွဲ|နေရာ)$/i;

function metaFieldKey(label: string): string {
  const normalized = label.replace(/:$/, "").trim().toLocaleLowerCase("en");
  if (/^(date|day)\b/.test(normalized) || normalized.includes("ရက်")) {
    return "Date";
  }
  if (/place|venue|location/.test(normalized) || normalized.includes("နေရာ")) {
    return "Place";
  }
  if (/estimate|duration/.test(normalized) || normalized.includes("ကြာချိန်")) {
    return "Estimate time";
  }
  if (/^estimate\s*time$|^estimated?\s*time$/.test(normalized)) {
    return "Estimate time";
  }
  return label.replace(/:$/, "").trim();
}

/**
 * Sheet-level event facts often sit beside/above the Participants table as
 * label|value pairs (Date / Place / Estimate time), not as per-person columns.
 */
function extractEventMetaRecords(values: unknown[][]): EventRecord[] {
  const records: EventRecord[] = [];
  const seen = new Set<string>();

  const pushMeta = (label: string, value: string) => {
    const field = metaFieldKey(label);
    const text = value.trim();
    if (!text || EVENT_META_LABEL.test(text.replace(/:$/, ""))) {
      return;
    }
    const key = `${field}::${text}`;
    if (seen.has(key)) {
      return;
    }
    seen.add(key);
    records.push({
      Section: "Event",
      Field: field,
      [field]: text,
      Value: text,
    });
  };

  for (let r = 0; r < Math.min(values.length, 40); r += 1) {
    const row = values[r] || [];
    for (let c = 0; c < row.length; c += 1) {
      const raw = cellToString(row[c]);
      if (!raw) {
        continue;
      }
      const inline = raw.match(/^([^:]{2,40}):\s*(.+)$/);
      if (inline && EVENT_META_LABEL.test(inline[1].trim())) {
        pushMeta(inline[1], inline[2]);
        continue;
      }
      const label = raw.replace(/:$/, "").trim();
      if (!EVENT_META_LABEL.test(label)) {
        continue;
      }
      const right = cellToString(row[c + 1]);
      if (right) {
        pushMeta(label, right);
        continue;
      }
      // Value sometimes sits one row below the label.
      const below = cellToString((values[r + 1] || [])[c]);
      if (below && !EVENT_META_LABEL.test(below.replace(/:$/, ""))) {
        pushMeta(label, below);
      }
    }
  }
  return records;
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

  // Title/banner row + classic No/Name/Location header → real table(s), not parallel lists.
  const classicHeader = values.findIndex((row) => looksLikeHeaderRow(row));
  if (classicHeader >= 0 && classicHeader !== headerIndex) {
    return null;
  }
  if (classicHeader >= 0 && findRepeatedFieldBlocks(values[classicHeader] || [])) {
    return null;
  }
  if (classicHeader === headerIndex) {
    return null;
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

function isSectionTitleRow(row: unknown[], nextRow: unknown[] | undefined): boolean {
  if (!rowHasValue(row) || looksLikeHeaderRow(row)) {
    return false;
  }
  if (!nextRow || !looksLikeHeaderRow(nextRow)) {
    return false;
  }
  const first = cellToString(row[0]);
  return /ferry|list|driver|route|အပြန်|ဖယ်ရီ/i.test(first) || nonEmptyCells(row).length <= 4;
}

/**
 * Vertically stacked tables (ferry lists): title → No/Name/Location header → rows, repeated.
 * Keeps Location/Gps Point on the same record as the passenger name.
 */
function matrixToStackedTableRecords(values: unknown[][]): EventRecord[] | null {
  const headerIndexes: number[] = [];
  for (let i = 0; i < values.length; i += 1) {
    if (looksLikeHeaderRow(values[i] || [])) {
      headerIndexes.push(i);
    }
  }
  if (!headerIndexes.length) {
    return null;
  }

  const firstHeader = headerIndexes[0];
  const hasTitleBefore =
    firstHeader > 0 &&
    rowHasValue(values[firstHeader - 1] || []) &&
    !looksLikeHeaderRow(values[firstHeader - 1] || []);
  if (headerIndexes.length < 2 && !hasTitleBefore) {
    return null;
  }

  const records: EventRecord[] = [];

  for (let h = 0; h < headerIndexes.length; h += 1) {
    const headerIndex = headerIndexes[h];
    const titleRow = headerIndex > 0 ? values[headerIndex - 1] || [] : [];
    const titleCells = nonEmptyCells(titleRow);
    const titleBlob = titleCells.map((cell) => cell.text).join(" ");
    const looksLikeFerryTitle =
      (/ferry|driver|route|အပြန်|ဖယ်ရီ|ကား/i.test(titleBlob) ||
        (/list/i.test(titleBlob) && !/volunteer|ကူညီ|ဝန်ထမ်း|notice|ကြေညာ/i.test(titleBlob))) &&
      !/volunteer|ကူညီ|ဝန်ထမ်း|notice|ကြေညာ/i.test(titleBlob);
    const ferry = looksLikeFerryTitle
      ? titleCells.find((cell) => /ferry|list|အပြန်|ဖယ်ရီ/i.test(cell.text))?.text ||
        titleCells[0]?.text ||
        ""
      : "";
    const driver = looksLikeFerryTitle
      ? titleCells.find((cell) => /driver/i.test(cell.text))?.text ||
        (titleCells.length > 1 ? titleCells[1].text : "")
      : "";
    const notice =
      !looksLikeFerryTitle && titleCells[0]?.text ? titleCells[0].text : "";

    const width = primaryTableWidth(values[headerIndex] || []);
    const headers = uniqueHeaders(
      (values[headerIndex] || []).slice(0, width).map((cell) => cellToString(cell))
    );
    let emptyStreak = 0;

    for (let r = headerIndex + 1; r < values.length; r += 1) {
      const row = values[r] || [];
      const next = values[r + 1];
      if (looksLikeHeaderRow(row) || isSectionTitleRow(row, next)) {
        break;
      }
      if (!row.slice(0, width).some((cell) => cellToString(cell))) {
        emptyStreak += 1;
        if (emptyStreak >= 3) {
          break;
        }
        continue;
      }
      emptyStreak = 0;

      const record: EventRecord = {};
      if (ferry) {
        record.Ferry = ferry;
      }
      if (driver) {
        record.Driver = driver;
      }
      if (notice) {
        record.Notice = notice;
      }
      headers.forEach((header, column) => {
        if (!header || header.startsWith("__EMPTY") || header.startsWith("Column_")) {
          return;
        }
        const text = cellToString(row[column]);
        if (text) {
          record[header] = text;
        }
      });

      // Skip bogus header-echo rows
      const name = record.Name || record.NAME || "";
      const no = record["No."] || record.No || "";
      if (/^name$/i.test(name) || /^no\.?$/i.test(no)) {
        continue;
      }

      if (Object.keys(record).some((key) => !["Ferry", "Driver"].includes(key))) {
        records.push(record);
      }
    }
  }

  return records.length ? records : null;
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

function isDrinkOrMenuLabel(text: string): boolean {
  return /^(juice|beer|cocktail|cocktial|stage|guest-\d+)$/i.test(text.trim()) ||
    /menu|appetizer|salad|soup|dessert|main course|dinner/i.test(text) ||
    text.startsWith("※") ||
    text.startsWith("အာလူး") ||
    text.includes("သုပ်") ||
    text.includes("ဟင်း") ||
    text.includes("ကြော်") ||
    text.includes("ပေါင်း") ||
    text.includes("ထမင်း") ||
    text.includes("အသီး");
}

function looksLikeSeatPersonName(text: string): boolean {
  const value = text.trim();
  if (!value || value.length > 60) {
    return false;
  }
  if (isSeatTableLabel(value)) {
    return false;
  }
  if (isDrinkOrMenuLabel(value)) {
    return false;
  }
  if (/^\d+(\.\d+)?$/.test(value)) {
    return false;
  }
  // Latin, Myanmar, or Japanese person names (reject pure punctuation).
  if (!/[A-Za-z\u1000-\u109F\u3040-\u30FF\u4E00-\u9FFF]/.test(value)) {
    return false;
  }
  const tokens = value.split(/\s+/);
  return tokens.length >= 1 && tokens.length <= 6;
}

/** Table-1, Table 1, Seat-3, 卓1, VIP — not locked to one event's naming. */
function isSeatTableLabel(text: string): boolean {
  const value = text.trim();
  if (/^vip$/i.test(value)) {
    return true;
  }
  return /^(table|seat|卓)\s*[-_]?\s*\d+$/i.test(value) || /^table-\d+$/i.test(value);
}

/**
 * Visual seating charts: each Table-N/VIP owns names from the row below it
 * down to (not including) the next label in the same column — never a fixed
 * ±N window (that pulls people from the table above into the table below).
 */
function matrixToSeatingChartRecords(values: unknown[][]): EventRecord[] | null {
  const labels: Array<{ row: number; col: number; table: string }> = [];
  const width = values.reduce((max, row) => Math.max(max, (row || []).length), 0);

  for (let r = 0; r < values.length; r += 1) {
    const row = values[r] || [];
    for (let c = 0; c < row.length; c += 1) {
      const text = cellToString(row[c]);
      if (isSeatTableLabel(text)) {
        labels.push({ row: r, col: c, table: text.trim() });
      }
    }
  }

  if (labels.length < 2) {
    return null;
  }

  const records: EventRecord[] = [];
  const seen = new Set<string>();
  const personClaimed = new Set<string>();

  for (const label of labels) {
    const sameRowRight = labels
      .filter((other) => other.row === label.row && other.col > label.col)
      .sort((a, b) => a.col - b.col)[0];
    const colEnd = sameRowRight
      ? sameRowRight.col - 1
      : Math.min(label.col + 5, width - 1);

    const belowCandidates = labels.filter(
      (other) =>
        other.row > label.row && other.col <= colEnd && other.col + 5 >= label.col
    );
    const sameColBelow = belowCandidates
      .filter((other) => other.col === label.col)
      .sort((a, b) => a.row - b.row)[0];
    const nextBelow =
      sameColBelow ||
      belowCandidates.sort((a, b) => a.row - b.row || a.col - b.col)[0];

    // Include the label row (names often sit beside Table-N) and rows below until the next label.
    const rowStart = label.row;
    const rowEnd = nextBelow ? nextBelow.row - 1 : values.length - 1;
    if (rowStart > rowEnd) {
      continue;
    }

    for (let r = rowStart; r <= rowEnd; r += 1) {
      const row = values[r] || [];
      for (let c = label.col; c <= colEnd; c += 1) {
        const text = cellToString(row[c]);
        if (isSeatTableLabel(text) || !looksLikeSeatPersonName(text)) {
          continue;
        }
        const personKey = text.toLocaleLowerCase("en");
        if (personClaimed.has(personKey)) {
          continue;
        }
        const key = `${label.table}::${text}`;
        if (seen.has(key)) {
          continue;
        }
        seen.add(key);
        personClaimed.add(personKey);
        records.push({ Table: label.table, Name: text });
      }
    }
  }

  let menuCol = -1;
  for (let r = 0; r < Math.min(values.length, 5); r += 1) {
    const row = values[r] || [];
    row.forEach((cell, c) => {
      if (/dinner menu/i.test(cellToString(cell))) {
        menuCol = c;
      }
    });
  }
  if (menuCol >= 0) {
    let category = "Menu";
    for (let r = 0; r < values.length; r += 1) {
      const text = cellToString((values[r] || [])[menuCol]);
      if (!text) {
        continue;
      }
      if (/:$/.test(text) || /appetizer|salad|soup|main course|dessert|menu/i.test(text)) {
        category = text.replace(/:$/, "").trim();
        continue;
      }
      records.push({ Section: "Menu", Category: category, Dish: text });
    }
  }

  return records.length >= 8 ? records : null;
}

const MENU_CATEGORY =
  /^(appetizer|salad|soup|main\s*course|dessert|starter|entree|entrée|side|drink|beverage)s?$/i;

/** Category in col A (often once per group) + dishes in col B+. */
function matrixToCourseMenuRecords(values: unknown[][]): EventRecord[] | null {
  if (values.length < 2) {
    return null;
  }

  let titleIndex = -1;
  for (let r = 0; r < Math.min(values.length, 5); r += 1) {
    const cells = (values[r] || []).map(cellToString).filter(Boolean);
    if (
      (cells.length === 1 && /^(dinner\s*)?menu$/i.test(cells[0])) ||
      cells.some((cell) => /^dinner\s*menu$/i.test(cell))
    ) {
      titleIndex = r;
      break;
    }
  }

  let categoryHits = 0;
  let dishHits = 0;
  for (let r = Math.max(0, titleIndex + 1); r < values.length; r += 1) {
    const row = values[r] || [];
    const left = cellToString(row[0]);
    const right = cellToString(row[1]) || cellToString(row[2]);
    if (MENU_CATEGORY.test(left)) {
      categoryHits += 1;
    }
    if (right && !MENU_CATEGORY.test(right) && !/^(dinner\s*)?menu$/i.test(right)) {
      dishHits += 1;
    }
  }
  if (titleIndex < 0 && categoryHits < 2) {
    return null;
  }
  if (dishHits < 2) {
    return null;
  }

  const records: EventRecord[] = [];
  let category = "Menu";
  for (let r = titleIndex >= 0 ? titleIndex + 1 : 0; r < values.length; r += 1) {
    const row = values[r] || [];
    const left = cellToString(row[0]);
    const dishes: string[] = [];
    for (let c = 1; c < row.length; c += 1) {
      const text = cellToString(row[c]);
      if (text) {
        dishes.push(text);
      }
    }
    if (!dishes.length && left && !MENU_CATEGORY.test(left)) {
      dishes.push(left);
    } else if (left && MENU_CATEGORY.test(left)) {
      category = left.replace(/\s+/g, " ").trim();
    } else if (left && !dishes.length) {
      continue;
    }
    for (const dish of dishes) {
      if (!dish || MENU_CATEGORY.test(dish) || /^(dinner\s*)?menu$/i.test(dish)) {
        continue;
      }
      records.push({ Section: "Menu", Category: category, Dish: dish });
    }
  }
  return records.length >= 2 ? records : null;
}

function looksLikeUnlabeledSeatingGrid(values: unknown[][]): boolean {
  let personCells = 0;
  let otherCells = 0;
  let tableLabels = 0;
  for (let r = 0; r < Math.min(values.length, 40); r += 1) {
    for (const cell of values[r] || []) {
      const text = cellToString(cell);
      if (!text) {
        continue;
      }
      if (isSeatTableLabel(text)) {
        tableLabels += 1;
        continue;
      }
      if (looksLikeSeatPersonName(text)) {
        personCells += 1;
      } else {
        otherCells += 1;
      }
    }
  }
  return tableLabels === 0 && personCells >= 12 && personCells > otherCells * 2;
}

/** Generic matrix → records for any workbook layout. */
export function matrixToRecords(values: unknown[][]): EventRecord[] {
  if (!values?.length) {
    return [];
  }

  const eventMeta = extractEventMetaRecords(values);

  const seating = matrixToSeatingChartRecords(values);
  if (seating) {
    return eventMeta.length ? [...eventMeta, ...seating] : seating;
  }

  const courseMenu = matrixToCourseMenuRecords(values);
  if (courseMenu) {
    return eventMeta.length ? [...eventMeta, ...courseMenu] : courseMenu;
  }

  // Try structured parsers before dropping name-heavy sheets (ferry routes, rosters).
  // Unlabeled seating grids are only suppressed when nothing structured matched.
  const sectioned = matrixToSectionedRecords(values);
  if (sectioned) {
    return eventMeta.length ? [...eventMeta, ...sectioned] : sectioned;
  }

  const stacked = matrixToStackedTableRecords(values);
  if (stacked) {
    return eventMeta.length ? [...eventMeta, ...stacked] : stacked;
  }

  const columnLists = matrixToColumnListRecords(values);
  if (columnLists) {
    return eventMeta.length ? [...eventMeta, ...columnLists] : columnLists;
  }

  const flat = matrixToFlatRecords(values);
  if (looksLikeUnlabeledSeatingGrid(values)) {
    // Keep structured people tables (No./Name); only drop pure visual name grids.
    const hasPeopleTable = flat.some(
      (row) =>
        Boolean(row.Name || row.NAME || row["Member name"] || row.Members) &&
        Boolean(row["No."] || row.No || row.Beverage || row.Description)
    );
    if (!hasPeopleTable) {
      return eventMeta;
    }
  }

  if (!eventMeta.length) {
    return flat;
  }
  // Prefer structured Event meta over raw Notice joins for the same facts.
  return [...eventMeta, ...flat];
}

const CHECKMARK_TOKENS = /[〇○●◯◎✓✔]/u;
const CROSS_TOKENS = /[×✗✘]/u;

export function isCheckedMark(value: string | undefined): boolean {
  const cell = (value || "").trim();
  if (!cell) {
    return false;
  }
  return CHECKMARK_TOKENS.test(cell) || /^(yes|y|true|1|v)$/i.test(cell);
}

function countCheckmarks(rows: EventRecord[], column: string): { checked: number; total: number } {
  let checked = 0;
  let total = 0;
  for (const row of rows) {
    const value = (row[column] || "").trim();
    if (!value) {
      continue;
    }
    if (isCheckedMark(value)) {
      checked += 1;
      total += 1;
    } else if (CROSS_TOKENS.test(value) || /^(no|n|false|0|x)$/i.test(value)) {
      total += 1;
    }
  }
  return { checked, total };
}

export function buildSheetSummary(sheetName: string, rows: EventRecord[]): SheetSummary {
  return defaultSummaryForRows(sheetName, rows);
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

  const checkmarkColumns: Record<string, { checked: number; total: number }> = {};
  for (const column of columns) {
    const counts = countCheckmarks(rows, column);
    if (counts.total >= 2) {
      checkmarkColumns[column] = counts;
    }
  }
  if (Object.keys(checkmarkColumns).length) {
    summary.checkmark_columns = checkmarkColumns;
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
  options?: { summaryOnly?: boolean; omitSummary?: boolean; keepCheckmarkColumns?: boolean }
): SheetData {
  const fieldOrder = fieldOrderFromRows(rows.length ? rows : sheet.rows);
  const compactRows = options?.summaryOnly
    ? rows.slice(0, 5).map((row) => compactRow(row, fieldOrder))
    : rows.map((row) => compactRow(row, fieldOrder));

  let summary: SheetSummary | undefined;
  if (options?.omitSummary) {
    if (options?.keepCheckmarkColumns && sheet.summary?.checkmark_columns) {
      summary = { checkmark_columns: sheet.summary.checkmark_columns };
    }
  } else {
    summary = sheet.summary;
  }

  return {
    sheet: sheet.sheet,
    numbered_item_count: options?.omitSummary
      ? compactRows.length
      : sheet.numbered_item_count ?? rows.filter((row) => isNumbered(row["No."] || row.No)).length,
    summary,
    rows: compactRows,
  };
}
