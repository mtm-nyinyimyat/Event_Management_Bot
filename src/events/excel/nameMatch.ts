import type { SheetData } from "./types";

const MYANMAR_CHAR = /[\u1000-\u109F]/;
const MYANMAR_RUN = /[\u1000-\u109F\u1040-\u1049]+/g;

/**
 * Common Myanmar name syllables → Latin variants (longest keys first).
 * Bridges Burmese questions to English NAME cells — not a prompt/hardcoded person list.
 */
const SYLLABLE_MAP: Array<[string, string[]]> = (
  [
  ["အားကာ", ["arka", "arkar"]],
  ["အာကာ", ["arka", "arkar"]],
  ["အောင်", ["aung"]],
  ["ကျော်", ["kyaw"]],
  ["ချစ်", ["chit", "kyit"]],
  ["ငွေ", ["ngwe"]],
  ["စန်း", ["san", "sann"]],
  ["စည်", ["si", "sii", "see", "syi"]],
  ["ဆွေ", ["swe", "swei"]],
  ["ညွန့်", ["nyunt"]],
  ["ထွန်း", ["htun", "tun", "htunn"]],
  ["နန္ဒာ", ["nanda", "nandah"]],
  ["နိုင်", ["naing", "nye"]],
  ["ပိုင်", ["paing", "pyae"]],
  ["မင်း", ["min", "minn", "ming"]],
  ["မြင့်", ["myint"]],
  ["မြတ်", ["myat"]],
  ["မိုး", ["moe", "mo"]],
  ["ယဉ်", ["yin", "yinn"]],
  ["ရန်", ["yan", "yann"]],
  ["လှ", ["hla", "la"]],
  ["ဝင်း", ["win", "winn"]],
  ["သန့်", ["thant", "thet"]],
  ["သူ", ["thu", "su", "thoo"]],
  ["သိန်း", ["thein"]],
  ["သီဟ", ["thiha", "thihah"]],
  ["ဇင်", ["zin", "zinn"]],
  ["အေး", ["aye", "ay"]],
  ["ဦး", ["u", "oo"]],
  ["ခန့်", ["khant"]],
  ["ခိုင်", ["khaing", "khine"]],
  ["ဇော်", ["zaw"]],
  ["ဖြူ", ["phyu", "phu"]],
  ["မာ", ["mar", "ma"]],
  ["လင်း", ["lin", "linn", "lynn"]],
  ["ရှင်း", ["shin"]],
  ["ဟိန်း", ["hein"]],
] as Array<[string, string[]]>
).sort((a, b) => b[0].length - a[0].length);

function normalizeLatin(value: string): string {
  return value
    .toLocaleLowerCase("en")
    .normalize("NFC")
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function greedyRomanize(myanmar: string): string[][] {
  const parts: string[][] = [];
  let i = 0;
  const text = myanmar.normalize("NFC");
  while (i < text.length) {
    let matched = false;
    for (const [syllable, variants] of SYLLABLE_MAP) {
      if (text.startsWith(syllable, i)) {
        parts.push(variants);
        i += syllable.length;
        matched = true;
        break;
      }
    }
    if (!matched) {
      i += 1;
    }
  }
  return parts;
}

function cartesianLimited(parts: string[][], limit = 24): string[] {
  if (!parts.length) {
    return [];
  }
  let combos = [""];
  for (const options of parts) {
    const next: string[] = [];
    for (const prefix of combos) {
      for (const option of options) {
        next.push(prefix ? `${prefix} ${option}` : option);
        if (next.length >= limit) {
          return next;
        }
      }
    }
    combos = next;
  }
  return combos;
}

function englishNameTokens(name: string): string[] {
  return normalizeLatin(name).split(/\s+/).filter(Boolean);
}

export function scoreBurmeseToEnglishName(burmese: string, english: string): number {
  const variants = cartesianLimited(greedyRomanize(burmese));
  if (!variants.length) {
    return 0;
  }
  const englishTokens = englishNameTokens(english);
  if (!englishTokens.length) {
    return 0;
  }
  const englishJoined = englishTokens.join(" ");

  let best = 0;
  for (const variant of variants) {
    const romanTokens = variant.split(/\s+/).filter(Boolean);
    if (!romanTokens.length) {
      continue;
    }
    if (englishJoined === romanTokens.join(" ")) {
      return 1;
    }
    let cursor = 0;
    let hits = 0;
    for (const token of romanTokens) {
      const found = englishTokens.findIndex((part, index) => index >= cursor && part === token);
      if (found >= 0) {
        hits += 1;
        cursor = found + 1;
      }
    }
    const score = hits / Math.max(romanTokens.length, englishTokens.length);
    if (score > best) {
      best = score;
    }
  }
  return best;
}

function levenshtein(a: string, b: string): number {
  if (a === b) {
    return 0;
  }
  if (!a.length) {
    return b.length;
  }
  if (!b.length) {
    return a.length;
  }
  const row = Array.from({ length: b.length + 1 }, (_, i) => i);
  for (let i = 1; i <= a.length; i += 1) {
    let prev = i - 1;
    row[0] = i;
    for (let j = 1; j <= b.length; j += 1) {
      const cur = row[j];
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      row[j] = Math.min(row[j] + 1, row[j - 1] + 1, prev + cost);
      prev = cur;
    }
  }
  return row[b.length];
}

/** Score two English names (handles model typos like Aung Zin Thu → Aung Si Thu). */
export function scoreEnglishNameSimilarity(left: string, right: string): number {
  const a = englishNameTokens(left);
  const b = englishNameTokens(right);
  if (!a.length || !b.length) {
    return 0;
  }
  if (a.join(" ") === b.join(" ")) {
    return 1;
  }
  // Same length token sequences with small edits (Zin vs Si)
  if (a.length === b.length) {
    let dist = 0;
    for (let i = 0; i < a.length; i += 1) {
      dist += levenshtein(a[i], b[i]);
    }
    const maxLen = Math.max(a.join("").length, b.join("").length, 1);
    return Math.max(0, 1 - dist / maxLen);
  }
  const joinedA = a.join(" ");
  const joinedB = b.join(" ");
  const dist = levenshtein(joinedA.replace(/\s+/g, ""), joinedB.replace(/\s+/g, ""));
  const maxLen = Math.max(joinedA.replace(/\s+/g, "").length, joinedB.replace(/\s+/g, "").length, 1);
  return Math.max(0, 1 - dist / maxLen);
}

function collectWorkbookPersonNames(sheets: SheetData[]): string[] {
  const names = new Set<string>();
  for (const sheet of sheets) {
    for (const row of sheet.rows) {
      for (const [key, value] of Object.entries(row)) {
        if (!value?.trim() || key.startsWith("__")) {
          continue;
        }
        if (/^(name|member name|team leader name)$/i.test(key) || key === "NAME") {
          names.add(value.trim());
          continue;
        }
        // Seating charts often store one person per cell under Table-* keys
        if (/^table/i.test(key) || /^vip$/i.test(key)) {
          const trimmed = value.trim();
          if (/[A-Za-z]/.test(trimmed) && trimmed.split(/\s+/).length <= 5) {
            names.add(trimmed);
          }
        }
        if (/^members$/i.test(key)) {
          for (const part of value.split(/[\n,;/]+/)) {
            const trimmed = part.trim();
            if (trimmed && /[A-Za-z]/.test(trimmed)) {
              names.add(trimmed);
            }
          }
        }
      }
    }
  }
  return [...names];
}

function stripBurmeseQuestionNoise(text: string): string {
  return text
    .normalize("NFC")
    .replace(/(က)?ဘာသောက်.*$/u, "")
    .replace(/ဘယ်မှာထိုင်.*$/u, "")
    .replace(/ဘယ်မှာ.*$/u, "")
    .replace(/ဘယ်သူ.*$/u, "")
    .replace(/ဘာလုပ်.*$/u, "")
    .replace(/ကိုရှာ.*$/u, "")
    .replace(/ထိုင်တာ.*$/u, "")
    .replace(/(လား|လဲ|နော်|ပါ|တယ်|တဲ့|ရဲ့|ကို|က)$/u, "")
    .trim();
}

function extractMyanmarNameCandidates(query: string): string[] {
  const stripped = stripBurmeseQuestionNoise(query);
  const runs = (stripped.match(MYANMAR_RUN) || [])
    .map((run) => run.trim())
    .filter((run) => run.length >= 3);
  if (runs.length) {
    return [...new Set(runs)];
  }
  const compact = stripped.replace(/[^\u1000-\u109F]/g, "");
  return compact.length >= 3 ? [compact] : [];
}

/** Pull likely English person-name phrases from a lookup query. */
function extractEnglishNameCandidates(query: string): string[] {
  const matches =
    query.match(/\b[A-Z][a-z]+(?:\s+[A-Z][a-z]+){1,3}\b/g) ||
    query.match(/\b[A-Za-z]+(?:\s+[A-Za-z]+){1,3}\b/g) ||
    [];
  const stop = new Set([
    "table",
    "layout",
    "ferry",
    "list",
    "seat",
    "seating",
    "where",
    "sitting",
    "drink",
    "beverage",
    "dinner",
    "menu",
    "participants",
    "volunteer",
    "guest",
  ]);
  return [
    ...new Set(
      matches
        .map((m) => m.trim())
        .filter((m) => {
          const tokens = normalizeLatin(m).split(/\s+/);
          return (
            tokens.length >= 2 &&
            tokens.length <= 4 &&
            !tokens.every((t) => stop.has(t))
          );
        })
    ),
  ];
}

function appendUnique(target: string[], values: string[]): void {
  for (const value of values) {
    if (value && !target.includes(value)) {
      target.push(value);
    }
  }
}

/**
 * Expand Burmese and near-miss English names to workbook NAME values
 * so RAG/lexical can hit English-only rows.
 */
export function expandPersonNamesAgainstWorkbook(
  query: string,
  sheets: SheetData[]
): string {
  if (!sheets.length) {
    return query;
  }

  const workbookNames = collectWorkbookPersonNames(sheets);
  if (!workbookNames.length) {
    return query;
  }

  const extras: string[] = [];

  if (MYANMAR_CHAR.test(query)) {
    for (const myanmar of extractMyanmarNameCandidates(query)) {
      const ranked = workbookNames
        .map((english) => ({
          english,
          score: scoreBurmeseToEnglishName(myanmar, english),
        }))
        .filter((item) => item.score >= 0.67)
        .sort((a, b) => b.score - a.score);
      appendUnique(
        extras,
        ranked.slice(0, 3).map((item) => item.english)
      );
    }
  }

  // Fix LLM typos / wrong romanization (Aung Zin Thu → Aung Si Thu)
  for (const candidate of extractEnglishNameCandidates(query)) {
    const exact = workbookNames.find(
      (name) => normalizeLatin(name) === normalizeLatin(candidate)
    );
    if (exact) {
      appendUnique(extras, [exact]);
      continue;
    }
    const ranked = workbookNames
      .map((english) => ({
        english,
        score: scoreEnglishNameSimilarity(candidate, english),
      }))
      .filter((item) => item.score >= 0.78)
      .sort((a, b) => b.score - a.score);
    appendUnique(
      extras,
      ranked.slice(0, 2).map((item) => item.english)
    );
  }

  return extras.length ? `${query} ${extras.join(" ")}` : query;
}
