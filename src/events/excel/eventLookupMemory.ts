/**
 * Per-conversation memory of the last successful workbook lookup.
 * Lets short follow-ups ("list", "ပို့ပေး", "yes") reuse the prior topic
 * without the user repeating every keyword.
 */

export type LastEventLookup = {
  query: string;
  retrieval?: string;
  matchCount?: number;
  at: number;
};

const STORE = new Map<string, LastEventLookup>();
const TTL_MS = 30 * 60 * 1000;

function prune(now = Date.now()): void {
  for (const [key, value] of STORE) {
    if (now - value.at > TTL_MS) {
      STORE.delete(key);
    }
  }
}

export function rememberEventLookup(
  conversationId: string | undefined,
  lookup: Omit<LastEventLookup, "at">
): void {
  const id = (conversationId || "").trim();
  if (!id || !lookup.query?.trim()) {
    return;
  }
  prune();
  STORE.set(id, { ...lookup, query: lookup.query.trim(), at: Date.now() });
}

export function getLastEventLookup(conversationId: string | undefined): LastEventLookup | null {
  const id = (conversationId || "").trim();
  if (!id) {
    return null;
  }
  prune();
  const entry = STORE.get(id);
  if (!entry) {
    return null;
  }
  if (Date.now() - entry.at > TTL_MS) {
    STORE.delete(id);
    return null;
  }
  return entry;
}

export function clearEventLookupMemory(conversationId?: string): void {
  if (!conversationId) {
    STORE.clear();
    return;
  }
  STORE.delete(conversationId.trim());
}

/** Short / affirming replies that almost always continue the previous event answer. */
export function isEventFollowUpAsk(text: string): boolean {
  const raw = String(text || "")
    .replace(/<\/?at>/gi, " ")
    .replace(/<at>[^<]*<\/at>/gi, " ")
    .trim();
  if (!raw || raw.length > 80) {
    return false;
  }
  const q = raw.toLocaleLowerCase("en-US").normalize("NFKC");
  // Explicit list/send/show continuations (EN + Burmese).
  if (
    /\b(list|send|show|names?|detail|details|full|all of them|the list)\b/.test(q) ||
    /စာရင်း|နာမည်|ပို့|ပေးပါ|ကြည့်|အကုန်|အားလုံး/.test(raw)
  ) {
    // Any concrete topic word means a new question ("show agenda"), not a follow-up.
    const topicNoise =
      /\b(table|ferry|menu|agenda|schedule|volunteer|karaoke|beer|juice|cocktail|wine|tea|seat|drink|place|date|participant)\b/i.test(
        q
      ) || /စားပွဲ|ဖယ်ရီ|မီနူး|စေတနာ|သောက်|အစီအစဉ်|နေရာ|ရက်/.test(raw);
    if (topicNoise) {
      return false;
    }
    return true;
  }
  // Bare confirmations after the bot offered more detail.
  return /^(yes|yeah|yep|y|ok|okay|sure|please|do it|go ahead|ဟုတ်|ရပါပြီ|ပေး|လိုချင်)[\s!.?]*$/i.test(
    q
  );
}

/**
 * If the user is continuing, fold the prior lookup keywords into this turn's search.
 */
export function resolveFollowUpLookupQuery(
  conversationId: string | undefined,
  toolQuery: string,
  userText: string
): string {
  const current = [toolQuery, userText]
    .map((part) => String(part || "").trim())
    .filter(Boolean)
    .join(" ")
    .trim();
  const last = getLastEventLookup(conversationId);
  if (!last?.query || !isEventFollowUpAsk(userText)) {
    return current;
  }
  // Keep prior topic primary; append short follow-up cues (list/show).
  const followBits = userText
    .replace(/<\/?at>/gi, " ")
    .trim();
  if (normalizeLoose(current).includes(normalizeLoose(last.query))) {
    return current;
  }
  return `${last.query} ${followBits}`.trim();
}

function normalizeLoose(value: string): string {
  return value.toLocaleLowerCase("en-US").normalize("NFKC").replace(/\s+/g, " ").trim();
}
