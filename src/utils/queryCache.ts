type CacheEntry<T> = {
  value: T;
  expiresAt: number;
};

function parseBool(value: string | undefined, fallback: boolean): boolean {
  if (value === undefined || value === "") {
    return fallback;
  }
  return !["0", "false", "no", "off"].includes(value.trim().toLowerCase());
}

function cacheTtlMs(): number {
  const raw = Number(process.env.QUERY_CACHE_TTL_MS || 15 * 60 * 1000);
  return Number.isFinite(raw) && raw >= 0 ? raw : 15 * 60 * 1000;
}

function cacheMaxEntries(): number {
  const raw = Number(process.env.QUERY_CACHE_MAX_ENTRIES || 200);
  return Number.isFinite(raw) && raw > 0 ? Math.floor(raw) : 200;
}

export function isQueryCacheEnabled(): boolean {
  return parseBool(process.env.QUERY_CACHE_ENABLED, true);
}

/** Normalize a user/tool question so slight wording/spacing differences still hit cache. */
export function normalizeCacheQuery(text: string): string {
  return String(text || "")
    .replace(/<\/?at>/gi, " ")
    .replace(/<at>[^<]*<\/at>/gi, " ")
    .toLocaleLowerCase("en-US")
    .normalize("NFKC")
    .replace(/[^\p{L}\p{N}\s&]+/gu, " ")
    .replace(/\s+/g, " ")
    .trim();
}

export class TtlCache<T> {
  private store = new Map<string, CacheEntry<T>>();

  constructor(private readonly name: string) {}

  get(key: string): T | undefined {
    if (!isQueryCacheEnabled() || !key) {
      return undefined;
    }
    const entry = this.store.get(key);
    if (!entry) {
      return undefined;
    }
    if (Date.now() > entry.expiresAt) {
      this.store.delete(key);
      return undefined;
    }
    // Refresh insertion order for simple LRU eviction
    this.store.delete(key);
    this.store.set(key, entry);
    return entry.value;
  }

  set(key: string, value: T, ttlMs = cacheTtlMs()): void {
    if (!isQueryCacheEnabled() || !key || ttlMs <= 0) {
      return;
    }
    while (this.store.size >= cacheMaxEntries()) {
      const oldest = this.store.keys().next().value;
      if (oldest === undefined) {
        break;
      }
      this.store.delete(oldest);
    }
    this.store.set(key, { value, expiresAt: Date.now() + ttlMs });
  }

  clear(): void {
    this.store.clear();
  }

  get size(): number {
    return this.store.size;
  }

  get label(): string {
    return this.name;
  }
}

/** Cached workbook lookup payloads (lookup_events / searchWorkbook). */
export const workbookSearchCache = new TtlCache<unknown>("workbook-search");

/** Cached final bot answers for identical normalized user questions. */
export const answerCache = new TtlCache<string>("answer");

export function clearQueryCaches(): void {
  workbookSearchCache.clear();
  answerCache.clear();
}
