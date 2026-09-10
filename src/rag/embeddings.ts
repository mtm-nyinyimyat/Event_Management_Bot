import type { EmbeddingProviderName } from "./types";

const LOCAL_DIMS = 384;

export interface EmbeddingClient {
  readonly provider: EmbeddingProviderName;
  readonly model: string;
  embed(texts: string[]): Promise<Float32Array[]>;
}

function normalizeText(value: string): string {
  return value.toLocaleLowerCase("my").normalize("NFC").trim();
}

function tokenize(value: string): string[] {
  const normalized = normalizeText(value);
  const tokens = normalized
    .split(/[^\p{L}\p{N}]+/u)
    .map((token) => token.trim())
    .filter((token) => token.length > 0);

  // Character bigrams help Burmese and fuzzy name matching
  const bigrams: string[] = [];
  const compact = normalized.replace(/\s+/g, "");
  for (let i = 0; i < compact.length - 1; i += 1) {
    bigrams.push(compact.slice(i, i + 2));
  }

  return [...tokens, ...bigrams];
}

function hashToken(token: string): number {
  let hash = 2166136261;
  for (let i = 0; i < token.length; i += 1) {
    hash ^= token.charCodeAt(i);
    hash = Math.imul(hash, 16777619);
  }
  return hash >>> 0;
}

/** Local hashed bag-of-tokens embedding (no external API). */
export function embedLocal(text: string, dims = LOCAL_DIMS): Float32Array {
  const vector = new Float32Array(dims);
  const tokens = tokenize(text);
  if (tokens.length === 0) {
    return vector;
  }

  for (const token of tokens) {
    const hash = hashToken(token);
    const index = hash % dims;
    const sign = (hash & 1) === 0 ? 1 : -1;
    const weight = token.length >= 4 ? 1.25 : 1;
    vector[index] += sign * weight;
  }

  return l2Normalize(vector);
}

export function l2Normalize(vector: Float32Array): Float32Array {
  let sumSquares = 0;
  for (let i = 0; i < vector.length; i += 1) {
    sumSquares += vector[i] * vector[i];
  }
  const norm = Math.sqrt(sumSquares);
  if (norm === 0) {
    return vector;
  }
  for (let i = 0; i < vector.length; i += 1) {
    vector[i] /= norm;
  }
  return vector;
}

export function cosineSimilarity(a: Float32Array, b: Float32Array): number {
  const length = Math.min(a.length, b.length);
  let dot = 0;
  for (let i = 0; i < length; i += 1) {
    dot += a[i] * b[i];
  }
  return dot;
}

class LocalEmbeddingClient implements EmbeddingClient {
  readonly provider: EmbeddingProviderName = "local";
  readonly model = `local-hash-${LOCAL_DIMS}`;

  async embed(texts: string[]): Promise<Float32Array[]> {
    return texts.map((text) => embedLocal(text));
  }
}

class OpenAICompatibleEmbeddingClient implements EmbeddingClient {
  readonly provider: EmbeddingProviderName;
  readonly model: string;
  private readonly apiKey: string;
  private readonly url: string;
  private readonly headers: Record<string, string>;

  constructor(options: {
    provider: EmbeddingProviderName;
    model: string;
    apiKey: string;
    url: string;
    headers?: Record<string, string>;
  }) {
    this.provider = options.provider;
    this.model = options.model;
    this.apiKey = options.apiKey;
    this.url = options.url;
    this.headers = options.headers || {};
  }

  async embed(texts: string[]): Promise<Float32Array[]> {
    if (texts.length === 0) {
      return [];
    }

    const response = await fetch(this.url, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${this.apiKey}`,
        ...this.headers,
      },
      body: JSON.stringify({
        model: this.model,
        input: texts,
      }),
    });

    if (!response.ok) {
      const body = await response.text();
      throw new Error(`Embedding request failed (${response.status}): ${body.slice(0, 400)}`);
    }

    const payload = (await response.json()) as {
      data?: Array<{ embedding?: number[]; index?: number }>;
    };

    const data = [...(payload.data || [])].sort(
      (a, b) => (a.index ?? 0) - (b.index ?? 0)
    );

    if (data.length !== texts.length) {
      throw new Error(`Embedding response size mismatch: expected ${texts.length}, got ${data.length}`);
    }

    return data.map((item) => {
      const values = item.embedding || [];
      return l2Normalize(Float32Array.from(values));
    });
  }
}

export interface EmbeddingConfig {
  provider: EmbeddingProviderName;
  model: string;
  apiKey: string;
  baseUrl?: string;
  endpoint?: string;
  apiVersion?: string;
}

export function createEmbeddingClient(config: EmbeddingConfig): EmbeddingClient {
  if (config.provider === "local") {
    return new LocalEmbeddingClient();
  }

  if (config.provider === "azure") {
    if (!config.endpoint) {
      throw new Error("Azure embeddings require AOAI_ENDPOINT / EMBEDDING_ENDPOINT");
    }
    const deployment = config.model;
    const apiVersion = config.apiVersion || "2024-02-01";
    const url = `${config.endpoint.replace(/\/$/, "")}/openai/deployments/${encodeURIComponent(
      deployment
    )}/embeddings?api-version=${encodeURIComponent(apiVersion)}`;

    return new OpenAICompatibleEmbeddingClient({
      provider: "azure",
      model: deployment,
      apiKey: config.apiKey,
      url,
      headers: { "api-key": config.apiKey },
    });
  }

  const baseUrl = (config.baseUrl || "https://api.openai.com/v1").replace(/\/$/, "");
  return new OpenAICompatibleEmbeddingClient({
    provider: "openai",
    model: config.model,
    apiKey: config.apiKey,
    url: `${baseUrl}/embeddings`,
  });
}

export async function embedInBatches(
  client: EmbeddingClient,
  texts: string[],
  batchSize = 64
): Promise<Float32Array[]> {
  const out: Float32Array[] = [];
  for (let i = 0; i < texts.length; i += batchSize) {
    const batch = texts.slice(i, i + batchSize);
    const vectors = await client.embed(batch);
    out.push(...vectors);
  }
  return out;
}
