import type { EmbeddingConfig } from "./embeddings";
import type { EmbeddingProviderName } from "./types";
import type { RagRuntimeConfig } from "./retriever";

function parseBool(value: string | undefined, fallback: boolean): boolean {
  if (value === undefined || value === "") {
    return fallback;
  }
  return !["0", "false", "no", "off"].includes(value.trim().toLowerCase());
}

function resolveApiKey(): string {
  return (
    process.env.EMBEDDING_API_KEY ||
    process.env.OPENAI_EMBEDDING_API_KEY ||
    process.env.AOAI_API_KEY ||
    process.env.OPENAI_API_KEY ||
    process.env.GROQ_API_KEY ||
    process.env.GEMINI_API_KEY ||
    ""
  );
}

function resolveEmbeddingProvider(): EmbeddingProviderName {
  const raw = (process.env.EMBEDDING_PROVIDER || "").trim().toLowerCase();
  if (raw === "openai" || raw === "azure" || raw === "local") {
    return raw;
  }
  if (process.env.AOAI_ENDPOINT || process.env.EMBEDDING_ENDPOINT) {
    return "azure";
  }
  if (process.env.EMBEDDING_API_KEY || process.env.OPENAI_EMBEDDING_API_KEY) {
    return "openai";
  }
  // Default: local hashed embeddings so RAG works with Groq-only chat setup
  return "local";
}

export function getEmbeddingConfig(): EmbeddingConfig {
  const provider = resolveEmbeddingProvider();
  const azureEndpoint = process.env.EMBEDDING_ENDPOINT || process.env.AOAI_ENDPOINT;
  const apiKey = resolveApiKey();

  if (provider === "azure") {
    return {
      provider,
      model:
        process.env.AOAI_EMBEDDING_DEPLOYMENT ||
        process.env.EMBEDDING_MODEL ||
        "text-embedding-3-small",
      apiKey,
      endpoint: azureEndpoint,
      apiVersion: process.env.EMBEDDING_API_VERSION || "2024-02-01",
    };
  }

  if (provider === "openai") {
    return {
      provider,
      model: process.env.EMBEDDING_MODEL || "text-embedding-3-small",
      apiKey,
      baseUrl: process.env.EMBEDDING_BASE_URL || "https://api.openai.com/v1",
    };
  }

  return {
    provider: "local",
    model: "local-hash-384",
    apiKey: "",
  };
}

export function getRagConfig(): RagRuntimeConfig {
  const topK = Number(process.env.RAG_TOP_K || 6);
  const hybridAlpha = Number(process.env.RAG_HYBRID_ALPHA || 0.65);
  const minScore = Number(process.env.RAG_MIN_SCORE || 0.12);

  return {
    enabled: parseBool(process.env.RAG_ENABLED, true),
    topK: Number.isFinite(topK) ? Math.max(1, Math.min(topK, 50)) : 6,
    hybridAlpha: Number.isFinite(hybridAlpha) ? Math.min(1, Math.max(0, hybridAlpha)) : 0.65,
    minScore: Number.isFinite(minScore) ? minScore : 0.12,
    embedding: getEmbeddingConfig(),
  };
}
