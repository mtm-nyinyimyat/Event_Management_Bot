import { OpenAIChatModel } from "@microsoft/teams.openai";
import { ILogger } from "@microsoft/teams.common";
import type { Fetch } from "openai/core";
import { getEventsSource } from "../events/excelStore";
import { assertGraphExcelConfig, describeGraphExcelConfig } from "../events/graphExcelClient";
import { getRagConfig } from "../rag/config";
import { resolveVectorBackend } from "../rag/vectorStore";
import { isPostgresConfigured, resolveDatabaseType } from "../storage/postgres";

export interface ModelConfig {
  model: string;
  apiKey: string;
  baseUrl?: string;
  endpoint?: string;
  apiVersion?: string;
}

export interface DatabaseConfig {
  type: "mssql" | "postgres";
  connectionString?: string;
  server?: string;
  database?: string;
  username?: string;
  password?: string;
}

export const DATABASE_CONFIG: DatabaseConfig = {
  type: resolveDatabaseType(),
  connectionString: process.env.SQL_CONNECTION_STRING || process.env.DATABASE_URL,
  server: process.env.SQL_SERVER || process.env.PGHOST,
  database: process.env.SQL_DATABASE || process.env.PGDATABASE,
  username: process.env.SQL_USERNAME || process.env.PGUSER,
  password: process.env.SQL_PASSWORD || process.env.PGPASSWORD,
};

const XKIRO_OPENAI_BASE_URL = "https://api.xkiro.com/v1";
const DEFAULT_MODEL = "qwen/qwen3.8-max:free";

function resolveApiKey(): string {
  const candidates = [
    process.env.XKIRO_API_KEY,
    process.env.OPENAI_API_KEY,
    process.env.GROQ_API_KEY,
    process.env.GEMINI_API_KEY,
    process.env.GOOGLE_API_KEY,
    process.env.AOAI_API_KEY,
  ];

  for (const raw of candidates) {
    const key = (raw || "").trim();
    if (!key) {
      continue;
    }
    // Do not send Groq keys to xKiro (or other non-Groq endpoints)
    if (key.startsWith("gsk_")) {
      continue;
    }
    return key;
  }

  return "";
}

function isGeminiBaseUrl(baseUrl?: string): boolean {
  return /generativelanguage\.googleapis\.com/i.test(baseUrl || "");
}

function ensureToolCallId(
  toolCall: { id?: string; function?: { name?: string } },
  index: number
): string {
  if (toolCall.id && toolCall.id.trim()) {
    return toolCall.id;
  }
  const name = toolCall.function?.name || "tool";
  const id = `${name}_${index}`;
  toolCall.id = id;
  return id;
}

function sanitizeGeminiTools(tools: unknown[]): void {
  for (const tool of tools) {
    if (!tool || typeof tool !== "object") {
      continue;
    }
    const fn = (tool as { function?: { parameters?: Record<string, unknown> } }).function;
    const params = fn?.parameters;
    if (!params || typeof params !== "object") {
      continue;
    }
    if (!params.type) {
      params.type = "object";
    }
    // Gemini rejects empty required arrays
    if (Array.isArray(params.required) && params.required.length === 0) {
      delete params.required;
    }
  }
}

function sanitizeGeminiMessages(messages: Array<Record<string, unknown>>): void {
  const idToName = new Map<string, string>();

  for (const message of messages) {
    if (message.role === "assistant" && Array.isArray(message.tool_calls)) {
      // Empty assistant content + tool_calls => 400 on Gemini OpenAI compat
      if (message.content === "" || message.content == null) {
        delete message.content;
      }

      (message.tool_calls as Array<{ id?: string; function?: { name?: string } }>).forEach(
        (toolCall, index) => {
          const id = ensureToolCallId(toolCall, index);
          if (toolCall.function?.name) {
            idToName.set(id, toolCall.function.name);
          }
        }
      );
    }

    if (message.role === "tool") {
      const toolCallId = typeof message.tool_call_id === "string" ? message.tool_call_id : "";
      if (!toolCallId.trim()) {
        // Recover from empty ids produced by Gemini
        const inferred =
          (typeof message.name === "string" && message.name) ||
          [...idToName.values()][0] ||
          "tool";
        message.tool_call_id = `${inferred}_0`;
      }

      const resolvedId = String(message.tool_call_id);
      const name =
        (typeof message.name === "string" && message.name) ||
        idToName.get(resolvedId) ||
        (resolvedId.includes("_") ? resolvedId.replace(/_\d+$/, "") : undefined);

      if (name) {
        message.name = name;
        if (!idToName.has(resolvedId)) {
          idToName.set(resolvedId, name);
        }
      }

      if (message.content == null) {
        message.content = "";
      }
    }
  }
}

/**
 * Gemini's OpenAI-compatible endpoint often returns empty tool_call ids and rejects
 * empty assistant content / empty required[]. Patch request+response bodies.
 */
function createGeminiCompatibleFetch(): Fetch {
  const patched = async (input: unknown, init?: RequestInit): Promise<Response> => {
    let url: RequestInfo | URL = input as RequestInfo | URL;
    let nextInit: RequestInit | undefined = init;

    // OpenAI SDK may pass a Request object; rebuild so we can safely rewrite the body.
    if (typeof Request !== "undefined" && input instanceof Request) {
      let bodyText: string | undefined;
      if (typeof init?.body === "string") {
        bodyText = init.body;
      } else if (input.method !== "GET" && input.method !== "HEAD") {
        bodyText = await input.clone().text();
      }

      nextInit = {
        method: init?.method || input.method,
        headers: init?.headers || Object.fromEntries(input.headers.entries()),
        body: bodyText,
        signal: init?.signal || input.signal,
      };
      url = input.url;
    }

    if (nextInit?.body && typeof nextInit.body === "string") {
      try {
        const body = JSON.parse(nextInit.body) as {
          messages?: Array<Record<string, unknown>>;
          tools?: unknown[];
        };
        if (Array.isArray(body.tools)) {
          sanitizeGeminiTools(body.tools);
        }
        if (Array.isArray(body.messages)) {
          sanitizeGeminiMessages(body.messages);
        }
        nextInit = {
          ...nextInit,
          body: JSON.stringify(body),
          headers: {
            ...(typeof nextInit.headers === "object" &&
            nextInit.headers &&
            !(nextInit.headers instanceof Headers)
              ? (nextInit.headers as Record<string, string>)
              : {}),
            "content-type": "application/json",
          },
        };
      } catch {
        // leave body unchanged
      }
    }

    const response = await globalThis.fetch(url, nextInit);
    const contentType = response.headers.get("content-type") || "";
    if (!contentType.includes("application/json")) {
      return response;
    }

    const text = await response.text();
    try {
      const payload = JSON.parse(text) as {
        choices?: Array<{
          message?: {
            content?: string | null;
            tool_calls?: Array<{ id?: string; function?: { name?: string } }>;
          };
        }>;
        error?: unknown;
      };

      if (!response.ok) {
        console.error(
          `❌ Gemini API ${response.status}: ${JSON.stringify(payload.error || payload).slice(0, 800)}`
        );
      }

      for (const choice of payload.choices || []) {
        const message = choice.message;
        if (!message?.tool_calls?.length) {
          continue;
        }
        message.tool_calls.forEach((toolCall, index) => ensureToolCallId(toolCall, index));
        if (message.content === "") {
          message.content = null;
        }
      }

      // Do NOT reuse original headers: fetch already decoded gzip, but headers may
      // still say content-encoding:gzip — that causes OpenAI SDK "Connection error".
      return new Response(JSON.stringify(payload), {
        status: response.status,
        statusText: response.statusText,
        headers: { "content-type": "application/json; charset=utf-8" },
      });
    } catch (error) {
      console.error(
        `❌ Gemini response parse failed: ${error instanceof Error ? error.message : String(error)} | body=${text.slice(0, 300)}`
      );
      return new Response(text, {
        status: response.status,
        statusText: response.statusText,
        headers: { "content-type": "application/json; charset=utf-8" },
      });
    }
  };

  return patched as unknown as Fetch;
}

function buildModelConfig(model: string): ModelConfig {
  const apiKey = resolveApiKey();
  const azureEndpoint = process.env.AOAI_ENDPOINT;

  if (azureEndpoint) {
    return {
      model: process.env.AOAI_MODEL || process.env.OPENAI_MODEL || model,
      apiKey,
      endpoint: azureEndpoint,
      apiVersion: "2025-04-01-preview",
    };
  }

  return {
    model: process.env.OPENAI_MODEL || model,
    apiKey,
    baseUrl: process.env.OPENAI_BASE_URL || XKIRO_OPENAI_BASE_URL,
  };
}

function getSharedModelConfig(): ModelConfig {
  return buildModelConfig(DEFAULT_MODEL);
}

export function getModelConfig(_capabilityType: string): ModelConfig {
  return getSharedModelConfig();
}

export function createChatModel(config: ModelConfig): OpenAIChatModel {
  if (config.endpoint) {
    return new OpenAIChatModel({
      model: config.model,
      apiKey: config.apiKey,
      endpoint: config.endpoint,
      apiVersion: config.apiVersion,
    });
  }

  return new OpenAIChatModel({
    model: config.model,
    apiKey: config.apiKey,
    baseUrl: config.baseUrl,
    fetch: isGeminiBaseUrl(config.baseUrl) ? createGeminiCompatibleFetch() : undefined,
  });
}

export { getRagConfig, getEmbeddingConfig } from "../rag/config";

export function validateEnvironment(logger: ILogger): void {
  if (!resolveApiKey()) {
    throw new Error(
      "Missing XKIRO_API_KEY (or OPENAI_API_KEY with sk-xt-…). Create a key at https://xkiro.com"
    );
  }

  const sessionMode = !["0", "false", "no", "off"].includes(
    (process.env.EVENTS_SESSION_MODE || "1").trim().toLowerCase()
  );
  if (sessionMode) {
    logger.debug(
      "📄 Excel source: SharePoint URL session mode (paste URL → /start → Q&A → /end). Set EVENTS_SESSION_MODE=0 for local/graph fallback."
    );
  } else if (getEventsSource() === "graph") {
    assertGraphExcelConfig();
    logger.debug(`📄 Excel source: Microsoft Graph (${describeGraphExcelConfig()})`);
  } else {
    logger.debug("📄 Excel source: local file (set EVENTS_SOURCE=graph to use Teams/SharePoint)");
  }

  const rag = getRagConfig();
  logger.debug(
    `🔎 RAG enabled (provider=${rag.embedding.provider}, model=${rag.embedding.model}, topK=${rag.topK})`
  );

  if (DATABASE_CONFIG.type === "postgres") {
    if (!isPostgresConfigured()) {
      throw new Error(
        "Postgres required but PG* / DATABASE_URL is incomplete. Set PGHOST, PGPORT, PGUSER, PGPASSWORD, PGDATABASE (or DATABASE_URL)."
      );
    }
    logger.debug(
      `✅ Postgres configuration validated (${process.env.PGDATABASE || "DATABASE_URL"})`
    );
  }

  if (DATABASE_CONFIG.type === "mssql") {
    const sqlRequiredVars = ["SQL_CONNECTION_STRING"];
    const sqlMissing = sqlRequiredVars.filter((envVar) => !process.env[envVar]);
    if (sqlMissing.length > 0) {
      throw new Error(
        `SQL Server configuration incomplete. Missing: ${sqlMissing.join(", ")}`
      );
    }
    logger.debug("✅ SQL Server configuration validated");
  }

  logger.debug(`📦 Using database: ${DATABASE_CONFIG.type}`);
  logger.debug(`📦 RAG vector store: ${resolveVectorBackend()}`);
  logger.debug("✅ Environment validation passed");
}

export function logModelConfigs(logger: ILogger): void {
  const config = getSharedModelConfig();
  const provider = config.endpoint
    ? `Azure OpenAI (${config.endpoint})`
    : `OpenAI-compatible (${config.baseUrl})`;
  logger.debug("🔧 AI Model Configuration:");
  logger.debug(`  Provider: ${provider}`);
  logger.debug(`  Model: ${config.model}`);
  if (isGeminiBaseUrl(config.baseUrl)) {
    logger.debug("  Gemini OpenAI-compat tool-call patches: enabled");
  }

  const rag = getRagConfig();
  logger.debug("🔧 RAG Configuration:");
  logger.debug(`  Enabled: ${rag.enabled}`);
  logger.debug(`  Embedding provider: ${rag.embedding.provider}`);
  logger.debug(`  Embedding model: ${rag.embedding.model}`);
  logger.debug(`  Top K: ${rag.topK}`);
  logger.debug(
    `  Vector store: postgres (${process.env.PGDATABASE || process.env.DATABASE_URL || "pg"})`
  );
}
