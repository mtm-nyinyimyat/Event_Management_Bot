import { OpenAIChatModel } from "@microsoft/teams.openai";
import { ILogger } from "@microsoft/teams.common";
import { getEventsSource } from "../events/excelStore";
import { assertGraphExcelConfig, describeGraphExcelConfig } from "../events/graphExcelClient";

export interface ModelConfig {
  model: string;
  apiKey: string;
  baseUrl?: string;
  endpoint?: string;
  apiVersion?: string;
}

export interface DatabaseConfig {
  type: "sqlite" | "mssql";
  connectionString?: string;
  server?: string;
  database?: string;
  username?: string;
  password?: string;
  sqlitePath?: string;
}

export const DATABASE_CONFIG: DatabaseConfig = {
  type: process.env.RUNNING_ON_AZURE === "1" ? "mssql" : "sqlite",
  connectionString: process.env.SQL_CONNECTION_STRING,
  server: process.env.SQL_SERVER,
  database: process.env.SQL_DATABASE,
  username: process.env.SQL_USERNAME,
  password: process.env.SQL_PASSWORD,
  sqlitePath: process.env.CONVERSATIONS_DB_PATH,
};

const GROQ_OPENAI_BASE_URL = "https://api.groq.com/openai/v1";

function resolveApiKey(): string {
  return (
    process.env.OPENAI_API_KEY ||
    process.env.GROQ_API_KEY ||
    process.env.GEMINI_API_KEY ||
    process.env.AOAI_API_KEY ||
    ""
  );
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
    baseUrl: process.env.OPENAI_BASE_URL || GROQ_OPENAI_BASE_URL,
  };
}

function getSharedModelConfig(): ModelConfig {
  return buildModelConfig("openai/gpt-oss-120b");
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
  });
}

export function validateEnvironment(logger: ILogger): void {
  if (!resolveApiKey()) {
    throw new Error(
      "Missing OPENAI_API_KEY. For Groq gpt-oss-120b, create a key at https://console.groq.com"
    );
  }

  if (getEventsSource() === "graph") {
    assertGraphExcelConfig();
    logger.debug(`📄 Excel source: Microsoft Graph (${describeGraphExcelConfig()})`);
  } else {
    logger.debug("📄 Excel source: local file (set EVENTS_SOURCE=graph to use Teams/SharePoint)");
  }

  if (DATABASE_CONFIG.type === "mssql") {
    const sqlRequiredVars = ["SQL_CONNECTION_STRING"];
    const sqlMissing = sqlRequiredVars.filter((envVar) => !process.env[envVar]);
    if (sqlMissing.length > 0) {
      logger.warn(
        `SQL Server configuration incomplete. Missing: ${sqlMissing.join(
          ", "
        )}. Falling back to SQLite.`
      );
      DATABASE_CONFIG.type = "sqlite";
    } else {
      logger.debug("✅ SQL Server configuration validated");
    }
  }

  logger.debug(`📦 Using database: ${DATABASE_CONFIG.type}`);
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
}
