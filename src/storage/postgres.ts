import { Pool, type PoolConfig } from "pg";

export function isPostgresConfigured(): boolean {
  return !!(
    process.env.DATABASE_URL?.trim() ||
    process.env.PGDATABASE?.trim() ||
    (process.env.PGHOST?.trim() && process.env.PGUSER?.trim())
  );
}

export function resolveDatabaseType(): "postgres" | "mssql" {
  const explicit = (process.env.DB_TYPE || process.env.DATABASE_TYPE || "").trim().toLowerCase();
  if (explicit === "mssql" || explicit === "sqlserver") {
    return "mssql";
  }
  if (
    explicit === "postgres" ||
    explicit === "postgresql" ||
    explicit === "pg" ||
    isPostgresConfigured()
  ) {
    return "postgres";
  }
  if (process.env.RUNNING_ON_AZURE === "1" || process.env.SQL_CONNECTION_STRING) {
    return "mssql";
  }
  // Default for this project is Postgres
  return "postgres";
}

export function getPostgresPoolConfig(): PoolConfig {
  if (process.env.DATABASE_URL?.trim()) {
    return { connectionString: process.env.DATABASE_URL.trim() };
  }

  return {
    host: process.env.PGHOST || "localhost",
    port: Number(process.env.PGPORT || 5432),
    user: process.env.PGUSER || "postgres",
    password: process.env.PGPASSWORD || "",
    database: process.env.PGDATABASE || "event_management",
  };
}

let sharedPool: Pool | null = null;

export function getPostgresPool(): Pool {
  if (!sharedPool) {
    sharedPool = new Pool(getPostgresPoolConfig());
    sharedPool.on("error", (err) => {
      console.error("Unexpected Postgres pool error:", err);
    });
  }
  return sharedPool;
}

export async function closePostgresPool(): Promise<void> {
  if (sharedPool) {
    await sharedPool.end();
    sharedPool = null;
  }
}

/** Format a Float32Array as a pgvector literal: [0.1,0.2,...] */
export function toPgVectorLiteral(embedding: Float32Array): string {
  return `[${Array.from(embedding).join(",")}]`;
}

export function parsePgVector(value: unknown, dims: number): Float32Array {
  if (value instanceof Float32Array) {
    return value;
  }
  if (Array.isArray(value)) {
    return Float32Array.from(value.map(Number));
  }
  if (typeof value === "string") {
    const trimmed = value.trim().replace(/^\[/, "").replace(/\]$/, "");
    if (!trimmed) {
      return new Float32Array(dims);
    }
    return Float32Array.from(trimmed.split(",").map((part) => Number(part.trim())));
  }
  return new Float32Array(dims);
}
