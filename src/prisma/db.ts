import "dotenv/config";
import "temporal-polyfill/full/global";
import pgvector from "@prisma/orm-extension-pgvector/runtime";
import postgres from "@prisma/orm-postgres/runtime";
import type { Contract } from "./contract.d";
import contractJson from "./contract.json" with { type: "json" };

/**
 * Prisma 8 client for optional typed access.
 * The bot still uses `pg` for RAG / sessions today; this client is ready for gradual adoption.
 * Runtime traffic should use the pooled DATABASE_URL.
 */
export const db = postgres<Contract>({
  contractJson,
  url: process.env.DATABASE_URL!,
  extensions: [pgvector],
});
