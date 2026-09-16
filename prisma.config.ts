import "dotenv/config";
import { definePrismaConfig } from "prisma/config";
import pgvector from "@prisma/orm-extension-pgvector/control";
import { defineConfig as ormConfig } from "@prisma/orm-postgres/config";

/**
 * Prisma 8 config.
 * - CLI (db init / migrate / verify): prefer DIRECT_URL (db.prisma.io).
 * - App runtime client in src/prisma/db.ts: use DATABASE_URL (pooled.db.prisma.io).
 */
const connection =
  process.env.DIRECT_URL?.trim() || process.env.DATABASE_URL?.trim();

if (!connection) {
  throw new Error("Set DIRECT_URL (preferred for CLI) or DATABASE_URL in .env");
}

export default definePrismaConfig({
  skills: {
    agents: [],
  },
  orm: ormConfig({
    contract: "./src/prisma/contract.prisma",
    extensions: [pgvector],
    migrations: {
      dir: "migrations",
    },
    db: {
      connection,
    },
  }),
});
