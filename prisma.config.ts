import "dotenv/config";
import { defineConfig, env } from "prisma/config";

/**
 * Prisma CLI (migrate, studio, db pull) must use the *direct* TCP URL.
 * App runtime should keep using DATABASE_URL (pooled.db.prisma.io).
 */
const migrateUrl =
  process.env.DIRECT_URL?.trim() ||
  process.env.DATABASE_URL?.trim();

if (!migrateUrl) {
  throw new Error("DIRECT_URL or DATABASE_URL must be defined");
}

export default defineConfig({
  schema: "prisma/schema.prisma",
  migrations: {
    path: "prisma/migrations",
  },
  datasource: {
    url: migrateUrl,
  },
});
