import { ILogger } from "@microsoft/teams.common";
import { DATABASE_CONFIG } from "../utils/config";
import { IDatabase } from "./database";
import { isPostgresConfigured } from "./postgres";
import { PostgresKVStore } from "./postgresStorage";

export class StorageFactory {
  static async createStorage(logger: ILogger): Promise<IDatabase> {
    if (DATABASE_CONFIG.type === "postgres" || isPostgresConfigured()) {
      logger.debug("🔧 Initializing Postgres storage...");
      const storage = new PostgresKVStore(logger.child("postgres"));
      await storage.initialize();
      logger.debug("✅ Postgres storage initialized successfully");
      return storage;
    }

    throw new Error(
      "Postgres is required. Set PGHOST, PGPORT, PGUSER, PGPASSWORD, PGDATABASE (or DATABASE_URL)."
    );
  }
}
