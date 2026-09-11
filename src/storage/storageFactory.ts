import { ILogger } from "@microsoft/teams.common";
import { DATABASE_CONFIG, DatabaseConfig } from "../utils/config";
import { IDatabase } from "./database";
import { MssqlKVStore } from "./mssqlStorage";
import { isPostgresConfigured } from "./postgres";
import { PostgresKVStore } from "./postgresStorage";

export class StorageFactory {
  static async createStorage(logger: ILogger, config?: DatabaseConfig): Promise<IDatabase> {
    const dbConfig = config || DATABASE_CONFIG;

    if (dbConfig.type === "postgres" || isPostgresConfigured()) {
      logger.debug("🔧 Initializing Postgres storage...");
      const storage = new PostgresKVStore(logger.child("postgres"));
      await storage.initialize();
      logger.debug("✅ Postgres storage initialized successfully");
      return storage;
    }

    if (dbConfig.type === "mssql") {
      logger.debug("🔧 Initializing MSSQL storage...");
      const storage = new MssqlKVStore(logger.child("mssql"), dbConfig);
      await storage.initialize();
      logger.debug("✅ MSSQL storage initialized successfully");
      return storage;
    }

    throw new Error(
      "No database configured. Set Postgres (PGHOST/PGDATABASE/PGUSER/PGPASSWORD or DATABASE_URL) or MSSQL (SQL_CONNECTION_STRING)."
    );
  }
}
