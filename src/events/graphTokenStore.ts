import { getPostgresPool } from "../storage/postgres";

export interface StoredGraphTokens {
  userId: string;
  accessToken: string;
  refreshToken: string | null;
  expiresAt: number;
  scope: string;
}

let schemaReady: Promise<void> | null = null;

export async function ensureGraphTokenSchema(): Promise<void> {
  if (!schemaReady) {
    schemaReady = (async () => {
      const pool = getPostgresPool();
      await pool.query(`
        CREATE TABLE IF NOT EXISTS graph_user_tokens (
          user_id TEXT PRIMARY KEY,
          access_token TEXT NOT NULL,
          refresh_token TEXT,
          expires_at TIMESTAMPTZ NOT NULL,
          scope TEXT NOT NULL DEFAULT '',
          updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
        );
      `);
    })();
  }
  await schemaReady;
}

export async function getStoredGraphTokens(
  userId: string
): Promise<StoredGraphTokens | null> {
  await ensureGraphTokenSchema();
  const result = await getPostgresPool().query<{
    user_id: string;
    access_token: string;
    refresh_token: string | null;
    expires_at: Date;
    scope: string;
  }>(
    `SELECT user_id, access_token, refresh_token, expires_at, scope
     FROM graph_user_tokens
     WHERE user_id = $1
     LIMIT 1`,
    [userId]
  );
  const row = result.rows[0];
  if (!row?.access_token) {
    return null;
  }
  return {
    userId: row.user_id,
    accessToken: row.access_token,
    refreshToken: row.refresh_token,
    expiresAt: new Date(row.expires_at).getTime(),
    scope: row.scope || "",
  };
}

export async function saveStoredGraphTokens(tokens: {
  userId: string;
  accessToken: string;
  refreshToken?: string | null;
  expiresAt: number;
  scope: string;
}): Promise<void> {
  await ensureGraphTokenSchema();
  await getPostgresPool().query(
    `INSERT INTO graph_user_tokens (
       user_id, access_token, refresh_token, expires_at, scope, updated_at
     ) VALUES ($1, $2, $3, to_timestamp($4 / 1000.0), $5, NOW())
     ON CONFLICT (user_id) DO UPDATE SET
       access_token = EXCLUDED.access_token,
       refresh_token = COALESCE(EXCLUDED.refresh_token, graph_user_tokens.refresh_token),
       expires_at = EXCLUDED.expires_at,
       scope = EXCLUDED.scope,
       updated_at = NOW()`,
    [
      tokens.userId,
      tokens.accessToken,
      tokens.refreshToken ?? null,
      tokens.expiresAt,
      tokens.scope,
    ]
  );
}

export async function clearStoredGraphTokens(userId: string): Promise<void> {
  await ensureGraphTokenSchema();
  await getPostgresPool().query(`DELETE FROM graph_user_tokens WHERE user_id = $1`, [
    userId,
  ]);
}

export async function hasStoredGraphTokens(userId: string): Promise<boolean> {
  await ensureGraphTokenSchema();
  const result = await getPostgresPool().query<{ exists: boolean }>(
    `SELECT EXISTS(
       SELECT 1 FROM graph_user_tokens WHERE user_id = $1
     ) AS exists`,
    [userId]
  );
  return Boolean(result.rows[0]?.exists);
}
