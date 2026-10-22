import { AsyncLocalStorage } from "node:async_hooks";

/**
 * Per-request Graph auth identity. Interactive handlers set the Teams user;
 * background jobs set the operator user id.
 *
 * `alternateUserIds` covers cases where tokens were saved under Teams `from.id`
 * (29:…) in one chat and looked up by `aadObjectId` in another (or the reverse).
 */
export interface GraphAuthStore {
  userId: string;
  alternateUserIds?: string[];
}

const graphAuthStorage = new AsyncLocalStorage<GraphAuthStore>();

export function runWithGraphAuthUser<T>(
  userId: string,
  fn: () => T,
  alternateUserIds?: string[]
): T {
  return graphAuthStorage.run({ userId, alternateUserIds }, fn);
}

export async function runWithGraphAuthUserAsync<T>(
  userId: string,
  fn: () => Promise<T>,
  alternateUserIds?: string[]
): Promise<T> {
  return graphAuthStorage.run({ userId, alternateUserIds }, fn);
}

export function getGraphAuthUserId(): string | undefined {
  return graphAuthStorage.getStore()?.userId;
}

export function getGraphAuthAlternateUserIds(): string[] {
  return graphAuthStorage.getStore()?.alternateUserIds || [];
}

/** Prefer Entra object id (stable across DM / group / channel); fall back to Teams id. */
export function resolveTeamsAuthUserIds(from?: {
  aadObjectId?: string;
  id?: string;
}): { canonical: string; alternates: string[] } {
  const aad = from?.aadObjectId?.trim() || "";
  const teams = from?.id?.trim() || "";
  if (aad && teams && aad !== teams) {
    return { canonical: aad, alternates: [teams] };
  }
  return { canonical: aad || teams, alternates: [] };
}
