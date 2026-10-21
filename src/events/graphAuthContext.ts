import { AsyncLocalStorage } from "node:async_hooks";

/**
 * Per-request Graph auth identity. Interactive handlers set the Teams user;
 * background jobs set the operator user id.
 */
export interface GraphAuthStore {
  userId: string;
}

const graphAuthStorage = new AsyncLocalStorage<GraphAuthStore>();

export function runWithGraphAuthUser<T>(userId: string, fn: () => T): T {
  return graphAuthStorage.run({ userId }, fn);
}

export async function runWithGraphAuthUserAsync<T>(
  userId: string,
  fn: () => Promise<T>
): Promise<T> {
  return graphAuthStorage.run({ userId }, fn);
}

export function getGraphAuthUserId(): string | undefined {
  return graphAuthStorage.getStore()?.userId;
}
