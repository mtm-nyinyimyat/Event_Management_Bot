/**
 * Lazy-load eventSession without hardcoding a .js path that breaks under
 * `ts-node/register` (dev) while still working from compiled `dist/`.
 */
export async function importEventSession(): Promise<typeof import("./eventSession")> {
  try {
    return await import("./eventSession");
  } catch (extensionlessError) {
    try {
      return await import("./eventSession.js");
    } catch {
      throw extensionlessError;
    }
  }
}
