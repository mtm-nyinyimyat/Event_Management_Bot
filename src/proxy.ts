// Routes outbound HTTPS traffic through a corporate proxy when HTTPS_PROXY /
// HTTP_PROXY / NO_PROXY are set. Node's global fetch (undici) ignores these
// variables by default, so we install a proxy-aware dispatcher explicitly.
// Imported first in index so it applies before any SDK opens a connection.
//
// undici@6 requires Node 18+ (ReadableStream global). Only load it when a
// proxy is configured, and polyfill web streams first for older runtimes.

const hasProxy = !!(
  process.env.HTTPS_PROXY ||
  process.env.HTTP_PROXY ||
  process.env.https_proxy ||
  process.env.http_proxy
);

if (hasProxy) {
  try {
    // Node 16 may lack global ReadableStream; undici@6 expects it.
    // eslint-disable-next-line @typescript-eslint/no-require-imports
    const webStreams = require("stream/web") as typeof import("stream/web");
    if (typeof (globalThis as { ReadableStream?: unknown }).ReadableStream === "undefined") {
      (globalThis as { ReadableStream?: unknown }).ReadableStream = webStreams.ReadableStream;
    }
    if (typeof (globalThis as { WritableStream?: unknown }).WritableStream === "undefined") {
      (globalThis as { WritableStream?: unknown }).WritableStream = webStreams.WritableStream;
    }
    if (typeof (globalThis as { TransformStream?: unknown }).TransformStream === "undefined") {
      (globalThis as { TransformStream?: unknown }).TransformStream = webStreams.TransformStream;
    }

    // eslint-disable-next-line @typescript-eslint/no-require-imports
    const undici = require("undici") as typeof import("undici");
    undici.setGlobalDispatcher(new undici.EnvHttpProxyAgent());
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    console.warn(
      `⚠️ Failed to enable HTTPS proxy via undici (${message}). ` +
        `Use Node.js 20+ (current: ${process.version}).`
    );
  }
}
