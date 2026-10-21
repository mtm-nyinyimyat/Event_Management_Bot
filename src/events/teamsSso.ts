import { randomUUID } from "node:crypto";
import {
  getAuthStartUrl,
  getGraphTokenExchangeResourceUri,
} from "./graphAuth";

/**
 * Card that triggers Teams silent SSO via tokenExchangeResource.
 * Sign-in button falls back to our Entra auth-code flow (/auth/start) —
 * no Azure Bot OAuth connection required.
 */
export function buildGraphSsoOAuthCard(options: {
  userId: string;
  text?: string;
  buttonTitle?: string;
}): {
  type: "message";
  attachments: Array<{
    contentType: string;
    content: Record<string, unknown>;
  }>;
} {
  const resourceUri = getGraphTokenExchangeResourceUri();
  const signInUrl = getAuthStartUrl(options.userId);
  return {
    type: "message",
    attachments: [
      {
        contentType: "application/vnd.microsoft.card.oauth",
        content: {
          text:
            options.text ||
            "Sign in to allow the bot to access Microsoft 365 on your behalf.",
          // Local label only — not looked up in Azure Bot OAuth Connections.
          connectionName: "GraphSSO",
          tokenExchangeResource: {
            id: randomUUID(),
            uri: resourceUri,
          },
          buttons: [
            {
              type: "signin",
              title: options.buttonTitle || "Sign in",
              value: signInUrl,
            },
          ],
        },
      },
    ],
  };
}

export function extractTokenExchangeAssertion(activity: {
  value?: unknown;
}): string | undefined {
  const value = (activity.value || {}) as {
    token?: string;
    id?: string;
  };
  const token = value.token?.trim();
  return token || undefined;
}
