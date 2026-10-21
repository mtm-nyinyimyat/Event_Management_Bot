# Teams SSO + OBO setup (Entra-only)

Device-code login and **Azure Bot OAuth Connection** are not used.  
`GRAPH_AUTH_MODE=sso` gets Graph tokens via:

1. **Silent Teams SSO** — `signin/tokenExchange` assertion → backend **On-Behalf-Of (OBO)**
2. **Consent fallback** — browser `/auth/start` → Entra authorize → `/auth/end` auth-code grant

Tokens (access ~1h, refresh ~90d) are stored per user in Postgres and auto-refreshed.

**Azure Subscription / Azure Bot OAuth Connection Settings are not required for this Graph SSO path.**

You still need a messaging endpoint the Teams bot can reach (local tunnel / Toolkit). That is separate from OAuth Connections.

## Environment

```env
GRAPH_AUTH_MODE=sso
CLIENT_ID=<entra-app-client-id>
CLIENT_SECRET=<entra-app-client-secret>
TENANT_ID=<entra-tenant-id>
PUBLIC_BASE_URL=https://<your-public-https-host>
GRAPH_OPERATOR_USER_ID=<operator-entra-object-id>
# Optional:
# BOT_ENDPOINT=https://<same-as-public-base>   # used if PUBLIC_BASE_URL unset
# GRAPH_SSO_RESOURCE_URI=api://botid-<CLIENT_ID>
# Production (admin consent usually required for *.Read.All):
# GRAPH_DELEGATED_SCOPES=https://graph.microsoft.com/Files.Read.All https://graph.microsoft.com/Sites.Read.All https://graph.microsoft.com/User.Read https://graph.microsoft.com/Chat.Read offline_access
# Local without admin consent (upload mode / limited Graph — see below):
# GRAPH_DELEGATED_SCOPES=https://graph.microsoft.com/User.Read https://graph.microsoft.com/Files.Read offline_access
```

`ChannelMessage.Read.All` is **not** in the default scope list. Channel message re-read via Graph is best-effort; without that permission the bot still works (Bot Framework payload only).

### Local testing without admin consent

Many Graph `*.Read.All` permissions (including `Files.Read.All`, `Sites.Read.All`, `ChannelMessage.Read.All`) require **admin consent**. For local tests without an admin:

1. Set ingest to chat upload (no SharePoint `/shares` call):
   ```env
   EVENTS_INGEST_MODE=upload
   GRAPH_DELEGATED_SCOPES=https://graph.microsoft.com/User.Read https://graph.microsoft.com/Files.Read offline_access
   ```
2. In Entra, only add those delegated permissions (user can consent).
3. Delete `graph_user_tokens` rows and sign in again.

SharePoint **sharing URL** ingest needs `Files.Read.All` (admin consent). Without it, use upload mode locally.

`GRAPH_OPERATOR_USER_ID` is required when SharePoint ingest, drive id, or webhooks are enabled. That user must complete SSO once; background delta/webhook calls use their stored refresh token.

`PUBLIC_BASE_URL` must be HTTPS and reachable by the user’s browser (and listed as an Entra redirect URI).

## Entra ID app (only identity setup needed)

1. **Expose an API** — Application ID URI: `api://botid-<CLIENT_ID>` (must match `webApplicationInfo.resource` in the Teams manifest).
2. Add scope `access_as_user` (Admins and users).
3. **Authorized client applications** — add Teams client IDs for that scope (see [bot SSO register AAD](https://learn.microsoft.com/microsoftteams/platform/bots/how-to/authentication/bot-sso-register-aad)).
4. **API permissions (delegated)** — pick by environment:
   - **Local (no admin):** `User.Read`, `Files.Read`, `offline_access` + `EVENTS_INGEST_MODE=upload`
   - **SharePoint URL ingest:** add `Files.Read.All` (+ usually `Sites.Read.All`) and **admin consent**
   - Optional later: `Chat.Read`, `ChannelMessage.Read.All` (channel Graph enrich; admin consent)
5. **Authentication** → add redirect URI: `https://<PUBLIC_BASE_URL>/auth/end` (Web platform).
6. Keep a **client secret** (confidential client for OBO and auth-code exchange).

Do **not** create Azure Bot → Configuration → OAuth Connection Settings for this flow.

### After changing Graph scopes

Old tokens in `graph_user_tokens` keep the previous scope set. Delete the user’s row (or truncate the table), restart the bot, and sign in again so OBO/auth-code requests `Files.Read.All`.

## Manifest

```json
"webApplicationInfo": {
  "id": "${{BOT_ID}}",
  "resource": "api://botid-${{BOT_ID}}"
}
```

Sideload/redeploy after Application ID URI changes. `BOT_ID` should be the same as the Entra app client id.

## Runtime flow

1. User action needs Graph → bot sends a card with `tokenExchangeResource` (`api://botid-...`).
2. Teams silent SSO → `signin/tokenExchange` → backend OBO → rows in `graph_user_tokens`.
3. If consent is required → user opens **Sign in** → `/auth/start` → Entra → `/auth/end` → tokens saved.
4. Later Graph calls use the stored access token; on expiry the backend refreshes with the refresh token (no login prompt).
5. Graph runs **on behalf of that user**, so private channels they cannot access stay inaccessible.

## Postgres

Table `graph_user_tokens` is created automatically on startup (`ensureGraphTokenSchema`).
