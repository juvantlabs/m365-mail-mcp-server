/**
 * MSAL Node client factory + cache plugin wiring.
 *
 * Uses ConfidentialClientApplication (we have a client_secret) with the
 * Authorization Code flow for delegated permissions. Per the handbook
 * spec § Auth, we never roll our own OAuth — MSAL Node is Microsoft's
 * official library and handles refresh, revocation, and edge cases.
 *
 * Tokens persist in the OS keychain via `src/auth/keyring.ts`. The MSAL
 * cache plugin pattern is: load on first access, save when it changes.
 *
 * NOTE: this server intentionally uses its OWN app registration (env
 * vars `M365_MAIL_CLIENT_ID`, `M365_MAIL_CLIENT_SECRET`,
 * `M365_MAIL_TENANT_ID`) distinct from the sibling `m365-graph-mcp-server`.
 * Isolation is deliberate — a scope misconfiguration in one server MUST
 * NOT reuse the other's app and silently escalate permissions.
 */

import {
  ConfidentialClientApplication,
  type Configuration,
  type ICachePlugin,
  type TokenCacheContext,
} from "@azure/msal-node";

import { getTokenStore } from "./keyring.js";

/**
 * Delegated scopes the MCP server requests. Order is irrelevant; MSAL
 * normalizes. `offline_access` is required to get a refresh token.
 *
 * v0.2 scope set — v0.1 baseline plus shared / delegate mailbox scopes:
 *   - User.Read              : identity ping (whoami / setup validation)
 *   - Mail.Read              : list folders/messages, get message + attachments
 *                              on the caller's own mailbox
 *   - Mail.ReadWrite         : create/update/move/delete drafts + messages
 *                              on the caller's own mailbox
 *   - Mail.Read.Shared       : read equivalents against a shared / delegate
 *                              mailbox the caller has been granted access
 *                              to (routed via `/users/{upn}/…`)
 *   - Mail.ReadWrite.Shared  : write-side equivalents (draft authoring,
 *                              mark-read, move, delete) against a shared /
 *                              delegate mailbox
 *   - offline_access         : refresh token so the server can run
 *                              unattended after `npm run setup`
 *
 * NOT included (deliberate — still out per v0.2 boundary):
 *   - Mail.Send                       → v0.3 (Shield-gated send)
 *   - Mail.Send.Shared                → beyond v0.3 (ADR 0001 §D7)
 *   - Files.* / Sites.* / Calendars.* → belong to `m365-graph-mcp-server`
 *
 * The `.Shared` scopes only permit routing — access is still enforced
 * per-mailbox by Exchange (the caller must have Full Access / delegate
 * permissions on the target mailbox). This server does NOT elevate.
 *
 * CI Layer A (permission-mutation surface invariant) treats any
 * *.Manage.All / *.FullControl.All / Application.ReadWrite.All scope as
 * a permission-mutation-class escalation and will fail the build.
 */
export const DELEGATED_SCOPES = [
  "User.Read",
  "Mail.Read",
  "Mail.ReadWrite",
  "Mail.Read.Shared",
  "Mail.ReadWrite.Shared",
  "offline_access",
];

/**
 * The redirect URI registered in the Entra app for the OAuth callback.
 * Must exactly match one of the redirect URIs configured in the Entra
 * app registration. See README § Local development.
 */
export const REDIRECT_URI = "http://localhost:3000/auth/callback";

/**
 * Build the MSAL cache plugin for a given tenant. Exported so tests
 * can drive the load/save lifecycle directly with a fake
 * TokenCacheContext + spied keychain store.
 */
export function makeCachePlugin(tenantId: string): ICachePlugin {
  const store = getTokenStore(tenantId);
  return {
    async beforeCacheAccess(cacheContext: TokenCacheContext): Promise<void> {
      const data = store.load();
      if (data) {
        cacheContext.tokenCache.deserialize(data);
      }
    },
    async afterCacheAccess(cacheContext: TokenCacheContext): Promise<void> {
      if (cacheContext.cacheHasChanged) {
        store.save(cacheContext.tokenCache.serialize());
      }
    },
  };
}

export function makeMsalClient(): ConfidentialClientApplication {
  const tenantId = process.env.M365_MAIL_TENANT_ID ?? "";
  const config: Configuration = {
    auth: {
      clientId: process.env.M365_MAIL_CLIENT_ID ?? "",
      clientSecret: process.env.M365_MAIL_CLIENT_SECRET ?? "",
      authority: `https://login.microsoftonline.com/${tenantId}`,
    },
    cache: {
      cachePlugin: makeCachePlugin(tenantId),
    },
  };
  return new ConfidentialClientApplication(config);
}

/**
 * Acquire an access token silently from the cache. Refreshes via the
 * cached refresh token if the access token is expired. Throws if no
 * cached account exists — the caller should run `npm run setup` to
 * complete the initial OAuth flow.
 */
export async function getAccessToken(
  client: ConfidentialClientApplication,
): Promise<string> {
  const cache = client.getTokenCache();
  const accounts = await cache.getAllAccounts();
  if (accounts.length === 0) {
    throw new Error(
      "No cached account found in the keychain. Run `npm run setup` (or " +
        "`m365-mail-mcp-server setup`) once to complete the OAuth flow.",
    );
  }
  const result = await client.acquireTokenSilent({
    account: accounts[0],
    scopes: DELEGATED_SCOPES,
  });
  if (!result?.accessToken) {
    throw new Error(
      "Silent token acquisition returned no access token. The refresh " +
        "token may have been revoked or expired. Re-run `npm run setup`.",
    );
  }
  return result.accessToken;
}
