/**
 * MSAL Node client factory + cache plugin wiring.
 *
 * The server runs on the user's own machine with delegated permissions,
 * so it is a PUBLIC client: `PublicClientApplication`, no client secret,
 * no certificate. Nothing secret lives on disk except the MSAL token
 * cache, which is kept in the OS keychain (`src/auth/keyring.ts`).
 *
 * Initial sign-in (`setup`):
 *   - default: interactive authorization code + PKCE on a loopback
 *     redirect (`http://localhost:<random free port>`), handled by
 *     MSAL Node's `acquireTokenInteractive`;
 *   - fallback for headless machines: device code flow
 *     (`setup --device-code`).
 * Afterwards the server acquires tokens silently from the cached
 * refresh token. We never roll our own OAuth: MSAL Node handles PKCE,
 * refresh, revocation and edge cases.
 *
 * NOTE: this server intentionally uses its OWN app registration (env
 * vars `M365_MAIL_TENANT_ID`, `M365_MAIL_CLIENT_ID`) distinct from the
 * sibling `m365-graph-mcp-server`. Isolation is deliberate: a scope
 * misconfiguration in one server MUST NOT reuse the other's app and
 * silently escalate permissions.
 */

import {
  PublicClientApplication,
  type Configuration,
  type DeviceCodeRequest,
  type ICachePlugin,
  type InteractiveRequest,
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
 *                              unattended after `setup`
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

const SUCCESS_HTML =
  "<html><body><h2>Authentication successful</h2>" +
  "<p>You can close this tab and return to the terminal.</p></body></html>";
const ERROR_HTML =
  "<html><body><h2>Authentication failed</h2>" +
  "<p>Return to the terminal for details.</p></body></html>";

/**
 * Interactive request used by `setup`. MSAL Node starts a one-shot
 * loopback listener on a free port of 127.0.0.1, builds the
 * authorization URL with redirect `http://localhost:<port>`, a PKCE
 * S256 challenge and `response_mode=form_post`, calls `openBrowser`,
 * then redeems the code with the PKCE verifier. No redirect URI is
 * passed: Entra accepts any port on `http://localhost` for the
 * "Mobile and desktop applications" platform. No nonce is sent, so the
 * strict ID-token nonce check (msal-node >= 5.6) has nothing to compare.
 * Pinned by tests/unit/msal_contract.test.ts.
 */
export function interactiveRequest(
  openBrowser: (url: string) => Promise<void>,
): InteractiveRequest {
  return {
    scopes: DELEGATED_SCOPES,
    openBrowser,
    prompt: "select_account",
    successTemplate: SUCCESS_HTML,
    errorTemplate: ERROR_HTML,
  };
}

/**
 * Device-code request used by `setup --device-code`, for machines with
 * no local browser. Requires "Allow public client flows" on the Entra
 * app registration.
 */
export function deviceCodeRequest(onMessage: (message: string) => void): DeviceCodeRequest {
  return {
    scopes: DELEGATED_SCOPES,
    deviceCodeCallback: (response) => {
      const message =
        response.message ||
        (response.userCode
          ? `Open ${response.verificationUri} and enter the code ${response.userCode}.`
          : "");
      if (message) onMessage(message);
    },
  };
}

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

/**
 * Optional, additive overrides. Production calls `makeMsalClient()` with
 * none; tests pass offline authority/cloud-discovery metadata and a fake
 * network module so the real MSAL code paths run without a tenant.
 * Client ID, authority and cache plugin are never overridable here.
 */
export interface MsalClientOverrides {
  authorityMetadata?: string;
  cloudDiscoveryMetadata?: string;
  system?: Configuration["system"];
}

export function makeMsalClient(overrides: MsalClientOverrides = {}): PublicClientApplication {
  const tenantId = process.env.M365_MAIL_TENANT_ID ?? "";
  const config: Configuration = {
    auth: {
      clientId: process.env.M365_MAIL_CLIENT_ID ?? "",
      authority: `https://login.microsoftonline.com/${tenantId}`,
      ...(overrides.authorityMetadata !== undefined
        ? { authorityMetadata: overrides.authorityMetadata }
        : {}),
      ...(overrides.cloudDiscoveryMetadata !== undefined
        ? { cloudDiscoveryMetadata: overrides.cloudDiscoveryMetadata }
        : {}),
    },
    cache: {
      cachePlugin: makeCachePlugin(tenantId),
    },
    ...(overrides.system !== undefined ? { system: overrides.system } : {}),
  };
  return new PublicClientApplication(config);
}

/**
 * Acquire an access token silently from the cache. Refreshes via the
 * cached refresh token if the access token is expired. Throws if no
 * cached account exists: the user should run `setup` once to complete
 * the initial sign-in.
 */
export async function getAccessToken(
  client: PublicClientApplication,
): Promise<string> {
  const cache = client.getTokenCache();
  const accounts = await cache.getAllAccounts();
  if (accounts.length === 0) {
    throw new Error(
      "No cached account found in the keychain. Run `npx -y " +
        "@juvantlabs/m365-mail-mcp-server setup` once (with the same " +
        "M365_MAIL_* environment variables) to sign in.",
    );
  }
  let result;
  try {
    result = await client.acquireTokenSilent({
      account: accounts[0],
      scopes: DELEGATED_SCOPES,
    });
  } catch (err) {
    // Typical causes: refresh token expired or revoked, consent changed,
    // or a token cached by a pre-0.3.0 (client-secret) version that the
    // public client cannot redeem. All are fixed by signing in again.
    const reason = err instanceof Error ? err.message : String(err);
    throw new Error(
      `Silent token acquisition failed: ${reason}. Re-run \`npx -y ` +
        "@juvantlabs/m365-mail-mcp-server setup` to sign in again (required " +
        "once after upgrading from a version older than 0.3.0).",
    );
  }
  if (!result?.accessToken) {
    throw new Error(
      "Silent token acquisition returned no access token. The refresh " +
        "token may have been revoked or expired. Re-run `setup`.",
    );
  }
  return result.accessToken;
}
