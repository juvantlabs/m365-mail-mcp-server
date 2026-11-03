/**
 * Contract tests against the REAL @azure/msal-node (no MSAL mocks).
 *
 * Only the edges are faked: the OS keychain (in-memory map behind the
 * real cache plugin), the identity provider (offline authority metadata
 * + a fake token endpoint) and, for the Graph tests, `fetch`. Everything
 * in between (request building, the auth-code exchange, ID-token
 * handling, cache (de)serialization, silent acquisition and refresh,
 * the Graph auth bridge) is the shipped code.
 *
 * These pin the behaviours an msal-node upgrade can silently break:
 *   1. setup's authorization URL: response_mode=query (the localhost
 *      listener reads GET ?code=) and no nonce.
 *   2. setup's code exchange -> keychain write -> silent token on restart.
 *   3. msal-node >= 5.6 strict ID-token nonce validation.
 *   4. a keychain cache written by an OLDER msal-node (fixtures from the
 *      5.x line, see tests/fixtures/msal-legacy-cache/) still returns the
 *      cached token for the full scope set, including the `.Shared`
 *      scopes, and refreshes with its refresh token.
 *   5. that token reaches Graph as the bearer for both the caller's own
 *      mailbox (`/me`) and a shared mailbox (`/users/{upn}`): one
 *      delegated token covers both, no second acquisition.
 */
import { readFileSync } from "node:fs";
import path from "node:path";

import type { INetworkModule, NetworkRequestOptions } from "@azure/msal-node";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

const { keychain } = vi.hoisted(() => ({ keychain: new Map<string, string>() }));

vi.mock("@napi-rs/keyring", () => ({
  // Vitest 4: constructor mocks need a `function`/class implementation.
  Entry: vi.fn().mockImplementation(function (service: string, account: string) {
    const key = `${service}/${account}`;
    return {
      getPassword: (): string => {
        const v = keychain.get(key);
        if (v === undefined) throw new Error("no entry");
        return v;
      },
      setPassword: (v: string): void => {
        keychain.set(key, v);
      },
      deletePassword: (): void => {
        keychain.delete(key);
      },
    };
  }),
}));

import {
  DELEGATED_SCOPES,
  authCodeTokenRequest,
  authCodeUrlRequest,
  getAccessToken,
  makeMsalClient,
} from "../../src/auth/msal.js";
import { makeGraphClient } from "../../src/client/graph.js";
import { listMessagesTool } from "../../src/tools/list_messages.js";

// Keep in sync with tests/fixtures/msal-legacy-cache/generate.mjs.
const TENANT = "11111111-1111-1111-1111-111111111111";
const CLIENT_ID = "22222222-2222-2222-2222-222222222222";
const OID = "33333333-3333-3333-3333-333333333333";
const HOST = "login.microsoftonline.com";
const KEYCHAIN_KEY = `juvantlabs-m365-mail-mcp-server/tenant:${TENANT}`;
const LEGACY_FIXTURES = ["5.3.1", "5.6.0"] as const;

const OFFLINE_METADATA = {
  authorityMetadata: JSON.stringify({
    token_endpoint: `https://${HOST}/${TENANT}/oauth2/v2.0/token`,
    authorization_endpoint: `https://${HOST}/${TENANT}/oauth2/v2.0/authorize`,
    end_session_endpoint: `https://${HOST}/${TENANT}/oauth2/v2.0/logout`,
    issuer: `https://${HOST}/${TENANT}/v2.0`,
    jwks_uri: `https://${HOST}/${TENANT}/discovery/v2.0/keys`,
  }),
  cloudDiscoveryMetadata: JSON.stringify({
    tenant_discovery_endpoint: `https://${HOST}/${TENANT}/v2.0/.well-known/openid-configuration`,
    "api-version": "1.1",
    metadata: [
      {
        preferred_network: HOST,
        preferred_cache: "login.windows.net",
        aliases: [HOST, "login.windows.net", "login.microsoft.com", "sts.windows.net"],
      },
    ],
  }),
};

const b64url = (o: unknown): string => Buffer.from(JSON.stringify(o)).toString("base64url");

function idToken(extraClaims: Record<string, unknown> = {}): string {
  const now = Math.floor(Date.now() / 1000);
  return [
    b64url({ alg: "none", typ: "JWT" }),
    b64url({
      aud: CLIENT_ID,
      iss: `https://${HOST}/${TENANT}/v2.0`,
      tid: TENANT,
      oid: OID,
      sub: "subject",
      preferred_username: "user@example.test",
      name: "Example User",
      iat: now,
      nbf: now,
      exp: now + 3600,
      ver: "2.0",
      ...extraClaims,
    }),
    "sig",
  ].join(".");
}

interface FakeIdp {
  networkClient: INetworkModule;
  posts: Array<{ url: string; body: URLSearchParams }>;
}

function fakeIdp(opts: { accessToken: string; idTokenClaims?: Record<string, unknown> }): FakeIdp {
  const posts: FakeIdp["posts"] = [];
  const networkClient: INetworkModule = {
    async sendGetRequestAsync(url: string) {
      throw new Error(`unexpected network GET ${url}`);
    },
    async sendPostRequestAsync(url: string, options?: NetworkRequestOptions) {
      posts.push({ url, body: new URLSearchParams(String(options?.body ?? "")) });
      return {
        status: 200,
        headers: {},
        body: {
          token_type: "Bearer",
          scope: DELEGATED_SCOPES.filter((s) => s !== "offline_access").join(" "),
          expires_in: 3600,
          ext_expires_in: 3600,
          access_token: opts.accessToken,
          refresh_token: "rotated-refresh-token",
          id_token: idToken(opts.idTokenClaims),
          client_info: b64url({ uid: OID, utid: TENANT }),
        },
      } as never;
    },
  };
  return { networkClient, posts };
}

function client(idp: FakeIdp): ReturnType<typeof makeMsalClient> {
  return makeMsalClient({ ...OFFLINE_METADATA, system: { networkClient: idp.networkClient } });
}

type CacheJson = Record<string, Record<string, Record<string, unknown>>>;

function loadFixture(version: string): { generated_by: string; cache: CacheJson } {
  return JSON.parse(
    readFileSync(
      path.resolve(__dirname, `../fixtures/msal-legacy-cache/cache-msal-node-${version}.json`),
      "utf8",
    ),
  ) as { generated_by: string; cache: CacheJson };
}

/**
 * The fixtures store a placeholder instead of the JWT-shaped synthetic
 * ID token (secret scanners treat it as a credential). Rebuild it here;
 * every key and other field is exactly as the old msal-node wrote it.
 */
function legacyCache(version: string): CacheJson {
  const cache = structuredClone(loadFixture(version).cache);
  for (const entry of Object.values(cache.IdToken)) {
    expect(entry.secret).toBe("<synthetic-id-token>");
    entry.secret = idToken();
  }
  return cache;
}

const originalEnv = { ...process.env };

beforeEach(() => {
  keychain.clear();
  process.env.M365_MAIL_TENANT_ID = TENANT;
  process.env.M365_MAIL_CLIENT_ID = CLIENT_ID;
  process.env.M365_MAIL_CLIENT_SECRET = "test-secret";
});

afterEach(() => {
  process.env = { ...originalEnv };
  vi.unstubAllGlobals();
});

describe("setup authorization URL (real MSAL)", () => {
  it("uses response_mode=query, sends no nonce, and carries the setup parameters", async () => {
    const idp = fakeIdp({ accessToken: "unused" });
    const url = new URL(await client(idp).getAuthCodeUrl(authCodeUrlRequest()));

    expect(`${url.origin}${url.pathname}`).toBe(`https://${HOST}/${TENANT}/oauth2/v2.0/authorize`);
    expect(url.searchParams.get("response_mode")).toBe("query");
    expect(url.searchParams.has("nonce")).toBe(false);
    expect(url.searchParams.get("redirect_uri")).toBe("http://localhost:3000/auth/callback");
    expect(url.searchParams.get("prompt")).toBe("select_account");
    expect(url.searchParams.get("client_id")).toBe(CLIENT_ID);
    const scopes = (url.searchParams.get("scope") ?? "").split(" ");
    for (const s of DELEGATED_SCOPES) expect(scopes).toContain(s);
    expect(scopes).toContain("Mail.Read.Shared");
    expect(scopes).toContain("Mail.ReadWrite.Shared");
    expect(idp.posts).toHaveLength(0);
  });
});

describe("setup code exchange -> keychain -> silent token (real MSAL)", () => {
  it("persists the cache to the keychain and serves the token silently after a restart", async () => {
    const idp = fakeIdp({ accessToken: "fresh-access-token" });
    const result = await client(idp).acquireTokenByCode(authCodeTokenRequest("auth-code-123"));

    expect(result.account?.username).toBe("user@example.test");
    expect(idp.posts).toHaveLength(1);
    const tokenUrl = new URL(idp.posts[0].url);
    expect(`${tokenUrl.origin}${tokenUrl.pathname}`).toBe(
      `https://${HOST}/${TENANT}/oauth2/v2.0/token`,
    );
    expect(idp.posts[0].body.get("grant_type")).toBe("authorization_code");
    expect(idp.posts[0].body.get("code")).toBe("auth-code-123");
    expect(idp.posts[0].body.get("redirect_uri")).toBe("http://localhost:3000/auth/callback");
    expect(keychain.has(KEYCHAIN_KEY)).toBe(true);

    // New client instance = server restart reading the keychain.
    const restarted = fakeIdp({ accessToken: "should-not-be-requested" });
    await expect(getAccessToken(client(restarted))).resolves.toBe("fresh-access-token");
    expect(restarted.posts).toHaveLength(0);
  });

  it("rejects an ID token carrying a nonce the request never sent (strict nonce validation)", async () => {
    const idp = fakeIdp({ accessToken: "x", idTokenClaims: { nonce: "unexpected" } });
    await expect(
      client(idp).acquireTokenByCode(authCodeTokenRequest("auth-code-123")),
    ).rejects.toMatchObject({ errorCode: "nonce_mismatch" });
    expect(keychain.has(KEYCHAIN_KEY)).toBe(false);
  });

  it("throws the setup hint when the keychain holds no account", async () => {
    await expect(getAccessToken(client(fakeIdp({ accessToken: "x" })))).rejects.toThrow(
      "Run `npm run setup`",
    );
  });
});

describe.each(LEGACY_FIXTURES)("keychain cache written by msal-node %s", (version) => {
  it(`was generated by msal-node ${version} with the full v0.2 scope set`, () => {
    const fixture = loadFixture(version);
    expect(fixture.generated_by).toBe(`@azure/msal-node ${version}`);
    const targets = Object.values(fixture.cache.AccessToken).map((at) => String(at.target));
    expect(targets).toHaveLength(1);
    for (const s of DELEGATED_SCOPES.filter((x) => x !== "offline_access")) {
      expect(targets[0].split(" ")).toContain(s);
    }
  });

  it("returns the cached access token with no network call", async () => {
    keychain.set(KEYCHAIN_KEY, JSON.stringify(legacyCache(version)));
    const idp = fakeIdp({ accessToken: "should-not-be-requested" });
    await expect(getAccessToken(client(idp))).resolves.toBe("legacy-access-token");
    expect(idp.posts).toHaveLength(0);
  });

  it("refreshes an expired legacy access token with the legacy refresh token", async () => {
    const cache = legacyCache(version);
    for (const at of Object.values(cache.AccessToken)) {
      at.expires_on = "1000000000";
      at.extended_expires_on = "1000000000";
    }
    keychain.set(KEYCHAIN_KEY, JSON.stringify(cache));

    const idp = fakeIdp({ accessToken: "refreshed-access-token" });
    await expect(getAccessToken(client(idp))).resolves.toBe("refreshed-access-token");
    expect(idp.posts).toHaveLength(1);
    expect(idp.posts[0].body.get("grant_type")).toBe("refresh_token");
    expect(idp.posts[0].body.get("refresh_token")).toBe("legacy-refresh-token");
    // The refresh requests the .Shared scopes too, so shared-mailbox
    // calls keep working after a refresh.
    const refreshScopes = (idp.posts[0].body.get("scope") ?? "").split(" ");
    expect(refreshScopes).toContain("Mail.Read.Shared");
    expect(refreshScopes).toContain("Mail.ReadWrite.Shared");
    // The refreshed cache is written back to the keychain.
    expect(keychain.get(KEYCHAIN_KEY)).toContain("refreshed-access-token");
  });
});

describe("legacy keychain token -> Graph bearer, own and shared mailbox (real MSAL)", () => {
  function stubGraphFetch(): Array<{ url: string; authorization: string | null }> {
    const calls: Array<{ url: string; authorization: string | null }> = [];
    vi.stubGlobal(
      "fetch",
      vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
        const url = typeof input === "string" ? input : input instanceof URL ? input.href : input.url;
        const headers = new Headers(
          init?.headers ?? (typeof input === "object" && "headers" in input ? input.headers : {}),
        );
        calls.push({ url, authorization: headers.get("Authorization") });
        return new Response(JSON.stringify({ value: [] }), {
          status: 200,
          headers: { "Content-Type": "application/json" },
        });
      }),
    );
    return calls;
  }

  it.each([
    { label: "own mailbox", sharedUser: undefined, pathPrefix: "/v1.0/me/messages" },
    {
      label: "shared mailbox",
      sharedUser: "shared@example.test",
      pathPrefix: "/v1.0/users/shared%40example.test/messages",
    },
  ])("routes $label with the cached delegated token", async ({ sharedUser, pathPrefix }) => {
    process.env.M365_MAIL_ALLOWED_SHARED_USERS = "shared@example.test";
    keychain.set(KEYCHAIN_KEY, JSON.stringify(legacyCache("5.3.1")));
    const idp = fakeIdp({ accessToken: "should-not-be-requested" });
    const calls = stubGraphFetch();

    const graph = makeGraphClient(client(idp));
    const res = await listMessagesTool.handler(graph, {
      limit: 1,
      ...(sharedUser !== undefined ? { shared_user: sharedUser } : {}),
    });

    expect(JSON.parse(res.content[0].text)).toMatchObject({
      shared_user: sharedUser ?? null,
      count: 0,
    });
    expect(calls).toHaveLength(1);
    const url = new URL(calls[0].url);
    expect(url.host).toBe("graph.microsoft.com");
    expect(url.pathname).toBe(pathPrefix);
    expect(calls[0].authorization).toBe("Bearer legacy-access-token");
    expect(idp.posts).toHaveLength(0);
  });
});
