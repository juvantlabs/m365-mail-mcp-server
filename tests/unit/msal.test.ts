import { afterEach, describe, expect, it, vi } from "vitest";

const { entryConstructor } = vi.hoisted(() => ({
  // Vitest 4: `new Entry(...)` needs a constructible implementation
  // (a `function`, not an arrow).
  entryConstructor: vi.fn().mockImplementation(function () {
    return {
      setPassword: vi.fn(),
      getPassword: vi.fn().mockImplementation(() => {
        throw new Error("no entry");
      }),
      deletePassword: vi.fn(),
    };
  }),
}));

vi.mock("@napi-rs/keyring", () => ({
  Entry: entryConstructor,
}));

import { PublicClientApplication } from "@azure/msal-node";

import {
  DELEGATED_SCOPES,
  deviceCodeRequest,
  getAccessToken,
  interactiveRequest,
  makeMsalClient,
} from "../../src/auth/msal.js";

afterEach(() => {
  vi.clearAllMocks();
});

describe("makeMsalClient", () => {
  it("constructs a PublicClientApplication from M365_MAIL_* env vars, with no secret", () => {
    const original = { ...process.env };
    process.env.M365_MAIL_TENANT_ID = "c557607d-995c-4eb7-967b-50c6361fbad9";
    process.env.M365_MAIL_CLIENT_ID = "00000000-0000-0000-0000-000000000001";
    delete process.env.M365_MAIL_CLIENT_SECRET;
    try {
      const client = makeMsalClient();
      expect(client).toBeInstanceOf(PublicClientApplication);
      expect(typeof client.acquireTokenSilent).toBe("function");
      expect(typeof client.acquireTokenInteractive).toBe("function");
      expect(typeof client.acquireTokenByDeviceCode).toBe("function");
    } finally {
      process.env = original;
    }
  });

  it("does not read a leftover M365_MAIL_CLIENT_SECRET", () => {
    const original = { ...process.env };
    process.env.M365_MAIL_TENANT_ID = "c557607d-995c-4eb7-967b-50c6361fbad9";
    process.env.M365_MAIL_CLIENT_ID = "00000000-0000-0000-0000-000000000001";
    process.env.M365_MAIL_CLIENT_SECRET = "leftover";
    try {
      const client = makeMsalClient();
      expect(client).toBeInstanceOf(PublicClientApplication);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const auth = (client as any).config.auth as Record<string, unknown>;
      expect(auth.clientSecret ?? "").toBe("");
    } finally {
      process.env = original;
    }
  });
});

describe("interactiveRequest", () => {
  it("requests DELEGATED_SCOPES with account selection and no fixed redirect URI", () => {
    const openBrowser = vi.fn();
    const req = interactiveRequest(openBrowser);
    expect(req.scopes).toEqual(DELEGATED_SCOPES);
    expect(req.openBrowser).toBe(openBrowser);
    expect(req.prompt).toBe("select_account");
    // MSAL picks a free loopback port; a fixed redirect URI is rejected
    // by PublicClientApplication.acquireTokenInteractive.
    expect(req.redirectUri).toBeUndefined();
    expect(req.nonce).toBeUndefined();
    expect(req.successTemplate).toMatch(/Authentication successful/);
  });
});

describe("deviceCodeRequest", () => {
  it("requests DELEGATED_SCOPES and forwards the user message", () => {
    const onMessage = vi.fn();
    const req = deviceCodeRequest(onMessage);
    expect(req.scopes).toEqual(DELEGATED_SCOPES);
    req.deviceCodeCallback({
      message: "To sign in, use a web browser to open the page and enter the code ABC.",
      userCode: "ABC",
      deviceCode: "dc",
      verificationUri: "https://microsoft.com/devicelogin",
      expiresIn: 900,
      interval: 5,
    });
    expect(onMessage).toHaveBeenCalledWith(expect.stringContaining("ABC"));
  });

  it("falls back to URI + code when the message is missing, and stays silent when both are", () => {
    const onMessage = vi.fn();
    const req = deviceCodeRequest(onMessage);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    req.deviceCodeCallback({ userCode: "XYZ", verificationUri: "https://microsoft.com/devicelogin" } as any);
    expect(onMessage).toHaveBeenCalledWith(
      "Open https://microsoft.com/devicelogin and enter the code XYZ.",
    );
    onMessage.mockClear();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    req.deviceCodeCallback({} as any);
    expect(onMessage).not.toHaveBeenCalled();
  });
});

describe("DELEGATED_SCOPES", () => {
  it("requests User.Read (identity ping)", () => {
    expect(DELEGATED_SCOPES).toContain("User.Read");
  });

  it("requests Mail.Read + Mail.ReadWrite (read + draft write surface)", () => {
    expect(DELEGATED_SCOPES).toContain("Mail.Read");
    expect(DELEGATED_SCOPES).toContain("Mail.ReadWrite");
  });

  it("includes offline_access for refresh tokens", () => {
    expect(DELEGATED_SCOPES).toContain("offline_access");
  });

  it("does NOT request Mail.Send (v0.3, Shield-gated)", () => {
    // Send is a distinct phase (v0.3) gated on human-in-the-loop review.
    // v0.2 MUST NOT be able to send: no scope, no tool.
    expect(DELEGATED_SCOPES).not.toContain("Mail.Send");
    // Mail.Send.Shared is explicitly beyond v0.3 (ADR 0001 §D7) — v0.2
    // adds .Shared read/write scopes but STOPS at send.
    expect(DELEGATED_SCOPES).not.toContain("Mail.Send.Shared");
  });

  it("v0.2 REQUESTS .Shared read/write variants for delegate mailboxes", () => {
    // v0.2 adds shared / delegate mailbox support via a `shared_user`
    // UPN parameter on every tool. Graph enforces per-mailbox access
    // via Exchange; these scopes only permit ROUTING the call. If this
    // assertion flips back, `shared_user` will 403 at Graph.
    expect(DELEGATED_SCOPES).toContain("Mail.Read.Shared");
    expect(DELEGATED_SCOPES).toContain("Mail.ReadWrite.Shared");
  });

  it("does NOT request scopes owned by m365-graph-mcp-server", () => {
    // Cross-server scope leakage is the failure mode this isolation
    // exists to prevent. m365-mail speaks Mail.* only.
    expect(DELEGATED_SCOPES).not.toContain("Files.ReadWrite");
    expect(DELEGATED_SCOPES).not.toContain("Sites.ReadWrite.All");
    expect(DELEGATED_SCOPES).not.toContain("Calendars.ReadWrite");
  });

  it("does NOT request permission-mutation-class scopes (decisions#210)", () => {
    expect(DELEGATED_SCOPES).not.toContain("Sites.Manage.All");
    expect(DELEGATED_SCOPES).not.toContain("Sites.FullControl.All");
    expect(DELEGATED_SCOPES).not.toContain("Application.ReadWrite.All");
  });
});

describe("getAccessToken", () => {
  function makeMsalMock(opts: {
    accounts: Array<{ username: string }>;
    silentResult: { accessToken: string } | null;
    silentThrows?: Error;
  }): {
    getTokenCache: () => { getAllAccounts: () => Promise<unknown[]> };
    acquireTokenSilent: ReturnType<typeof vi.fn>;
  } {
    const acquireTokenSilent = vi.fn();
    if (opts.silentThrows) {
      acquireTokenSilent.mockRejectedValue(opts.silentThrows);
    } else {
      acquireTokenSilent.mockResolvedValue(opts.silentResult);
    }
    return {
      getTokenCache: () => ({
        getAllAccounts: vi.fn().mockResolvedValue(opts.accounts),
      }),
      acquireTokenSilent,
    };
  }

  it("throws when no cached account exists (setup not run)", async () => {
    const msal = makeMsalMock({ accounts: [], silentResult: null });
    await expect(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      getAccessToken(msal as any),
    ).rejects.toThrow(/No cached account/);
  });

  it("returns the access token from acquireTokenSilent", async () => {
    const msal = makeMsalMock({
      accounts: [{ username: "alice@x.com" }],
      silentResult: { accessToken: "the-token" },
    });
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const tok = await getAccessToken(msal as any);
    expect(tok).toBe("the-token");
  });

  it("throws if silent result has no accessToken (refresh failed)", async () => {
    const msal = makeMsalMock({
      accounts: [{ username: "alice@x.com" }],
      silentResult: null,
    });
    await expect(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      getAccessToken(msal as any),
    ).rejects.toThrow(/Silent token acquisition/);
  });

  it("wraps a silent-acquisition failure with the re-run-setup hint", async () => {
    const msal = makeMsalMock({
      accounts: [{ username: "alice@x.com" }],
      silentResult: null,
      silentThrows: new Error("invalid_grant: AADSTS7000218"),
    });
    await expect(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      getAccessToken(msal as any),
    ).rejects.toThrow(/Silent token acquisition failed: invalid_grant: AADSTS7000218\. Re-run .*setup/);
  });

  it("requests the configured DELEGATED_SCOPES", async () => {
    const msal = makeMsalMock({
      accounts: [{ username: "alice@x.com" }],
      silentResult: { accessToken: "tok" },
    });
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    await getAccessToken(msal as any);
    expect(msal.acquireTokenSilent).toHaveBeenCalledWith(
      expect.objectContaining({ scopes: DELEGATED_SCOPES }),
    );
  });
});
