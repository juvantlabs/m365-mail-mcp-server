import type { Client } from "@microsoft/microsoft-graph-client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import {
  _injectExpiredConfirmation,
  _resetConfirmationTokens,
} from "../../src/auth/confirmation_tokens.js";
import { deleteMessageTool } from "../../src/tools/delete_message.js";

beforeEach(() => _resetConfirmationTokens());
afterEach(() => _resetConfirmationTokens());

function makeClient(messageMetadata: unknown): {
  apiCalls: string[];
  deleteCalled: number;
  client: Client;
} {
  const apiCalls: string[] = [];
  let deleteCalled = 0;
  const get = vi.fn().mockResolvedValue(messageMetadata);
  const select = vi.fn().mockReturnValue({ get });
  const del = vi.fn().mockImplementation(() => {
    deleteCalled++;
    return Promise.resolve(undefined);
  });
  const api = vi.fn().mockImplementation((p: string) => {
    apiCalls.push(p);
    return { get, select, delete: del };
  });
  return {
    apiCalls,
    get deleteCalled() {
      return deleteCalled;
    },
    client: { api } as unknown as Client,
  };
}

describe("deleteMessageTool — phase 1 (preview)", () => {
  it("requires message_id", async () => {
    const { client } = makeClient({});
    await expect(deleteMessageTool.handler(client, {})).rejects.toThrow(
      "'message_id' must be a non-empty string",
    );
  });

  it("returns a preview + confirmation_token, does NOT delete", async () => {
    const m = {
      id: "m1",
      subject: "Invoice",
      from: { emailAddress: { address: "supplier@x.com" } },
      receivedDateTime: "2026-06-15T10:00:00Z",
      isDraft: false,
      hasAttachments: true,
      parentFolderId: "inbox",
      webLink: "https://outlook/m1",
    };
    const c = makeClient(m);
    const resp = await deleteMessageTool.handler(c.client, { message_id: "m1" });

    const parsed = JSON.parse((resp.content[0] as { type: string; text: string }).text);
    expect(parsed.preview).toBeDefined();
    expect(parsed.preview.item.id).toBe("m1");
    expect(parsed.preview.item.subject).toBe("Invoice");
    expect(parsed.preview.item.from_email).toBe("supplier@x.com");
    expect(parsed.preview.item.has_attachments).toBe(true);
    expect(parsed.preview.confirmation_token).toMatch(/^[0-9a-f]{32}$/);
    expect(parsed.preview.expires_in_seconds).toBe(300);
    expect(c.deleteCalled).toBe(0);
  });

  it("treats an empty-string confirmation_token as absent (phase 1)", async () => {
    const c = makeClient({ id: "m1", subject: "x" });
    const resp = await deleteMessageTool.handler(c.client, {
      message_id: "m1",
      confirmation_token: "",
    });
    const parsed = JSON.parse((resp.content[0] as { type: string; text: string }).text);
    expect(parsed.preview).toBeDefined();
    expect(c.deleteCalled).toBe(0);
  });
});

describe("deleteMessageTool — phase 2 (execute)", () => {
  it("rejects an unknown token with token_unknown", async () => {
    const c = makeClient({});
    await expect(
      deleteMessageTool.handler(c.client, {
        message_id: "m1",
        confirmation_token: "deadbeef".repeat(4),
      }),
    ).rejects.toThrow("token_unknown");
    expect(c.deleteCalled).toBe(0);
  });

  it("rejects a token issued for a different message_id (spec_mismatch)", async () => {
    const m = { id: "m1", subject: "x" };
    const c = makeClient(m);

    // Phase 1 for message A
    const phase1 = await deleteMessageTool.handler(c.client, { message_id: "A" });
    const { confirmation_token } = JSON.parse(
      (phase1.content[0] as { type: string; text: string }).text,
    ).preview;

    // Phase 2: try to reuse the token for message B
    await expect(
      deleteMessageTool.handler(c.client, {
        message_id: "B",
        confirmation_token,
      }),
    ).rejects.toThrow("spec_mismatch");
    expect(c.deleteCalled).toBe(0);
  });

  it("rejects an expired token with token_expired", async () => {
    const c = makeClient({ id: "m1" });

    // Inject an already-expired token directly, bypassing phase 1.
    // v0.2: the spec MUST include shared_user (null for own mailbox)
    // to match what the handler will canonicalize on lookup.
    _injectExpiredConfirmation(
      "expired-tok",
      "m365-mail:delete_message",
      { message_id: "m1", shared_user: null },
    );

    await expect(
      deleteMessageTool.handler(c.client, {
        message_id: "m1",
        confirmation_token: "expired-tok",
      }),
    ).rejects.toThrow("token_expired");
    expect(c.deleteCalled).toBe(0);
  });

  it("executes DELETE on a valid token + matching spec", async () => {
    const m = { id: "m1", subject: "x" };
    const c = makeClient(m);

    const phase1 = await deleteMessageTool.handler(c.client, { message_id: "m1" });
    const { confirmation_token } = JSON.parse(
      (phase1.content[0] as { type: string; text: string }).text,
    ).preview;

    const phase2 = await deleteMessageTool.handler(c.client, {
      message_id: "m1",
      confirmation_token,
    });
    const parsed = JSON.parse((phase2.content[0] as { type: string; text: string }).text);
    expect(parsed.deleted.message_id).toBe("m1");
    expect(c.deleteCalled).toBe(1);
  });

  it("token is single-use — second attempt fails token_unknown", async () => {
    const m = { id: "m1", subject: "x" };
    const c = makeClient(m);

    const phase1 = await deleteMessageTool.handler(c.client, { message_id: "m1" });
    const { confirmation_token } = JSON.parse(
      (phase1.content[0] as { type: string; text: string }).text,
    ).preview;

    await deleteMessageTool.handler(c.client, {
      message_id: "m1",
      confirmation_token,
    });
    await expect(
      deleteMessageTool.handler(c.client, {
        message_id: "m1",
        confirmation_token,
      }),
    ).rejects.toThrow("token_unknown");
    expect(c.deleteCalled).toBe(1);
  });

  it("classifies as write_irreversible (CI enforcement + registry test)", () => {
    expect(deleteMessageTool.category).toBe("write_irreversible");
  });
});

// ─── v0.2: shared_user in the delete_message spec ────────────────────────
describe("deleteMessageTool — v0.2 shared_user routing + spec hash", () => {
  it("phase-1 preview against a shared mailbox routes to /users/{upn}/messages/{id}", async () => {
    const c = makeClient({
      id: "m1",
      subject: "Invoice",
      from: { emailAddress: { address: "supplier@x.com" } },
    });
    await deleteMessageTool.handler(c.client, {
      message_id: "m1",
      shared_user: "finance@juvant.io",
    });
    expect(c.apiCalls[0]).toBe("/users/finance%40juvant.io/messages/m1");
  });

  it("phase-1 preview echoes shared_user on the preview", async () => {
    const c = makeClient({ id: "m1", subject: "x" });
    const resp = await deleteMessageTool.handler(c.client, {
      message_id: "m1",
      shared_user: "finance@juvant.io",
    });
    const parsed = JSON.parse((resp.content[0] as { type: string; text: string }).text);
    expect(parsed.preview.shared_user).toBe("finance@juvant.io");
  });

  it("phase-2 executes DELETE on the shared mailbox path", async () => {
    const c = makeClient({ id: "m1", subject: "x" });

    const phase1 = await deleteMessageTool.handler(c.client, {
      message_id: "m1",
      shared_user: "finance@juvant.io",
    });
    const { confirmation_token } = JSON.parse(
      (phase1.content[0] as { type: string; text: string }).text,
    ).preview;

    const phase2 = await deleteMessageTool.handler(c.client, {
      message_id: "m1",
      shared_user: "finance@juvant.io",
      confirmation_token,
    });
    const parsed = JSON.parse((phase2.content[0] as { type: string; text: string }).text);
    expect(parsed.deleted.message_id).toBe("m1");
    expect(parsed.deleted.shared_user).toBe("finance@juvant.io");
    expect(c.deleteCalled).toBe(1);
    // Both phase 1 preview GET + phase 2 DELETE go through the same
    // shared mailbox root.
    expect(c.apiCalls[0]).toBe("/users/finance%40juvant.io/messages/m1");
    expect(c.apiCalls[1]).toBe("/users/finance%40juvant.io/messages/m1");
  });

  it("rejects a token issued for own-mailbox but replayed with shared_user (spec_mismatch)", async () => {
    // The load-bearing v0.2 security invariant. Phase-1 preview against
    // /me/messages/A issues a token bound to { message_id: "A",
    // shared_user: null }. Phase-2 with the SAME message_id but a
    // shared_user MUST NOT go through, because that would authorise a
    // delete against a different mailbox from the one previewed.
    const c = makeClient({ id: "A", subject: "own-mailbox message" });

    const phase1 = await deleteMessageTool.handler(c.client, {
      message_id: "A",
    });
    const { confirmation_token } = JSON.parse(
      (phase1.content[0] as { type: string; text: string }).text,
    ).preview;

    await expect(
      deleteMessageTool.handler(c.client, {
        message_id: "A",
        shared_user: "finance@juvant.io",
        confirmation_token,
      }),
    ).rejects.toThrow("spec_mismatch");
    expect(c.deleteCalled).toBe(0);
  });

  it("rejects a token issued for shared_user but replayed against own-mailbox (spec_mismatch)", async () => {
    // Symmetric case: phase-1 against a shared mailbox, phase-2
    // dropping the shared_user, MUST fail.
    const c = makeClient({ id: "A", subject: "shared-mailbox message" });

    const phase1 = await deleteMessageTool.handler(c.client, {
      message_id: "A",
      shared_user: "finance@juvant.io",
    });
    const { confirmation_token } = JSON.parse(
      (phase1.content[0] as { type: string; text: string }).text,
    ).preview;

    await expect(
      deleteMessageTool.handler(c.client, {
        message_id: "A",
        confirmation_token,
      }),
    ).rejects.toThrow("spec_mismatch");
    expect(c.deleteCalled).toBe(0);
  });

  it("rejects a token issued for one shared_user but replayed against another (spec_mismatch)", async () => {
    // Two shared mailboxes on the same tenant with the same message id
    // — the token issued for one MUST NOT authorise a delete on the
    // other.
    const c = makeClient({ id: "A", subject: "cross-mailbox replay" });

    const phase1 = await deleteMessageTool.handler(c.client, {
      message_id: "A",
      shared_user: "finance@juvant.io",
    });
    const { confirmation_token } = JSON.parse(
      (phase1.content[0] as { type: string; text: string }).text,
    ).preview;

    await expect(
      deleteMessageTool.handler(c.client, {
        message_id: "A",
        shared_user: "legal@juvant.io",
        confirmation_token,
      }),
    ).rejects.toThrow("spec_mismatch");
    expect(c.deleteCalled).toBe(0);
  });

  it("rejects a malformed shared_user before hitting Graph in phase 1", async () => {
    const c = makeClient({});
    await expect(
      deleteMessageTool.handler(c.client, {
        message_id: "m1",
        shared_user: "not-a-upn",
      }),
    ).rejects.toThrow(/UPN/);
    expect(c.apiCalls).toEqual([]);
    expect(c.deleteCalled).toBe(0);
  });
});
