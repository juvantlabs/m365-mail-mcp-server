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
    _injectExpiredConfirmation(
      "expired-tok",
      "m365-mail:delete_message",
      { message_id: "m1" },
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
