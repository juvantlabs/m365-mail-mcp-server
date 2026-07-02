import type { Client } from "@microsoft/microsoft-graph-client";
import { describe, expect, it, vi } from "vitest";

import { createReplyDraftTool } from "../../src/tools/create_reply_draft.js";

function makeClient(opts: {
  createReplyReturns: Record<string, unknown>;
  patchReturns: Record<string, unknown>;
}): {
  apiCalls: string[];
  patchBodies: Record<string, unknown>[];
  client: Client;
} {
  const apiCalls: string[] = [];
  const patchBodies: Record<string, unknown>[] = [];
  const post = vi.fn().mockResolvedValue(opts.createReplyReturns);
  const patch = vi.fn().mockImplementation((b: Record<string, unknown>) => {
    patchBodies.push(b);
    return Promise.resolve(opts.patchReturns);
  });
  const api = vi.fn().mockImplementation((p: string) => {
    apiCalls.push(p);
    return { post, patch };
  });
  return { apiCalls, patchBodies, client: { api } as unknown as Client };
}

describe("createReplyDraftTool handler", () => {
  it("requires message_id", async () => {
    const { client } = makeClient({
      createReplyReturns: { id: "d1" },
      patchReturns: { id: "d1" },
    });
    await expect(createReplyDraftTool.handler(client, {})).rejects.toThrow(
      "'message_id' must be a non-empty string",
    );
  });

  it("POSTs to /me/messages/{id}/createReply by default", async () => {
    const { apiCalls, client } = makeClient({
      createReplyReturns: { id: "d1", subject: "RE: Original" },
      patchReturns: { id: "d1", subject: "[agent-draft] RE: Original" },
    });
    await createReplyDraftTool.handler(client, { message_id: "m1" });
    expect(apiCalls[0]).toBe("/me/messages/m1/createReply");
  });

  it("POSTs to /me/messages/{id}/createReplyAll when reply_all=true", async () => {
    const { apiCalls, client } = makeClient({
      createReplyReturns: { id: "d1", subject: "RE: x" },
      patchReturns: { id: "d1", subject: "[agent-draft] RE: x" },
    });
    await createReplyDraftTool.handler(client, { message_id: "m1", reply_all: true });
    expect(apiCalls[0]).toBe("/me/messages/m1/createReplyAll");
  });

  it("Shield C4 — unconditionally PATCHes subject with marker when caller omits subject", async () => {
    const { patchBodies, client } = makeClient({
      createReplyReturns: { id: "d1", subject: "RE: Weekly sync" },
      patchReturns: { id: "d1", subject: "[agent-draft] RE: Weekly sync" },
    });
    await createReplyDraftTool.handler(client, { message_id: "m1" });
    expect(patchBodies[0].subject).toBe("[agent-draft] RE: Weekly sync");
  });

  it("Shield C4 — marks the caller-supplied subject", async () => {
    const { patchBodies, client } = makeClient({
      createReplyReturns: { id: "d1", subject: "RE: x" },
      patchReturns: { id: "d1", subject: "[agent-draft] My reply" },
    });
    await createReplyDraftTool.handler(client, {
      message_id: "m1",
      subject: "My reply",
    });
    expect(patchBodies[0].subject).toBe("[agent-draft] My reply");
  });

  it("throws if createReply returns no draft id", async () => {
    const { client } = makeClient({
      createReplyReturns: {},
      patchReturns: {},
    });
    await expect(
      createReplyDraftTool.handler(client, { message_id: "m1" }),
    ).rejects.toThrow(/no draft id/);
  });

  it("rejects invalid reply_all type", async () => {
    const { client } = makeClient({
      createReplyReturns: { id: "d1" },
      patchReturns: { id: "d1" },
    });
    await expect(
      createReplyDraftTool.handler(client, { message_id: "m1", reply_all: "yes" }),
    ).rejects.toThrow("'reply_all' must be a boolean");
  });

  it("category is 'write_idempotent'", () => {
    expect(createReplyDraftTool.category).toBe("write_idempotent");
  });

  // ─── v0.2: shared_user routing ─────────────────────────────────────
  it("POSTs createReply on /users/{upn}/messages when shared_user is set", async () => {
    const { apiCalls, client } = makeClient({
      createReplyReturns: { id: "d1", subject: "RE: x" },
      patchReturns: { id: "d1", subject: "[agent-draft] RE: x" },
    });
    await createReplyDraftTool.handler(client, {
      message_id: "m1",
      shared_user: "finance@juvant.io",
    });
    expect(apiCalls[0]).toBe("/users/finance%40juvant.io/messages/m1/createReply");
  });

  it("PATCHes the resulting draft under the same shared mailbox root", async () => {
    // The subject PATCH that applies the Shield marker must hit the
    // shared mailbox's copy of the draft, not /me/messages/{id}.
    const { apiCalls, client } = makeClient({
      createReplyReturns: { id: "reply-draft-1", subject: "RE: x" },
      patchReturns: { id: "reply-draft-1", subject: "[agent-draft] RE: x" },
    });
    await createReplyDraftTool.handler(client, {
      message_id: "m1",
      shared_user: "finance@juvant.io",
    });
    // apiCalls[1] is the PATCH target.
    expect(apiCalls[1]).toBe(
      "/users/finance%40juvant.io/messages/reply-draft-1",
    );
  });

  it("echoes shared_user in the response", async () => {
    const { client } = makeClient({
      createReplyReturns: { id: "d1", subject: "RE: x" },
      patchReturns: { id: "d1", subject: "[agent-draft] RE: x" },
    });
    const resp = await createReplyDraftTool.handler(client, {
      message_id: "m1",
      shared_user: "finance@juvant.io",
    });
    const parsed = JSON.parse((resp.content[0] as { type: string; text: string }).text);
    expect(parsed.shared_user).toBe("finance@juvant.io");
  });

  it("rejects a malformed shared_user before hitting Graph", async () => {
    const { apiCalls, client } = makeClient({
      createReplyReturns: { id: "d1" },
      patchReturns: { id: "d1" },
    });
    await expect(
      createReplyDraftTool.handler(client, {
        message_id: "m1",
        shared_user: "bad",
      }),
    ).rejects.toThrow(/UPN/);
    expect(apiCalls).toEqual([]);
  });
});
