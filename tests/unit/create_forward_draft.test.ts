import type { Client } from "@microsoft/microsoft-graph-client";
import { describe, expect, it, vi } from "vitest";

import { createForwardDraftTool } from "../../src/tools/create_forward_draft.js";

function makeClient(opts: {
  createForwardReturns: Record<string, unknown>;
  patchReturns: Record<string, unknown>;
}): {
  apiCalls: string[];
  patchBodies: Record<string, unknown>[];
  client: Client;
} {
  const apiCalls: string[] = [];
  const patchBodies: Record<string, unknown>[] = [];
  const post = vi.fn().mockResolvedValue(opts.createForwardReturns);
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

describe("createForwardDraftTool handler", () => {
  it("requires message_id", async () => {
    const { client } = makeClient({
      createForwardReturns: { id: "d1" },
      patchReturns: { id: "d1" },
    });
    await expect(createForwardDraftTool.handler(client, {})).rejects.toThrow(
      "'message_id' must be a non-empty string",
    );
  });

  it("POSTs to /me/messages/{id}/createForward", async () => {
    const { apiCalls, client } = makeClient({
      createForwardReturns: { id: "d1", subject: "FW: Original" },
      patchReturns: { id: "d1", subject: "[agent-draft] FW: Original" },
    });
    await createForwardDraftTool.handler(client, { message_id: "m1" });
    expect(apiCalls[0]).toBe("/me/messages/m1/createForward");
  });

  it("Shield C4 — unconditionally marks the subject even when caller omits it", async () => {
    const { patchBodies, client } = makeClient({
      createForwardReturns: { id: "d1", subject: "FW: Weekly sync" },
      patchReturns: { id: "d1" },
    });
    await createForwardDraftTool.handler(client, { message_id: "m1" });
    expect(patchBodies[0].subject).toBe("[agent-draft] FW: Weekly sync");
  });

  it("Shield C4 — marks caller-supplied subject", async () => {
    const { patchBodies, client } = makeClient({
      createForwardReturns: { id: "d1", subject: "FW: x" },
      patchReturns: { id: "d1" },
    });
    await createForwardDraftTool.handler(client, {
      message_id: "m1",
      subject: "FYI please review",
    });
    expect(patchBodies[0].subject).toBe("[agent-draft] FYI please review");
  });

  it("PATCHes recipients when caller supplies them", async () => {
    const { patchBodies, client } = makeClient({
      createForwardReturns: { id: "d1", subject: "FW: x" },
      patchReturns: { id: "d1" },
    });
    await createForwardDraftTool.handler(client, {
      message_id: "m1",
      to: ["ext@partner.com"],
    });
    expect(patchBodies[0].toRecipients).toEqual([
      { emailAddress: { address: "ext@partner.com" } },
    ]);
  });

  it("throws if createForward returns no draft id", async () => {
    const { client } = makeClient({
      createForwardReturns: {},
      patchReturns: {},
    });
    await expect(
      createForwardDraftTool.handler(client, { message_id: "m1" }),
    ).rejects.toThrow(/no draft id/);
  });

  it("category is 'write_idempotent'", () => {
    expect(createForwardDraftTool.category).toBe("write_idempotent");
  });

  // ─── v0.2: shared_user routing ─────────────────────────────────────
  it("POSTs createForward on /users/{upn}/messages when shared_user is set", async () => {
    const { apiCalls, client } = makeClient({
      createForwardReturns: { id: "d1", subject: "FW: x" },
      patchReturns: { id: "d1" },
    });
    await createForwardDraftTool.handler(client, {
      message_id: "m1",
      shared_user: "finance@juvant.io",
    });
    expect(apiCalls[0]).toBe(
      "/users/finance%40juvant.io/messages/m1/createForward",
    );
  });

  it("PATCHes the resulting draft under the same shared mailbox root", async () => {
    const { apiCalls, client } = makeClient({
      createForwardReturns: { id: "fwd-draft-1", subject: "FW: x" },
      patchReturns: { id: "fwd-draft-1" },
    });
    await createForwardDraftTool.handler(client, {
      message_id: "m1",
      shared_user: "finance@juvant.io",
    });
    expect(apiCalls[1]).toBe(
      "/users/finance%40juvant.io/messages/fwd-draft-1",
    );
  });

  it("echoes shared_user in the response", async () => {
    const { client } = makeClient({
      createForwardReturns: { id: "d1", subject: "FW: x" },
      patchReturns: { id: "d1" },
    });
    const resp = await createForwardDraftTool.handler(client, {
      message_id: "m1",
      shared_user: "finance@juvant.io",
    });
    const parsed = JSON.parse((resp.content[0] as { type: string; text: string }).text);
    expect(parsed.shared_user).toBe("finance@juvant.io");
  });

  it("rejects a malformed shared_user before hitting Graph", async () => {
    const { apiCalls, client } = makeClient({
      createForwardReturns: { id: "d1" },
      patchReturns: { id: "d1" },
    });
    await expect(
      createForwardDraftTool.handler(client, {
        message_id: "m1",
        shared_user: "bad",
      }),
    ).rejects.toThrow(/UPN/);
    expect(apiCalls).toEqual([]);
  });
});
