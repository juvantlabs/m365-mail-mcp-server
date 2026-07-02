import type { Client } from "@microsoft/microsoft-graph-client";
import { describe, expect, it, vi } from "vitest";

import { updateDraftTool } from "../../src/tools/update_draft.js";

function makeClient(opts: {
  currentIsDraft: boolean;
  patched?: unknown;
}): {
  apiCalls: string[];
  patchBodies: Record<string, unknown>[];
  client: Client;
} {
  const apiCalls: string[] = [];
  const patchBodies: Record<string, unknown>[] = [];
  const patched = opts.patched ?? { id: "m1", subject: "[agent-draft] Updated" };
  const get = vi.fn().mockResolvedValue({
    id: "m1",
    isDraft: opts.currentIsDraft,
    subject: "current",
  });
  const select = vi.fn().mockReturnValue({ get });
  const patch = vi.fn().mockImplementation((b: Record<string, unknown>) => {
    patchBodies.push(b);
    return Promise.resolve(patched);
  });
  const api = vi.fn().mockImplementation((p: string) => {
    apiCalls.push(p);
    return { select, patch };
  });
  return { apiCalls, patchBodies, client: { api } as unknown as Client };
}

describe("updateDraftTool handler", () => {
  it("requires message_id", async () => {
    const { client } = makeClient({ currentIsDraft: true });
    await expect(updateDraftTool.handler(client, {})).rejects.toThrow(
      "'message_id' must be a non-empty string",
    );
  });

  it("rejects when nothing to patch (only message_id supplied)", async () => {
    const { client } = makeClient({ currentIsDraft: true });
    await expect(
      updateDraftTool.handler(client, { message_id: "m1" }),
    ).rejects.toThrow("at least one field to update");
  });

  it("refuses to PATCH a non-draft message (Shield safety)", async () => {
    const { client } = makeClient({ currentIsDraft: false });
    await expect(
      updateDraftTool.handler(client, { message_id: "m1", subject: "x" }),
    ).rejects.toThrow(/not a draft/);
  });

  it("Shield C4 — re-applies [agent-draft] to any patched subject", async () => {
    const { patchBodies, client } = makeClient({ currentIsDraft: true });
    await updateDraftTool.handler(client, { message_id: "m1", subject: "Revised" });
    expect(patchBodies[0].subject).toBe("[agent-draft] Revised");
  });

  it("Shield C4 — idempotent on already-marked subject", async () => {
    const { patchBodies, client } = makeClient({ currentIsDraft: true });
    await updateDraftTool.handler(client, {
      message_id: "m1",
      subject: "[agent-draft] Revised",
    });
    expect(patchBodies[0].subject).toBe("[agent-draft] Revised");
  });

  it("does NOT touch subject if the caller did not patch subject", async () => {
    const { patchBodies, client } = makeClient({ currentIsDraft: true });
    await updateDraftTool.handler(client, {
      message_id: "m1",
      body: "new body",
    });
    expect(patchBodies[0].subject).toBeUndefined();
  });

  it("category is 'write_idempotent'", () => {
    expect(updateDraftTool.category).toBe("write_idempotent");
  });

  // ─── v0.2: shared_user routing ─────────────────────────────────────
  it("routes both the pre-flight GET and the PATCH via /users/{upn}/messages/{id}", async () => {
    const { apiCalls, client } = makeClient({ currentIsDraft: true });
    await updateDraftTool.handler(client, {
      message_id: "m1",
      subject: "Revised",
      shared_user: "finance@juvant.io",
    });
    // Two api() calls: one for pre-flight GET, one for PATCH — both
    // MUST hit the shared mailbox path.
    expect(apiCalls[0]).toBe("/users/finance%40juvant.io/messages/m1");
    expect(apiCalls[1]).toBe("/users/finance%40juvant.io/messages/m1");
  });

  it("echoes shared_user in the response", async () => {
    const { client } = makeClient({ currentIsDraft: true });
    const resp = await updateDraftTool.handler(client, {
      message_id: "m1",
      subject: "x",
      shared_user: "finance@juvant.io",
    });
    const parsed = JSON.parse((resp.content[0] as { type: string; text: string }).text);
    expect(parsed.shared_user).toBe("finance@juvant.io");
  });

  it("rejects a malformed shared_user before hitting Graph", async () => {
    const { apiCalls, client } = makeClient({ currentIsDraft: true });
    await expect(
      updateDraftTool.handler(client, {
        message_id: "m1",
        subject: "x",
        shared_user: "bad",
      }),
    ).rejects.toThrow(/UPN/);
    expect(apiCalls).toEqual([]);
  });
});
