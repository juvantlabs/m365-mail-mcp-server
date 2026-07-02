import type { Client } from "@microsoft/microsoft-graph-client";
import { describe, expect, it, vi } from "vitest";

import { markReadTool } from "../../src/tools/mark_read.js";

function captureRequest(returnValue: unknown): {
  apiCalls: string[];
  patchBodies: Record<string, unknown>[];
  client: Client;
} {
  const apiCalls: string[] = [];
  const patchBodies: Record<string, unknown>[] = [];
  const patch = vi.fn().mockImplementation((b: Record<string, unknown>) => {
    patchBodies.push(b);
    return Promise.resolve(returnValue);
  });
  const api = vi.fn().mockImplementation((p: string) => {
    apiCalls.push(p);
    return { patch };
  });
  return { apiCalls, patchBodies, client: { api } as unknown as Client };
}

describe("markReadTool handler", () => {
  it("requires message_id", async () => {
    const { client } = captureRequest({ isRead: true });
    await expect(markReadTool.handler(client, {})).rejects.toThrow(
      "'message_id' must be a non-empty string",
    );
  });

  it("PATCHes { isRead: true } by default", async () => {
    const { patchBodies, client } = captureRequest({ isRead: true });
    await markReadTool.handler(client, { message_id: "m1" });
    expect(patchBodies[0]).toEqual({ isRead: true });
  });

  it("PATCHes { isRead: false } when is_read=false", async () => {
    const { patchBodies, client } = captureRequest({ isRead: false });
    await markReadTool.handler(client, { message_id: "m1", is_read: false });
    expect(patchBodies[0]).toEqual({ isRead: false });
  });

  it("URL-encodes message_id in the path", async () => {
    const { apiCalls, client } = captureRequest({ isRead: true });
    await markReadTool.handler(client, { message_id: "with space" });
    expect(apiCalls[0]).toBe("/me/messages/with%20space");
  });

  it("rejects non-boolean is_read", async () => {
    const { client } = captureRequest({ isRead: true });
    await expect(
      markReadTool.handler(client, { message_id: "m1", is_read: "yes" }),
    ).rejects.toThrow("'is_read' must be a boolean");
  });

  it("returns message_id + is_read from the server response", async () => {
    const { client } = captureRequest({ isRead: false });
    const resp = await markReadTool.handler(client, { message_id: "m1", is_read: false });
    const parsed = JSON.parse((resp.content[0] as { type: string; text: string }).text);
    expect(parsed.message_id).toBe("m1");
    expect(parsed.is_read).toBe(false);
  });

  it("category is 'write_idempotent'", () => {
    expect(markReadTool.category).toBe("write_idempotent");
  });

  // ─── v0.2: shared_user routing ─────────────────────────────────────
  it("PATCHes /users/{upn}/messages/{id} when shared_user is set", async () => {
    const { apiCalls, client } = captureRequest({ isRead: true });
    await markReadTool.handler(client, {
      message_id: "m1",
      shared_user: "finance@juvant.io",
    });
    expect(apiCalls[0]).toBe("/users/finance%40juvant.io/messages/m1");
  });

  it("echoes shared_user in the response", async () => {
    const { client } = captureRequest({ isRead: true });
    const resp = await markReadTool.handler(client, {
      message_id: "m1",
      shared_user: "finance@juvant.io",
    });
    const parsed = JSON.parse((resp.content[0] as { type: string; text: string }).text);
    expect(parsed.shared_user).toBe("finance@juvant.io");
  });

  it("rejects a malformed shared_user before hitting Graph", async () => {
    const { apiCalls, client } = captureRequest({ isRead: true });
    await expect(
      markReadTool.handler(client, { message_id: "m1", shared_user: "bad" }),
    ).rejects.toThrow(/UPN/);
    expect(apiCalls).toEqual([]);
  });
});
