import type { Client } from "@microsoft/microsoft-graph-client";
import { describe, expect, it, vi } from "vitest";

import { searchMessagesTool } from "../../src/tools/search_messages.js";

function captureRequest(returnValue: unknown): {
  apiCalls: string[];
  searchArgs: string[];
  headers: Array<[string, string]>;
  client: Client;
} {
  const apiCalls: string[] = [];
  const searchArgs: string[] = [];
  const headers: Array<[string, string]> = [];
  const get = vi.fn().mockResolvedValue(returnValue);
  const top = vi.fn().mockReturnValue({ get });
  const select = vi.fn().mockReturnValue({ top });
  const search = vi.fn().mockImplementation((q: string) => {
    searchArgs.push(q);
    return { select, top };
  });
  const header = vi.fn().mockImplementation((k: string, v: string) => {
    headers.push([k, v]);
    return { search };
  });
  const api = vi.fn().mockImplementation((path: string) => {
    apiCalls.push(path);
    return { header };
  });
  return { apiCalls, searchArgs, headers, client: { api } as unknown as Client };
}

describe("searchMessagesTool handler", () => {
  it("requires query", async () => {
    const { client } = captureRequest({ value: [] });
    await expect(searchMessagesTool.handler(client, {})).rejects.toThrow(
      "'query' must be a non-empty string",
    );
  });

  it("calls /me/messages when no folder_id", async () => {
    const { apiCalls, client } = captureRequest({ value: [] });
    await searchMessagesTool.handler(client, { query: "from:alice" });
    expect(apiCalls[0]).toBe("/me/messages");
  });

  it("scopes to /me/mailFolders/{id}/messages when folder_id given", async () => {
    const { apiCalls, client } = captureRequest({ value: [] });
    await searchMessagesTool.handler(client, { query: "x", folder_id: "inbox" });
    expect(apiCalls[0]).toBe("/me/mailFolders/inbox/messages");
  });

  it("sends ConsistencyLevel: eventual header (Graph $search requirement)", async () => {
    const { headers, client } = captureRequest({ value: [] });
    await searchMessagesTool.handler(client, { query: "x" });
    expect(headers).toContainEqual(["ConsistencyLevel", "eventual"]);
  });

  it("wraps the query in double-quotes for KQL", async () => {
    const { searchArgs, client } = captureRequest({ value: [] });
    await searchMessagesTool.handler(client, { query: "from:alice" });
    expect(searchArgs[0]).toBe('"from:alice"');
  });

  it("escapes embedded double-quotes in the query", async () => {
    const { searchArgs, client } = captureRequest({ value: [] });
    await searchMessagesTool.handler(client, { query: 'subject:"quarterly"' });
    expect(searchArgs[0]).toBe('"subject:\\"quarterly\\""');
  });

  it("returns count + messages + query in response", async () => {
    const { client } = captureRequest({
      value: [{ id: "m1", subject: "Q1 review" }],
    });
    const resp = await searchMessagesTool.handler(client, { query: "review" });
    const parsed = JSON.parse((resp.content[0] as { type: string; text: string }).text);
    expect(parsed.query).toBe("review");
    expect(parsed.count).toBe(1);
  });

  it("category is 'read'", () => {
    expect(searchMessagesTool.category).toBe("read");
  });

  // ─── v0.2: shared_user routing ─────────────────────────────────────
  it("routes to /users/{upn}/messages when shared_user is set", async () => {
    const { apiCalls, client } = captureRequest({ value: [] });
    await searchMessagesTool.handler(client, {
      query: "invoice",
      shared_user: "finance@juvant.io",
    });
    expect(apiCalls[0]).toBe("/users/finance%40juvant.io/messages");
  });

  it("routes to /users/{upn}/mailFolders/{f}/messages when both shared_user + folder_id", async () => {
    const { apiCalls, client } = captureRequest({ value: [] });
    await searchMessagesTool.handler(client, {
      query: "invoice",
      shared_user: "finance@juvant.io",
      folder_id: "inbox",
    });
    expect(apiCalls[0]).toBe(
      "/users/finance%40juvant.io/mailFolders/inbox/messages",
    );
  });

  it("still sends ConsistencyLevel: eventual on shared mailboxes", async () => {
    const { headers, client } = captureRequest({ value: [] });
    await searchMessagesTool.handler(client, {
      query: "x",
      shared_user: "finance@juvant.io",
    });
    expect(headers).toContainEqual(["ConsistencyLevel", "eventual"]);
  });

  it("echoes shared_user in the response", async () => {
    const { client } = captureRequest({ value: [] });
    const resp = await searchMessagesTool.handler(client, {
      query: "x",
      shared_user: "finance@juvant.io",
    });
    const parsed = JSON.parse((resp.content[0] as { type: string; text: string }).text);
    expect(parsed.shared_user).toBe("finance@juvant.io");
  });

  it("rejects a malformed shared_user before hitting Graph", async () => {
    const { apiCalls, client } = captureRequest({ value: [] });
    await expect(
      searchMessagesTool.handler(client, { query: "x", shared_user: "bad" }),
    ).rejects.toThrow(/UPN/);
    expect(apiCalls).toEqual([]);
  });
});
