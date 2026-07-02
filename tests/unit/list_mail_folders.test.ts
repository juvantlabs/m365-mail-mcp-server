import type { Client } from "@microsoft/microsoft-graph-client";
import { describe, expect, it, vi } from "vitest";

import {
  listMailFoldersTool,
  summarizeFolder,
} from "../../src/tools/list_mail_folders.js";

function captureRequest(returnValue: unknown): {
  apiCalls: string[];
  client: Client;
} {
  const apiCalls: string[] = [];
  const get = vi.fn().mockResolvedValue(returnValue);
  const top = vi.fn().mockReturnValue({ get });
  const select = vi.fn().mockReturnValue({ top });
  const api = vi.fn().mockImplementation((path: string) => {
    apiCalls.push(path);
    return { select };
  });
  return { apiCalls, client: { api } as unknown as Client };
}

describe("summarizeFolder", () => {
  it("extracts folder metadata", () => {
    const s = summarizeFolder({
      id: "AAMk",
      displayName: "Inbox",
      parentFolderId: "root",
      childFolderCount: 3,
      unreadItemCount: 7,
      totalItemCount: 42,
    });
    expect(s).toEqual({
      id: "AAMk",
      display_name: "Inbox",
      parent_folder_id: "root",
      child_folder_count: 3,
      unread_item_count: 7,
      total_item_count: 42,
    });
  });

  it("defaults missing fields to sensible values", () => {
    const s = summarizeFolder({ id: "x" });
    expect(s.display_name).toBe("");
    expect(s.parent_folder_id).toBeNull();
    expect(s.child_folder_count).toBe(0);
    expect(s.unread_item_count).toBe(0);
    expect(s.total_item_count).toBe(0);
  });
});

describe("listMailFoldersTool handler", () => {
  it("calls /me/mailFolders", async () => {
    const { apiCalls, client } = captureRequest({ value: [] });
    await listMailFoldersTool.handler(client, {});
    expect(apiCalls).toEqual(["/me/mailFolders"]);
  });

  it("returns count + folders in the response", async () => {
    const { client } = captureRequest({
      value: [
        { id: "f1", displayName: "Inbox" },
        { id: "f2", displayName: "Sent Items" },
      ],
    });
    const resp = await listMailFoldersTool.handler(client, {});
    const parsed = JSON.parse((resp.content[0] as { type: string; text: string }).text);
    expect(parsed.count).toBe(2);
    expect(parsed.folders[0].display_name).toBe("Inbox");
    expect(parsed.folders[1].display_name).toBe("Sent Items");
  });

  it("handles missing response.value as empty list", async () => {
    const { client } = captureRequest({});
    const resp = await listMailFoldersTool.handler(client, {});
    const parsed = JSON.parse((resp.content[0] as { type: string; text: string }).text);
    expect(parsed.count).toBe(0);
    expect(parsed.folders).toEqual([]);
  });

  it("validates limit boundary", async () => {
    const { client } = captureRequest({ value: [] });
    await expect(
      listMailFoldersTool.handler(client, { limit: 0 }),
    ).rejects.toThrow("between 1 and 100");
    await expect(
      listMailFoldersTool.handler(client, { limit: 101 }),
    ).rejects.toThrow("between 1 and 100");
  });

  it("category is 'read'", () => {
    expect(listMailFoldersTool.category).toBe("read");
  });

  // ─── v0.2: shared_user routing ─────────────────────────────────────
  it("routes to /users/{upn}/mailFolders when shared_user is set", async () => {
    const { apiCalls, client } = captureRequest({ value: [] });
    await listMailFoldersTool.handler(client, { shared_user: "finance@juvant.io" });
    expect(apiCalls).toEqual(["/users/finance%40juvant.io/mailFolders"]);
  });

  it("echoes shared_user in the response (null when omitted)", async () => {
    const { client } = captureRequest({ value: [] });
    const own = await listMailFoldersTool.handler(client, {});
    const ownParsed = JSON.parse((own.content[0] as { type: string; text: string }).text);
    expect(ownParsed.shared_user).toBeNull();

    const { client: c2 } = captureRequest({ value: [] });
    const shared = await listMailFoldersTool.handler(c2, {
      shared_user: "finance@juvant.io",
    });
    const sharedParsed = JSON.parse(
      (shared.content[0] as { type: string; text: string }).text,
    );
    expect(sharedParsed.shared_user).toBe("finance@juvant.io");
  });

  it("rejects a malformed shared_user before hitting Graph", async () => {
    const { apiCalls, client } = captureRequest({ value: [] });
    await expect(
      listMailFoldersTool.handler(client, { shared_user: "not-a-upn" }),
    ).rejects.toThrow(/UPN/);
    expect(apiCalls).toEqual([]);
  });
});
