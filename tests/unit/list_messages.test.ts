import type { Client } from "@microsoft/microsoft-graph-client";
import { describe, expect, it, vi } from "vitest";

import {
  listMessagesTool,
  summarizeMessage,
} from "../../src/tools/list_messages.js";

function captureRequest(returnValue: unknown): {
  apiCalls: string[];
  client: Client;
} {
  const apiCalls: string[] = [];
  const get = vi.fn().mockResolvedValue(returnValue);
  const orderby = vi.fn().mockReturnValue({ get });
  const top = vi.fn().mockReturnValue({ orderby });
  const select = vi.fn().mockReturnValue({ top });
  const api = vi.fn().mockImplementation((path: string) => {
    apiCalls.push(path);
    return { select };
  });
  return { apiCalls, client: { api } as unknown as Client };
}

describe("summarizeMessage", () => {
  it("extracts headers, recipients, timestamps, flags", () => {
    const m = {
      id: "m1",
      conversationId: "c1",
      subject: "Hi",
      from: { emailAddress: { name: "Alice", address: "alice@x.com" } },
      toRecipients: [{ emailAddress: { name: "Bob", address: "bob@x.com" } }],
      ccRecipients: [{ emailAddress: { address: "carol@x.com" } }],
      receivedDateTime: "2026-06-01T10:00:00Z",
      sentDateTime: "2026-06-01T09:59:00Z",
      isRead: false,
      isDraft: false,
      hasAttachments: true,
      importance: "high",
      bodyPreview: "Hello there",
      webLink: "https://outlook/m1",
      parentFolderId: "inbox",
    };
    const s = summarizeMessage(m);
    expect(s.id).toBe("m1");
    expect(s.subject).toBe("Hi");
    expect(s.from).toEqual({ name: "Alice", email: "alice@x.com" });
    expect(s.to).toEqual([{ name: "Bob", email: "bob@x.com" }]);
    expect(s.cc).toEqual([{ name: "", email: "carol@x.com" }]);
    expect(s.is_read).toBe(false);
    expect(s.has_attachments).toBe(true);
    expect(s.importance).toBe("high");
    expect(s.parent_folder_id).toBe("inbox");
  });

  it("returns null from when the message has no from field (drafts)", () => {
    expect(summarizeMessage({ id: "x" }).from).toBeNull();
  });

  it("returns empty arrays for missing to/cc lists", () => {
    const s = summarizeMessage({ id: "x" });
    expect(s.to).toEqual([]);
    expect(s.cc).toEqual([]);
  });
});

describe("listMessagesTool handler", () => {
  it("calls /me/messages when no folder_id", async () => {
    const { apiCalls, client } = captureRequest({ value: [] });
    await listMessagesTool.handler(client, {});
    expect(apiCalls).toEqual(["/me/messages"]);
  });

  it("scopes to /me/mailFolders/{id}/messages when folder_id given", async () => {
    const { apiCalls, client } = captureRequest({ value: [] });
    await listMessagesTool.handler(client, { folder_id: "inbox" });
    expect(apiCalls[0]).toBe("/me/mailFolders/inbox/messages");
  });

  it("URL-encodes folder_id (defense-in-depth for well-known-name shortcuts)", async () => {
    const { apiCalls, client } = captureRequest({ value: [] });
    await listMessagesTool.handler(client, { folder_id: "folder with space" });
    expect(apiCalls[0]).toBe("/me/mailFolders/folder%20with%20space/messages");
  });

  it("returns count + messages + folder_id in response", async () => {
    const { client } = captureRequest({
      value: [
        { id: "m1", subject: "One" },
        { id: "m2", subject: "Two" },
      ],
    });
    const resp = await listMessagesTool.handler(client, { folder_id: "inbox" });
    const parsed = JSON.parse((resp.content[0] as { type: string; text: string }).text);
    expect(parsed.count).toBe(2);
    expect(parsed.folder_id).toBe("inbox");
    expect(parsed.messages).toHaveLength(2);
  });

  it("rejects invalid limit", async () => {
    const { client } = captureRequest({ value: [] });
    await expect(
      listMessagesTool.handler(client, { limit: 0 }),
    ).rejects.toThrow("between 1 and 100");
    await expect(
      listMessagesTool.handler(client, { limit: 101 }),
    ).rejects.toThrow();
  });

  it("category is 'read'", () => {
    expect(listMessagesTool.category).toBe("read");
  });
});
