import type { Client } from "@microsoft/microsoft-graph-client";
import { describe, expect, it, vi } from "vitest";

import { expandMessage, getMessageTool } from "../../src/tools/get_message.js";

function mockClientWithMessage(m: unknown): {
  apiCalls: string[];
  client: Client;
} {
  const apiCalls: string[] = [];
  const get = vi.fn().mockResolvedValue(m);
  const select = vi.fn().mockReturnValue({ get });
  const api = vi.fn().mockImplementation((path: string) => {
    apiCalls.push(path);
    return { select };
  });
  return { apiCalls, client: { api } as unknown as Client };
}

describe("expandMessage", () => {
  it("keeps short bodies intact and marks body_truncated=false", () => {
    const r = expandMessage({
      id: "m1",
      subject: "x",
      body: { contentType: "text", content: "short body" },
    });
    expect(r.body).toBe("short body");
    expect(r.body_truncated).toBe(false);
    expect(r.body_content_type).toBe("text");
  });

  it("caps long bodies at 16 000 chars and sets body_truncated=true", () => {
    const long = "a".repeat(20_000);
    const r = expandMessage({
      id: "m1",
      subject: "x",
      body: { contentType: "html", content: long },
    });
    expect(r.body.length).toBe(16_000);
    expect(r.body_truncated).toBe(true);
    expect(r.body_content_type).toBe("html");
  });

  it("caps exactly at 16 000 chars — 16 001 flips the truncation flag", () => {
    const at = "a".repeat(16_000);
    const over = "a".repeat(16_001);
    expect(expandMessage({ id: "m1", subject: "x", body: { content: at } }).body_truncated).toBe(
      false,
    );
    expect(
      expandMessage({ id: "m1", subject: "x", body: { content: over } }).body_truncated,
    ).toBe(true);
  });

  it("returns [] for missing reply_to", () => {
    const r = expandMessage({ id: "m1", subject: "x" });
    expect(r.reply_to).toEqual([]);
  });

  it("maps reply_to entries via emailAddress", () => {
    const r = expandMessage({
      id: "m1",
      subject: "x",
      replyTo: [{ emailAddress: { name: "Bob", address: "bob@x.com" } }],
    });
    expect(r.reply_to).toEqual([{ name: "Bob", email: "bob@x.com" }]);
  });
});

describe("getMessageTool handler", () => {
  it("requires message_id", async () => {
    const { client } = mockClientWithMessage({});
    await expect(getMessageTool.handler(client, {})).rejects.toThrow(
      "'message_id' must be a non-empty string",
    );
  });

  it("calls /me/messages/{id} with URL-encoded id", async () => {
    const { apiCalls, client } = mockClientWithMessage({ id: "m1", subject: "x" });
    await getMessageTool.handler(client, { message_id: "some/id" });
    expect(apiCalls[0]).toBe("/me/messages/some%2Fid");
  });

  it("returns the expanded message as JSON", async () => {
    const { client } = mockClientWithMessage({
      id: "m1",
      subject: "Hello",
      body: { content: "hi", contentType: "text" },
    });
    const resp = await getMessageTool.handler(client, { message_id: "m1" });
    const parsed = JSON.parse((resp.content[0] as { type: string; text: string }).text);
    expect(parsed.subject).toBe("Hello");
    expect(parsed.body).toBe("hi");
    expect(parsed.body_truncated).toBe(false);
  });

  it("category is 'read'", () => {
    expect(getMessageTool.category).toBe("read");
  });
});
