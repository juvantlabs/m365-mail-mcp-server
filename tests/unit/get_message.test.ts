import type { Client } from "@microsoft/microsoft-graph-client";
import { describe, expect, it, vi } from "vitest";

import {
  applyBodyPagination,
  expandMessage,
  getMessageTool,
} from "../../src/tools/get_message.js";

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

function parseResponse(
  resp: Awaited<ReturnType<typeof getMessageTool.handler>>,
): Record<string, unknown> {
  return JSON.parse((resp.content[0] as { type: string; text: string }).text);
}

describe("expandMessage", () => {
  it("returns the full body as-is (no cap)", () => {
    const long = "a".repeat(50_000);
    const r = expandMessage({
      id: "m1",
      subject: "x",
      body: { contentType: "html", content: long },
    });
    expect(r.body.length).toBe(50_000);
    expect(r.body_content_type).toBe("html");
  });

  it("short bodies pass through unchanged", () => {
    const r = expandMessage({
      id: "m1",
      subject: "x",
      body: { contentType: "text", content: "short body" },
    });
    expect(r.body).toBe("short body");
    expect(r.body_content_type).toBe("text");
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

  it("empty body content defaults to empty string", () => {
    const r = expandMessage({ id: "m1", subject: "x" });
    expect(r.body).toBe("");
    expect(r.body_content_type).toBe("text");
  });
});

describe("applyBodyPagination", () => {
  const base = expandMessage({
    id: "m1",
    subject: "x",
    body: { contentType: "text", content: "0123456789" }, // len 10
  });

  it("no maxBodyChars → returns full body, next_offset=null, body_truncated=false", () => {
    const r = applyBodyPagination(base, 0, undefined);
    expect(r.body).toBe("0123456789");
    expect(r.body_offset).toBe(0);
    expect(r.body_char_count).toBe(10);
    expect(r.total_body_char_count).toBe(10);
    expect(r.next_offset).toBeNull();
    expect(r.body_truncated).toBe(false);
  });

  it("no maxBodyChars + nonzero offset → returns tail, next_offset=null", () => {
    const r = applyBodyPagination(base, 4, undefined);
    expect(r.body).toBe("456789");
    expect(r.body_offset).toBe(4);
    expect(r.body_char_count).toBe(6);
    expect(r.total_body_char_count).toBe(10);
    expect(r.next_offset).toBeNull();
    expect(r.body_truncated).toBe(false);
  });

  it("maxBodyChars mid-body → slices and returns next_offset", () => {
    const r = applyBodyPagination(base, 0, 4);
    expect(r.body).toBe("0123");
    expect(r.body_char_count).toBe(4);
    expect(r.next_offset).toBe(4);
    expect(r.body_truncated).toBe(true);
  });

  it("maxBodyChars reaching exactly the end → next_offset=null (terminates)", () => {
    const r = applyBodyPagination(base, 6, 4);
    expect(r.body).toBe("6789");
    expect(r.body_char_count).toBe(4);
    expect(r.next_offset).toBeNull();
    expect(r.body_truncated).toBe(false);
  });

  it("maxBodyChars past the end → returns remainder, next_offset=null", () => {
    const r = applyBodyPagination(base, 8, 100);
    expect(r.body).toBe("89");
    expect(r.body_char_count).toBe(2);
    expect(r.next_offset).toBeNull();
    expect(r.body_truncated).toBe(false);
  });

  it("body_offset beyond end → empty body, next_offset=null (terminates cleanly)", () => {
    const r = applyBodyPagination(base, 999, 100);
    expect(r.body).toBe("");
    expect(r.body_offset).toBe(10);
    expect(r.body_char_count).toBe(0);
    expect(r.next_offset).toBeNull();
    expect(r.body_truncated).toBe(false);
  });

  it("iterative pagination walks the full body and terminates cleanly", () => {
    const collected: string[] = [];
    let offset: number | null = 0;
    let guard = 0;
    while (offset !== null) {
      if (guard++ > 100) throw new Error("pagination did not terminate");
      const chunk = applyBodyPagination(base, offset, 3);
      collected.push(chunk.body);
      offset = chunk.next_offset;
    }
    expect(collected.join("")).toBe("0123456789");
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

  it("paramless call returns the FULL body untruncated for a >16 000-char body", async () => {
    const HUGE = 40_000;
    const long = "x".repeat(HUGE);
    const { client } = mockClientWithMessage({
      id: "m1",
      subject: "long one",
      body: { content: long, contentType: "html" },
    });
    const resp = await getMessageTool.handler(client, { message_id: "m1" });
    const parsed = parseResponse(resp);
    expect(parsed.body.length).toBe(HUGE);
    expect(parsed.body).toBe(long);
    expect(parsed.body_offset).toBe(0);
    expect(parsed.body_char_count).toBe(HUGE);
    expect(parsed.total_body_char_count).toBe(HUGE);
    expect(parsed.next_offset).toBeNull();
    expect(parsed.body_truncated).toBe(false);
  });

  it("returns the expanded message as JSON with pagination metadata", async () => {
    const { client } = mockClientWithMessage({
      id: "m1",
      subject: "Hello",
      body: { content: "hi", contentType: "text" },
    });
    const resp = await getMessageTool.handler(client, { message_id: "m1" });
    const parsed = parseResponse(resp);
    expect(parsed.subject).toBe("Hello");
    expect(parsed.body).toBe("hi");
    expect(parsed.body_offset).toBe(0);
    expect(parsed.body_char_count).toBe(2);
    expect(parsed.total_body_char_count).toBe(2);
    expect(parsed.next_offset).toBeNull();
    expect(parsed.body_truncated).toBe(false);
  });

  it("max_body_chars + body_offset slices correctly and returns next_offset", async () => {
    const body = "abcdefghijklmnop"; // len 16
    const { client } = mockClientWithMessage({
      id: "m1",
      subject: "s",
      body: { content: body, contentType: "text" },
    });
    const resp = await getMessageTool.handler(client, {
      message_id: "m1",
      body_offset: 4,
      max_body_chars: 5,
    });
    const parsed = parseResponse(resp);
    expect(parsed.body).toBe("efghi");
    expect(parsed.body_offset).toBe(4);
    expect(parsed.body_char_count).toBe(5);
    expect(parsed.total_body_char_count).toBe(16);
    expect(parsed.next_offset).toBe(9);
    expect(parsed.body_truncated).toBe(true);
  });

  it("pagination terminates cleanly: the final chunk has next_offset=null", async () => {
    const body = "abcdefghij"; // len 10
    const { client } = mockClientWithMessage({
      id: "m1",
      subject: "s",
      body: { content: body, contentType: "text" },
    });
    const resp = await getMessageTool.handler(client, {
      message_id: "m1",
      body_offset: 7,
      max_body_chars: 100,
    });
    const parsed = parseResponse(resp);
    expect(parsed.body).toBe("hij");
    expect(parsed.body_offset).toBe(7);
    expect(parsed.body_char_count).toBe(3);
    expect(parsed.total_body_char_count).toBe(10);
    expect(parsed.next_offset).toBeNull();
    expect(parsed.body_truncated).toBe(false);
  });

  it("rejects invalid body_offset (negative)", async () => {
    const { client } = mockClientWithMessage({ id: "m1", subject: "s" });
    await expect(
      getMessageTool.handler(client, { message_id: "m1", body_offset: -1 }),
    ).rejects.toThrow(/body_offset/);
  });

  it("rejects invalid max_body_chars (zero)", async () => {
    const { client } = mockClientWithMessage({ id: "m1", subject: "s" });
    await expect(
      getMessageTool.handler(client, { message_id: "m1", max_body_chars: 0 }),
    ).rejects.toThrow(/max_body_chars/);
  });

  it("category is 'read'", () => {
    expect(getMessageTool.category).toBe("read");
  });

  // ─── v0.2: shared_user routing ─────────────────────────────────────
  it("routes to /users/{upn}/messages/{id} when shared_user is set", async () => {
    const { apiCalls, client } = mockClientWithMessage({
      id: "m1",
      subject: "x",
    });
    await getMessageTool.handler(client, {
      message_id: "m1",
      shared_user: "finance@juvant.io",
    });
    expect(apiCalls[0]).toBe("/users/finance%40juvant.io/messages/m1");
  });

  it("echoes shared_user in the response (own → null; shared → UPN)", async () => {
    const { client: c1 } = mockClientWithMessage({ id: "m1", subject: "x" });
    const own = await getMessageTool.handler(c1, { message_id: "m1" });
    expect(parseResponse(own).shared_user).toBeNull();

    const { client: c2 } = mockClientWithMessage({ id: "m1", subject: "x" });
    const shared = await getMessageTool.handler(c2, {
      message_id: "m1",
      shared_user: "finance@juvant.io",
    });
    expect(parseResponse(shared).shared_user).toBe("finance@juvant.io");
  });

  it("rejects a malformed shared_user before hitting Graph", async () => {
    const { apiCalls, client } = mockClientWithMessage({ id: "m1", subject: "x" });
    await expect(
      getMessageTool.handler(client, { message_id: "m1", shared_user: "bad" }),
    ).rejects.toThrow(/UPN/);
    expect(apiCalls).toEqual([]);
  });
});
