import type { Client } from "@microsoft/microsoft-graph-client";
import { describe, expect, it, vi } from "vitest";

import { createDraftTool } from "../../src/tools/create_draft.js";

function captureRequest(returnValue: unknown): {
  apiCalls: string[];
  bodies: Record<string, unknown>[];
  client: Client;
} {
  const apiCalls: string[] = [];
  const bodies: Record<string, unknown>[] = [];
  const post = vi.fn().mockImplementation((b: Record<string, unknown>) => {
    bodies.push(b);
    return Promise.resolve(returnValue);
  });
  const api = vi.fn().mockImplementation((path: string) => {
    apiCalls.push(path);
    return { post };
  });
  return { apiCalls, bodies, client: { api } as unknown as Client };
}

const validResponse = { id: "draft-1", subject: "[agent-draft] Hello", isDraft: true };

describe("createDraftTool handler", () => {
  it("POSTs to /me/messages", async () => {
    const { apiCalls, client } = captureRequest(validResponse);
    await createDraftTool.handler(client, { subject: "Hello" });
    expect(apiCalls).toEqual(["/me/messages"]);
  });

  it("Shield C4 — prepends [agent-draft] to the subject", async () => {
    const { bodies, client } = captureRequest(validResponse);
    await createDraftTool.handler(client, { subject: "Weekly update" });
    expect(bodies[0].subject).toBe("[agent-draft] Weekly update");
  });

  it("Shield C4 — marks even when subject is omitted", async () => {
    const { bodies, client } = captureRequest(validResponse);
    await createDraftTool.handler(client, {});
    expect(bodies[0].subject).toBe("[agent-draft] ");
  });

  it("Shield C4 — attaches internetMessageHeaders with X-Juvant-Agent-Author", async () => {
    const { bodies, client } = captureRequest(validResponse);
    await createDraftTool.handler(client, { subject: "x" });
    expect(bodies[0].internetMessageHeaders).toEqual([
      { name: "X-Juvant-Agent-Author", value: "@juvantlabs/m365-mail-mcp-server" },
    ]);
  });

  it("Shield C4 — idempotent when caller already prefixed the subject", async () => {
    const { bodies, client } = captureRequest(validResponse);
    await createDraftTool.handler(client, { subject: "[agent-draft] Hello" });
    expect(bodies[0].subject).toBe("[agent-draft] Hello");
  });

  it("builds toRecipients/ccRecipients/bccRecipients", async () => {
    const { bodies, client } = captureRequest(validResponse);
    await createDraftTool.handler(client, {
      subject: "x",
      to: ["alice@x.com"],
      cc: [{ email: "bob@x.com", name: "Bob" }],
      bcc: ["carol@x.com"],
    });
    expect(bodies[0].toRecipients).toEqual([{ emailAddress: { address: "alice@x.com" } }]);
    expect(bodies[0].ccRecipients).toEqual([
      { emailAddress: { address: "bob@x.com", name: "Bob" } },
    ]);
    expect(bodies[0].bccRecipients).toEqual([{ emailAddress: { address: "carol@x.com" } }]);
  });

  it("includes body when supplied; defaults contentType to text", async () => {
    const { bodies, client } = captureRequest(validResponse);
    await createDraftTool.handler(client, { subject: "x", body: "Hello world" });
    expect(bodies[0].body).toEqual({ contentType: "text", content: "Hello world" });
  });

  it("defaults body to empty when omitted (Graph requires body on create)", async () => {
    const { bodies, client } = captureRequest(validResponse);
    await createDraftTool.handler(client, { subject: "x" });
    expect(bodies[0].body).toEqual({ contentType: "text", content: "" });
  });

  it("rejects invalid importance / body_content_type", async () => {
    const { client } = captureRequest(validResponse);
    await expect(
      createDraftTool.handler(client, { importance: "critical" }),
    ).rejects.toThrow("must be one of");
    await expect(
      createDraftTool.handler(client, { body: "x", body_content_type: "md" }),
    ).rejects.toThrow("must be one of");
  });

  it("returns the created draft summary + Shield markers metadata", async () => {
    const { client } = captureRequest(validResponse);
    const resp = await createDraftTool.handler(client, { subject: "x" });
    const parsed = JSON.parse((resp.content[0] as { type: string; text: string }).text);
    expect(parsed.created.id).toBe("draft-1");
    expect(parsed.agent_draft_markers.subject_prefix).toBe("[agent-draft] ");
    expect(parsed.agent_draft_markers.header).toBe("X-Juvant-Agent-Author");
  });

  it("category is 'write_idempotent'", () => {
    expect(createDraftTool.category).toBe("write_idempotent");
  });
});
