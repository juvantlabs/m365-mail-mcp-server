import type { Client } from "@microsoft/microsoft-graph-client";
import { describe, expect, it, vi } from "vitest";

import { createForwardDraftTool } from "../../src/tools/create_forward_draft.js";
import { createReplyDraftTool } from "../../src/tools/create_reply_draft.js";
import { prependToPreparedBody, readMessageBody } from "../../src/tools/_shared.js";

// Graph's createReply / createReplyAll / createForward return a draft
// whose body already contains the original message. These tests pin
// that a caller-supplied body goes ABOVE that content, never replaces it.

const QUOTED_HTML =
  '<html><head><meta charset="utf-8"></head><body><div id="appendonsend"></div>' +
  '<hr><div><b>From:</b> Sender &lt;sender@example.com&gt;</div>' +
  "<div>Original message text</div></body></html>";

const QUOTED_TEXT =
  "________________________________\r\nFrom: Sender <sender@example.com>\r\n\r\nOriginal message text";

function makeClient(opts: {
  postReturns: Record<string, unknown>;
  getReturns?: Record<string, unknown>;
}): {
  apiCalls: string[];
  posts: unknown[];
  patchBodies: Record<string, unknown>[];
  selects: string[];
  client: Client;
} {
  const apiCalls: string[] = [];
  const posts: unknown[] = [];
  const patchBodies: Record<string, unknown>[] = [];
  const selects: string[] = [];
  const post = vi.fn().mockImplementation((b: unknown) => {
    posts.push(b);
    return Promise.resolve(opts.postReturns);
  });
  const patch = vi.fn().mockImplementation((b: Record<string, unknown>) => {
    patchBodies.push(b);
    return Promise.resolve({ id: opts.postReturns.id, ...b });
  });
  const get = vi.fn().mockResolvedValue(opts.getReturns ?? {});
  const request = {
    post,
    patch,
    get,
    select: vi.fn().mockImplementation((s: string) => {
      selects.push(s);
      return request;
    }),
  };
  const api = vi.fn().mockImplementation((p: string) => {
    apiCalls.push(p);
    return request;
  });
  return { apiCalls, posts, patchBodies, selects, client: { api } as unknown as Client };
}

function patchedBody(patchBodies: Record<string, unknown>[]): { contentType: string; content: string } {
  return patchBodies[0].body as { contentType: string; content: string };
}

describe("create_reply_draft keeps the quoted original", () => {
  it("text body on an HTML reply: caller text on top, quote kept below", async () => {
    const { patchBodies, client } = makeClient({
      postReturns: {
        id: "d1",
        subject: "RE: Plan",
        body: { contentType: "html", content: QUOTED_HTML },
      },
    });
    await createReplyDraftTool.handler(client, {
      message_id: "m1",
      body: "Thanks, see you <Monday> & bring notes.\nBest",
    });
    const body = patchedBody(patchBodies);
    expect(body.contentType).toBe("html");
    expect(body.content).toContain("Original message text");
    expect(body.content).toContain("sender@example.com");
    expect(body.content).toContain(
      "<div>Thanks, see you &lt;Monday&gt; &amp; bring notes.<br>Best</div><br>",
    );
    // Caller text is right after <body>, before the quoted part.
    const callerAt = body.content.indexOf("Thanks, see you");
    expect(callerAt).toBeGreaterThan(body.content.indexOf("<body>"));
    expect(callerAt).toBeLessThan(body.content.indexOf('<div id="appendonsend">'));
    expect(body.content.endsWith("</body></html>")).toBe(true);
  });

  it("HTML body on an HTML reply: inserted verbatim above the quote", async () => {
    const { patchBodies, client } = makeClient({
      postReturns: {
        id: "d1",
        subject: "RE: Plan",
        body: { contentType: "HTML", content: QUOTED_HTML },
      },
    });
    await createReplyDraftTool.handler(client, {
      message_id: "m1",
      body: "<p>Agreed.</p>",
      body_content_type: "html",
    });
    const body = patchedBody(patchBodies);
    expect(body.contentType).toBe("html");
    expect(body.content).toContain("<body><p>Agreed.</p><br><div id=\"appendonsend\">");
    expect(body.content).toContain("Original message text");
  });

  it("text body on a text reply: caller text, blank line, then the quote", async () => {
    const { patchBodies, client } = makeClient({
      postReturns: {
        id: "d1",
        subject: "RE: Plan",
        body: { contentType: "text", content: QUOTED_TEXT },
      },
    });
    await createReplyDraftTool.handler(client, { message_id: "m1", body: "Sounds good." });
    expect(patchedBody(patchBodies)).toEqual({
      contentType: "text",
      content: `Sounds good.\r\n\r\n${QUOTED_TEXT}`,
    });
  });

  it("reply_all with a body keeps the quote too", async () => {
    const { apiCalls, patchBodies, client } = makeClient({
      postReturns: {
        id: "d1",
        subject: "RE: Plan",
        body: { contentType: "html", content: QUOTED_HTML },
      },
    });
    await createReplyDraftTool.handler(client, {
      message_id: "m1",
      reply_all: true,
      body: "Adding everyone.",
    });
    expect(apiCalls[0]).toBe("/me/messages/m1/createReplyAll");
    const body = patchedBody(patchBodies);
    expect(body.content).toContain("<div>Adding everyone.</div><br>");
    expect(body.content).toContain("Original message text");
  });

  it("without a body, the prepared body is not touched", async () => {
    const { patchBodies, apiCalls, client } = makeClient({
      postReturns: {
        id: "d1",
        subject: "RE: Plan",
        body: { contentType: "html", content: QUOTED_HTML },
      },
    });
    await createReplyDraftTool.handler(client, { message_id: "m1" });
    expect(patchBodies[0]).toEqual({ subject: "[agent-draft] RE: Plan" });
    expect(apiCalls).toEqual(["/me/messages/m1/createReply", "/me/messages/d1"]);
  });

  it("keeps the [agent-draft] subject behaviour when a body is supplied", async () => {
    const { patchBodies, client } = makeClient({
      postReturns: {
        id: "d1",
        subject: "RE: Plan",
        body: { contentType: "html", content: QUOTED_HTML },
      },
    });
    await createReplyDraftTool.handler(client, { message_id: "m1", body: "Ok" });
    expect(patchBodies[0].subject).toBe("[agent-draft] RE: Plan");
  });

  it("still POSTs an empty createReply payload (body is merged on PATCH)", async () => {
    const { posts, client } = makeClient({
      postReturns: { id: "d1", subject: "RE: x", body: { contentType: "html", content: QUOTED_HTML } },
    });
    await createReplyDraftTool.handler(client, { message_id: "m1", body: "Ok" });
    expect(posts).toEqual([{}]);
  });

  it("reads the body back from the draft when createReply omits it", async () => {
    const { apiCalls, selects, patchBodies, client } = makeClient({
      postReturns: { id: "d1", subject: "RE: Plan" },
      getReturns: { body: { contentType: "html", content: QUOTED_HTML } },
    });
    await createReplyDraftTool.handler(client, {
      message_id: "m1",
      body: "Ok",
      shared_user: "team@example.com",
    });
    expect(apiCalls[1]).toBe("/users/team%40example.com/messages/d1");
    expect(selects).toEqual(["body"]);
    expect(patchedBody(patchBodies).content).toContain("Original message text");
    expect(patchedBody(patchBodies).content).toContain("<div>Ok</div><br>");
  });
});

describe("create_forward_draft keeps the forwarded message", () => {
  it("text note on an HTML forward: note on top, forwarded content kept", async () => {
    const { patchBodies, client } = makeClient({
      postReturns: {
        id: "f1",
        subject: "FW: Plan",
        body: { contentType: "html", content: QUOTED_HTML },
      },
    });
    await createForwardDraftTool.handler(client, {
      message_id: "m1",
      to: ["colleague@example.com"],
      body: "FYI",
    });
    const body = patchedBody(patchBodies);
    expect(body.contentType).toBe("html");
    expect(body.content).toContain("<body><div>FYI</div><br><div id=\"appendonsend\">");
    expect(body.content).toContain("Original message text");
    expect(patchBodies[0].subject).toBe("[agent-draft] FW: Plan");
    expect(patchBodies[0].toRecipients).toEqual([
      { emailAddress: { address: "colleague@example.com" } },
    ]);
  });

  it("HTML note on a text forward: result is HTML with the original escaped below", async () => {
    const { patchBodies, client } = makeClient({
      postReturns: {
        id: "f1",
        subject: "FW: Plan",
        body: { contentType: "text", content: QUOTED_TEXT },
      },
    });
    await createForwardDraftTool.handler(client, {
      message_id: "m1",
      body: "<p>Please review</p>",
      body_content_type: "html",
    });
    const body = patchedBody(patchBodies);
    expect(body.contentType).toBe("html");
    expect(body.content.startsWith("<p>Please review</p><br><div>")).toBe(true);
    expect(body.content).toContain("From: Sender &lt;sender@example.com&gt;<br><br>Original message text");
  });

  it("without a body, the forwarded content is not touched", async () => {
    const { patchBodies, client } = makeClient({
      postReturns: {
        id: "f1",
        subject: "FW: Plan",
        body: { contentType: "html", content: QUOTED_HTML },
      },
    });
    await createForwardDraftTool.handler(client, { message_id: "m1" });
    expect(patchBodies[0].body).toBeUndefined();
  });
});

describe("prependToPreparedBody / readMessageBody", () => {
  it("returns the caller body unchanged when nothing was prepared", () => {
    const caller = { contentType: "text" as const, content: "Hi" };
    expect(prependToPreparedBody(caller, undefined)).toEqual(caller);
    expect(prependToPreparedBody(caller, { contentType: "html", content: "  " })).toEqual(caller);
  });

  it("prepends to HTML that has no <body> tag", () => {
    expect(
      prependToPreparedBody(
        { contentType: "text", content: "Hi" },
        { contentType: "html", content: "<div>quoted</div>" },
      ),
    ).toEqual({ contentType: "html", content: "<div>Hi</div><br><div>quoted</div>" });
  });

  it("handles a <body> tag with attributes", () => {
    const out = prependToPreparedBody(
      { contentType: "text", content: "Hi" },
      { contentType: "html", content: '<html><BODY dir="ltr"><div>quoted</div></BODY></html>' },
    );
    expect(out.content).toBe('<html><BODY dir="ltr"><div>Hi</div><br><div>quoted</div></BODY></html>');
  });

  it("readMessageBody normalizes contentType and rejects malformed bodies", () => {
    expect(readMessageBody({ body: { contentType: "HTML", content: "x" } })).toEqual({
      contentType: "html",
      content: "x",
    });
    expect(readMessageBody({})).toBeUndefined();
    expect(readMessageBody({ body: { contentType: "html" } })).toBeUndefined();
    expect(readMessageBody({ body: { contentType: "rtf", content: "x" } })).toBeUndefined();
  });
});
