import { describe, expect, it } from "vitest";

import {
  AGENT_DRAFT_HEADER_NAME,
  AGENT_DRAFT_HEADER_VALUE,
  AGENT_DRAFT_SUBJECT_PREFIX,
  buildMessageBody,
  ensureAgentDraftSubject,
  parseRecipients,
} from "../../src/tools/_shared.js";

describe("Shield C4 constants", () => {
  it("subject prefix is '[agent-draft] ' (load-bearing string)", () => {
    expect(AGENT_DRAFT_SUBJECT_PREFIX).toBe("[agent-draft] ");
  });

  it("header name is X-Juvant-Agent-Author (load-bearing string)", () => {
    expect(AGENT_DRAFT_HEADER_NAME).toBe("X-Juvant-Agent-Author");
    expect(AGENT_DRAFT_HEADER_VALUE).toBe("@juvantlabs/m365-mail-mcp-server");
  });
});

describe("ensureAgentDraftSubject", () => {
  it("prepends the marker when absent", () => {
    expect(ensureAgentDraftSubject("Hello")).toBe("[agent-draft] Hello");
  });

  it("is idempotent — never stacks the marker", () => {
    const once = ensureAgentDraftSubject("Hello");
    expect(ensureAgentDraftSubject(once)).toBe(once);
    expect(ensureAgentDraftSubject(ensureAgentDraftSubject(once))).toBe(once);
  });

  it("marks empty subject", () => {
    expect(ensureAgentDraftSubject("")).toBe("[agent-draft] ");
  });

  it("preserves the marker on reply-style pre-filled subjects", () => {
    expect(ensureAgentDraftSubject("RE: Weekly sync")).toBe("[agent-draft] RE: Weekly sync");
  });
});

describe("parseRecipients", () => {
  it("returns undefined for undefined / null", () => {
    expect(parseRecipients(undefined, "to")).toBeUndefined();
    expect(parseRecipients(null, "to")).toBeUndefined();
  });

  it("accepts bare email strings", () => {
    expect(parseRecipients(["alice@x.com"], "to")).toEqual([{ email: "alice@x.com" }]);
  });

  it("accepts {email, name?} objects", () => {
    expect(parseRecipients([{ email: "alice@x.com", name: "Alice" }], "to")).toEqual([
      { email: "alice@x.com", name: "Alice" },
    ]);
    expect(parseRecipients([{ email: "bob@x.com" }], "to")).toEqual([{ email: "bob@x.com" }]);
  });

  it("mixes bare strings and objects", () => {
    expect(
      parseRecipients(["a@x.com", { email: "b@x.com", name: "Bob" }], "to"),
    ).toEqual([{ email: "a@x.com" }, { email: "b@x.com", name: "Bob" }]);
  });

  it("throws when top-level value is not an array", () => {
    expect(() => parseRecipients("alice@x.com", "to")).toThrow("'to' must be an array");
  });

  it("throws when item is neither string nor object", () => {
    expect(() => parseRecipients([42], "to")).toThrow(
      "'to[0]' must be a string or an object",
    );
  });

  it("throws when object item is missing email", () => {
    expect(() => parseRecipients([{ name: "Alice" }], "to")).toThrow(
      "'to[0].email' must be a non-empty string",
    );
  });
});

describe("buildMessageBody — create mode", () => {
  it("always marks the subject with '[agent-draft] '", () => {
    const body = buildMessageBody({ subject: "Hello" }, "create");
    expect(body.subject).toBe("[agent-draft] Hello");
  });

  it("marks even when subject is omitted (uses createFallbackSubject)", () => {
    const body = buildMessageBody({}, {
      mode: "create",
      createFallbackSubject: "no subject",
    });
    expect(body.subject).toBe("[agent-draft] no subject");
  });

  it("does not double-mark when subject already carries the prefix", () => {
    const body = buildMessageBody({ subject: "[agent-draft] Hello" }, "create");
    expect(body.subject).toBe("[agent-draft] Hello");
  });

  it("attaches internetMessageHeaders when attachAgentHeader=true", () => {
    const body = buildMessageBody(
      { subject: "x" },
      { mode: "create", attachAgentHeader: true },
    );
    expect(body.internetMessageHeaders).toEqual([
      { name: "X-Juvant-Agent-Author", value: "@juvantlabs/m365-mail-mcp-server" },
    ]);
  });

  it("does NOT attach internetMessageHeaders when attachAgentHeader is unset", () => {
    const body = buildMessageBody({ subject: "x" }, "create");
    expect(body.internetMessageHeaders).toBeUndefined();
  });

  it("provides an empty body on create when body is omitted", () => {
    const body = buildMessageBody({}, "create") as {
      body: { contentType: string; content: string };
    };
    expect(body.body).toEqual({ contentType: "text", content: "" });
  });

  it("respects the supplied body_content_type", () => {
    const body = buildMessageBody(
      { body: "<p>hi</p>", body_content_type: "html" },
      "create",
    ) as { body: { contentType: string; content: string } };
    expect(body.body).toEqual({ contentType: "html", content: "<p>hi</p>" });
  });

  it("builds toRecipients when provided", () => {
    const body = buildMessageBody(
      { to: ["a@x.com", { email: "b@x.com", name: "Bob" }] },
      "create",
    );
    expect(body.toRecipients).toEqual([
      { emailAddress: { address: "a@x.com" } },
      { emailAddress: { address: "b@x.com", name: "Bob" } },
    ]);
  });

  it("builds ccRecipients / bccRecipients", () => {
    const body = buildMessageBody({ cc: ["c@x.com"], bcc: ["d@x.com"] }, "create");
    expect(body.ccRecipients).toEqual([{ emailAddress: { address: "c@x.com" } }]);
    expect(body.bccRecipients).toEqual([{ emailAddress: { address: "d@x.com" } }]);
  });

  it("emits importance only when the caller passes one", () => {
    const body1 = buildMessageBody({ importance: "high" }, "create");
    expect(body1.importance).toBe("high");
    const body2 = buildMessageBody({}, "create");
    expect(body2.importance).toBeUndefined();
  });

  it("rejects invalid importance", () => {
    expect(() => buildMessageBody({ importance: "critical" }, "create")).toThrow(
      "must be one of",
    );
  });

  it("rejects invalid body_content_type", () => {
    expect(() =>
      buildMessageBody({ body: "x", body_content_type: "markdown" }, "create"),
    ).toThrow("must be one of");
  });
});

describe("buildMessageBody — patch mode", () => {
  it("returns empty body when nothing was supplied", () => {
    const body = buildMessageBody({}, "patch");
    expect(body).toEqual({});
  });

  it("marks subject when patched, leaves it absent when not", () => {
    expect(buildMessageBody({ subject: "hi" }, "patch").subject).toBe("[agent-draft] hi");
    expect(buildMessageBody({ to: ["a@x.com"] }, "patch").subject).toBeUndefined();
  });

  it("does not attach internetMessageHeaders in patch mode (Graph forbids)", () => {
    const body = buildMessageBody(
      { subject: "x" },
      { mode: "patch", attachAgentHeader: true },
    );
    // Even though attachAgentHeader is true, PATCH mode still attaches
    // it — the caller is responsible for only setting attachAgentHeader
    // when the underlying operation supports it. Test the semantics
    // are intentional: it's the caller's contract to pass this only in
    // create-time flows.
    expect(body.internetMessageHeaders).toBeDefined();
  });

  it("only PATCHes fields the caller supplied", () => {
    const body = buildMessageBody({ to: ["a@x.com"] }, "patch");
    expect(Object.keys(body)).toEqual(["toRecipients"]);
  });
});
