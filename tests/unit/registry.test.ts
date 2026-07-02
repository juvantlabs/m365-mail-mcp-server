import { describe, expect, it } from "vitest";

import { ALL_TOOLS, buildHandlerMap } from "../../src/tools/index.js";

describe("ALL_TOOLS registry", () => {
  it("registers the v0.2 tool set — still 13 tools (shared_user added as parameter, not new tools)", () => {
    // v0.2 does not add new tools — it widens the 13 existing tools with
    // an optional `shared_user` parameter. If this count changes, either
    // v0.3's `send_draft` has landed (bump to 14) or something regressed.
    expect(ALL_TOOLS).toHaveLength(13);
  });

  it("has unique tool names (dispatcher requires uniqueness for handler map)", () => {
    const names = ALL_TOOLS.map((t) => t.definition.name);
    expect(new Set(names).size).toBe(names.length);
  });

  it("every tool name is prefixed 'm365-mail:' (vendor namespace)", () => {
    for (const t of ALL_TOOLS) {
      expect(t.definition.name.startsWith("m365-mail:")).toBe(true);
    }
  });

  it("every tool has a non-empty description + inputSchema.type='object'", () => {
    for (const t of ALL_TOOLS) {
      expect(t.definition.description.length).toBeGreaterThan(0);
      expect(t.definition.inputSchema.type).toBe("object");
      expect(t.definition.inputSchema.properties).toBeDefined();
      expect(Array.isArray(t.definition.inputSchema.required)).toBe(true);
    }
  });

  it("registers every documented v0.1 tool by name", () => {
    const expected = new Set([
      "m365-mail:list_mail_folders",
      "m365-mail:list_messages",
      "m365-mail:search_messages",
      "m365-mail:get_message",
      "m365-mail:list_attachments",
      "m365-mail:download_attachment",
      "m365-mail:create_draft",
      "m365-mail:update_draft",
      "m365-mail:create_reply_draft",
      "m365-mail:create_forward_draft",
      "m365-mail:mark_read",
      "m365-mail:move_message",
      "m365-mail:delete_message",
    ]);
    const registered = new Set(ALL_TOOLS.map((t) => t.definition.name));
    expect(registered).toEqual(expected);
  });

  it("v0.2 must still NOT ship any send / send_draft tool (Shield-gated for v0.3)", () => {
    for (const t of ALL_TOOLS) {
      expect(t.definition.name).not.toMatch(/send/i);
    }
  });

  it("v0.2 exposes `shared_user` as an optional parameter on every tool (delegate mailboxes)", () => {
    // v0.2 widens the 13 tools with an optional `shared_user` UPN
    // parameter. This flipped from v0.1's assert-absent invariant.
    // Every tool MUST expose the parameter, and it MUST NOT be
    // required (own-mailbox behaviour is the default). If a new tool
    // lands without wiring shared_user through, this test fails —
    // catching the "silently drops to /me" failure mode.
    for (const t of ALL_TOOLS) {
      const props = t.definition.inputSchema.properties as Record<string, unknown>;
      const requiredList = t.definition.inputSchema.required as string[];
      expect(
        Object.keys(props),
        `${t.definition.name} is missing shared_user in inputSchema.properties`,
      ).toContain("shared_user");
      expect(
        requiredList,
        `${t.definition.name} incorrectly marks shared_user as required — must remain optional so callers get /me behaviour by default`,
      ).not.toContain("shared_user");
    }
  });

  it("v0.2 `shared_user` property carries the shared-scope description on every tool", () => {
    // Every tool sources the `shared_user` property from the SAME
    // schema constant (SHARED_USER_SCHEMA_PROPERTY in _mailbox.ts).
    // Pin the description string so a fork or copy-paste change to
    // one tool's schema is caught immediately.
    for (const t of ALL_TOOLS) {
      const props = t.definition.inputSchema.properties as Record<
        string,
        { type: string; description: string }
      >;
      const desc = props.shared_user?.description;
      expect(desc).toBeDefined();
      expect(desc).toContain("Mail.Read.Shared");
      expect(desc).toContain("Mail.ReadWrite.Shared");
      expect(props.shared_user.type).toBe("string");
    }
  });

  it("has exactly one write_irreversible tool: delete_message", () => {
    const irreversible = ALL_TOOLS.filter((t) => t.category === "write_irreversible");
    expect(irreversible).toHaveLength(1);
    expect(irreversible[0].definition.name).toBe("m365-mail:delete_message");
  });

  it("has zero permission_mutating tools in the v0.1 baseline", () => {
    const permMut = ALL_TOOLS.filter((t) => t.category === "permission_mutating");
    expect(permMut).toHaveLength(0);
  });
});

describe("buildHandlerMap", () => {
  it("maps every tool name to its handler", () => {
    const map = buildHandlerMap(ALL_TOOLS);
    expect(map.size).toBe(ALL_TOOLS.length);
    for (const t of ALL_TOOLS) {
      expect(map.get(t.definition.name)).toBe(t.handler);
    }
  });

  it("returns an empty map for an empty tool list", () => {
    const map = buildHandlerMap([]);
    expect(map.size).toBe(0);
  });
});
