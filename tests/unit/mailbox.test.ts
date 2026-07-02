/**
 * Unit tests for `src/tools/_mailbox.ts` — the v0.2 shared_user helper.
 *
 * Covers the two load-bearing exports:
 *   - `validateSharedUser` — UPN-shape validation (accept / reject /
 *     coerce).
 *   - `mailboxRoot` — `/me` vs `/users/{encoded-upn}` router.
 *
 * And the schema constant:
 *   - `SHARED_USER_SCHEMA_PROPERTY` — pinned wording so the JSON schema
 *     description doesn't silently drift.
 *
 * These are the SINGLE points of composition every tool depends on. If
 * the shape here regresses, every tool regresses.
 */

import { afterEach, beforeEach, describe, expect, it } from "vitest";

import {
  SHARED_USER_SCHEMA_PROPERTY,
  mailboxRoot,
  validateSharedUser,
} from "../../src/tools/_mailbox.js";

// v0.2.1 — most tests in this file predate the server-side allowlist.
// To keep them focused on the shape / normalization behaviour they
// were written for, we install a wildcard allowlist by default. Tests
// that specifically exercise the allowlist below manage
// `M365_MAIL_ALLOWED_SHARED_USERS` explicitly per-case.
const ALLOWLIST_ENV = "M365_MAIL_ALLOWED_SHARED_USERS";
let savedAllowlist: string | undefined;

beforeEach(() => {
  savedAllowlist = process.env[ALLOWLIST_ENV];
  process.env[ALLOWLIST_ENV] = "*";
});

afterEach(() => {
  if (savedAllowlist === undefined) {
    delete process.env[ALLOWLIST_ENV];
  } else {
    process.env[ALLOWLIST_ENV] = savedAllowlist;
  }
});

describe("validateSharedUser", () => {
  it("returns undefined for undefined / null / empty string", () => {
    expect(validateSharedUser(undefined)).toBeUndefined();
    expect(validateSharedUser(null)).toBeUndefined();
    expect(validateSharedUser("")).toBeUndefined();
    // Blank-string-after-trim is treated as absent, not as a bad UPN.
    expect(validateSharedUser("   ")).toBeUndefined();
  });

  it("accepts a plain UPN", () => {
    expect(validateSharedUser("finance@juvant.io")).toBe("finance@juvant.io");
  });

  it("trims incidental whitespace before validating", () => {
    expect(validateSharedUser("  finance@juvant.io  ")).toBe("finance@juvant.io");
  });

  it("rejects a string without an @ character", () => {
    expect(() => validateSharedUser("financejuvantio")).toThrow(/UPN/);
  });

  it("rejects a string without a domain suffix (no dot)", () => {
    expect(() => validateSharedUser("finance@localhost")).toThrow(/UPN/);
  });

  it("rejects a string with embedded whitespace", () => {
    expect(() => validateSharedUser("finance @juvant.io")).toThrow(/UPN/);
    expect(() => validateSharedUser("finance@juvant .io")).toThrow(/UPN/);
  });

  it("rejects a string with multiple @ characters", () => {
    expect(() => validateSharedUser("a@b@juvant.io")).toThrow(/UPN/);
  });

  it("rejects a GUID (proposal: UPN only, not GUID)", () => {
    // A GUID would be a valid Graph user id, but v0.2 restricts to UPN
    // shape as the caller-facing contract. If this changes, the ADR
    // and description must move with it.
    expect(() =>
      validateSharedUser("c557607d-995c-4eb7-967b-50c6361fbad9"),
    ).toThrow(/UPN/);
  });

  it("rejects non-string inputs (defense-in-depth)", () => {
    expect(() => validateSharedUser(42)).toThrow(/non-empty string/);
    expect(() => validateSharedUser({ upn: "x@y.z" })).toThrow(/non-empty string/);
    expect(() => validateSharedUser(true)).toThrow(/non-empty string/);
  });

  it("uses the fieldName argument in error messages (so callers get a specific field)", () => {
    expect(() => validateSharedUser("bad", "custom_field")).toThrow(/'custom_field'/);
  });

  // ─── v0.2 (FUP-2): case normalization at the boundary ──────────────
  //
  // UPNs are case-insensitive in Entra / Exchange, but the value flows
  // into two byte-sensitive client-side hashes (delete_message
  // confirmation-token spec-hash; download_attachment sandbox path).
  // Without normalization here, `Finance@juvant.io` and
  // `finance@juvant.io` would issue non-interchangeable tokens and
  // write to two separate sandbox files even though they name the same
  // mailbox. The invariant is: validateSharedUser is the SINGLE point
  // where the raw value enters the system; every downstream consumer
  // must see the canonical (lowercase) form.
  it("lowercases the local-part of the UPN", () => {
    expect(validateSharedUser("Finance@juvant.io")).toBe("finance@juvant.io");
  });

  it("lowercases the domain of the UPN", () => {
    expect(validateSharedUser("finance@JUVANT.IO")).toBe("finance@juvant.io");
  });

  it("lowercases mixed-case in both parts", () => {
    expect(validateSharedUser("Finance@Juvant.Io")).toBe("finance@juvant.io");
  });

  it("trims first, then lowercases (order-independent for shape but pinned here)", () => {
    expect(validateSharedUser("  Finance@Juvant.Io  ")).toBe("finance@juvant.io");
  });

  it("mixed-case and lowercase inputs return the SAME value (byte-equal)", () => {
    const mixed = validateSharedUser("Finance@Juvant.Io");
    const lower = validateSharedUser("finance@juvant.io");
    expect(mixed).toBe(lower);
  });
});

describe("mailboxRoot", () => {
  it("returns /me when sharedUser is undefined", () => {
    expect(mailboxRoot(undefined)).toBe("/me");
  });

  it("returns /users/{encoded-upn} when sharedUser is set", () => {
    expect(mailboxRoot("finance@juvant.io")).toBe(
      "/users/finance%40juvant.io",
    );
  });

  it("URL-encodes special characters in the UPN", () => {
    // Rare but valid: a UPN with a plus-sign or apostrophe. Graph
    // accepts either raw or encoded, but encoded is safer.
    expect(mailboxRoot("finance+ap@juvant.io")).toBe(
      "/users/finance%2Bap%40juvant.io",
    );
  });

  it("never returns a path with double-slash (composes cleanly with `/messages`)", () => {
    // Downstream tools compose `${mailboxRoot(u)}/messages/...`. This
    // must never produce `//messages` or `//users//…`.
    expect(mailboxRoot(undefined) + "/messages").toBe("/me/messages");
    expect(mailboxRoot("a@b.c") + "/messages").toBe("/users/a%40b.c/messages");
  });

  // ─── v0.2 (FUP-2): mailboxRoot passthrough is case-preserving ──────
  //
  // Case normalization is `validateSharedUser`'s job — `mailboxRoot`
  // is a pure formatter and MUST NOT do a second normalization pass.
  // Every tool calls `mailboxRoot(validateSharedUser(...))` in that
  // order, and Graph itself is case-insensitive on the URL, so this
  // just documents the layering: normalize once at the input boundary,
  // pass through everywhere else.
  it("is case-preserving (normalization is validateSharedUser's job, not this one)", () => {
    expect(mailboxRoot("Finance@Juvant.Io")).toBe(
      "/users/Finance%40Juvant.Io",
    );
  });
});

describe("SHARED_USER_SCHEMA_PROPERTY", () => {
  it("exposes exactly the shared_user property (no other siblings)", () => {
    expect(Object.keys(SHARED_USER_SCHEMA_PROPERTY)).toEqual(["shared_user"]);
  });

  it("shared_user is a string type with a description mentioning Mail.*.Shared", () => {
    const prop = SHARED_USER_SCHEMA_PROPERTY.shared_user;
    expect(prop.type).toBe("string");
    expect(prop.description).toContain("User Principal Name");
    expect(prop.description).toContain("Mail.Read.Shared");
    expect(prop.description).toContain("Mail.ReadWrite.Shared");
  });

  it("description spells out that access is enforced by Exchange (not by the parameter)", () => {
    // Security-model reminder — we want callers (agents included) to
    // understand that setting the parameter does NOT grant access.
    expect(SHARED_USER_SCHEMA_PROPERTY.shared_user.description).toMatch(
      /Access is enforced by Exchange/,
    );
  });

  it("description mentions the v0.2.1 server-side allowlist (M365_MAIL_ALLOWED_SHARED_USERS)", () => {
    // Callers need to be told upfront that the allowlist exists and
    // is fail-closed — so a rejection is a policy artefact, not a
    // shape bug they will loop trying to fix by re-typing.
    expect(SHARED_USER_SCHEMA_PROPERTY.shared_user.description).toContain(
      "M365_MAIL_ALLOWED_SHARED_USERS",
    );
    expect(SHARED_USER_SCHEMA_PROPERTY.shared_user.description).toContain(
      "fail-closed",
    );
  });
});

// ────────────────────────────────────────────────────────────────────
// v0.2.1 — server-side allowlist (M365_MAIL_ALLOWED_SHARED_USERS)
//
// Second, fail-closed enforcement layer on top of Exchange delegation.
// The invariants tested here are load-bearing:
//   - the gate sits AFTER UPN-shape validation and case normalization
//     (a mis-shaped UPN throws with the shape error, not the allowlist
//     error — the caller sees the specific bug they need to fix);
//   - the gate sits BEFORE anything the caller does with the return
//     value (i.e. before mailboxRoot / any Graph call);
//   - a supplied-but-unallowed UPN is REJECTED, never silently
//     downgraded to `/me` (that would be a scope surprise);
//   - unset / empty env is fail-closed for shared calls but
//     unaffected for own-mailbox calls (omitting `shared_user`);
//   - `*` restores v0.2 Exchange-only behaviour;
//   - env-var entries are compared case-insensitively so the operator
//     can list `Finance@juvant.io` and match a caller `finance@…`.
// ────────────────────────────────────────────────────────────────────
describe("validateSharedUser — allowlist (v0.2.1)", () => {
  it("accepts a UPN that is on the allowlist", () => {
    process.env[ALLOWLIST_ENV] = "finance@juvant.io,legal@juvant.io";
    expect(validateSharedUser("finance@juvant.io")).toBe("finance@juvant.io");
    expect(validateSharedUser("legal@juvant.io")).toBe("legal@juvant.io");
  });

  it("rejects a shape-valid UPN that is NOT on the allowlist", () => {
    process.env[ALLOWLIST_ENV] = "finance@juvant.io";
    expect(() => validateSharedUser("press@juvant.io")).toThrow(
      /M365_MAIL_ALLOWED_SHARED_USERS/,
    );
  });

  it("rejects ALL shared_user values when the env var is unset (fail-closed default)", () => {
    delete process.env[ALLOWLIST_ENV];
    expect(() => validateSharedUser("finance@juvant.io")).toThrow(
      /M365_MAIL_ALLOWED_SHARED_USERS/,
    );
    expect(() => validateSharedUser("anything@juvant.io")).toThrow(
      /M365_MAIL_ALLOWED_SHARED_USERS/,
    );
  });

  it("rejects ALL shared_user values when the env var is empty / whitespace-only", () => {
    process.env[ALLOWLIST_ENV] = "";
    expect(() => validateSharedUser("finance@juvant.io")).toThrow(
      /M365_MAIL_ALLOWED_SHARED_USERS/,
    );
    process.env[ALLOWLIST_ENV] = "   ";
    expect(() => validateSharedUser("finance@juvant.io")).toThrow(
      /M365_MAIL_ALLOWED_SHARED_USERS/,
    );
  });

  it("still returns undefined for omitted shared_user even when the env var is unset (own-mailbox `/me` still works)", () => {
    // The allowlist gates ONLY the shared path. An agent that
    // never sets `shared_user` must still be able to call every
    // tool against its own mailbox on a server with no allowlist
    // configured (v0.1 baseline).
    delete process.env[ALLOWLIST_ENV];
    expect(validateSharedUser(undefined)).toBeUndefined();
    expect(validateSharedUser(null)).toBeUndefined();
    expect(validateSharedUser("")).toBeUndefined();
    expect(validateSharedUser("   ")).toBeUndefined();
  });

  it("`*` sentinel restores v0.2 Exchange-only behaviour (any UPN accepted)", () => {
    process.env[ALLOWLIST_ENV] = "*";
    expect(validateSharedUser("finance@juvant.io")).toBe("finance@juvant.io");
    expect(validateSharedUser("press@juvant.io")).toBe("press@juvant.io");
    expect(validateSharedUser("random@example.com")).toBe("random@example.com");
  });

  it("`*` with surrounding whitespace is still the wildcard sentinel", () => {
    process.env[ALLOWLIST_ENV] = "  *  ";
    expect(validateSharedUser("finance@juvant.io")).toBe("finance@juvant.io");
  });

  it("compares case-insensitively — env var UPN in mixed case matches caller in lowercase", () => {
    process.env[ALLOWLIST_ENV] = "Finance@Juvant.Io";
    expect(validateSharedUser("finance@juvant.io")).toBe("finance@juvant.io");
  });

  it("compares case-insensitively — caller UPN in mixed case matches env var in lowercase", () => {
    process.env[ALLOWLIST_ENV] = "finance@juvant.io";
    expect(validateSharedUser("Finance@Juvant.Io")).toBe("finance@juvant.io");
  });

  it("trims whitespace around comma-separated entries", () => {
    process.env[ALLOWLIST_ENV] = "  finance@juvant.io  ,  legal@juvant.io  ";
    expect(validateSharedUser("finance@juvant.io")).toBe("finance@juvant.io");
    expect(validateSharedUser("legal@juvant.io")).toBe("legal@juvant.io");
  });

  it("ignores empty entries (e.g. trailing comma) without matching them to caller input", () => {
    process.env[ALLOWLIST_ENV] = "finance@juvant.io,,legal@juvant.io,";
    expect(validateSharedUser("finance@juvant.io")).toBe("finance@juvant.io");
    expect(() => validateSharedUser("press@juvant.io")).toThrow(
      /M365_MAIL_ALLOWED_SHARED_USERS/,
    );
  });

  it("rejects a mis-SHAPED UPN with the shape error, NOT the allowlist error (shape check fires first)", () => {
    // Load-bearing order: shape → normalize → allowlist. A caller
    // fixing a typo should see "UPN shape", not "not on allowlist" —
    // the latter would send them chasing the wrong bug.
    process.env[ALLOWLIST_ENV] = "*";
    expect(() => validateSharedUser("not-a-upn")).toThrow(/UPN/);
    expect(() => validateSharedUser("not-a-upn")).not.toThrow(
      /M365_MAIL_ALLOWED_SHARED_USERS/,
    );
  });

  it("allowlist error message names M365_MAIL_ALLOWED_SHARED_USERS so operators know where to fix policy", () => {
    process.env[ALLOWLIST_ENV] = "finance@juvant.io";
    try {
      validateSharedUser("press@juvant.io");
      throw new Error("expected validateSharedUser to throw");
    } catch (err) {
      const message = (err as Error).message;
      expect(message).toContain("M365_MAIL_ALLOWED_SHARED_USERS");
      expect(message).toContain("press@juvant.io");
    }
  });

  it("unset-env error message names the env var AND explains own-mailbox is unaffected", () => {
    delete process.env[ALLOWLIST_ENV];
    try {
      validateSharedUser("finance@juvant.io");
      throw new Error("expected validateSharedUser to throw");
    } catch (err) {
      const message = (err as Error).message;
      expect(message).toContain("M365_MAIL_ALLOWED_SHARED_USERS");
      // Operator hint: they can either enumerate mailboxes or use '*'.
      expect(message).toContain("*");
      // Caller hint: omitting the parameter still works.
      expect(message).toMatch(/omit/i);
    }
  });

  it("never silently downgrades a rejected shared_user to `/me` — the return is a THROW, not undefined", () => {
    // This is the confused-deputy invariant. If we returned
    // undefined on rejection, mailboxRoot would compose `/me` and
    // the agent's shared-mailbox intent would silently execute
    // against the wrong mailbox. Rejection MUST be visible.
    process.env[ALLOWLIST_ENV] = "finance@juvant.io";
    expect(() => validateSharedUser("press@juvant.io")).toThrow();
    // The throw path never yields undefined; if the implementation
    // is ever "softened" to a downgrade this test flips red.
    let returned: unknown = "sentinel";
    try {
      returned = validateSharedUser("press@juvant.io");
    } catch {
      returned = "threw";
    }
    expect(returned).toBe("threw");
  });
});
