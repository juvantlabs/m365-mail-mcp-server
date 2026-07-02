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

import { describe, expect, it } from "vitest";

import {
  SHARED_USER_SCHEMA_PROPERTY,
  mailboxRoot,
  validateSharedUser,
} from "../../src/tools/_mailbox.js";

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
});
