import { describe, expect, it } from "vitest";

import {
  sanitizeFilename,
  validateOptionalBoolean,
  validateOptionalEnum,
  validateOptionalInteger,
  validateOptionalIntegerOrUndefined,
  validateOptionalString,
  validateRequiredString,
} from "../../src/types/validators.js";

describe("validateRequiredString", () => {
  it("returns the string when valid", () => {
    expect(validateRequiredString("foo", "field")).toBe("foo");
  });

  it("throws on empty string", () => {
    expect(() => validateRequiredString("", "field")).toThrow(
      "'field' must be a non-empty string",
    );
  });

  it("throws on undefined / null / non-string", () => {
    for (const bad of [undefined, null, 42, true, {}, []]) {
      expect(() => validateRequiredString(bad, "field")).toThrow();
    }
  });
});

describe("validateOptionalString", () => {
  it("returns undefined for undefined / null", () => {
    expect(validateOptionalString(undefined, "f")).toBeUndefined();
    expect(validateOptionalString(null, "f")).toBeUndefined();
  });

  it("returns the string when valid", () => {
    expect(validateOptionalString("foo", "f")).toBe("foo");
  });

  it("throws on empty / non-string", () => {
    expect(() => validateOptionalString("", "f")).toThrow();
    expect(() => validateOptionalString(42, "f")).toThrow();
  });
});

describe("validateOptionalInteger", () => {
  const opts = { min: 1, max: 10, default: 5 };

  it("returns default for undefined / null", () => {
    expect(validateOptionalInteger(undefined, "n", opts)).toBe(5);
    expect(validateOptionalInteger(null, "n", opts)).toBe(5);
  });

  it("returns the integer when in range (incl. boundaries)", () => {
    expect(validateOptionalInteger(1, "n", opts)).toBe(1);
    expect(validateOptionalInteger(7, "n", opts)).toBe(7);
    expect(validateOptionalInteger(10, "n", opts)).toBe(10);
  });

  it("throws below min / above max / non-integer / non-number", () => {
    expect(() => validateOptionalInteger(0, "n", opts)).toThrow("between 1 and 10");
    expect(() => validateOptionalInteger(11, "n", opts)).toThrow();
    expect(() => validateOptionalInteger(3.5, "n", opts)).toThrow();
    expect(() => validateOptionalInteger("5", "n", opts)).toThrow();
  });
});

describe("validateOptionalIntegerOrUndefined", () => {
  const opts = { min: 1, max: 100 };

  it("returns undefined for undefined / null (no forced default)", () => {
    expect(validateOptionalIntegerOrUndefined(undefined, "n", opts)).toBeUndefined();
    expect(validateOptionalIntegerOrUndefined(null, "n", opts)).toBeUndefined();
  });

  it("returns the integer when in range (incl. boundaries)", () => {
    expect(validateOptionalIntegerOrUndefined(1, "n", opts)).toBe(1);
    expect(validateOptionalIntegerOrUndefined(50, "n", opts)).toBe(50);
    expect(validateOptionalIntegerOrUndefined(100, "n", opts)).toBe(100);
  });

  it("throws below min / above max / non-integer / non-number", () => {
    expect(() => validateOptionalIntegerOrUndefined(0, "n", opts)).toThrow(
      "between 1 and 100",
    );
    expect(() => validateOptionalIntegerOrUndefined(101, "n", opts)).toThrow();
    expect(() => validateOptionalIntegerOrUndefined(3.5, "n", opts)).toThrow();
    expect(() => validateOptionalIntegerOrUndefined("5", "n", opts)).toThrow();
  });
});

describe("validateOptionalEnum", () => {
  const allowed = ["text", "html"] as const;

  it("returns default for undefined / null", () => {
    expect(validateOptionalEnum(undefined, "f", allowed, "text")).toBe("text");
    expect(validateOptionalEnum(null, "f", allowed, "text")).toBe("text");
  });

  it("returns the value when allowed", () => {
    expect(validateOptionalEnum("html", "f", allowed, "text")).toBe("html");
  });

  it("throws when not in allowed set", () => {
    expect(() => validateOptionalEnum("markdown", "f", allowed, "text")).toThrow(
      "must be one of",
    );
  });

  it("throws on non-string", () => {
    expect(() => validateOptionalEnum(42, "f", allowed, "text")).toThrow();
  });
});

describe("validateOptionalBoolean", () => {
  it("returns undefined for undefined / null", () => {
    expect(validateOptionalBoolean(undefined, "f")).toBeUndefined();
    expect(validateOptionalBoolean(null, "f")).toBeUndefined();
  });

  it("returns the boolean when valid", () => {
    expect(validateOptionalBoolean(true, "f")).toBe(true);
    expect(validateOptionalBoolean(false, "f")).toBe(false);
  });

  it("throws on non-boolean", () => {
    expect(() => validateOptionalBoolean("true", "f")).toThrow("must be a boolean");
    expect(() => validateOptionalBoolean(1, "f")).toThrow();
  });
});

describe("sanitizeFilename", () => {
  it("replaces slashes / backslashes / null bytes with underscores", () => {
    expect(sanitizeFilename("foo/bar/baz")).toBe("foo_bar_baz");
    expect(sanitizeFilename("foo\\bar")).toBe("foo_bar");
    expect(sanitizeFilename("foo\0bar")).toBe("foo_bar");
  });

  it("strips leading dots (path traversal defense)", () => {
    expect(sanitizeFilename("..hidden")).toBe("_hidden");
    expect(sanitizeFilename("...etc")).toBe("_etc");
  });

  it("preserves dots elsewhere in the name", () => {
    expect(sanitizeFilename("doc.v1.pdf")).toBe("doc.v1.pdf");
  });

  it("caps length at 200 chars", () => {
    const long = "a".repeat(500);
    const result = sanitizeFilename(long);
    expect(result.length).toBe(200);
  });

  it("neutralizes classic ../../etc/passwd payload", () => {
    const result = sanitizeFilename("../../etc/passwd");
    expect(result).not.toContain("/");
    expect(result.startsWith(".")).toBe(false);
  });
});
