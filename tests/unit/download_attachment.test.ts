import os from "node:os";
import path from "node:path";

import type { Client } from "@microsoft/microsoft-graph-client";
import { afterEach, describe, expect, it, vi } from "vitest";

import {
  attachmentKind,
  deriveSafeLocalPath,
  downloadAttachmentTool,
  getSandboxRoot,
} from "../../src/tools/download_attachment.js";

describe("getSandboxRoot env-var precedence", () => {
  const originalEnv = { ...process.env };

  afterEach(() => {
    process.env = { ...originalEnv };
  });

  it("uses M365_MAIL_DOWNLOAD_DIR when set", () => {
    process.env.M365_MAIL_DOWNLOAD_DIR = "/custom/dir";
    expect(getSandboxRoot("tenant-1")).toBe(path.resolve("/custom/dir", "tenant-1"));
  });

  it("falls back to XDG_CACHE_HOME/m365-mail-mcp-server", () => {
    delete process.env.M365_MAIL_DOWNLOAD_DIR;
    process.env.XDG_CACHE_HOME = "/xdg-cache";
    expect(getSandboxRoot("tenant-1")).toBe(
      path.resolve("/xdg-cache", "m365-mail-mcp-server", "tenant-1"),
    );
  });

  it("falls back to ~/.cache/m365-mail-mcp-server when neither override is set", () => {
    delete process.env.M365_MAIL_DOWNLOAD_DIR;
    delete process.env.XDG_CACHE_HOME;
    expect(getSandboxRoot("tenant-1")).toBe(
      path.resolve(os.homedir(), ".cache", "m365-mail-mcp-server", "tenant-1"),
    );
  });

  it("uses a directory distinct from the m365-graph sandbox (isolation)", () => {
    delete process.env.M365_MAIL_DOWNLOAD_DIR;
    delete process.env.XDG_CACHE_HOME;
    const mailRoot = getSandboxRoot("tenant-1");
    const graphRoot = path.resolve(
      os.homedir(),
      ".cache",
      "m365-graph-mcp-server",
      "tenant-1",
    );
    expect(mailRoot).not.toBe(graphRoot);
  });

  it("scopes path per-tenant", () => {
    process.env.M365_MAIL_DOWNLOAD_DIR = "/custom";
    expect(getSandboxRoot("a")).not.toBe(getSandboxRoot("b"));
  });
});

describe("deriveSafeLocalPath", () => {
  const sandbox = "/sandbox/tenant-x";

  it("constructs filename as <hash>-<sanitized name>", () => {
    const p = deriveSafeLocalPath(sandbox, "msg-1", "att-1", "report.pdf");
    expect(p.startsWith(sandbox + "/")).toBe(true);
    expect(p).toMatch(/^.+\/[0-9a-f]{16}-report\.pdf$/);
  });

  it("sanitizes path-traversal payloads", () => {
    const p = deriveSafeLocalPath(sandbox, "m", "a", "../../etc/passwd");
    expect(p.startsWith(sandbox + "/")).toBe(true);
    expect(p).not.toContain("../");
    expect(p).not.toContain("/etc/passwd");
  });

  it("produces a different path when message_id changes but attachment_id + name are equal", () => {
    const a = deriveSafeLocalPath(sandbox, "msg-A", "att-shared", "same.pdf");
    const b = deriveSafeLocalPath(sandbox, "msg-B", "att-shared", "same.pdf");
    expect(a).not.toBe(b);
  });

  it("produces a different path when attachment_id changes but message_id + name are equal", () => {
    const a = deriveSafeLocalPath(sandbox, "msg-1", "att-A", "same.pdf");
    const b = deriveSafeLocalPath(sandbox, "msg-1", "att-B", "same.pdf");
    expect(a).not.toBe(b);
  });

  it("never escapes the sandbox root even with adversarial names", () => {
    for (const name of [
      "../malicious",
      "../../../../etc/passwd",
      "/absolute/path",
      "name\\with\\backslash",
    ]) {
      const p = deriveSafeLocalPath(sandbox, "m", "a", name);
      expect(p.startsWith(sandbox + "/")).toBe(true);
    }
  });
});

describe("attachmentKind", () => {
  it("extracts the trailing type name from @odata.type", () => {
    expect(attachmentKind({ "@odata.type": "#microsoft.graph.fileAttachment" })).toBe(
      "fileAttachment",
    );
  });

  it("returns 'unknown' when @odata.type is missing / empty", () => {
    expect(attachmentKind({})).toBe("unknown");
    expect(attachmentKind({ "@odata.type": "" })).toBe("unknown");
  });
});

describe("downloadAttachmentTool handler — pre-flight rejection paths", () => {
  function mockClient(meta: unknown): Client {
    const get = vi.fn().mockResolvedValue(meta);
    const select = vi.fn().mockReturnValue({ get });
    const api = vi.fn().mockReturnValue({ select });
    return { api } as unknown as Client;
  }

  it("requires message_id + attachment_id", async () => {
    const client = mockClient({});
    await expect(downloadAttachmentTool.handler(client, {})).rejects.toThrow(
      "'message_id' must be a non-empty string",
    );
    await expect(
      downloadAttachmentTool.handler(client, { message_id: "m1" }),
    ).rejects.toThrow("'attachment_id' must be a non-empty string");
  });

  it("rejects itemAttachment with a structured error message", async () => {
    const client = mockClient({
      id: "a1",
      name: "embedded",
      size: 100,
      "@odata.type": "#microsoft.graph.itemAttachment",
    });
    await expect(
      downloadAttachmentTool.handler(client, {
        message_id: "m1",
        attachment_id: "a1",
      }),
    ).rejects.toThrow(/attachment kind 'itemAttachment' is not downloadable/);
  });

  it("rejects referenceAttachment with a pointer to m365-graph-mcp-server", async () => {
    const client = mockClient({
      id: "a1",
      name: "link",
      size: 100,
      "@odata.type": "#microsoft.graph.referenceAttachment",
    });
    await expect(
      downloadAttachmentTool.handler(client, {
        message_id: "m1",
        attachment_id: "a1",
      }),
    ).rejects.toThrow(/referenceAttachment.*m365-graph/);
  });

  it("rejects attachments exceeding the 200 MB cap", async () => {
    const client = mockClient({
      id: "a1",
      name: "huge.bin",
      size: 250 * 1024 * 1024,
      "@odata.type": "#microsoft.graph.fileAttachment",
    });
    await expect(
      downloadAttachmentTool.handler(client, {
        message_id: "m1",
        attachment_id: "a1",
      }),
    ).rejects.toThrow(/exceeds the 200 MB cap/);
  });

  it("category is 'read'", () => {
    expect(downloadAttachmentTool.category).toBe("read");
  });
});
