import type { Client } from "@microsoft/microsoft-graph-client";
import { describe, expect, it, vi } from "vitest";

import {
  listAttachmentsTool,
  summarizeAttachment,
} from "../../src/tools/list_attachments.js";

function captureRequest(returnValue: unknown): {
  apiCalls: string[];
  client: Client;
} {
  const apiCalls: string[] = [];
  const get = vi.fn().mockResolvedValue(returnValue);
  const top = vi.fn().mockReturnValue({ get });
  const select = vi.fn().mockReturnValue({ top });
  const api = vi.fn().mockImplementation((path: string) => {
    apiCalls.push(path);
    return { select };
  });
  return { apiCalls, client: { api } as unknown as Client };
}

describe("summarizeAttachment", () => {
  it("normalizes @odata.type to short kind name", () => {
    const s = summarizeAttachment({
      id: "a1",
      name: "report.pdf",
      size: 2048,
      contentType: "application/pdf",
      isInline: false,
      "@odata.type": "#microsoft.graph.fileAttachment",
    });
    expect(s.attachment_type).toBe("fileAttachment");
    expect(s.name).toBe("report.pdf");
    expect(s.size).toBe(2048);
    expect(s.is_inline).toBe(false);
  });

  it("maps itemAttachment / referenceAttachment to their short names", () => {
    expect(
      summarizeAttachment({ "@odata.type": "#microsoft.graph.itemAttachment" }).attachment_type,
    ).toBe("itemAttachment");
    expect(
      summarizeAttachment({ "@odata.type": "#microsoft.graph.referenceAttachment" }).attachment_type,
    ).toBe("referenceAttachment");
  });

  it("falls back to 'unknown' when @odata.type is missing", () => {
    expect(summarizeAttachment({}).attachment_type).toBe("unknown");
  });
});

describe("listAttachmentsTool handler", () => {
  it("requires message_id", async () => {
    const { client } = captureRequest({ value: [] });
    await expect(listAttachmentsTool.handler(client, {})).rejects.toThrow(
      "'message_id' must be a non-empty string",
    );
  });

  it("calls /me/messages/{id}/attachments with URL-encoded id", async () => {
    const { apiCalls, client } = captureRequest({ value: [] });
    await listAttachmentsTool.handler(client, { message_id: "m/1" });
    expect(apiCalls[0]).toBe("/me/messages/m%2F1/attachments");
  });

  it("returns count + attachments + message_id", async () => {
    const { client } = captureRequest({
      value: [
        {
          id: "a1",
          name: "x.pdf",
          size: 100,
          contentType: "application/pdf",
          "@odata.type": "#microsoft.graph.fileAttachment",
        },
      ],
    });
    const resp = await listAttachmentsTool.handler(client, { message_id: "m1" });
    const parsed = JSON.parse((resp.content[0] as { type: string; text: string }).text);
    expect(parsed.message_id).toBe("m1");
    expect(parsed.count).toBe(1);
    expect(parsed.attachments[0].attachment_type).toBe("fileAttachment");
  });

  it("category is 'read'", () => {
    expect(listAttachmentsTool.category).toBe("read");
  });
});
