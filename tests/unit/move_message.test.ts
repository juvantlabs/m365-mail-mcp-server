import type { Client } from "@microsoft/microsoft-graph-client";
import { describe, expect, it, vi } from "vitest";

import { moveMessageTool } from "../../src/tools/move_message.js";

function captureRequest(returnValue: unknown): {
  apiCalls: string[];
  postBodies: Record<string, unknown>[];
  client: Client;
} {
  const apiCalls: string[] = [];
  const postBodies: Record<string, unknown>[] = [];
  const post = vi.fn().mockImplementation((b: Record<string, unknown>) => {
    postBodies.push(b);
    return Promise.resolve(returnValue);
  });
  const api = vi.fn().mockImplementation((p: string) => {
    apiCalls.push(p);
    return { post };
  });
  return { apiCalls, postBodies, client: { api } as unknown as Client };
}

describe("moveMessageTool handler", () => {
  it("requires message_id + destination_folder_id", async () => {
    const { client } = captureRequest({ id: "m1-new" });
    await expect(moveMessageTool.handler(client, {})).rejects.toThrow("'message_id'");
    await expect(
      moveMessageTool.handler(client, { message_id: "m1" }),
    ).rejects.toThrow("'destination_folder_id'");
  });

  it("POSTs to /me/messages/{id}/move with destinationId body", async () => {
    const { apiCalls, postBodies, client } = captureRequest({
      id: "m1-new",
      subject: "x",
    });
    await moveMessageTool.handler(client, {
      message_id: "m1",
      destination_folder_id: "archive",
    });
    expect(apiCalls[0]).toBe("/me/messages/m1/move");
    expect(postBodies[0]).toEqual({ destinationId: "archive" });
  });

  it("URL-encodes the message_id in the path", async () => {
    const { apiCalls, client } = captureRequest({ id: "n" });
    await moveMessageTool.handler(client, {
      message_id: "with space",
      destination_folder_id: "inbox",
    });
    expect(apiCalls[0]).toBe("/me/messages/with%20space/move");
  });

  it("returns moved summary + source_message_id + destination_folder_id", async () => {
    const { client } = captureRequest({ id: "m1-new", subject: "x" });
    const resp = await moveMessageTool.handler(client, {
      message_id: "m1",
      destination_folder_id: "archive",
    });
    const parsed = JSON.parse((resp.content[0] as { type: string; text: string }).text);
    expect(parsed.source_message_id).toBe("m1");
    expect(parsed.destination_folder_id).toBe("archive");
    expect(parsed.moved.id).toBe("m1-new");
  });

  it("category is 'write_idempotent'", () => {
    expect(moveMessageTool.category).toBe("write_idempotent");
  });
});
