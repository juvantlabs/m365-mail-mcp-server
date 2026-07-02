/**
 * Tool: m365-mail:mark_read
 *
 * Toggle the read/unread state of a message. Wraps Graph
 * `PATCH /me/messages/{id}` with `{ isRead: <bool> }`.
 *
 * Required Graph scope: `Mail.ReadWrite` (delegated).
 *
 * Idempotent by construction: passing the same is_read value twice
 * leaves the mailbox in the same state. No spec/approval pattern
 * needed — the change is trivially reversible by calling this tool
 * again with the opposite value.
 */

import type { Client } from "@microsoft/microsoft-graph-client";

import {
  SHARED_USER_SCHEMA_PROPERTY,
  mailboxRoot,
  validateSharedUser,
} from "./_mailbox.js";
import {
  validateOptionalBoolean,
  validateRequiredString,
} from "../types/validators.js";
import type { Tool, ToolDefinition, ToolHandler, ToolResponse } from "../types/tool.js";

const definition: ToolDefinition = {
  name: "m365-mail:mark_read",
  description:
    "Mark a message as read (default) or unread. Idempotent. Returns the message id + new is_read state. Pass `shared_user` to mark a message in a shared / delegate mailbox (v0.2, requires Mail.ReadWrite.Shared).",
  inputSchema: {
    type: "object",
    properties: {
      message_id: {
        type: "string",
        description: "Message id to update.",
      },
      is_read: {
        type: "boolean",
        description: "Target state. Default true (mark as read).",
      },
      ...SHARED_USER_SCHEMA_PROPERTY,
    },
    required: ["message_id"],
  },
};

const handler: ToolHandler = async (
  graph: Client,
  args: Record<string, unknown>,
): Promise<ToolResponse> => {
  const messageId = validateRequiredString(args.message_id, "message_id");
  const isRead = validateOptionalBoolean(args.is_read, "is_read") ?? true;
  const sharedUser = validateSharedUser(args.shared_user);
  const root = mailboxRoot(sharedUser);

  const updated = (await graph
    .api(`${root}/messages/${encodeURIComponent(messageId)}`)
    .patch({ isRead })) as Record<string, unknown>;

  return {
    content: [
      {
        type: "text",
        text: JSON.stringify(
          {
            message_id: messageId,
            shared_user: sharedUser ?? null,
            is_read: Boolean(updated.isRead ?? isRead),
          },
          null,
          2,
        ),
      },
    ],
  };
};

export const markReadTool: Tool = {
  category: "write_idempotent",
  definition,
  handler,
};
