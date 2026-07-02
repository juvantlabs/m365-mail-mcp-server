/**
 * Tool: m365-mail:move_message
 *
 * Move a message to a different mail folder. Wraps Graph
 * `POST /me/messages/{id}/move` with `{ destinationId: <folder-id> }`.
 * Graph reassigns the message's parentFolderId synchronously and
 * returns the moved message resource.
 *
 * Required Graph scope: `Mail.ReadWrite` (delegated).
 *
 * `destination_folder_id` accepts either a folder id from
 * list_mail_folders OR a well-known name shortcut ('inbox',
 * 'drafts', 'sentitems', 'deleteditems', 'archive', 'junkemail').
 *
 * Move-to-Deleted-Items is equivalent to a soft delete (recoverable
 * from the recycle bin). For hard delete, use delete_message (which
 * is two-phase gated).
 *
 * No spec/approval pattern — moves are reversible by calling
 * move_message with the original folder id.
 */

import type { Client } from "@microsoft/microsoft-graph-client";

import { summarizeMessage } from "./list_messages.js";
import { validateRequiredString } from "../types/validators.js";
import type { Tool, ToolDefinition, ToolHandler, ToolResponse } from "../types/tool.js";

const definition: ToolDefinition = {
  name: "m365-mail:move_message",
  description:
    "Move a message to a different mail folder. destination_folder_id accepts folder ids from list_mail_folders or well-known names ('inbox', 'drafts', 'sentitems', 'deleteditems', 'archive', 'junkemail'). Reversible.",
  inputSchema: {
    type: "object",
    properties: {
      message_id: {
        type: "string",
        description: "Message id to move.",
      },
      destination_folder_id: {
        type: "string",
        description: "Destination folder id or well-known name.",
      },
    },
    required: ["message_id", "destination_folder_id"],
  },
};

const handler: ToolHandler = async (
  graph: Client,
  args: Record<string, unknown>,
): Promise<ToolResponse> => {
  const messageId = validateRequiredString(args.message_id, "message_id");
  const destinationFolderId = validateRequiredString(
    args.destination_folder_id,
    "destination_folder_id",
  );

  const moved = (await graph
    .api(`/me/messages/${encodeURIComponent(messageId)}/move`)
    .post({ destinationId: destinationFolderId })) as Record<string, unknown>;

  const summary = summarizeMessage(moved);

  return {
    content: [
      {
        type: "text",
        text: JSON.stringify(
          {
            moved: summary,
            source_message_id: messageId,
            destination_folder_id: destinationFolderId,
            note: "The moved message has a NEW id (Graph reissues ids on move). Use the returned `moved.id` for follow-up calls.",
          },
          null,
          2,
        ),
      },
    ],
  };
};

export const moveMessageTool: Tool = {
  category: "write_idempotent",
  definition,
  handler,
};
