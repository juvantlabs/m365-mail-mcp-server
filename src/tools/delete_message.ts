/**
 * Tool: m365-mail:delete_message
 *
 * Delete a message. Two-phase spec/approval pattern per handbook
 * ADR 0002 § Delete-class operations:
 *
 *   Call 1 (no confirmation_token):
 *     - Tool fetches message metadata for preview.
 *     - Issues a confirmation_token tied to the spec {message_id}
 *       and to this tool name.
 *     - Returns the preview + token + expiry hint.
 *
 *   Call 2 (with confirmation_token):
 *     - Tool verifies the token (exists, not expired, tied to this
 *       tool, spec matches).
 *     - DELETEs the message.
 *     - Consumes the token (single-use).
 *
 * The DELETE endpoint (/me/messages/{id}) sends the item to Deleted
 * Items in Outlook — recoverable from the recycle bin (default
 * ~30-day retention). This is the same semantics as clicking Delete
 * in Outlook: not "shift-delete permanent", but not undoable by this
 * server either. Hence the two-phase gate.
 *
 * Required Graph scope: `Mail.ReadWrite` (delegated).
 */

import type { Client } from "@microsoft/microsoft-graph-client";

import {
  consumeConfirmation,
  issueConfirmation,
} from "../auth/confirmation_tokens.js";
import { validateRequiredString } from "../types/validators.js";
import type { Tool, ToolDefinition, ToolHandler, ToolResponse } from "../types/tool.js";

const TOOL_NAME = "m365-mail:delete_message";

const definition: ToolDefinition = {
  name: TOOL_NAME,
  description:
    "Delete a message (sends it to Deleted Items — recoverable from the Outlook recycle bin, not a hard delete). Two-phase: first call returns a preview + confirmation_token; second call (with the token + same args) executes the delete. Token is single-use, expires in 5 minutes, and tied to the exact spec — passing a different message_id with someone else's token fails.",
  inputSchema: {
    type: "object",
    properties: {
      message_id: {
        type: "string",
        description: "Message id to delete.",
      },
      confirmation_token: {
        type: "string",
        description:
          "Omit on the first call (preview mode). Include the token returned by the preview call to execute the delete.",
      },
    },
    required: ["message_id"],
  },
};

interface DeletePreviewItem {
  id: string;
  subject: string;
  from_email: string | null;
  received_at: string;
  is_draft: boolean;
  has_attachments: boolean;
  parent_folder_id: string | null;
  web_link: string;
}

function summarizePreview(m: Record<string, unknown>): DeletePreviewItem {
  const from = m.from as Record<string, unknown> | undefined;
  const fromEmail = (from?.emailAddress as Record<string, unknown> | undefined)?.address;
  return {
    id: String(m.id ?? ""),
    subject: String(m.subject ?? ""),
    from_email: (fromEmail as string | undefined) ?? null,
    received_at: String(m.receivedDateTime ?? ""),
    is_draft: Boolean(m.isDraft),
    has_attachments: Boolean(m.hasAttachments),
    parent_folder_id: (m.parentFolderId as string | undefined) ?? null,
    web_link: String(m.webLink ?? ""),
  };
}

const handler: ToolHandler = async (
  graph: Client,
  args: Record<string, unknown>,
): Promise<ToolResponse> => {
  const messageId = validateRequiredString(args.message_id, "message_id");
  // Optional: token may be absent (phase 1) or present (phase 2). We
  // read it defensively without validateOptionalString because we want
  // to distinguish "omitted" from "empty string" cleanly.
  const rawToken = args.confirmation_token;
  const confirmationToken =
    typeof rawToken === "string" && rawToken.length > 0 ? rawToken : undefined;

  const apiPath = `/me/messages/${encodeURIComponent(messageId)}`;
  const spec: Record<string, unknown> = { message_id: messageId };

  // ---------- Phase 1: preview ----------
  if (!confirmationToken) {
    const message = (await graph
      .api(apiPath)
      .select("id,subject,from,receivedDateTime,isDraft,hasAttachments,parentFolderId,webLink")
      .get()) as Record<string, unknown>;
    const summary = summarizePreview(message);
    const issued = issueConfirmation(TOOL_NAME, spec);
    return {
      content: [
        {
          type: "text",
          text: JSON.stringify(
            {
              preview: {
                item: summary,
                ...issued,
                instructions: `Re-call ${TOOL_NAME} with the SAME args (message_id) plus this confirmation_token to execute the delete.`,
              },
            },
            null,
            2,
          ),
        },
      ],
    };
  }

  // ---------- Phase 2: execute ----------
  const verdict = consumeConfirmation(confirmationToken, TOOL_NAME, spec);
  if (!verdict.ok) {
    throw new Error(
      `Refusing to delete: confirmation_token ${verdict.error}. ` +
        `Re-run without confirmation_token to get a fresh preview + token.`,
    );
  }

  await graph.api(apiPath).delete();

  return {
    content: [
      {
        type: "text",
        text: JSON.stringify(
          {
            deleted: { message_id: messageId },
            note: "DELETE executed against Graph. The message was moved to Deleted Items and may be recoverable from the user's Outlook recycle bin (default ~30-day retention).",
          },
          null,
          2,
        ),
      },
    ],
  };
};

export const deleteMessageTool: Tool = {
  category: "write_irreversible",
  definition,
  handler,
};
