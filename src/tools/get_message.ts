/**
 * Tool: m365-mail:get_message
 *
 * Fetch a single message's full detail, including the body. Wraps
 * Graph `GET /me/messages/{id}`.
 *
 * Required Graph scope: `Mail.Read` (delegated). Read-only.
 *
 * Body cap: ~16 KB of characters. Outlook messages are almost always
 * short by human standards, but HTML bodies with quoted history +
 * signatures + tracking pixels can bloat to hundreds of KB. Capping
 * at 16 000 chars keeps the MCP response payload predictable and the
 * agent's context budget bounded. `body_truncated: true` in the
 * response tells the agent bytes were dropped so it can, if needed,
 * open the `web_link` (or drop back to `body_preview`).
 */

import type { Client } from "@microsoft/microsoft-graph-client";

import { summarizeMessage } from "./list_messages.js";
import { validateRequiredString } from "../types/validators.js";
import type { Tool, ToolDefinition, ToolHandler, ToolResponse } from "../types/tool.js";

const BODY_CHAR_CAP = 16_000;

const definition: ToolDefinition = {
  name: "m365-mail:get_message",
  description:
    "Fetch full details for a single message — headers, recipients, and body. Body is capped at 16 000 characters; use the returned web_link for the untruncated view in Outlook. Read-only.",
  inputSchema: {
    type: "object",
    properties: {
      message_id: {
        type: "string",
        description: "Message id from list_messages or search_messages.",
      },
    },
    required: ["message_id"],
  },
};

interface FullMessage extends ReturnType<typeof summarizeMessage> {
  body_content_type: string;
  body: string;
  body_truncated: boolean;
  reply_to: Array<{ name: string; email: string }>;
  internet_message_id: string;
}

export function expandMessage(m: Record<string, unknown>): FullMessage {
  const summary = summarizeMessage(m);
  const body = m.body as Record<string, unknown> | undefined;
  const bodyContent = String(body?.content ?? "");
  const truncated = bodyContent.length > BODY_CHAR_CAP;
  const replyTo = Array.isArray(m.replyTo) ? m.replyTo : [];
  return {
    ...summary,
    body_content_type: String(body?.contentType ?? "text"),
    body: truncated ? bodyContent.slice(0, BODY_CHAR_CAP) : bodyContent,
    body_truncated: truncated,
    reply_to: replyTo.map((r) => {
      const email = (r as Record<string, unknown>).emailAddress as
        | Record<string, unknown>
        | undefined;
      return {
        name: String(email?.name ?? ""),
        email: String(email?.address ?? ""),
      };
    }),
    internet_message_id: String(m.internetMessageId ?? ""),
  };
}

const SELECT_FIELDS =
  "id,conversationId,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,isRead,isDraft,hasAttachments,importance,bodyPreview,webLink,parentFolderId,body,replyTo,internetMessageId";

const handler: ToolHandler = async (
  graph: Client,
  args: Record<string, unknown>,
): Promise<ToolResponse> => {
  const messageId = validateRequiredString(args.message_id, "message_id");

  const message = await graph
    .api(`/me/messages/${encodeURIComponent(messageId)}`)
    .select(SELECT_FIELDS)
    .get();

  const result = expandMessage(message);

  return {
    content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
  };
};

export const getMessageTool: Tool = { category: "read", definition, handler };
