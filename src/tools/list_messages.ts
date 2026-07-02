/**
 * Tool: m365-mail:list_messages
 *
 * List messages in a mail folder, ordered by receivedDateTime desc
 * (newest first). Wraps Graph
 * `GET /me/mailFolders/{folder}/messages` (or `/me/messages` when no
 * folder_id is given).
 *
 * Required Graph scope: `Mail.Read` (delegated). Read-only.
 *
 * Body is intentionally NOT returned by this tool — use `get_message`
 * to fetch a single message's full body. The `body_preview` field
 * (Graph-provided, ~255 char text preview) is returned instead so the
 * agent can triage without pulling every full body into context.
 */

import type { Client } from "@microsoft/microsoft-graph-client";

import {
  validateOptionalInteger,
  validateOptionalString,
} from "../types/validators.js";
import type { Tool, ToolDefinition, ToolHandler, ToolResponse } from "../types/tool.js";

const definition: ToolDefinition = {
  name: "m365-mail:list_messages",
  description:
    "List messages in a mail folder, newest first (by receivedDateTime desc). Returns metadata + body_preview only — use get_message for a specific message's full body. Read-only. Well-known folder shortcuts accepted: 'inbox', 'drafts', 'sentitems', 'deleteditems'.",
  inputSchema: {
    type: "object",
    properties: {
      folder_id: {
        type: "string",
        description:
          "Optional folder id from list_mail_folders, or a well-known name ('inbox', 'drafts', 'sentitems', 'deleteditems'). Omit to search across all folders (/me/messages).",
      },
      limit: {
        type: "integer",
        minimum: 1,
        maximum: 100,
        description: "Maximum messages to return (default 25).",
      },
    },
    required: [],
  },
};

interface RecipientSummary {
  name: string;
  email: string;
}

interface MessageSummary {
  id: string;
  conversation_id: string;
  subject: string;
  from: RecipientSummary | null;
  to: RecipientSummary[];
  cc: RecipientSummary[];
  received_at: string;
  sent_at: string;
  is_read: boolean;
  is_draft: boolean;
  has_attachments: boolean;
  importance: string;
  body_preview: string;
  web_link: string;
  parent_folder_id: string | null;
}

function summarizeRecipient(r: Record<string, unknown>): RecipientSummary {
  const email = r.emailAddress as Record<string, unknown> | undefined;
  return {
    name: String(email?.name ?? ""),
    email: String(email?.address ?? ""),
  };
}

export function summarizeMessage(m: Record<string, unknown>): MessageSummary {
  const from = m.from as Record<string, unknown> | undefined;
  const toList = Array.isArray(m.toRecipients) ? m.toRecipients : [];
  const ccList = Array.isArray(m.ccRecipients) ? m.ccRecipients : [];
  return {
    id: String(m.id ?? ""),
    conversation_id: String(m.conversationId ?? ""),
    subject: String(m.subject ?? ""),
    from: from ? summarizeRecipient(from) : null,
    to: toList.map((r) => summarizeRecipient(r as Record<string, unknown>)),
    cc: ccList.map((r) => summarizeRecipient(r as Record<string, unknown>)),
    received_at: String(m.receivedDateTime ?? ""),
    sent_at: String(m.sentDateTime ?? ""),
    is_read: Boolean(m.isRead),
    is_draft: Boolean(m.isDraft),
    has_attachments: Boolean(m.hasAttachments),
    importance: String(m.importance ?? "normal"),
    body_preview: String(m.bodyPreview ?? ""),
    web_link: String(m.webLink ?? ""),
    parent_folder_id: (m.parentFolderId as string | undefined) ?? null,
  };
}

const SELECT_FIELDS =
  "id,conversationId,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,isRead,isDraft,hasAttachments,importance,bodyPreview,webLink,parentFolderId";

const handler: ToolHandler = async (
  graph: Client,
  args: Record<string, unknown>,
): Promise<ToolResponse> => {
  const folderId = validateOptionalString(args.folder_id, "folder_id");
  const limit = validateOptionalInteger(args.limit, "limit", {
    min: 1,
    max: 100,
    default: 25,
  });

  const apiBase = folderId
    ? `/me/mailFolders/${encodeURIComponent(folderId)}/messages`
    : "/me/messages";

  const response = await graph
    .api(apiBase)
    .select(SELECT_FIELDS)
    .top(limit)
    .orderby("receivedDateTime desc")
    .get();

  const items = Array.isArray(response?.value) ? response.value : [];
  const messages = items.map((m: Record<string, unknown>) => summarizeMessage(m));

  const result = {
    folder_id: folderId ?? null,
    count: messages.length,
    messages,
  };

  return {
    content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
  };
};

export const listMessagesTool: Tool = { category: "read", definition, handler };
