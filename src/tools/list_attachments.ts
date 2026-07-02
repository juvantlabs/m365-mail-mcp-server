/**
 * Tool: m365-mail:list_attachments
 *
 * List the attachments on a message (metadata only — no content).
 * Wraps Graph `GET /me/messages/{id}/attachments` with a $select to
 * skip the `contentBytes` field (which would otherwise bloat the
 * response for every attachment, since the default projection
 * includes it).
 *
 * Required Graph scope: `Mail.Read` (delegated). Read-only.
 *
 * The `attachment_type` field distinguishes:
 *   - fileAttachment       — a real file → downloadable via
 *                            download_attachment
 *   - itemAttachment       — an embedded Outlook item (a message /
 *                            event forwarded as an attachment). Not
 *                            downloadable as a file.
 *   - referenceAttachment  — a link to a cloud file. Not downloadable
 *                            via /$value; agents should use the
 *                            attachment's sourceUrl in another tool.
 */

import type { Client } from "@microsoft/microsoft-graph-client";

import {
  validateOptionalInteger,
  validateRequiredString,
} from "../types/validators.js";
import type { Tool, ToolDefinition, ToolHandler, ToolResponse } from "../types/tool.js";

const definition: ToolDefinition = {
  name: "m365-mail:list_attachments",
  description:
    "List the attachments on a message (metadata only — no content bytes). Use download_attachment to fetch a fileAttachment's bytes to a local sandbox. itemAttachment / referenceAttachment kinds are listed but not downloadable via this server. Read-only.",
  inputSchema: {
    type: "object",
    properties: {
      message_id: {
        type: "string",
        description: "Message id from list_messages / search_messages / get_message.",
      },
      limit: {
        type: "integer",
        minimum: 1,
        maximum: 100,
        description: "Maximum attachments to list (default 25).",
      },
    },
    required: ["message_id"],
  },
};

interface AttachmentSummary {
  id: string;
  name: string;
  size: number;
  content_type: string;
  is_inline: boolean;
  attachment_type: string;
}

export function summarizeAttachment(a: Record<string, unknown>): AttachmentSummary {
  const rawType = String(a["@odata.type"] ?? "");
  // Graph returns e.g. "#microsoft.graph.fileAttachment" — normalize
  // to the trailing type name for a stable, agent-friendly value.
  const short = rawType.split(".").pop() ?? "unknown";
  return {
    id: String(a.id ?? ""),
    name: String(a.name ?? ""),
    size: Number(a.size ?? 0),
    content_type: String(a.contentType ?? ""),
    is_inline: Boolean(a.isInline),
    attachment_type: short || "unknown",
  };
}

const handler: ToolHandler = async (
  graph: Client,
  args: Record<string, unknown>,
): Promise<ToolResponse> => {
  const messageId = validateRequiredString(args.message_id, "message_id");
  const limit = validateOptionalInteger(args.limit, "limit", {
    min: 1,
    max: 100,
    default: 25,
  });

  const response = await graph
    .api(`/me/messages/${encodeURIComponent(messageId)}/attachments`)
    .select("id,name,size,contentType,isInline")
    .top(limit)
    .get();

  const items = Array.isArray(response?.value) ? response.value : [];
  const attachments = items.map((a: Record<string, unknown>) => summarizeAttachment(a));

  const result = {
    message_id: messageId,
    count: attachments.length,
    attachments,
  };

  return {
    content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
  };
};

export const listAttachmentsTool: Tool = { category: "read", definition, handler };
