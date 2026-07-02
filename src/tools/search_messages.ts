/**
 * Tool: m365-mail:search_messages
 *
 * Full-text search across the mailbox. Wraps Graph's `$search`
 * KQL-style query parameter on `/me/messages` (or on a specific
 * folder), routed via `mailboxRoot(shared_user)` (v0.2). Graph
 * forbids combining `$search` with `$orderby`, so results come back
 * relevance-ordered by Graph.
 *
 * Required Graph scope: `Mail.Read` (delegated). Read-only.
 * When `shared_user` is set, also requires `Mail.Read.Shared` (v0.2).
 *
 * KQL cheat sheet (agent-facing description mentions the common ones):
 *   from:alice@x.com
 *   subject:"invoice"
 *   received:>2026-06-01
 *   hasattachments:true
 *   "exact phrase"
 *
 * The query is passed through verbatim; agents are expected to craft
 * KQL, not free-form natural language.
 */

import type { Client } from "@microsoft/microsoft-graph-client";

import { summarizeMessage } from "./list_messages.js";
import {
  SHARED_USER_SCHEMA_PROPERTY,
  mailboxRoot,
  validateSharedUser,
} from "./_mailbox.js";
import {
  validateOptionalInteger,
  validateOptionalString,
  validateRequiredString,
} from "../types/validators.js";
import type { Tool, ToolDefinition, ToolHandler, ToolResponse } from "../types/tool.js";

const definition: ToolDefinition = {
  name: "m365-mail:search_messages",
  description:
    "Full-text search across the mailbox (or within a folder). Uses Graph $search / KQL — e.g. 'from:alice@x.com subject:\"invoice\"' or 'received:>2026-06-01 hasattachments:true'. Returns metadata + body_preview only (fetch full body with get_message). Read-only. Pass `shared_user` to search a shared / delegate mailbox instead of the caller's own (v0.2, requires Mail.Read.Shared).",
  inputSchema: {
    type: "object",
    properties: {
      query: {
        type: "string",
        description:
          "KQL search expression. Common fields: from, to, subject, body, received, hasattachments. Free-text terms match subject + body.",
      },
      folder_id: {
        type: "string",
        description:
          "Optional folder id (from list_mail_folders) or well-known name ('inbox', 'drafts', 'sentitems', 'deleteditems') to scope the search. Omit to search the whole mailbox.",
      },
      limit: {
        type: "integer",
        minimum: 1,
        maximum: 100,
        description: "Maximum matches to return (default 25).",
      },
      ...SHARED_USER_SCHEMA_PROPERTY,
    },
    required: ["query"],
  },
};

const SELECT_FIELDS =
  "id,conversationId,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,isRead,isDraft,hasAttachments,importance,bodyPreview,webLink,parentFolderId";

const handler: ToolHandler = async (
  graph: Client,
  args: Record<string, unknown>,
): Promise<ToolResponse> => {
  const query = validateRequiredString(args.query, "query");
  const folderId = validateOptionalString(args.folder_id, "folder_id");
  const limit = validateOptionalInteger(args.limit, "limit", {
    min: 1,
    max: 100,
    default: 25,
  });
  const sharedUser = validateSharedUser(args.shared_user);
  const root = mailboxRoot(sharedUser);

  const apiBase = folderId
    ? `${root}/mailFolders/${encodeURIComponent(folderId)}/messages`
    : `${root}/messages`;

  // $search accepts a quoted KQL string. Graph requires ConsistencyLevel: eventual
  // for $search on message endpoints.
  const response = await graph
    .api(apiBase)
    .header("ConsistencyLevel", "eventual")
    .search(`"${query.replace(/"/g, '\\"')}"`)
    .select(SELECT_FIELDS)
    .top(limit)
    .get();

  const items = Array.isArray(response?.value) ? response.value : [];
  const messages = items.map((m: Record<string, unknown>) => summarizeMessage(m));

  const result = {
    query,
    folder_id: folderId ?? null,
    shared_user: sharedUser ?? null,
    count: messages.length,
    messages,
  };

  return {
    content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
  };
};

export const searchMessagesTool: Tool = { category: "read", definition, handler };
