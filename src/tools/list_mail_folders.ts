/**
 * Tool: m365-mail:list_mail_folders
 *
 * List the top-level mail folders in the user's mailbox (Inbox,
 * Drafts, Sent Items, Deleted Items, plus user-created folders).
 * Wraps Graph `GET /me/mailFolders` (or `/users/{upn}/mailFolders`
 * when `shared_user` is set — v0.2).
 *
 * Required Graph scope: `Mail.Read` (delegated). Read-only.
 * When `shared_user` is set, also requires `Mail.Read.Shared` (v0.2).
 *
 * NOTE: this returns top-level folders only. Nested child folders are
 * not expanded in v0.1 — pass a folder's id to `list_messages` to see
 * its contents. Discovery of child folders can be added later without
 * a breaking change.
 *
 * The tool's only scopable input parameter is `shared_user` (v0.2).
 * Without it, the tool enumerates the caller's full folder surface;
 * with it, the shared user's — but the shape "enumerate a mailbox's
 * top-level folders" is still non-message-scopable, so the tool
 * remains on the permission-surface NO_TARGET_TOOLS allowlist. See
 * tests/unit/permission_surface.test.ts.
 */

import type { Client } from "@microsoft/microsoft-graph-client";

import {
  SHARED_USER_SCHEMA_PROPERTY,
  mailboxRoot,
  validateSharedUser,
} from "./_mailbox.js";
import { validateOptionalInteger } from "../types/validators.js";
import type { Tool, ToolDefinition, ToolHandler, ToolResponse } from "../types/tool.js";

const definition: ToolDefinition = {
  name: "m365-mail:list_mail_folders",
  description:
    "List the top-level mail folders (Inbox, Drafts, Sent Items, Deleted Items, plus user-created folders). Read-only. Use the returned folder id with list_messages / move_message to scope to a specific folder. Well-known folder names (e.g. 'inbox', 'drafts', 'sentitems', 'deleteditems') also work as folder_id shortcuts in downstream tools. Pass `shared_user` to enumerate folders on a shared / delegate mailbox instead of the caller's own (v0.2, requires Mail.Read.Shared).",
  inputSchema: {
    type: "object",
    properties: {
      limit: {
        type: "integer",
        minimum: 1,
        maximum: 100,
        description: "Maximum number of folders to return (default 50).",
      },
      ...SHARED_USER_SCHEMA_PROPERTY,
    },
    required: [],
  },
};

interface FolderSummary {
  id: string;
  display_name: string;
  parent_folder_id: string | null;
  child_folder_count: number;
  unread_item_count: number;
  total_item_count: number;
}

export function summarizeFolder(folder: Record<string, unknown>): FolderSummary {
  return {
    id: String(folder.id ?? ""),
    display_name: String(folder.displayName ?? ""),
    parent_folder_id: (folder.parentFolderId as string | undefined) ?? null,
    child_folder_count: Number(folder.childFolderCount ?? 0),
    unread_item_count: Number(folder.unreadItemCount ?? 0),
    total_item_count: Number(folder.totalItemCount ?? 0),
  };
}

const handler: ToolHandler = async (
  graph: Client,
  args: Record<string, unknown>,
): Promise<ToolResponse> => {
  const limit = validateOptionalInteger(args.limit, "limit", {
    min: 1,
    max: 100,
    default: 50,
  });
  const sharedUser = validateSharedUser(args.shared_user);
  const root = mailboxRoot(sharedUser);

  const response = await graph
    .api(`${root}/mailFolders`)
    .select("id,displayName,parentFolderId,childFolderCount,unreadItemCount,totalItemCount")
    .top(limit)
    .get();
  const items = Array.isArray(response?.value) ? response.value : [];
  const folders = items.map((f: Record<string, unknown>) => summarizeFolder(f));

  const result = {
    count: folders.length,
    folders,
    shared_user: sharedUser ?? null,
  };

  return {
    content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
  };
};

export const listMailFoldersTool: Tool = { category: "read", definition, handler };
