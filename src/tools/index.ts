/**
 * Tool registry — single source of truth for which tools the server
 * exposes. Both the tools/list handler and the tools/call dispatcher
 * read from this module.
 *
 * Each new tool: add an import here + push into ALL_TOOLS. The
 * dispatcher in src/index.ts builds a map from tool name to handler
 * once at startup; no per-call registration overhead.
 *
 * v0.1 surface: 13 tools, own mailbox only. No send, no shared mailbox.
 */

import type { Tool } from "../types/tool.js";

// Read
import { listMailFoldersTool } from "./list_mail_folders.js";
import { listMessagesTool } from "./list_messages.js";
import { searchMessagesTool } from "./search_messages.js";
import { getMessageTool } from "./get_message.js";

// Attachments
import { listAttachmentsTool } from "./list_attachments.js";
import { downloadAttachmentTool } from "./download_attachment.js";

// Write — idempotent (draft authoring + mailbox curation)
import { createDraftTool } from "./create_draft.js";
import { updateDraftTool } from "./update_draft.js";
import { createReplyDraftTool } from "./create_reply_draft.js";
import { createForwardDraftTool } from "./create_forward_draft.js";
import { markReadTool } from "./mark_read.js";
import { moveMessageTool } from "./move_message.js";

// Write — irreversible (two-phase gated)
import { deleteMessageTool } from "./delete_message.js";

export const ALL_TOOLS: ReadonlyArray<Tool> = [
  // Read
  listMailFoldersTool,
  listMessagesTool,
  searchMessagesTool,
  getMessageTool,
  // Attachments
  listAttachmentsTool,
  downloadAttachmentTool,
  // Write — idempotent
  createDraftTool,
  updateDraftTool,
  createReplyDraftTool,
  createForwardDraftTool,
  markReadTool,
  moveMessageTool,
  // Write — irreversible
  deleteMessageTool,
];

export function buildHandlerMap(tools: ReadonlyArray<Tool>): Map<string, Tool["handler"]> {
  const m = new Map<string, Tool["handler"]>();
  for (const t of tools) {
    m.set(t.definition.name, t.handler);
  }
  return m;
}
