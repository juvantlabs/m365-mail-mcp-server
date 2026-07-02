/**
 * Tool: m365-mail:create_draft
 *
 * Create a new draft message in the user's Drafts folder. Wraps Graph
 * `POST /me/messages` with `isDraft: true` implied (Graph sets it for
 * anything created via POST /me/messages).
 *
 * Required Graph scope: `Mail.ReadWrite` (delegated).
 *
 * Shield C4 (v0.1) — every draft this server creates carries:
 *   - subject prefix `[agent-draft] ` (added by buildMessageBody)
 *   - custom internet-message header `X-Juvant-Agent-Author` set to
 *     `@juvantlabs/m365-mail-mcp-server` (attached only in create_draft
 *     — Graph forbids setting internetMessageHeaders on PATCH; the
 *     reply/forward variants rely on the subject prefix alone).
 *
 * v0.1 boundary — send is NOT here. This tool creates the draft in
 * Outlook's Drafts folder; the user (or a future v0.3 send_draft tool
 * gated by Shield) is the only actor that pushes the message onto the
 * wire. That separation is the point: the agent can prep, the user
 * ships. See docs/adr/0001-v0.3-send-gate-contract.md.
 *
 * Attachments are v0.3+ (Graph upload session flow). v0.1 drafts are
 * text/HTML body only.
 *
 * No spec/approval pattern — creation is reversible (delete_message)
 * and doesn't send anything to third parties.
 */

import type { Client } from "@microsoft/microsoft-graph-client";

import { summarizeMessage } from "./list_messages.js";
import {
  DRAFT_BODY_SCHEMA_PROPERTIES,
  buildMessageBody,
} from "./_shared.js";
import type { Tool, ToolDefinition, ToolHandler, ToolResponse } from "../types/tool.js";

const definition: ToolDefinition = {
  name: "m365-mail:create_draft",
  description:
    "Create a new draft message in the user's Drafts folder. Does NOT send — the draft sits in Drafts until the user (or a later Shield-gated send_draft tool in v0.3) sends it. Every created draft is marked (subject prefix '[agent-draft] ' + custom header X-Juvant-Agent-Author) so agent-authored drafts are visually distinguishable. Returns the created draft's id + summary. Attachments are not supported in v0.1.",
  inputSchema: {
    type: "object",
    properties: {
      ...DRAFT_BODY_SCHEMA_PROPERTIES,
    },
    required: [],
  },
};

const handler: ToolHandler = async (
  graph: Client,
  args: Record<string, unknown>,
): Promise<ToolResponse> => {
  const body = buildMessageBody(args, {
    mode: "create",
    attachAgentHeader: true,
    createFallbackSubject: "",
  });

  const created = await graph.api("/me/messages").post(body);
  const summary = summarizeMessage(created);

  return {
    content: [
      {
        type: "text",
        text: JSON.stringify(
          {
            created: summary,
            agent_draft_markers: {
              subject_prefix: "[agent-draft] ",
              header: "X-Juvant-Agent-Author",
            },
            note: "Draft saved to Drafts folder. Not sent. Use update_draft to edit; use delete_message (two-phase) to remove.",
          },
          null,
          2,
        ),
      },
    ],
  };
};

export const createDraftTool: Tool = {
  category: "write_idempotent",
  definition,
  handler,
};
