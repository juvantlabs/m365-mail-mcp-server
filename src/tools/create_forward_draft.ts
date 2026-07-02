/**
 * Tool: m365-mail:create_forward_draft
 *
 * Create a forward draft from an existing message. Two-step Graph
 * flow (mirrors create_reply_draft):
 *   1. POST /me/messages/{id}/createForward → Graph returns a draft
 *      with quoted history and (by default) no To recipients.
 *   2. PATCH the draft with the caller's body/subject/recipient
 *      overrides, PLUS an unconditional subject PATCH to attach the
 *      Shield C4 `[agent-draft] ` marker to Graph's pre-filled
 *      `FW: <subject>`.
 *
 * Required Graph scope: `Mail.ReadWrite` (delegated).
 *
 * Unlike reply, forward has no pre-filled To — the caller should
 * supply `to` (and optionally `cc` / `bcc`) unless intentionally
 * leaving the draft addressee-less for the user to complete in
 * Outlook.
 *
 * Shield C4: forward drafts cannot carry the X-Juvant-Agent-Author
 * header (same Graph limitation as reply). The subject prefix is the
 * marker on this path.
 *
 * v0.1 does not send. Returns the newly created draft.
 */

import type { Client } from "@microsoft/microsoft-graph-client";

import { summarizeMessage } from "./list_messages.js";
import {
  AGENT_DRAFT_SUBJECT_PREFIX,
  DRAFT_BODY_SCHEMA_PROPERTIES,
  buildMessageBody,
  ensureAgentDraftSubject,
} from "./_shared.js";
import { validateRequiredString } from "../types/validators.js";
import type { Tool, ToolDefinition, ToolHandler, ToolResponse } from "../types/tool.js";

const definition: ToolDefinition = {
  name: "m365-mail:create_forward_draft",
  description:
    "Create a forward draft from an existing message. Graph pre-fills quoted history; supply `to` (and optionally `cc` / `bcc`) to address it. Optional `body` prepends a note above the quoted history. Does NOT send. The Shield '[agent-draft] ' subject prefix is applied unconditionally.",
  inputSchema: {
    type: "object",
    properties: {
      message_id: {
        type: "string",
        description: "Parent message id to forward.",
      },
      ...DRAFT_BODY_SCHEMA_PROPERTIES,
    },
    required: ["message_id"],
  },
};

const handler: ToolHandler = async (
  graph: Client,
  args: Record<string, unknown>,
): Promise<ToolResponse> => {
  const messageId = validateRequiredString(args.message_id, "message_id");

  const draft = (await graph
    .api(`/me/messages/${encodeURIComponent(messageId)}/createForward`)
    .post({})) as Record<string, unknown>;

  const draftId = String(draft.id ?? "");
  if (!draftId) {
    throw new Error(
      "Graph createForward returned no draft id — refusing to proceed with PATCH.",
    );
  }

  const patch = buildMessageBody(args, { mode: "patch" });
  if (patch.subject === undefined) {
    const currentSubject = String(draft.subject ?? "");
    patch.subject = ensureAgentDraftSubject(currentSubject);
  }

  const finalMessage = (await graph
    .api(`/me/messages/${encodeURIComponent(draftId)}`)
    .patch(patch)) as Record<string, unknown>;

  const summary = summarizeMessage(finalMessage);

  return {
    content: [
      {
        type: "text",
        text: JSON.stringify(
          {
            created: summary,
            parent_message_id: messageId,
            agent_draft_markers: {
              subject_prefix: AGENT_DRAFT_SUBJECT_PREFIX.trimEnd(),
              header: "not applicable — createForward-derived drafts cannot carry internetMessageHeaders",
            },
            note: "Forward draft saved to Drafts. Not sent.",
          },
          null,
          2,
        ),
      },
    ],
  };
};

export const createForwardDraftTool: Tool = {
  category: "write_idempotent",
  definition,
  handler,
};
