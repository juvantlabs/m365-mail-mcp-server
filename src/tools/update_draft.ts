/**
 * Tool: m365-mail:update_draft
 *
 * Update an existing draft message. Wraps Graph
 * `PATCH /me/messages/{id}`.
 *
 * Required Graph scope: `Mail.ReadWrite` (delegated).
 *
 * Guardrails:
 *   - The tool refuses PATCH if the message is NOT a draft (isDraft
 *     === false), because Graph will silently mutate a *sent* message
 *     otherwise (subject/body of an item already delivered — very
 *     confusing UX). One extra GET before the PATCH is worth it.
 *   - When updating recipients (to / cc / bcc), Graph REPLACES the
 *     entire list. Pass the full intended list, not a delta.
 *
 * Shield C4: if the caller patches the subject, the '[agent-draft] '
 * prefix is re-applied idempotently by buildMessageBody. The custom
 * X-Juvant-Agent-Author header is preserved from the original
 * create_draft POST (Graph does not overwrite immutable headers on
 * PATCH).
 *
 * No spec/approval pattern — updating a draft is reversible (the user
 * can revert in Outlook, or the agent can call update_draft again).
 */

import type { Client } from "@microsoft/microsoft-graph-client";

import { summarizeMessage } from "./list_messages.js";
import {
  DRAFT_BODY_SCHEMA_PROPERTIES,
  buildMessageBody,
} from "./_shared.js";
import { validateRequiredString } from "../types/validators.js";
import type { Tool, ToolDefinition, ToolHandler, ToolResponse } from "../types/tool.js";

const definition: ToolDefinition = {
  name: "m365-mail:update_draft",
  description:
    "Update an existing draft (subject, body, recipients, importance). Refuses to PATCH a message that is not a draft. WARNING: passing `to` / `cc` / `bcc` REPLACES the entire list — pass the full intended list, not a delta. The '[agent-draft] ' subject marker is re-applied idempotently on any subject update.",
  inputSchema: {
    type: "object",
    properties: {
      message_id: {
        type: "string",
        description: "Draft message id from list_messages / search_messages / create_draft.",
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
  const patch = buildMessageBody(args, { mode: "patch" });
  if (Object.keys(patch).length === 0) {
    throw new Error(
      "update_draft requires at least one field to update besides message_id",
    );
  }

  // Pre-flight: refuse non-drafts. One extra GET, but it prevents the
  // silent-mutation-of-sent-mail failure mode.
  const current = (await graph
    .api(`/me/messages/${encodeURIComponent(messageId)}`)
    .select("id,isDraft,subject")
    .get()) as Record<string, unknown>;
  if (current.isDraft !== true) {
    throw new Error(
      `Refusing to update: message ${messageId} is not a draft ` +
        `(isDraft=${String(current.isDraft)}). update_draft only edits messages in the Drafts folder.`,
    );
  }

  const updated = await graph
    .api(`/me/messages/${encodeURIComponent(messageId)}`)
    .patch(patch);
  const summary = summarizeMessage(updated);

  return {
    content: [
      {
        type: "text",
        text: JSON.stringify({ updated: summary }, null, 2),
      },
    ],
  };
};

export const updateDraftTool: Tool = {
  category: "write_idempotent",
  definition,
  handler,
};
