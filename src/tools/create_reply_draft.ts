/**
 * Tool: m365-mail:create_reply_draft
 *
 * Create a reply draft against an existing message, optionally
 * populating body/subject/recipients in the same call. Two-step Graph
 * flow:
 *   1. POST /me/messages/{id}/createReply (or createReplyAll) → Graph
 *      returns a fully populated draft with quoted history + To/Cc
 *      pre-filled from the parent.
 *   2. PATCH /me/messages/{new-id} with any body/subject/recipients
 *      overrides the caller supplied, PLUS an unconditional subject
 *      PATCH to attach the Shield C4 `[agent-draft] ` marker to
 *      Graph's pre-filled `RE: <subject>`.
 *
 * Required Graph scope: `Mail.ReadWrite` (delegated).
 *
 * `reply_all` boolean picks which Graph action is invoked. `to` /
 * `cc` / `bcc` OVERRIDE the pre-filled recipients when supplied; omit
 * them to keep Graph's defaults (reply → original sender only;
 * reply_all → sender + all recipients).
 *
 * Shield C4: reply drafts cannot carry the X-Juvant-Agent-Author
 * header — Graph forbids setting internetMessageHeaders on PATCH, and
 * the createReply POST does not accept a `message` body. The subject
 * prefix is the marker on this path.
 *
 * v0.1 does not send. Returns the newly created draft message.
 */

import type { Client } from "@microsoft/microsoft-graph-client";

import { summarizeMessage } from "./list_messages.js";
import {
  AGENT_DRAFT_SUBJECT_PREFIX,
  DRAFT_BODY_SCHEMA_PROPERTIES,
  buildMessageBody,
  ensureAgentDraftSubject,
} from "./_shared.js";
import {
  SHARED_USER_SCHEMA_PROPERTY,
  mailboxRoot,
  validateSharedUser,
} from "./_mailbox.js";
import {
  validateOptionalBoolean,
  validateRequiredString,
} from "../types/validators.js";
import type { Tool, ToolDefinition, ToolHandler, ToolResponse } from "../types/tool.js";

const definition: ToolDefinition = {
  name: "m365-mail:create_reply_draft",
  description:
    "Create a reply (or reply-all) draft against an existing message. Graph pre-fills quoted history + To/Cc; supply `body` / `subject` / `to` / `cc` / `bcc` to override. Does NOT send. Set reply_all=true for reply-all. The Shield '[agent-draft] ' subject prefix is applied unconditionally. Pass `shared_user` to reply to a message in a shared / delegate mailbox (v0.2, requires Mail.ReadWrite.Shared); the reply draft lands in the same mailbox's Drafts folder.",
  inputSchema: {
    type: "object",
    properties: {
      message_id: {
        type: "string",
        description: "Parent message id to reply to.",
      },
      reply_all: {
        type: "boolean",
        description: "If true, use Graph createReplyAll. Default false (reply to sender only).",
      },
      ...DRAFT_BODY_SCHEMA_PROPERTIES,
      ...SHARED_USER_SCHEMA_PROPERTY,
    },
    required: ["message_id"],
  },
};

const handler: ToolHandler = async (
  graph: Client,
  args: Record<string, unknown>,
): Promise<ToolResponse> => {
  const messageId = validateRequiredString(args.message_id, "message_id");
  const replyAll = validateOptionalBoolean(args.reply_all, "reply_all") ?? false;
  const sharedUser = validateSharedUser(args.shared_user);
  const root = mailboxRoot(sharedUser);

  const action = replyAll ? "createReplyAll" : "createReply";
  const draft = (await graph
    .api(`${root}/messages/${encodeURIComponent(messageId)}/${action}`)
    .post({})) as Record<string, unknown>;

  const draftId = String(draft.id ?? "");
  if (!draftId) {
    throw new Error(
      `Graph ${action} returned no draft id — refusing to proceed with PATCH.`,
    );
  }

  // Build the caller's requested patch (subject prefix is applied
  // idempotently by buildMessageBody if the caller passed a subject).
  const patch = buildMessageBody(args, { mode: "patch" });

  // Shield C4: unconditionally mark the subject. If the caller didn't
  // patch subject, we mark Graph's pre-filled `RE: <original>`.
  if (patch.subject === undefined) {
    const currentSubject = String(draft.subject ?? "");
    patch.subject = ensureAgentDraftSubject(currentSubject);
  }

  const finalMessage = (await graph
    .api(`${root}/messages/${encodeURIComponent(draftId)}`)
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
            reply_all: replyAll,
            shared_user: sharedUser ?? null,
            agent_draft_markers: {
              subject_prefix: AGENT_DRAFT_SUBJECT_PREFIX.trimEnd(),
              header: "not applicable — createReply-derived drafts cannot carry internetMessageHeaders",
            },
            note: sharedUser
              ? `Reply draft saved to ${sharedUser}'s Drafts folder. Not sent.`
              : "Reply draft saved to Drafts. Not sent.",
          },
          null,
          2,
        ),
      },
    ],
  };
};

export const createReplyDraftTool: Tool = {
  category: "write_idempotent",
  definition,
  handler,
};
