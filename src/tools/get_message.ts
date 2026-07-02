/**
 * Tool: m365-mail:get_message
 *
 * Fetch a single message's full detail, including the body. Wraps
 * Graph `GET /me/messages/{id}`.
 *
 * Required Graph scope: `Mail.Read` (delegated). Read-only.
 *
 * Body handling — v0.1 behavior (no silent truncation):
 *   - Default (no pagination params): return the FULL body. `body_truncated`
 *     is always `false`. There is no silent cap.
 *   - Optional caller-driven pagination via `body_offset` +
 *     `max_body_chars` (mirrors the sibling m365-graph `get_transcript`
 *     offset/max_chars contract). When `max_body_chars` is set the
 *     response returns the slice `body[body_offset : body_offset +
 *     max_body_chars]` and includes `next_offset` iff there is more
 *     content beyond the returned slice. `next_offset` is `null` when
 *     the slice reaches the end.
 *
 * `body_truncated` is retained for readability but is strictly equivalent
 * to `next_offset !== null` — it is NEVER `true` for a paramless call.
 */

import type { Client } from "@microsoft/microsoft-graph-client";

import { summarizeMessage } from "./list_messages.js";
import {
  validateOptionalInteger,
  validateOptionalIntegerOrUndefined,
  validateRequiredString,
} from "../types/validators.js";
import type { Tool, ToolDefinition, ToolHandler, ToolResponse } from "../types/tool.js";

// Sanity bounds on the pagination inputs. These are NOT a body cap — the
// body itself is never truncated silently. They just prevent absurd
// caller inputs from generating NaN math.
const MAX_OFFSET = 2_000_000_000;
const MAX_PAGE_CHARS = 2_000_000_000;

const definition: ToolDefinition = {
  name: "m365-mail:get_message",
  description:
    "Fetch full details for a single message — headers, recipients, and body. " +
    "By default the entire body is returned untruncated. Callers may optionally " +
    "page through very long bodies by passing max_body_chars (and body_offset for " +
    "the continuation offset returned as next_offset in the previous response). " +
    "Read-only.",
  inputSchema: {
    type: "object",
    properties: {
      message_id: {
        type: "string",
        description: "Message id from list_messages or search_messages.",
      },
      body_offset: {
        type: "integer",
        minimum: 0,
        maximum: MAX_OFFSET,
        description:
          "Character offset into the message body (default 0). Pass the previous " +
          "response's next_offset to fetch the next chunk. Only meaningful when " +
          "max_body_chars is also set.",
      },
      max_body_chars: {
        type: "integer",
        minimum: 1,
        maximum: MAX_PAGE_CHARS,
        description:
          "Optional. When set, the response body is sliced to at most this many " +
          "characters starting at body_offset; next_offset in the response points " +
          "to the next chunk (or is null when the slice reaches the end). When " +
          "omitted, the full body is returned untruncated.",
      },
    },
    required: ["message_id"],
  },
};

/**
 * Base expansion of the Graph message payload — the full body, no
 * pagination applied. The body-slicing logic lives in the handler so
 * this stays a pure transform of Graph's response shape.
 */
export interface ExpandedMessage extends ReturnType<typeof summarizeMessage> {
  body_content_type: string;
  body: string;
  reply_to: Array<{ name: string; email: string }>;
  internet_message_id: string;
}

export function expandMessage(m: Record<string, unknown>): ExpandedMessage {
  const summary = summarizeMessage(m);
  const body = m.body as Record<string, unknown> | undefined;
  const bodyContent = String(body?.content ?? "");
  const replyTo = Array.isArray(m.replyTo) ? m.replyTo : [];
  return {
    ...summary,
    body_content_type: String(body?.contentType ?? "text"),
    body: bodyContent,
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

/**
 * Apply caller-driven pagination to an already-expanded message.
 *
 * Contract:
 *   - `maxBodyChars === undefined` → return the full body starting at
 *     `bodyOffset`. `next_offset` is `null`. `body_truncated` is `false`.
 *   - `maxBodyChars !== undefined` → slice
 *     `body[bodyOffset : bodyOffset + maxBodyChars]`. If that slice
 *     reaches the end, `next_offset` is `null`; otherwise it points at
 *     the next offset to fetch.
 *   - `bodyOffset` past the end returns an empty body with
 *     `next_offset: null` (terminating cleanly rather than throwing).
 */
export function applyBodyPagination(
  expanded: ExpandedMessage,
  bodyOffset: number,
  maxBodyChars: number | undefined,
): ExpandedMessage & {
  body_offset: number;
  body_char_count: number;
  total_body_char_count: number;
  next_offset: number | null;
  body_truncated: boolean;
} {
  const fullBody = expanded.body;
  const total = fullBody.length;
  const startClamped = Math.min(bodyOffset, total);
  const endExclusive =
    maxBodyChars === undefined
      ? total
      : Math.min(startClamped + maxBodyChars, total);
  const slice = fullBody.slice(startClamped, endExclusive);
  const nextOffset = endExclusive < total ? endExclusive : null;

  return {
    ...expanded,
    body: slice,
    body_offset: startClamped,
    body_char_count: slice.length,
    total_body_char_count: total,
    next_offset: nextOffset,
    // Kept for readability. Strictly equivalent to `next_offset !== null`.
    // Never true for a default (paramless) call — no silent truncation.
    body_truncated: nextOffset !== null,
  };
}

const SELECT_FIELDS =
  "id,conversationId,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,isRead,isDraft,hasAttachments,importance,bodyPreview,webLink,parentFolderId,body,replyTo,internetMessageId";

const handler: ToolHandler = async (
  graph: Client,
  args: Record<string, unknown>,
): Promise<ToolResponse> => {
  const messageId = validateRequiredString(args.message_id, "message_id");
  const bodyOffset = validateOptionalInteger(args.body_offset, "body_offset", {
    min: 0,
    max: MAX_OFFSET,
    default: 0,
  });
  const maxBodyChars = validateOptionalIntegerOrUndefined(
    args.max_body_chars,
    "max_body_chars",
    { min: 1, max: MAX_PAGE_CHARS },
  );

  const message = await graph
    .api(`/me/messages/${encodeURIComponent(messageId)}`)
    .select(SELECT_FIELDS)
    .get();

  const expanded = expandMessage(message);
  const result = applyBodyPagination(expanded, bodyOffset, maxBodyChars);

  return {
    content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
  };
};

export const getMessageTool: Tool = { category: "read", definition, handler };
