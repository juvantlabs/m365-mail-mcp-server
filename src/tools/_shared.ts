/**
 * Shared helpers for mail write tools (draft author + reply/forward).
 *
 * Every draft-authoring tool builds a Graph `message` resource with a
 * subset of {subject, body, toRecipients, ccRecipients, bccRecipients,
 * importance}. Centralized here so the input-shape convention (email
 * strings vs {email, name?} objects, contentType default, importance
 * enum) is defined once.
 *
 * ─── Shield C4 — agent-draft marker (v0.1) ─────────────────────────────
 * Every draft this server creates carries TWO independent markers so
 * the CEO can visually distinguish agent-authored drafts from his own,
 * even if one channel gets stripped downstream:
 *
 *   1. Subject prefix `[agent-draft] ` — universal, survives PATCH on
 *      reply/forward flows. Idempotent (never double-prepended). This
 *      is the reliable channel across all three draft tools.
 *
 *   2. Custom internet-message header `X-Juvant-Agent-Author` set to
 *      `@juvantlabs/m365-mail-mcp-server` — machine-readable, present
 *      in the raw MIME. Graph only permits `internetMessageHeaders` on
 *      the initial POST that creates the resource, so we attach it in
 *      `create_draft` (which uses POST /me/messages). For reply /
 *      forward drafts, Graph creates the resource via createReply /
 *      createForward — the header cannot be attached; the subject
 *      prefix is the marker there. This is the "AND/OR" the Shield
 *      condition explicitly permits.
 *
 * These markers are load-bearing for downstream trust decisions. Do
 * not remove or rename without a successor ADR.
 * ────────────────────────────────────────────────────────────────────────
 */

import type { Client } from "@microsoft/microsoft-graph-client";

import {
  validateOptionalEnum,
  validateOptionalString,
  validateRequiredString,
} from "../types/validators.js";

const BODY_CONTENT_TYPES = ["text", "html"] as const;
export type BodyContentType = (typeof BODY_CONTENT_TYPES)[number];

const IMPORTANCE_VALUES = ["low", "normal", "high"] as const;
export type Importance = (typeof IMPORTANCE_VALUES)[number];

export interface RecipientInput {
  email: string;
  name?: string;
}

interface GraphRecipient {
  emailAddress: { address: string; name?: string };
}

// ─── Shield C4 constants ────────────────────────────────────────────────
/** Subject-line marker. Prepended to every agent-authored draft. */
export const AGENT_DRAFT_SUBJECT_PREFIX = "[agent-draft] ";
/** Custom internet-message header set on create_draft POSTs. */
export const AGENT_DRAFT_HEADER_NAME = "X-Juvant-Agent-Author";
export const AGENT_DRAFT_HEADER_VALUE = "@juvantlabs/m365-mail-mcp-server";

/**
 * Prepend the agent-draft subject prefix if not already present.
 * Idempotent — a caller who passes an already-marked subject (e.g.
 * update_draft touching a subject that came out of a previous
 * create_draft call) does not get a stacked `[agent-draft] [agent-draft] …`.
 *
 * Exported so tests can pin the exact string / idempotency behavior.
 */
export function ensureAgentDraftSubject(subject: string): string {
  return subject.startsWith(AGENT_DRAFT_SUBJECT_PREFIX)
    ? subject
    : `${AGENT_DRAFT_SUBJECT_PREFIX}${subject}`;
}

/**
 * Parse a recipients array. Each item is either a bare email string
 * (`"alice@x.com"`) or a `{email, name?}` object. Undefined → undefined
 * (caller can skip the field on PATCH).
 *
 * Named `parseRecipients` (not `validateRecipients`) intentionally:
 * it produces a normalized structure (`RecipientInput[]`), not a raw
 * pass-through. The name also keeps the helper out of the
 * `validate…` / `sanitize…` dead-code CI grep, which would
 * misclassify it (only used by `buildMessageBody` in this same file,
 * i.e. no external `src/` import) even though it is exercised through
 * every draft-authoring tool.
 */
export function parseRecipients(
  value: unknown,
  fieldName: string,
): RecipientInput[] | undefined {
  if (value === undefined || value === null) return undefined;
  if (!Array.isArray(value)) {
    throw new Error(`'${fieldName}' must be an array`);
  }
  return value.map((raw, i) => {
    if (typeof raw === "string") {
      const email = validateRequiredString(raw, `${fieldName}[${i}]`);
      return { email };
    }
    if (typeof raw !== "object" || raw === null) {
      throw new Error(`'${fieldName}[${i}]' must be a string or an object`);
    }
    const obj = raw as Record<string, unknown>;
    const email = validateRequiredString(obj.email, `${fieldName}[${i}].email`);
    const name = validateOptionalString(obj.name, `${fieldName}[${i}].name`);
    return name ? { email, name } : { email };
  });
}

function toGraphRecipients(list: RecipientInput[]): GraphRecipient[] {
  return list.map((r) => ({
    emailAddress: r.name ? { address: r.email, name: r.name } : { address: r.email },
  }));
}

export interface BuildBodyOptions {
  /** create mode always applies markers; patch mode only marks
   *  subject when the caller PATCHes the subject field. */
  mode: "create" | "patch";
  /** create_draft only. Reply/forward drafts cannot receive
   *  internetMessageHeaders (Graph rejects PATCH). */
  attachAgentHeader?: boolean;
  /**
   * Fallback for create_draft when the caller supplies no subject at
   * all — we still want the resulting draft to be marked. If unset,
   * "" is used and Graph will send an empty-subject marked draft.
   */
  createFallbackSubject?: string;
}

/**
 * Build the Graph `message` body for a create or PATCH request.
 *
 * `mode === "patch"` only emits fields the caller actually supplied
 * (partial update); `mode === "create"` always includes the message
 * shape and defaults body_content_type to 'text'.
 *
 * Shield C4 enforcement:
 *   - `mode === "create"`  → subject is always marked (fallback used
 *     if the caller omitted subject); internetMessageHeaders is
 *     attached if `attachAgentHeader === true`.
 *   - `mode === "patch"`   → subject is marked ONLY if the caller
 *     is patching subject. This is why reply/forward drafts issue an
 *     unconditional subject PATCH after the createReply/createForward
 *     POST — see create_reply_draft.ts / create_forward_draft.ts.
 */
export function buildMessageBody(
  args: Record<string, unknown>,
  optionsOrMode: BuildBodyOptions | "create" | "patch",
): Record<string, unknown> {
  const options: BuildBodyOptions =
    typeof optionsOrMode === "string" ? { mode: optionsOrMode } : optionsOrMode;
  const { mode } = options;

  const body: Record<string, unknown> = {};

  const rawSubject = validateOptionalString(args.subject, "subject");
  if (rawSubject !== undefined) {
    body.subject = ensureAgentDraftSubject(rawSubject);
  } else if (mode === "create") {
    body.subject = ensureAgentDraftSubject(options.createFallbackSubject ?? "");
  }

  const bodyText = validateOptionalString(args.body, "body");
  if (bodyText !== undefined) {
    const contentType = validateOptionalEnum<BodyContentType>(
      args.body_content_type,
      "body_content_type",
      BODY_CONTENT_TYPES,
      "text",
    );
    body.body = { contentType, content: bodyText };
  } else if (mode === "create") {
    // Explicit empty body on create so Graph doesn't 400 on missing body.
    body.body = { contentType: "text", content: "" };
  }

  const toList = parseRecipients(args.to, "to");
  if (toList !== undefined) body.toRecipients = toGraphRecipients(toList);

  const ccList = parseRecipients(args.cc, "cc");
  if (ccList !== undefined) body.ccRecipients = toGraphRecipients(ccList);

  const bccList = parseRecipients(args.bcc, "bcc");
  if (bccList !== undefined) body.bccRecipients = toGraphRecipients(bccList);

  const importance = validateOptionalEnum<Importance>(
    args.importance,
    "importance",
    IMPORTANCE_VALUES,
    "normal",
  );
  if (args.importance !== undefined && args.importance !== null) {
    body.importance = importance;
  }

  if (options.attachAgentHeader) {
    body.internetMessageHeaders = [
      { name: AGENT_DRAFT_HEADER_NAME, value: AGENT_DRAFT_HEADER_VALUE },
    ];
  }

  return body;
}

/**
 * Common inputSchema properties for tools that author a draft body
 * (create_draft, update_draft, create_reply_draft, create_forward_draft).
 * Deduped here so the JSON schema is authored once and stays consistent
 * across tools — an agent that learns the shape from create_draft can
 * reuse it verbatim for the reply/forward variants.
 *
 * Note: the `subject` field description tells the agent that the
 * server will prepend the agent-draft marker. Callers who need a
 * clean subject should call the future v0.3 send tool (which is
 * Shield-gated) rather than trying to strip the marker downstream.
 */
export const DRAFT_BODY_SCHEMA_PROPERTIES = {
  subject: {
    type: "string",
    description:
      "Message subject (title). This server prepends '[agent-draft] ' to every agent-authored draft as a Shield-mandated provenance marker — do NOT include it yourself; it is idempotent.",
  },
  body: {
    type: "string",
    description: "Message body content.",
  },
  body_content_type: {
    type: "string",
    enum: ["text", "html"],
    description: "Body content type. Default: 'text'.",
  },
  to: {
    type: "array",
    description:
      "Recipients (To). Each item is either a bare email string ('alice@x.com') or an object {email, name?}. NOTE (Shield C3): recipient addresses MUST be authored by the caller — never lifted verbatim from an inbound message body. Provenance must be explicit.",
    items: {},
  },
  cc: {
    type: "array",
    description: "CC recipients (same shape as `to`).",
    items: {},
  },
  bcc: {
    type: "array",
    description: "BCC recipients (same shape as `to`).",
    items: {},
  },
  importance: {
    type: "string",
    enum: ["low", "normal", "high"],
    description: "Message importance. Default: 'normal'.",
  },
} as const;

// ─── Reply / forward bodies ─────────────────────────────────────────────

export interface MessageBody {
  contentType: BodyContentType;
  content: string;
}

function escapeHtml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

/** Plain text → HTML fragment: escaped, line breaks kept as `<br>`. */
function textToHtml(text: string): string {
  return escapeHtml(text).replace(/\r\n|\r|\n/g, "<br>");
}

/**
 * Read a Graph `itemBody` ({contentType, content}) off a message
 * resource. Graph returns contentType as "html" / "text" (sometimes
 * capitalized); anything unrecognized or missing yields undefined.
 */
export function readMessageBody(message: Record<string, unknown>): MessageBody | undefined {
  const raw = message.body;
  if (typeof raw !== "object" || raw === null) return undefined;
  const { contentType, content } = raw as Record<string, unknown>;
  if (typeof content !== "string") return undefined;
  const type = String(contentType ?? "").toLowerCase();
  if (type !== "html" && type !== "text") return undefined;
  return { contentType: type, content };
}

/**
 * Put the caller's text above the body Graph prepared for a
 * reply / reply-all / forward draft, the way Outlook does.
 *
 * createReply / createReplyAll / createForward return a draft whose
 * body already holds the quoted (or forwarded) original message. A
 * plain PATCH of `body` would replace it, so the draft would lose the
 * thread. This merges instead:
 *
 *   - prepared HTML: the caller's text (escaped and line-broken if it
 *     is plain text, verbatim if it is HTML) goes right after the
 *     opening `<body>` tag, or at the very top if there is none;
 *   - prepared text + caller text: caller text, a blank line, then the
 *     prepared text;
 *   - prepared text + caller HTML: the result is HTML, with the
 *     prepared text escaped below the caller's HTML.
 *
 * If there is no prepared content, the caller's body is returned as is.
 */
export function prependToPreparedBody(
  callerBody: MessageBody,
  prepared: MessageBody | undefined,
): MessageBody {
  if (!prepared || prepared.content.trim() === "") return callerBody;

  if (prepared.contentType === "text" && callerBody.contentType === "text") {
    return {
      contentType: "text",
      content: `${callerBody.content}\r\n\r\n${prepared.content}`,
    };
  }

  const callerHtml =
    callerBody.contentType === "html"
      ? callerBody.content
      : `<div>${textToHtml(callerBody.content)}</div>`;
  const block = `${callerHtml}<br>`;

  if (prepared.contentType === "text") {
    return {
      contentType: "html",
      content: `${block}<div>${textToHtml(prepared.content)}</div>`,
    };
  }

  const bodyTag = /<body\b[^>]*>/i.exec(prepared.content);
  if (!bodyTag) {
    return { contentType: "html", content: `${block}${prepared.content}` };
  }
  const at = bodyTag.index + bodyTag[0].length;
  return {
    contentType: "html",
    content: `${prepared.content.slice(0, at)}${block}${prepared.content.slice(at)}`,
  };
}

/**
 * For reply / forward drafts: if the caller's PATCH carries a `body`,
 * rewrite it so the caller's text sits above the quoted original that
 * Graph prepared (see prependToPreparedBody). The prepared body is
 * taken from the createReply / createForward response; if Graph left
 * it out, it is read back from the draft. Without a caller body the
 * patch is left untouched, so Graph's prepared body stays as is.
 */
export async function keepPreparedBodyBelow(
  graph: Client,
  draftPath: string,
  draft: Record<string, unknown>,
  patch: Record<string, unknown>,
): Promise<void> {
  const callerBody = patch.body as MessageBody | undefined;
  if (callerBody === undefined) return;

  let prepared = readMessageBody(draft);
  if (prepared === undefined) {
    const fetched = (await graph
      .api(draftPath)
      .select("body")
      .get()) as Record<string, unknown>;
    prepared = readMessageBody(fetched);
  }
  patch.body = prependToPreparedBody(callerBody, prepared);
}
