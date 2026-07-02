/**
 * Tool: m365-mail:download_attachment
 *
 * Download a message attachment's bytes to a per-tenant local sandbox
 * directory, returning the local path. Mirrors the download_file
 * pattern from the sibling m365-graph server; agents read the file
 * via a filesystem-aware tool.
 *
 * Why path return, not content:
 *   - Attachments can be large (spec cap: 200 MB).
 *   - MCP responses are JSON-typed; binary via base64 inflates by
 *     33% and stresses the agent's context.
 *   - Composability: M365-aware fetch + provider-agnostic read are
 *     two concerns that cleanly separate.
 *
 * Attachment kind handling:
 *   - fileAttachment       → downloadable. Bytes fetched via
 *                            `/attachments/{aid}/$value`.
 *   - itemAttachment       → REJECTED with a structured error field.
 *                            These are embedded Outlook items, not
 *                            files with content; agents should fetch
 *                            via message endpoints instead.
 *   - referenceAttachment  → REJECTED with a structured error field.
 *                            These are links to cloud files; the
 *                            correct fetch path lives in a different
 *                            MCP server (m365-graph for OneDrive
 *                            items). This server does not chase links.
 *
 * Sandboxing (handbook spec § Sandboxing + anti-pattern #1):
 *   - All downloads land under <sandbox-root>/<tenant-id>/ with a
 *     server-controlled filename:
 *       <sha256(message_id||"::"||attachment_id)[:16]>-<sanitized name>
 *   - sandbox-root is M365_MAIL_DOWNLOAD_DIR (env), or
 *     $XDG_CACHE_HOME/m365-mail-mcp-server, or
 *     ~/.cache/m365-mail-mcp-server. Distinct from the m365-graph
 *     server's cache by directory name, so cross-server confusion is
 *     structurally prevented.
 *   - 0o700 dir mode + 0o600 file mode.
 *   - Resolved path verified to start with the sandbox root.
 *
 * Size cap: 200 MB. The metadata GET runs BEFORE the byte fetch so
 * the tool refuses large downloads without streaming them.
 *
 * Required Graph scope: `Mail.Read` (delegated). Read-only on the
 * cloud side; writes only to the sandbox locally.
 */

import { createWriteStream } from "node:fs";
import { mkdir } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import crypto from "node:crypto";
import { Readable } from "node:stream";
import { pipeline } from "node:stream/promises";

import type { Client } from "@microsoft/microsoft-graph-client";

import {
  SHARED_USER_SCHEMA_PROPERTY,
  mailboxRoot,
  validateSharedUser,
} from "./_mailbox.js";
import {
  sanitizeFilename,
  validateRequiredString,
} from "../types/validators.js";
import type { Tool, ToolDefinition, ToolHandler, ToolResponse } from "../types/tool.js";

const MAX_FILE_SIZE_BYTES = 200 * 1024 * 1024;
const SANDBOX_DIR_MODE = 0o700;
const SANDBOX_FILE_MODE = 0o600;

const definition: ToolDefinition = {
  name: "m365-mail:download_attachment",
  description:
    "Download a fileAttachment's bytes to a local sandbox directory. Returns the local path, not the file content — agents read the file via a filesystem-aware tool. Size capped at 200 MB. Rejects itemAttachment / referenceAttachment kinds with an explicit error. Read-only on the cloud side. Pass `shared_user` to download an attachment from a message in a shared / delegate mailbox (v0.2, requires Mail.Read.Shared).",
  inputSchema: {
    type: "object",
    properties: {
      message_id: {
        type: "string",
        description: "Parent message id.",
      },
      attachment_id: {
        type: "string",
        description: "Attachment id from list_attachments.",
      },
      ...SHARED_USER_SCHEMA_PROPERTY,
    },
    required: ["message_id", "attachment_id"],
  },
};

export function getSandboxRoot(tenantId: string): string {
  const override = process.env.M365_MAIL_DOWNLOAD_DIR;
  if (override) return path.resolve(override, tenantId);

  const xdg = process.env.XDG_CACHE_HOME;
  if (xdg) return path.resolve(xdg, "m365-mail-mcp-server", tenantId);

  return path.resolve(os.homedir(), ".cache", "m365-mail-mcp-server", tenantId);
}

/**
 * Derive the safe local path for a downloaded attachment.
 *
 * v0.2 change: `sharedUser` is folded into the hash key so a download
 * of message A/attachment X from the caller's own mailbox and the same
 * ids from a shared mailbox produce distinct local paths. Graph message
 * ids are unique per mailbox in practice, but hashing the mailbox
 * routing key ("me" or the UPN) makes the deterministic path collision-
 * proof even in the pathological case.
 *
 * Backward-compatibility note: the hash prefix format is preserved
 * (still `<hex16>-<sanitized-name>`), so callers can still identify
 * downloaded files by name shape. The bytes of the hash change from
 * v0.1 (no mailbox routing key was included) — this is not a semver
 * break because the local sandbox is treated as an implementation
 * detail (each subprocess writes to it; agents read via the returned
 * `local_path`; no caller pins the exact hash).
 */
export function deriveSafeLocalPath(
  sandboxRoot: string,
  messageId: string,
  attachmentId: string,
  originalName: string,
  sharedUser?: string,
): string {
  const mailboxKey = sharedUser ?? "me";
  const key = `${mailboxKey}::${messageId}::${attachmentId}`;
  const hash = crypto.createHash("sha256").update(key).digest("hex").slice(0, 16);
  const safe = sanitizeFilename(originalName);
  const filename = `${hash}-${safe}`;
  const resolved = path.resolve(sandboxRoot, filename);

  // Defense-in-depth even though the filename is server-constructed.
  if (!resolved.startsWith(sandboxRoot + path.sep) && resolved !== sandboxRoot) {
    throw new Error(`Refusing to write outside sandbox: ${resolved}`);
  }
  return resolved;
}

/**
 * Normalize the `@odata.type` field into a short kind name
 * (`fileAttachment`, `itemAttachment`, `referenceAttachment`,
 * `unknown`). Exported for unit tests.
 */
export function attachmentKind(meta: Record<string, unknown>): string {
  const raw = String(meta["@odata.type"] ?? "");
  const short = raw.split(".").pop();
  return short && short.length > 0 ? short : "unknown";
}

const handler: ToolHandler = async (
  graph: Client,
  args: Record<string, unknown>,
): Promise<ToolResponse> => {
  const messageId = validateRequiredString(args.message_id, "message_id");
  const attachmentId = validateRequiredString(args.attachment_id, "attachment_id");
  const sharedUser = validateSharedUser(args.shared_user);
  const root = mailboxRoot(sharedUser);

  const metaPath =
    `${root}/messages/${encodeURIComponent(messageId)}` +
    `/attachments/${encodeURIComponent(attachmentId)}`;

  const meta = (await graph
    .api(metaPath)
    .select("id,name,size,contentType,isInline")
    .get()) as Record<string, unknown>;

  const kind = attachmentKind(meta);
  if (kind !== "fileAttachment") {
    // Structured error: throw so the dispatcher wraps as isError.
    // Include the kind so the agent can act on it.
    throw new Error(
      `attachment kind '${kind}' is not downloadable via this tool ` +
        `(only fileAttachment is supported). ` +
        `itemAttachment is an embedded Outlook item; ` +
        `referenceAttachment is a cloud-file link (use m365-graph-mcp-server for those).`,
    );
  }

  const size = Number(meta.size ?? 0);
  if (size > MAX_FILE_SIZE_BYTES) {
    throw new Error(
      `Attachment size (${size} bytes) exceeds the 200 MB cap. Refusing to download.`,
    );
  }

  const tenantId = process.env.M365_MAIL_TENANT_ID ?? "unknown";
  const sandboxRoot = getSandboxRoot(tenantId);
  await mkdir(sandboxRoot, { recursive: true, mode: SANDBOX_DIR_MODE });

  const localPath = deriveSafeLocalPath(
    sandboxRoot,
    messageId,
    attachmentId,
    String(meta.name ?? "attachment"),
    sharedUser,
  );

  // Stream bytes from `/$value`. Graph client returns a Node Readable
  // for raw endpoints; we pipeline into the sandbox file.
  const stream = (await graph
    .api(`${metaPath}/$value`)
    .getStream()) as unknown as NodeJS.ReadableStream;

  const out = createWriteStream(localPath, { mode: SANDBOX_FILE_MODE });
  await pipeline(Readable.from(stream), out);

  const result = {
    message_id: messageId,
    attachment_id: attachmentId,
    shared_user: sharedUser ?? null,
    local_path: localPath,
    size_bytes: size,
    name: String(meta.name ?? ""),
    content_type: String(meta.contentType ?? ""),
    attachment_type: kind,
  };

  return {
    content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
  };
};

export const downloadAttachmentTool: Tool = { category: "read", definition, handler };
