# Architecture — M365 Mail MCP Server

Design rationale for `@juvantlabs/m365-mail-mcp-server`. Read alongside
the
[handbook MCP server spec](https://github.com/juvantlabs/handbook/blob/main/docs/repo-types/mcp-server.md)
for the cross-cutting conventions; this doc covers what's specific to
this server.

## Purpose

Microsoft 365 Outlook Mail via delegated Microsoft Graph:

- read mail folders, messages, and attachments;
- author, edit, reply-to, and forward drafts (does **not** send);
- move / mark-read / delete (two-phase) messages.

## Scope

### In scope for v0.1

- Own mailbox reads: folders, messages, attachments (metadata +
  bytes).
- Own mailbox writes: draft creation / update / reply / forward,
  mark-read toggle, folder moves.
- Own mailbox delete: soft delete via `DELETE /me/messages/{id}` (→
  Deleted Items). Two-phase confirmation-token gate.

### Out of scope for v0.1 (deliberate)

- **Sending mail** (`Mail.Send`). Deferred to v0.3, gated by ADR 0001
  (see `docs/adr/0001-v0-3-send-gate-contract.md`). Neither the scope
  nor any send tool is registered here; a CI invariant fails the
  build if a send path leaks in without a consumed confirmation-token.
- **Shared / delegate mailboxes** (`Mail.Read.Shared`,
  `Mail.ReadWrite.Shared`). Deferred to v0.2. The v0.1 tool schema
  intentionally has NO `shared_user` parameter — this is enforced by
  `tests/unit/registry.test.ts`.
- **Draft attachment upload**. Requires the Graph upload-session flow
  (large attachments) and its own security review of local-file
  provenance. Deferred to v0.3+.
- **OneDrive / SharePoint / Calendar / Meeting transcripts**. Handled
  by the sibling `@juvantlabs/m365-graph-mcp-server`. Cross-server
  scope leakage would be a security regression; see § Isolation from
  sibling server.

## Authentication

- **Flow**: OAuth 2.0 Authorization Code with a Confidential Client
  Application (`@azure/msal-node`). The client secret is required.
- **Setup**: interactive one-time flow via `m365-mail-mcp-server setup`
  → opens the browser at Microsoft's consent screen → local
  `http://localhost:3000/auth/callback` catches the code → MSAL
  exchanges for tokens → tokens persist in the OS keychain.
- **Runtime**: subsequent server starts read the cached refresh token
  from the keychain and refresh access tokens silently. MSAL handles
  refresh, revocation detection, and expiry edges.
- **Storage**: `@napi-rs/keyring` (not the archived `keytar`).
  Service name is `juvantlabs-m365-mail-mcp-server`, account keyed by
  tenant ID. Distinct from the m365-graph server's service name so
  cached tokens live in separate keychain entries.

### Delegated scopes

```
User.Read
Mail.Read
Mail.ReadWrite
offline_access
```

Deliberately narrow. `Mail.ReadWrite` subsumes `Mail.Read` at the
scope level, but Microsoft's admin consent screen renders both when
requested, so we list only the ones actually needed. `offline_access`
is required to receive a refresh token.

### Isolation from sibling server

The two servers must not silently share credentials or token caches.
Independent isolation layers:

| Boundary | m365-mail | m365-graph |
|---|---|---|
| Env-var prefix | `M365_MAIL_*` | `M365_*` |
| Keychain service name | `juvantlabs-m365-mail-mcp-server` | `juvantlabs-m365-graph-mcp-server` |
| Download-sandbox subdir | `m365-mail-mcp-server/<tenant>/` | `m365-graph-mcp-server/<tenant>/` |
| Delegated scopes requested | Only `Mail.*` (+ `User.Read`, `offline_access`) | Only `Files.*` / `Sites.*` / `Calendars.*` / meeting scopes |

The env-var namespacing is the load-bearing guard: a config-file typo
that reaches the wrong server is caught by startup validation
(`checkEnv`), which requires the `M365_MAIL_*` names specifically.

## Threat model

### Universal Boundaries (per `SYSTEM_INVARIANTS.md` §4)

- No general-purpose URL forwarder primitive — this server speaks
  only Microsoft Graph endpoints reachable via
  `@microsoft/microsoft-graph-client`.
- No write tools beyond what's explicitly authorized. `Mail.Send`
  is deliberately absent (see ADR 0001).
- Per-tenant subprocess — no shared cache state across tenants; the
  confirmation-token store lives in a per-process in-memory Map that
  dies with the subprocess.
- Stdout discipline: `console.error` only outside the JSON-RPC
  protocol path. Enforced by ESLint (`no-console` with `allow:
  [error, warn]`) and a CI grep step as defense-in-depth.

### Shield conditions (v0.1)

Per the 2026-07-02 Shield review of the Azure app PR:

- **C2 — deniable mutating tools**. Each mutating tool has its own
  namespaced identity (`mcp__m365-mail__delete_message`,
  `mcp__m365-mail__move_message`, `mcp__m365-mail__mark_read`, etc.)
  so the harness pre-tool-use hook can deny individually.

- **C3 — inbox-as-untrusted-data**. Message content received via
  Graph (bodies, subjects, headers) is treated as untrusted.
  Tool-call parameters (recipient addresses, draft bodies) MUST be
  caller-authored, never lifted verbatim from inbound content. This
  is a policy boundary documented in tool descriptions; the server
  does not attempt to detect prompt injection in message bodies.

- **C4 — agent-draft marker**. Every draft this server creates
  carries two independent markers:
  1. Subject prefix `[agent-draft] ` — universal, idempotent,
     applied by `create_draft`, `update_draft` (on subject patch),
     `create_reply_draft`, `create_forward_draft`. See
     `src/tools/_shared.ts::ensureAgentDraftSubject`.
  2. Custom header `X-Juvant-Agent-Author:
     @juvantlabs/m365-mail-mcp-server` — attached to the initial
     `POST /me/messages` in `create_draft`. Graph forbids setting
     `internetMessageHeaders` on `PATCH`, so reply/forward flows
     rely on the subject prefix alone.

- **C5 — port-3000 op note**. The OAuth callback listens on port
  3000, shared with `m365-graph-mcp-server`. Only one interactive
  setup can run at a time. Documented in `README.md § Local
  development`.

### Anti-patterns actively defended against

| # | Anti-pattern | Defense |
|---|---|---|
| S1 | Dead security helpers (`validate*`, `sanitize*`, `guard*`) that CI can't detect | CI grep in `.github/workflows/ci.yml` fails the build if any exported helper is unused in `src/` |
| S2 | Env-var documented but not wired | CI parses `README.md § Environment variables` and requires each name to appear as `process.env.<NAME>` in `src/` |
| S3 | Silent `console.log` in the stdio path | ESLint `no-console` + CI grep + strict lint job |
| S4 | Path traversal via caller-supplied filenames | `deriveSafeLocalPath` constructs filenames from a SHA-256 hash + `sanitizeFilename(originalName)`; result is verified to start with the sandbox root |
| S5 | `keytar` (archived since 2022) | `@napi-rs/keyring` |
| S6 | Silent mutation of a sent message | `update_draft` refuses PATCH if `isDraft !== true` |

## Performance characteristics

- **Attachment download**: streamed via `pipeline` — no whole-file
  buffering. Metadata GET runs BEFORE the byte fetch; files >200 MB
  are rejected pre-flight. Cap is defense against runaway memory /
  disk on adversarial or misconfigured inputs.
- **Message body**: capped at 16 000 characters in `get_message`
  (`body_truncated: true` when hit). Outlook bodies with quoted
  history + HTML signatures + tracking pixels can bloat well past
  what an agent needs to reason about; the cap keeps context budgets
  predictable.
- **Search**: uses Graph `$search` with the required
  `ConsistencyLevel: eventual` header. Graph forbids
  `$search`+`$orderby`, so results come back relevance-ordered.

## Tool catalog

### Read tools

| Tool | Underlying API call | Key input | Output shape | Notes |
|---|---|---|---|---|
| `m365-mail:list_mail_folders` | `GET /me/mailFolders` | `limit?` | `{count, folders[]}` | Top-level only in v0.1 |
| `m365-mail:list_messages` | `GET /me/mailFolders/{f}/messages` \| `/me/messages` | `folder_id?`, `limit?` | `{folder_id, count, messages[]}` | Ordered `receivedDateTime desc` |
| `m365-mail:search_messages` | `GET /me/messages?$search=…` | `query`, `folder_id?`, `limit?` | `{query, folder_id, count, messages[]}` | KQL; `ConsistencyLevel: eventual` |
| `m365-mail:get_message` | `GET /me/messages/{id}` | `message_id` | Full message + capped body | Body cap 16 000 chars |
| `m365-mail:list_attachments` | `GET /me/messages/{id}/attachments` | `message_id`, `limit?` | `{message_id, count, attachments[]}` | Metadata only (no `contentBytes`) |
| `m365-mail:download_attachment` | `GET /me/messages/{id}/attachments/{aid}/$value` | `message_id`, `attachment_id` | `{local_path, size_bytes, name, …}` | Sandbox pattern; rejects itemAttachment / referenceAttachment; 200 MB cap |

### Write — idempotent tools

| Tool | Underlying API call | Key input | Output shape | Notes |
|---|---|---|---|---|
| `m365-mail:create_draft` | `POST /me/messages` | `subject?`, `body?`, `to?`, `cc?`, `bcc?`, `importance?` | `{created, agent_draft_markers, note}` | Marks subject + attaches `X-Juvant-Agent-Author` header |
| `m365-mail:update_draft` | `PATCH /me/messages/{id}` | `message_id`, + any of `create_draft`'s fields | `{updated}` | Refuses PATCH if `isDraft !== true`; re-marks subject idempotently |
| `m365-mail:create_reply_draft` | `POST /me/messages/{id}/createReply\|createReplyAll` + `PATCH` | `message_id`, `reply_all?`, + draft fields | `{created, parent_message_id, reply_all, agent_draft_markers}` | Unconditional subject PATCH ensures marker survives Graph's `RE:` prefill |
| `m365-mail:create_forward_draft` | `POST /me/messages/{id}/createForward` + `PATCH` | `message_id`, + draft fields | `{created, parent_message_id, agent_draft_markers}` | Same subject-PATCH pattern as reply |
| `m365-mail:mark_read` | `PATCH /me/messages/{id}` `{isRead}` | `message_id`, `is_read?` (default `true`) | `{message_id, is_read}` | Idempotent by construction |
| `m365-mail:move_message` | `POST /me/messages/{id}/move` | `message_id`, `destination_folder_id` | `{moved, source_message_id, destination_folder_id, note}` | Graph reissues the message ID on move; response returns the new id |

### Write — irreversible (two-phase) tool

| Tool | Underlying API call | Key input | Output shape | Notes |
|---|---|---|---|---|
| `m365-mail:delete_message` | `DELETE /me/messages/{id}` | `message_id`, `confirmation_token?` | Preview (phase 1) or `{deleted, note}` (phase 2) | Two-phase; token single-use, 5-min TTL, canonical-spec hash |

## Scope map

The following table is the ground truth for which delegated Graph
scope each tool depends on. CI Layer A (permission-mutation surface
invariant) fails the build if a permission-mutation-class scope is
introduced without a `permission_mutating`-classified tool.

| Tool | Scope | Category |
|---|---|---|
| `list_mail_folders` | `Mail.Read` | `read` |
| `list_messages` | `Mail.Read` | `read` |
| `search_messages` | `Mail.Read` | `read` |
| `get_message` | `Mail.Read` | `read` |
| `list_attachments` | `Mail.Read` | `read` |
| `download_attachment` | `Mail.Read` | `read` |
| `create_draft` | `Mail.ReadWrite` | `write_idempotent` |
| `update_draft` | `Mail.ReadWrite` | `write_idempotent` |
| `create_reply_draft` | `Mail.ReadWrite` | `write_idempotent` |
| `create_forward_draft` | `Mail.ReadWrite` | `write_idempotent` |
| `mark_read` | `Mail.ReadWrite` | `write_idempotent` |
| `move_message` | `Mail.ReadWrite` | `write_idempotent` |
| `delete_message` | `Mail.ReadWrite` | `write_irreversible` |

Zero tools carry `permission_mutating` in v0.1 — this is the CI Layer
A baseline (see `tests/unit/permission_surface.test.ts`).

## Dependencies

| Dependency | Version | Why |
|---|---|---|
| `@modelcontextprotocol/sdk` | `^1.25.2` | MCP framing; ≥ 1.25.2 required (ReDoS + DNS rebinding fixes) |
| `@azure/msal-node` | `^5.0.0` | OAuth Confidential Client + token cache. Microsoft's official library — never roll your own |
| `@microsoft/microsoft-graph-client` | `^3.0.7` | Graph SDK. Uses `initWithMiddleware` + a custom `AuthenticationProvider` bridging MSAL |
| `@napi-rs/keyring` | `^1.3.0` | Cross-platform OS keychain; replacement for archived `keytar` |
| `isomorphic-fetch` | `^3.0.0` | Peer dep of the Graph client. Imported once at module load |

## References

- [Handbook MCP server spec](https://github.com/juvantlabs/handbook/blob/main/docs/repo-types/mcp-server.md)
- [Handbook MCP abstract roles ADR 0002](https://github.com/juvantlabs/handbook/blob/main/docs/adr/0002-mcp-abstract-roles.md)
- [Juvant OS MCP_INVENTORY.md](https://github.com/juvantlabs/juvant-os/blob/main/docs/MCP_INVENTORY.md)
- [Handbook ADR 0004 — agent action guardrails](https://github.com/juvantlabs/handbook/blob/main/docs/adr/0004-agent-action-guardrails.md)
- [Sibling reference: `juvantlabs/m365-graph-mcp-server`](https://github.com/juvantlabs/m365-graph-mcp-server)
- [In-repo `docs/adr/0001-v0-3-send-gate-contract.md`](docs/adr/0001-v0-3-send-gate-contract.md)
