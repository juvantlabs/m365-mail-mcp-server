# M365 Mail MCP Server

`@juvantlabs/m365-mail-mcp-server` — Model Context Protocol server wrapping
**Microsoft 365 Outlook Mail** via delegated Microsoft Graph:

- read mail folders, messages, and attachments;
- author, edit, reply-to, and forward drafts (**does NOT send**);
- move / mark-read / delete (two-phase confirmation) messages.

Send is a later, Shield-gated phase — see
[`docs/adr/0001-v0-3-send-gate-contract.md`](docs/adr/0001-v0-3-send-gate-contract.md).

## Status

**v0.1 — 13 tools, own mailbox only.** Designed to be consumed by Juvant
OS agents (or any MCP-aware client) via `npx`.

Conforms to the handbook
[`mcp-server.md`](https://github.com/juvantlabs/handbook/blob/main/docs/repo-types/mcp-server.md)
spec.

## Install + run

```bash
npx @juvantlabs/m365-mail-mcp-server
```

For development / first-time setup, see `§ Local development` below.

## Environment variables

Required (all namespaced `M365_MAIL_*` so a typo cannot silently reuse
the sibling `m365-graph-mcp-server`'s app credentials):

| Variable | Purpose |
|---|---|
| `M365_MAIL_TENANT_ID` | Microsoft Entra tenant ID (UUID), or `common` / `organizations` / `consumers` for multi-tenant flows. |
| `M365_MAIL_CLIENT_ID` | Application (client) ID from the Entra app registration. |
| `M365_MAIL_CLIENT_SECRET` | Client secret VALUE from Entra → Certificates & secrets. Treat as a password. |

Optional:

| Variable | Purpose |
|---|---|
| `M365_MAIL_DOWNLOAD_DIR` | Root for the attachment download sandbox. Default: `$XDG_CACHE_HOME/m365-mail-mcp-server` or `~/.cache/m365-mail-mcp-server`. |
| `XDG_CACHE_HOME` | Standard XDG cache root; used only as a fallback for the sandbox. |
| `MCP_SERVER_LOG_LEVEL` | Log level for diagnostics on stderr (default `info`). |

> CI enforces that every variable documented in this section is
> actually read from `process.env.<NAME>` somewhere in `src/` —
> placeholder names containing `<>` are skipped. Documenting an env
> var without wiring it up will fail the build.

## Local development

### 1. Register a Microsoft Entra app

1. Sign in to the [Azure Portal](https://portal.azure.com/) → **Microsoft Entra ID** → **App registrations** → **New registration**.
2. Name: `m365-mail-mcp-server` (or your team's convention).
3. Supported account types: single-tenant (recommended) or as appropriate.
4. Redirect URI (Web): `http://localhost:3000/auth/callback`.
5. Register.
6. **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated**. Add:
   - `User.Read`
   - `Mail.Read`
   - `Mail.ReadWrite`
   - `offline_access`

   Do NOT add `Mail.Send` (Shield-gated for v0.3), `Mail.*.Shared`
   (v0.2 delegate mailboxes), or any `Files.*` / `Sites.*` / `Calendars.*`
   scopes (those belong to `m365-graph-mcp-server`).

7. **Grant admin consent** if your tenant policy requires it.
8. **Certificates & secrets** → **New client secret** → copy the secret **Value** immediately (Microsoft only shows it once).

### 2. Copy env-var template

```bash
cp .env.example .env.local
# Fill in the three M365_MAIL_* values from your Entra app.
```

### 3. First-time OAuth (populate the keychain)

```bash
npm install
npm run setup
```

This opens your default browser at Microsoft's consent screen. Once
consent is granted, the browser is redirected to
`http://localhost:3000/auth/callback` where a one-shot listener
exchanges the code for tokens and writes them to your OS keychain
(macOS Keychain, Linux Secret Service, or Windows Credential Manager,
via `@napi-rs/keyring`). Tokens are refreshed silently on subsequent
runs.

> **Operational note (Shield C5)** — the setup flow listens on port
> **3000**, the same port the sibling `m365-graph-mcp-server` uses
> for its OAuth callback. Only one setup can run at a time. If you
> already have `m365-graph`'s setup dance in progress, finish that
> one first (or kill any process bound to `:3000`) before running
> `npm run setup` here.

### 4. Run the MCP server

```bash
npm run dev
```

Or, after publish:

```bash
npx @juvantlabs/m365-mail-mcp-server
```

The server speaks MCP over stdio. All diagnostic output goes to
stderr; stdout is reserved for JSON-RPC framing.

## Tools

The v0.1 tool surface is **13 tools**, scoped to the caller's own
mailbox (no shared / delegate mailboxes yet — v0.2). Every tool name
is `mcp__m365-mail__<tool>` from the harness's point of view, so
mutating tools are individually targetable by deny-lists (Shield C2).

### Read

| Tool | Underlying Graph call | Scope required |
|---|---|---|
| `list_mail_folders` | `GET /me/mailFolders` | `Mail.Read` |
| `list_messages` | `GET /me/mailFolders/{f}/messages` or `/me/messages` (ordered `receivedDateTime desc`) | `Mail.Read` |
| `search_messages` | `GET /me/messages?$search=…` (KQL) | `Mail.Read` |
| `get_message` | `GET /me/messages/{id}` (full body by default; optional `body_offset` + `max_body_chars` pagination) | `Mail.Read` |
| `list_attachments` | `GET /me/messages/{id}/attachments` (metadata only) | `Mail.Read` |
| `download_attachment` | `GET /me/messages/{id}/attachments/{aid}/$value` → local sandbox path | `Mail.Read` |

### Write — idempotent

| Tool | Underlying Graph call | Scope required |
|---|---|---|
| `create_draft` | `POST /me/messages` (draft; marked with `[agent-draft]` + `X-Juvant-Agent-Author`) | `Mail.ReadWrite` |
| `update_draft` | `PATCH /me/messages/{id}` (refuses non-drafts) | `Mail.ReadWrite` |
| `create_reply_draft` | `POST /me/messages/{id}/createReply\|createReplyAll` + `PATCH` (marked subject) | `Mail.ReadWrite` |
| `create_forward_draft` | `POST /me/messages/{id}/createForward` + `PATCH` (marked subject) | `Mail.ReadWrite` |
| `mark_read` | `PATCH /me/messages/{id}` `{ isRead }` | `Mail.ReadWrite` |
| `move_message` | `POST /me/messages/{id}/move` | `Mail.ReadWrite` |

### Write — irreversible (two-phase confirmation-token)

| Tool | Underlying Graph call | Scope required |
|---|---|---|
| `delete_message` | `DELETE /me/messages/{id}` (→ Deleted Items) | `Mail.ReadWrite` |

For `delete_message`, the first call returns a preview + a
`confirmation_token`; the second call (with the token + same args)
executes the delete. Tokens are single-use, expire in 5 minutes, and
tied to the exact spec — passing a different `message_id` with
someone else's token fails with `spec_mismatch`.

### Agent-draft markers (Shield C4)

Every draft this server creates carries **both** of these markers so
they survive across Graph flows and can be visually distinguished from
mail authored by the mailbox owner:

1. **Subject prefix `[agent-draft] `** — universal, idempotent
   (never double-stacked). Applied by `create_draft`, `update_draft`
   (when patching subject), `create_reply_draft`, and
   `create_forward_draft`.
2. **Custom header `X-Juvant-Agent-Author: @juvantlabs/m365-mail-mcp-server`** — attached on `create_draft` only (Graph forbids setting
   `internetMessageHeaders` on `PATCH`, so reply/forward drafts rely
   on the subject prefix alone).

### Inbox as untrusted data (Shield C3)

Inbound message content (bodies, subjects, headers) is treated as
untrusted data. Tool parameters (recipient addresses, draft bodies)
must be **explicitly authored by the caller** — never lifted verbatim
from an inbound message. This provenance boundary is documented at the
input-schema description level and enforced by policy, not by code.

## Binding

The Juvant OS adopter binds this server in `.juvant/config.json`:

```json
{
  "mail_reader": {
    "provider": "m365-mail",
    "mcp_server": "npx @juvantlabs/m365-mail-mcp-server",
    "scope": "rw"
  }
}
```

See [handbook MCP_INVENTORY.md](https://github.com/juvantlabs/juvant-os/blob/main/docs/MCP_INVENTORY.md)
for the abstract role this server fulfills + the canonical config shape.

## Architecture

See [`ARCHITECTURE.md`](ARCHITECTURE.md) for design rationale (scope,
auth model, threat model, performance characteristics, tool catalog).

## Roadmap

- **v0.2** — shared / delegate mailboxes (`Mail.*.Shared` scopes,
  `shared_user` tool parameter).
- **v0.3** — `send_draft` tool, Shield-gated per
  [ADR 0001](docs/adr/0001-v0-3-send-gate-contract.md).
- **v0.3+** — draft attachment upload.

## Contributing

See [`CONTRIBUTING.md`](CONTRIBUTING.md). The repo follows the
[`juvantlabs/handbook`](https://github.com/juvantlabs/handbook)
conventions for MCP server repos.

## Security

See [`SECURITY.md`](SECURITY.md) for the disclosure process. Per the
[handbook security disclosure process](https://github.com/juvantlabs/handbook/blob/main/docs/security/disclosure-process.md),
report vulnerabilities privately via GitHub Security Advisory or
`security@juvant.io`.

## License

[MIT](LICENSE). Copyright (c) 2026 Juvant Srls.
