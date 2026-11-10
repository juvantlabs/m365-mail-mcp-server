# M365 Mail MCP Server

`@juvantlabs/m365-mail-mcp-server` — Model Context Protocol server wrapping
**Microsoft 365 Outlook Mail** via delegated Microsoft Graph:

- read mail folders, messages, and attachments;
- author, edit, reply-to, and forward drafts (**does NOT send**);
- move / mark-read / delete (two-phase confirmation) messages.

Send is a later, Shield-gated phase — see
[`docs/adr/0001-v0-3-send-gate-contract.md`](docs/adr/0001-v0-3-send-gate-contract.md).

## Status

**v0.3.0 — public-client auth (no client secret); 13 tools, own mailbox + shared / delegate mailboxes (with a
server-side allowlist gate).** Designed to be consumed by Juvant OS
agents (or any MCP-aware client) via `npx`. Every tool accepts an
optional `shared_user` UPN parameter to operate against a delegated /
shared mailbox the caller is authorised on; when omitted, tools default
to the caller's own mailbox (v0.1 semantics). A server-side
`M365_MAIL_ALLOWED_SHARED_USERS` allowlist (v0.2.1) gates every
non-empty `shared_user` value before any Graph call is made — see the
env-var table and the `Shared / delegate mailboxes` section below.

Conforms to the handbook
[`mcp-server.md`](https://github.com/juvantlabs/handbook/blob/main/docs/repo-types/mcp-server.md)
spec.

## Install + run

Nothing to clone. Register an Entra app once (§ 1 below), sign in once,
then point your MCP client at the npm package:

```bash
# 1. One-time sign-in: opens the browser, caches tokens in the OS keychain
M365_MAIL_TENANT_ID=<tenant-id> M365_MAIL_CLIENT_ID=<client-id> \
  npx -y @juvantlabs/m365-mail-mcp-server@0.3.0 setup

# 2. Register the server with Claude Code (or any MCP client)
claude mcp add m365-mail -s local \
  -e M365_MAIL_TENANT_ID=<tenant-id> \
  -e M365_MAIL_CLIENT_ID=<client-id> \
  -e M365_MAIL_ALLOWED_SHARED_USERS=shared-mailbox@example.com \
  -- npx -y @juvantlabs/m365-mail-mcp-server@0.3.0
```

The server is a **public client**: it holds no client secret. The tenant
ID and client ID are identifiers, not credentials, so they can live in
your MCP client configuration. The only sensitive material on disk is
the MSAL token cache, kept in the OS keychain. Drop the
`M365_MAIL_ALLOWED_SHARED_USERS` line if you only use your own mailbox.

## Environment variables

Required (all namespaced `M365_MAIL_*` so a typo cannot silently reuse
the sibling `m365-graph-mcp-server`'s app registration). Both `setup`
and the server read them from the process environment only; no file in
the current directory is needed.

| Variable | Purpose |
|---|---|
| `M365_MAIL_TENANT_ID` | Microsoft Entra tenant ID (UUID), or `common` / `organizations` / `consumers` for multi-tenant flows. |
| `M365_MAIL_CLIENT_ID` | Application (client) ID from the Entra app registration. |

No client secret: since v0.3.0 a leftover secret variable is ignored and
a one-line deprecation warning is printed on stderr (see § Migrating
from v0.2.x).

Optional:

| Variable | Purpose |
|---|---|
| `M365_MAIL_ALLOWED_SHARED_USERS` | **v0.2.1 — server-side shared-mailbox allowlist.** Comma-separated list of UPNs a caller is permitted to pass as `shared_user` (e.g. `finance@juvant.io,legal@juvant.io`). Case-insensitive. `*` (single value) restores v0.2 behaviour (any UPN accepted; Exchange is the sole gate). **Unset or empty → fail-closed: every non-empty `shared_user` value is rejected before any Graph call.** Own-mailbox calls (omit `shared_user`) are unaffected. See `§ Shared / delegate mailboxes`. |
| `M365_MAIL_DOWNLOAD_DIR` | Root for the attachment download sandbox. Default: `$XDG_CACHE_HOME/m365-mail-mcp-server` or `~/.cache/m365-mail-mcp-server`. |
| `XDG_CACHE_HOME` | Standard XDG cache root; used only as a fallback for the sandbox. |
| `MCP_SERVER_LOG_LEVEL` | Log level for diagnostics on stderr (default `info`). |

> CI enforces that every variable documented in this section is
> actually read from `process.env.<NAME>` somewhere in `src/` —
> placeholder names containing `<>` are skipped. Documenting an env
> var without wiring it up will fail the build.

## Setup

### 1. Register a Microsoft Entra app

1. Sign in to the [Azure Portal](https://portal.azure.com/) → **Microsoft Entra ID** → **App registrations** → **New registration**.
2. Name: `m365-mail-mcp-server` (or your team's convention).
3. Supported account types: single-tenant (recommended) or as appropriate.
4. Redirect URI: platform **Public client/native (mobile & desktop)**, URI `http://localhost`. Register.
   (On an existing app: **Authentication** → **Add a platform** → **Mobile and desktop applications** → custom redirect URI `http://localhost`.)
   Any port on `http://localhost` is accepted: `setup` listens on a free port chosen at run time.
5. **Authentication** → **Allow public client flows**: set to **Yes** only if you will use `setup --device-code` (machines without a browser). Leave **No** otherwise.
6. **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated**. Add:
   - `User.Read`
   - `Mail.Read`
   - `Mail.ReadWrite`
   - `Mail.Read.Shared` — v0.2, enables `shared_user` on read tools
   - `Mail.ReadWrite.Shared` — v0.2, enables `shared_user` on write tools
   - `offline_access`

   Do NOT add `Mail.Send` (the Shield-gated send phase of ADR 0001, not in
   v0.3.0), `Mail.Send.Shared` (explicitly beyond that phase per
   [ADR 0001 §D7](docs/adr/0001-v0-3-send-gate-contract.md)), or any
   `Files.*` / `Sites.*` / `Calendars.*` scopes (those belong to
   `m365-graph-mcp-server`).

7. **Grant admin consent** if your tenant policy requires it.
8. **Certificates & secrets**: nothing to create. The server never uses a client secret or certificate.

### 2. Sign in once (populate the keychain)

```bash
export M365_MAIL_TENANT_ID=<tenant-id>
export M365_MAIL_CLIENT_ID=<client-id>
npx -y @juvantlabs/m365-mail-mcp-server@0.3.0 setup
```

`setup` uses the OAuth 2.0 authorization code flow with PKCE: MSAL
starts a one-shot listener on a free port of `127.0.0.1`, opens your
default browser at Microsoft's sign-in page (the URL is also printed
on stderr), receives the code on `http://localhost:<port>` and redeems
it with the PKCE verifier. Tokens are written to your OS keychain
(macOS Keychain, Linux Secret Service, or Windows Credential Manager,
via `@napi-rs/keyring`), keyed by tenant ID, and refreshed silently on
subsequent runs. The keychain entry does not depend on the current
directory.

On a machine without a browser, use the device code flow (requires
**Allow public client flows** = Yes on the app):

```bash
npx -y @juvantlabs/m365-mail-mcp-server@0.3.0 setup --device-code
```

It prints a code to enter at `https://microsoft.com/devicelogin` from
any device.

Because the port is chosen at run time, `setup` can run while the
sibling `m365-graph-mcp-server` setup is also in progress.

### 3. Register the server with your MCP client

Claude Code:

```bash
claude mcp add m365-mail -s local \
  -e M365_MAIL_TENANT_ID=<tenant-id> \
  -e M365_MAIL_CLIENT_ID=<client-id> \
  -e M365_MAIL_ALLOWED_SHARED_USERS=shared-mailbox@example.com \
  -- npx -y @juvantlabs/m365-mail-mcp-server@0.3.0
```

Other MCP clients: run `npx -y @juvantlabs/m365-mail-mcp-server@0.3.0`
over stdio with the same environment variables. All diagnostic output
goes to stderr; stdout is reserved for JSON-RPC framing.

### Migrating from v0.2.x (client secret → public client)

v0.2.x used a confidential client with a client secret and a fixed
`http://localhost:3000/auth/callback` Web redirect. v0.3.0 needs no
secret. On your existing app registration:

1. **Authentication** → **Add a platform** → **Mobile and desktop
   applications** → custom redirect URI `http://localhost`.
2. Set **Allow public client flows** to **Yes** only if you will use
   `setup --device-code`.
3. Remove the old **Web** redirect `http://localhost:3000/auth/callback`.
4. **Certificates & secrets** → delete the client secret.
5. Remove `M365_MAIL_CLIENT_SECRET` from your MCP client configuration
   and any `.env.local` (if it is still set, the server ignores it and
   prints a one-line warning on stderr).
6. Sign in again: `npx -y @juvantlabs/m365-mail-mcp-server@0.3.0 setup`.
   A refresh token cached by v0.2.x was issued to a confidential client
   and may not be redeemable without the secret; signing in again
   replaces it in the same keychain entry.
7. Update the version pin in your MCP client configuration to `0.3.0`.

### Local development

```bash
npm install
cp .env.example .env.local   # tenant ID + client ID; gitignored
npm run setup                # = tsx --env-file=.env.local src/index.ts setup
npm run dev                  # stdio server from source
```

`.env.local` is only a convenience for the `npm run setup` / `npm run
dev` scripts in a checkout. The published package never reads it.

## Tools

The v0.2 tool surface is **13 tools** — the same 13 as v0.1, each
widened with an optional `shared_user` parameter for shared / delegate
mailbox access. When `shared_user` is omitted, tools operate on the
caller's own mailbox (v0.1 semantics, bit-for-bit). When set to a UPN
(e.g. `finance@juvant.io`), the Graph call routes to
`/users/{shared_user}/…`. Access is enforced by Exchange — passing a
UPN the caller has no permission on returns 403 from Graph.

Every tool name is `mcp__m365-mail__<tool>` from the harness's point of
view, so mutating tools are individually targetable by deny-lists
(Shield C2).

In the tables below, the paths are shown as `/me/…`; every tool also
accepts an optional `shared_user` UPN which reroutes the call to
`/users/{shared_user}/…`. The extra scope column notes the shared
scope needed when `shared_user` is set.

### Read

| Tool | Underlying Graph call | Scope required | Shared scope |
|---|---|---|---|
| `list_mail_folders` | `GET /me/mailFolders` | `Mail.Read` | `Mail.Read.Shared` |
| `list_messages` | `GET /me/mailFolders/{f}/messages` or `/me/messages` (ordered `receivedDateTime desc`) | `Mail.Read` | `Mail.Read.Shared` |
| `search_messages` | `GET /me/messages?$search=…` (KQL) | `Mail.Read` | `Mail.Read.Shared` |
| `get_message` | `GET /me/messages/{id}` (full body by default; optional `body_offset` + `max_body_chars` pagination) | `Mail.Read` | `Mail.Read.Shared` |
| `list_attachments` | `GET /me/messages/{id}/attachments` (metadata only) | `Mail.Read` | `Mail.Read.Shared` |
| `download_attachment` | `GET /me/messages/{id}/attachments/{aid}/$value` → local sandbox path | `Mail.Read` | `Mail.Read.Shared` |

### Write — idempotent

| Tool | Underlying Graph call | Scope required | Shared scope |
|---|---|---|---|
| `create_draft` | `POST /me/messages` (draft; marked with `[agent-draft]` + `X-Juvant-Agent-Author`) | `Mail.ReadWrite` | `Mail.ReadWrite.Shared` |
| `update_draft` | `PATCH /me/messages/{id}` (refuses non-drafts) | `Mail.ReadWrite` | `Mail.ReadWrite.Shared` |
| `create_reply_draft` | `POST /me/messages/{id}/createReply\|createReplyAll` + `PATCH` (marked subject) | `Mail.ReadWrite` | `Mail.ReadWrite.Shared` |
| `create_forward_draft` | `POST /me/messages/{id}/createForward` + `PATCH` (marked subject) | `Mail.ReadWrite` | `Mail.ReadWrite.Shared` |
| `mark_read` | `PATCH /me/messages/{id}` `{ isRead }` | `Mail.ReadWrite` | `Mail.ReadWrite.Shared` |
| `move_message` | `POST /me/messages/{id}/move` | `Mail.ReadWrite` | `Mail.ReadWrite.Shared` |

### Write — irreversible (two-phase confirmation-token)

| Tool | Underlying Graph call | Scope required | Shared scope |
|---|---|---|---|
| `delete_message` | `DELETE /me/messages/{id}` (→ Deleted Items) | `Mail.ReadWrite` | `Mail.ReadWrite.Shared` |

For `delete_message`, the first call returns a preview + a
`confirmation_token`; the second call (with the token + same args)
executes the delete. Tokens are single-use, expire in 5 minutes, and
tied to the exact spec — passing a different `message_id` (or a
different `shared_user`) with someone else's token fails with
`spec_mismatch`. See
[ADR 0002](docs/adr/0002-v0-2-shared-mailbox-parameter.md) §D6 for the
spec-hash invariant.

### Shared / delegate mailboxes (v0.2, allowlist gate v0.2.1)

Every tool accepts an optional `shared_user` parameter:

- **Shape**: a User Principal Name (UPN) like `finance@juvant.io`.
  GUID user ids are not accepted; a malformed value raises a loud
  error rather than silently defaulting to `/me`.
- **Omitted**: tool operates on the caller's own mailbox (v0.1
  behaviour). Not affected by the allowlist below.
- **Set**: Graph call routes to `/users/{shared_user}/…` and requires
  the corresponding `Mail.Read.Shared` / `Mail.ReadWrite.Shared`
  delegated scope granted at the app registration.
- **Access model — two layers.** In v0.2.1 access is enforced by BOTH
  of the following, in order:
    1. **Server-side allowlist** (`M365_MAIL_ALLOWED_SHARED_USERS`
       env var). Rejection happens at the tool boundary, before any
       Graph call. **Fail-closed default: if the env var is unset or
       empty, EVERY non-empty `shared_user` value is rejected.**
       Configure the env var to a comma-separated list of UPNs the
       agent may route to, e.g.
       `M365_MAIL_ALLOWED_SHARED_USERS=finance@juvant.io,legal@juvant.io`.
       Matches are case-insensitive (`Finance@juvant.io` in the list
       matches a caller-supplied `finance@juvant.io` and vice versa).
       The single value `*` disables the gate and restores v0.2
       Exchange-only enforcement — explicit opt-in for adopters who
       want Exchange as the sole gate.
    2. **Exchange itself**. Even for an allow-listed UPN, Exchange
       enforces per-mailbox access. Passing an allow-listed UPN
       does NOT grant access; it only routes the call, and Graph
       returns 403 when the caller has no delegated permission on
       that mailbox. Rejection at layer 2 is a Graph-level error
       (see the "Common errors" table).
- **Shield C4 markers**: still applied on drafts landing in a shared
  mailbox's Drafts folder.
- **Shield C3 (untrusted-data)**: still applies. Inbound content from a
  shared mailbox is untrusted; recipient addresses on drafts MUST be
  caller-authored, never lifted verbatim.
- **Non-goal**: `Mail.Send.Shared` (sending *as* another mailbox) is
  explicitly out per ADR 0001 §D7 — even beyond v0.3.

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

## Common errors

| Symptom (error / message) | Root cause | Fix |
| --- | --- | --- |
| `AADSTS65001` / `interaction_required` from MSAL on the FIRST call after upgrading to a version with wider Graph scopes (v0.1 → v0.2, or any future scope widening) | The cached delegated token still carries the OLD scope set. Silent-refresh cannot mint a token for the newly-requested `Mail.Read.Shared` / `Mail.ReadWrite.Shared` scopes until the user has re-consented under the widened set. | Re-run `setup` to re-consent under the widened scope set. If admin consent is required for the shared scopes on your tenant, ask a Global / Cloud App Admin to grant it in the Entra app registration first, then re-run `setup`. |
| `Silent token acquisition failed: … AADSTS7000218 …` (or another `invalid_grant`) after upgrading to v0.3.0 | The keychain still holds a refresh token issued to the v0.2.x confidential client, which the public client cannot redeem without a secret; or the refresh token expired / was revoked. | Run `npx -y @juvantlabs/m365-mail-mcp-server@0.3.0 setup` once (see § Migrating from v0.2.x). |
| `AADSTS7000218: The request body must contain … client_assertion or client_secret` during `setup` | The app registration still treats the redirect as a confidential (Web) client. | Add the **Mobile and desktop applications** platform with redirect `http://localhost` (§ 1 step 4); for `--device-code`, also set **Allow public client flows** = Yes. |
| `403 ErrorAccessDenied` on a `/users/{shared_user}/…` call, `shared_user` correctly shaped | The signed-in user has no Exchange permission on that mailbox (not a "Full Access" delegate, or not a shared-mailbox member). The `shared_user` parameter only routes the call — it does not grant access. | Have the mailbox owner (or an Exchange admin) add the signed-in user as a "Full Access" delegate / shared-mailbox member in Exchange Admin Center. |
| `shared_user` throws `must be a User Principal Name (UPN)` before any network call | The value is a GUID user id, or is missing an `@`, a domain suffix, or contains whitespace / a second `@`. GUID user ids are explicitly rejected by design (see ADR 0002). | Pass the UPN Graph resolves the user to (e.g. `finance@juvant.io`). Casing does not matter — the server lowercases at the input boundary. |
| `shared_user` throws `M365_MAIL_ALLOWED_SHARED_USERS` before any network call | The env var is unset / empty (fail-closed default, v0.2.1), or is set but does not include the supplied UPN. The allowlist is a server-side gate that fires BEFORE any Graph call, so a rejected UPN never reaches Graph and is never silently downgraded to `/me`. | On the deploying host, set `M365_MAIL_ALLOWED_SHARED_USERS` to a comma-separated list of UPNs the agent may route to (e.g. `finance@juvant.io,legal@juvant.io`), then restart the server. To reproduce v0.2 behaviour (Exchange as sole gate), set the env var to the single value `*`. |
| `confirmation_token spec_mismatch` on `delete_message` phase-2 | Phase-2 args do not match phase-1 args on `message_id` and/or `shared_user`. The token is bound to the exact spec, including the mailbox routing key. | Re-run phase 1 (omit `confirmation_token`) with the args you actually intend to execute against. Never carry a token across mailboxes. |

## Binding

The Juvant OS adopter binds this server in `.juvant/config.json`:

```json
{
  "mail_reader": {
    "provider": "m365-mail",
    "mcp_server": "npx -y @juvantlabs/m365-mail-mcp-server@0.3.0",
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

- **v0.2 — LANDED.** Shared / delegate mailboxes (`Mail.*.Shared`
  scopes + optional `shared_user` UPN parameter on all 13 tools).
  See [ADR 0002](docs/adr/0002-v0-2-shared-mailbox-parameter.md).
- **v0.2.1 — LANDED.** Server-side `M365_MAIL_ALLOWED_SHARED_USERS`
  allowlist gating `shared_user` before any Graph call, fail-closed
  when unset (Shield C1 defense-in-depth on top of Exchange
  delegation).
- **v0.3.0 — public client.** No client secret: interactive auth
  code + PKCE on a loopback redirect, device code fallback; install
  and configure straight from npm.
- **next** — `send_draft` tool, Shield-gated per
  [ADR 0001](docs/adr/0001-v0-3-send-gate-contract.md).
- **later** — draft attachment upload.

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
