# Changelog

All notable changes to `@juvantlabs/m365-mail-mcp-server` will be documented in this
file.

The format is based on
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this
project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [0.1.0] — 2026-07-02

Initial functional release. 13 tools, own-mailbox only, no send.

### Added

- **Delegated OAuth auth.** MSAL Node confidential-client + Authorization
  Code flow, wired to OS keychain via `@napi-rs/keyring`. Env vars are
  namespaced `M365_MAIL_*` for isolation from the sibling
  `@juvantlabs/m365-graph-mcp-server`. Interactive setup dance via
  `m365-mail-mcp-server setup` writes the initial token cache; runtime
  refreshes silently.

- **Delegated scopes** — `User.Read`, `Mail.Read`, `Mail.ReadWrite`,
  `offline_access`. Deliberately narrow. `Mail.Send`,
  `Mail.*.Shared`, and all `Files.*`/`Sites.*`/`Calendars.*` scopes
  are explicitly NOT requested; unit tests
  (`tests/unit/msal.test.ts`) enforce their absence.

- **Read tools (6):**
  - `m365-mail:list_mail_folders` — `GET /me/mailFolders`
  - `m365-mail:list_messages` — `GET /me/mailFolders/{f}/messages` or
    `/me/messages`, `receivedDateTime desc`
  - `m365-mail:search_messages` — `GET /me/messages?$search=…` with
    `ConsistencyLevel: eventual`
  - `m365-mail:get_message` — full body returned untruncated by default.
    Optional caller-driven pagination via `body_offset` + `max_body_chars`
    (mirrors the sibling `m365-graph:get_transcript` offset/max_chars
    contract); the response includes `next_offset` for the next chunk when
    the slice does not reach the end. `body_truncated` is retained as a
    boolean alias for `next_offset !== null` and is NEVER `true` for a
    paramless call — no silent data loss.
  - `m365-mail:list_attachments` — metadata only
  - `m365-mail:download_attachment` — sandboxed path return (not
    bytes); rejects `itemAttachment` and `referenceAttachment` with
    structured errors; 200 MB cap; streaming pipeline

- **Idempotent write tools (6):**
  - `m365-mail:create_draft` — `POST /me/messages`
  - `m365-mail:update_draft` — `PATCH /me/messages/{id}`, refuses PATCH
    when `isDraft !== true`
  - `m365-mail:create_reply_draft` — `createReply`/`createReplyAll` +
    `PATCH`
  - `m365-mail:create_forward_draft` — `createForward` + `PATCH`
  - `m365-mail:mark_read` — `PATCH /me/messages/{id}` with `isRead`
  - `m365-mail:move_message` — `POST /me/messages/{id}/move`

- **Irreversible write tool (1):**
  - `m365-mail:delete_message` — `DELETE /me/messages/{id}`, gated by
    the two-phase confirmation-token pattern (single-use, 5-min TTL,
    canonical-JSON SHA-256 spec hash)

- **Shield C4 agent-draft markers.** Every draft created by this
  server carries two independent provenance markers:
  - Subject prefix `[agent-draft] ` (idempotent; applied on
    create_draft, update_draft, and the reply/forward variants).
  - Custom internet-message header `X-Juvant-Agent-Author:
    @juvantlabs/m365-mail-mcp-server` on `create_draft` (Graph forbids
    setting `internetMessageHeaders` on `PATCH`, so reply/forward
    drafts rely on the subject prefix).

- **Sandboxed attachment downloads.** `download_attachment` writes
  under `M365_MAIL_DOWNLOAD_DIR` / `$XDG_CACHE_HOME/m365-mail-mcp-server`
  / `~/.cache/m365-mail-mcp-server`. Filenames derived from
  `sha256(message_id || "::" || attachment_id)[:16] + sanitized name`
  with 0o700 dir / 0o600 file mode + prefix-resolve verification.

- **CI enforcement layers** (mirrored from
  `juvantlabs/m365-graph-mcp-server`):
  - Stdout discipline (no `console.log` in `src/`).
  - Dead-code check (every exported `validate*`/`sanitize*` helper
    must be imported elsewhere in `src/`).
  - README env-var accuracy (every documented env var must appear in
    `src/` via `process.env.<NAME>`).
  - Confirmation-token enforcement for every `write_irreversible`
    tool.
  - Permission-mutation surface invariant (CI Layer A) —
    v0.1 baseline: zero `permission_mutating` tools, zero
    permission-mutation-class scopes.
  - **New:** v0.3 send-gate invariant (ADR 0001 D8) — fails the build
    if any tool references `/sendMail`, `/send`, `/reply` (send-form),
    `/replyAll` (send-form), or `/forward` (send-form) without a
    consumed `confirmation_token`.

- **`docs/adr/0001-v0-3-send-gate-contract.md`** — records the v0.3
  `Mail.Send` gate contract as an ADR so the v0.3 PR shape is
  pre-approved. Not implemented in v0.1.

### Security

- Isolated from sibling `m365-graph-mcp-server`: distinct env-var
  prefix (`M365_MAIL_*` vs `M365_*`), distinct keychain service name,
  distinct download-sandbox subdirectory. Prevents silent scope
  escalation via config typo.
- 238 unit tests (98.9% line coverage, 91.2% branch). Zero flaky
  tests, no `.skip`.
- `npm audit --audit-level=moderate` clean.

### Release engineering

- **First-publish token bootstrap.** The publish workflow
  (`.github/workflows/publish.yml`) was copied from
  `m365-graph-mcp-server` in its post-migration OIDC-only shape (npm
  Trusted Publishing, no `NODE_AUTH_TOKEN`). npm's Trusted Publisher
  UI, however, cannot pre-register a package that does not yet exist
  on npmjs.com. v0.1.0's publish step therefore temporarily carries a
  `NODE_AUTH_TOKEN: ${{ secrets.NPM_TOKEN }}` env — a Granular Access
  Token stored as the repo secret `NPM_TOKEN` — while the rest of the
  workflow (npm upgrade, `--provenance`, `id-token: write`, the
  `production` environment approval gate) is unchanged. Provenance
  attestation still runs via OIDC and is independent of publish auth.
  Immediately after v0.1.0 is live we register the Trusted Publisher
  on npmjs.com and land a follow-up PR that deletes the `env` block,
  reverting to pure OIDC — no other workflow changes. This mirrors
  the sibling `m365-graph-mcp-server`'s v0.1.0 → v0.1.3 migration
  (token bootstrap → OIDC attempt → npm-version fix), the difference
  being that here the migration will be one PR instead of three
  releases because we already know the `npm install -g npm@latest`
  step is required.

### Planned / not implemented

- **v0.2** — shared / delegate mailboxes (`Mail.*.Shared` scopes,
  `shared_user` parameter on every applicable tool). Requires Shield
  review of untrusted-header-forgery threat model on delegate
  responses.
- **v0.3** — `send_draft` tool per ADR 0001. Two-phase
  confirmation-token model: draft tool issues token → `send_draft`
  consumes it. NO one-shot `send_mail(to, subject, body)` signature.
  Full To/Cc/Bcc surfacing (including BCC). Rate cap ~5/hour.
  Harness-side deny in non-interactive contexts.
- **v0.3+** — draft attachment upload via Graph upload session.

---

[0.1.0]: https://github.com/juvantlabs/m365-mail-mcp-server/releases/tag/v0.1.0
