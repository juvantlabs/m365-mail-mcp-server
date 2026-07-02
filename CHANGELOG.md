# Changelog

All notable changes to `@juvantlabs/m365-mail-mcp-server` will be documented in this
file.

The format is based on
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this
project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [Unreleased]

## [0.2.1] — server-side shared-mailbox allowlist (Shield C1 defense-in-depth)

### Added — v0.2.1

- **`M365_MAIL_ALLOWED_SHARED_USERS` env-var allowlist.** New optional
  environment variable that gates the v0.2 `shared_user` parameter at
  the tool boundary, BEFORE any Graph call is made. This is a second,
  fail-closed enforcement layer on top of Exchange delegation:
    - **Unset / empty → fail-closed.** Every non-empty `shared_user`
      value is rejected with a loud, specific error naming the env var.
      Own-mailbox `/me` calls (`shared_user` omitted) are unaffected.
    - **`*` sentinel → v0.2 posture.** Explicit opt-in for adopters who
      want Exchange as the sole gate: any UPN-shaped value is accepted
      and only Exchange delegation stops the call.
    - **Comma-separated UPN list → exact allowlist.** Each entry is
      trimmed and lowercased before membership check, so
      `Finance@juvant.io` in the env var matches a caller
      `finance@juvant.io` and vice versa. Non-matching UPN → rejected
      before any Graph call, never silently downgraded to `/me`.
  Enforcement point: `validateSharedUser` in `src/tools/_mailbox.ts`,
  after D1 UPN-shape validation and FUP-2 case normalization, and
  before the value is returned to any downstream consumer
  (`mailboxRoot`, spec-hash, sandbox key). Because normalization runs
  first, the caller sees a shape error for a mis-formed UPN (not a
  policy error) — the two failure modes are distinct so a mistyped
  UPN doesn't send the operator chasing an allowlist misconfiguration.
- **16 new unit tests** in `tests/unit/mailbox.test.ts` pin the
  invariants: allow-listed UPN passes; non-listed UPN rejected before
  Graph; unset env → all shared rejected but `/me` still works;
  empty / whitespace-only env same as unset; `*` restores v0.2
  Exchange-only behaviour; case-insensitive match in both directions;
  whitespace trimmed around comma-separated entries; empty entries
  (trailing comma) don't match empty input; shape check fires BEFORE
  allowlist check; error messages name `M365_MAIL_ALLOWED_SHARED_USERS`
  so operators know where to fix policy; unset-env message calls out
  the `*` opt-in AND that own-mailbox calls are unaffected; and the
  load-bearing "never silently downgraded to `/me`" confused-deputy
  invariant.
- **`tests/setup.ts`** — Vitest global setup file that seeds a
  `M365_MAIL_ALLOWED_SHARED_USERS=*` default for the test suite. Unit
  tests that exercise the 13-tool routing composition pass
  `shared_user: "finance@juvant.io"` as a fixture; they are NOT
  testing the policy layer, so a test-only wildcard keeps their
  focus on the behaviour they were written for. Dedicated allowlist
  tests manage the env var explicitly with `beforeEach` / `afterEach`.
  Production semantics stay fail-closed on unset — that invariant is
  exercised by the dedicated tests, not diluted by this default.
- **README updates.** New env-var row for
  `M365_MAIL_ALLOWED_SHARED_USERS`; expanded `Shared / delegate
  mailboxes` section describing the two-layer enforcement model
  (allowlist → Exchange); new row in the `Common errors` table for the
  allowlist rejection failure mode; roadmap line for v0.2.1 marked
  landed.
- **Schema description.** `SHARED_USER_SCHEMA_PROPERTY.description`
  now spells out the two-layer enforcement (Exchange + v0.2.1
  server-side allowlist, fail-closed on unset). Pinned by a new unit
  test so the wording doesn't silently drift.

### Security — v0.2.1

- The allowlist is defense-in-depth. It DOES NOT replace Exchange
  delegation; it adds a second, operator-controlled gate that is
  configured at deploy time (via env var) rather than in Exchange
  Admin Center. An agent-side compromise or a mis-scoped tool call
  therefore fails at the server boundary before Graph is touched,
  even if Exchange would have accepted the delegation.
- Fail-closed default (unset env → all shared rejected) is a
  deliberate posture change from v0.2: the safe direction for a
  server that's forgotten to be configured is to route to `/me`
  only, not to route anywhere Exchange happens to allow.
- Own-mailbox `/me` calls are unaffected. An agent that never
  supplies `shared_user` continues to work bit-for-bit as in v0.2
  regardless of allowlist configuration.

### Approved via

- CEO decision #227 (adopting the server-side allowlist).
- `security_audit_log` #84 — Shield concurrence, condition C1
  (allowlist MUST land before any agent invokes `shared_user` in
  production).
- Tracking issue: [#7](https://github.com/juvantlabs/m365-mail-mcp-server/issues/7).

---

## [0.2.0] — shared / delegate mailboxes

### Added — v0.2 (shared / delegate mailboxes)

- **`shared_user` optional parameter on all 13 tools.** Every tool
  (read, idempotent-write, and irreversible-write) now accepts an
  optional `shared_user` UPN parameter (e.g. `finance@juvant.io`).
  When set, the underlying Graph call routes to `/users/{shared_user}/…`
  instead of `/me/…`, letting the caller operate on a shared /
  delegated mailbox they have been granted access to at the Exchange
  level. Omitted → v0.1 behaviour (bit-for-bit identical to v0.1
  when the parameter is absent). Composition point:
  `src/tools/_mailbox.ts::mailboxRoot`. Design record:
  [ADR 0002](docs/adr/0002-v0-2-shared-mailbox-parameter.md).

- **`Mail.Read.Shared` + `Mail.ReadWrite.Shared` delegated scopes**
  added to `DELEGATED_SCOPES` in `src/auth/msal.ts`. Requires the app
  registration to be updated (add the two scopes + admin consent) —
  this is an eng-platform action, tracked separately in
  `juvant-shared-infra` Terraform. Until that lands, calls with
  `shared_user` set will 403 at Graph.

- **`shared_user` folded into `delete_message` confirmation-token spec
  hash.** The canonical-JSON spec now includes `shared_user` (null
  for own-mailbox), so a token issued against own-mailbox message A
  cannot authorise a delete of shared-mailbox message B with the
  same id. Three new unit tests pin the invariant (own→shared,
  shared→own, shared A→shared B).

- **`shared_user` folded into `download_attachment` sandbox hash.**
  `deriveSafeLocalPath` now includes the mailbox routing key
  (`"me"` or the UPN) in the SHA-256 that derives the local
  filename. Defense-in-depth against filesystem collisions where the
  same `message_id`+`attachment_id` pair could belong to different
  mailboxes.

- **`docs/adr/0002-v0-2-shared-mailbox-parameter.md`** records the
  six design decisions from [#5](https://github.com/juvantlabs/m365-mail-mcp-server/issues/5):
  UPN-only parameter shape, delegated-vs-shared distinction not
  surfaced, single composition point, Shield C4 markers preserved on
  shared drafts, threat-model boundary unchanged (still Mail-only),
  and the delete-spec-hash invariant.

### Changed — v0.2

- `tests/unit/registry.test.ts` — invariant flipped from "v0.1 must
  NOT ship any `shared_user` parameter" to "v0.2 exposes `shared_user`
  as an optional parameter on every tool". Every tool's schema is
  now asserted to include the parameter with the shared-scope-aware
  description (the description constant is
  `SHARED_USER_SCHEMA_PROPERTY` in `_mailbox.ts`).
- `README.md` and `ARCHITECTURE.md` — reflect the v0.2 status, the
  widened Entra scope list, the new `shared_user` parameter on all 13
  tool tables, and the roadmap line for v0.2 marked landed. The
  "Do NOT add `Mail.*.Shared`" line in the app-registration steps
  is replaced with the corresponding "Add `Mail.Read.Shared` /
  `Mail.ReadWrite.Shared`" instruction.

### Follow-ups folded into the v0.2 PR (post deep-review)

- **UPN case normalization at the input boundary** (FUP-2).
  `validateSharedUser` in `src/tools/_mailbox.ts` now lowercases the
  UPN after trim + shape validation. UPNs are semantically
  case-insensitive in Entra / Exchange, but the value flows into two
  byte-sensitive client-side hashes — the `delete_message`
  confirmation-token spec-hash (SHA-256 over canonical JSON) and the
  `download_attachment` sandbox-path hash — where `Finance@juvant.io`
  and `finance@juvant.io` would otherwise fork into two different
  tokens and two different sandbox files for what a human reads as the
  same mailbox. Normalizing at the SINGLE point where `shared_user`
  enters the system means every downstream consumer (routing, token
  spec, sandbox key) sees the same canonical value. New unit tests
  pin: mixed-case → lowercase normalization; cross-case delete-token
  spec-hash equality (mixed→lower and lower→mixed both validate);
  cross-case sandbox-path equality; `mailboxRoot` is case-preserving
  (normalization is the validator's job, not the formatter's).
- **README "Common errors" table** (FUP-3). Maps four common
  failure modes to their fixes: MSAL `interaction_required` after a
  scope widening (fix: re-run `npm run setup`), `403` on a
  correctly-shaped shared-mailbox call, invalid UPN shape, and
  `delete_message` `spec_mismatch`. The MSAL row is the load-bearing
  one for the v0.2 rollout — first-time callers hitting the widened
  scope set will otherwise loop through a confusing silent-refresh
  failure without knowing they need to re-consent.
- **`canonicalize()` upgrade warning** (FUP-4).
  `src/auth/confirmation_tokens.ts` comment on `canonicalize()`
  clarifies that the shallow (top-level-only) key sort is intentional
  and sufficient for the current flat-primitive spec shape used by
  `delete_message` (`{message_id, shared_user}`), and warns that any
  future `write_irreversible` tool shipping a nested spec MUST upgrade
  canonicalization to recurse — sort keys at every level and
  canonicalize arrays element-wise — before it ships. Docs-only; no
  behavior change.

### Security — v0.2 boundaries preserved

- `Mail.Send` and `Mail.Send.Shared` remain out of scope. `Mail.Send`
  is a v0.3 concern (Shield-gated per ADR 0001). `Mail.Send.Shared` is
  explicitly beyond v0.3 per ADR 0001 §D7.
- Application (app-only) permissions remain out of scope. v0.2 stays
  delegated-only.
- `Files.*` / `Sites.*` / `Calendars.*` scopes and their `.Shared`
  equivalents remain the sibling `m365-graph-mcp-server`'s territory.

### Release engineering

- **Publish workflow migrated to pure npm Trusted Publishing (OIDC).**
  `.github/workflows/publish.yml` — the temporary
  `NODE_AUTH_TOKEN: ${{ secrets.NPM_TOKEN }}` env block on the publish
  step (added for v0.1.0's first-publish bootstrap, since npm's "Add
  Trusted Publisher" UI cannot pre-register a package that doesn't yet
  exist on npmjs.com) is now removed. After v0.1.0 shipped, the Trusted
  Publisher was registered on npmjs.com — Publisher: GitHub Actions,
  Org: `juvantlabs`, Repo: `m365-mail-mcp-server`, Workflow:
  `publish.yml`, Environment: `production`, Permission: `npm publish` —
  and from v0.1.1 onward every publish authenticates via the GitHub
  Actions OIDC token. No static secret to rotate; scope tied to the
  exact (repo, workflow, environment) tuple. This mirrors the sibling
  `m365-graph-mcp-server`'s v0.1.0 → v0.1.3 migration but collapsed
  into a single follow-up PR because the `npm install -g npm@latest`
  step (which was the pitfall for m365-graph's v0.1.2) was already
  present in v0.1.0's workflow. Repo secret `NPM_TOKEN` is now
  redundant and can be revoked/deleted independently.

### Prerequisite for full v0.2 rollout

End-to-end verification against a live shared mailbox
(`finance@juvant.io`) is BLOCKED until eng-platform lands the app
registration scope change in `juvant-shared-infra`
(`terraform/shared/agent-appregistrations/m365_mail_mcp.tf` → add
`Mail.Read.Shared` + `Mail.ReadWrite.Shared` delegated scopes → grant
admin consent → re-run `npm run setup` to widen the cached token).
The code in this branch merges "implemented, not yet e2e-verified
against a live shared mailbox."

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

[0.2.1]: https://github.com/juvantlabs/m365-mail-mcp-server/releases/tag/v0.2.1
[0.2.0]: https://github.com/juvantlabs/m365-mail-mcp-server/releases/tag/v0.2.0
[0.1.0]: https://github.com/juvantlabs/m365-mail-mcp-server/releases/tag/v0.1.0
