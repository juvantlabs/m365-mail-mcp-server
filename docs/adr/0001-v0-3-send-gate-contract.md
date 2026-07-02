# ADR 0001 — v0.3 Mail.Send gate contract

- **Status**: Proposed (not yet implemented)
- **Date**: 2026-07-02
- **Author**: Mercury (m365-mail maintainer)
- **Reviewers pending ratification for v0.3**: Arch, Shield, CEO
- **Applies to**: `juvantlabs/m365-mail-mcp-server` from v0.3 onward
- **v0.1 status of this contract**: recorded only. No `send_draft` tool
  ships in v0.1. CI enforces this via the `no send tool` invariant in
  `.github/workflows/ci.yml` and the registry unit test.

## Context

The v0.1 tool surface exposes read + attachments + draft authoring +
delete (two-phase). It deliberately does NOT expose a `Mail.Send`
capability: the delegated scope is not requested, and no tool routes
to `POST /me/messages/{id}/send`, `POST /me/sendMail`, or
`POST /me/messages/{id}/reply` (the send-form of reply).

The scaffolding must, however, anticipate v0.3 cleanly. The types
(`ToolCategory`, `ConsumeResult`), the confirmation-token infra, and
the tool-registry shape are already in place; what remains for v0.3
is a specific gate contract that Shield will enforce and that CI can
statically verify.

This ADR pins that contract so the v0.1 → v0.3 transition is a
purely-additive PR — no scaffolding regret, no ambiguity about what
Shield expects.

## Decision

The v0.3 send capability MUST ship as a **single tool**,
`m365-mail:send_draft`, with the following invariants:

### D1. Two-phase confirmation-token, session-bound

A send is authorised by a `confirmation_token` that:

1. Is issued only by a preceding `create_draft` / `create_reply_draft` /
   `create_forward_draft` / `update_draft` call in the **same MCP
   session** (same subprocess, same MSAL cache lifetime).
2. Is **single-use** — the same token cannot authorise two sends.
3. Expires **≤ 5 minutes** after issuance (existing
   `src/auth/confirmation_tokens.ts` semantics — no new expiry class).
4. Is tied to `{draft_id}` — the canonical-JSON spec hash prevents a
   token issued for draft A from authorising a send of draft B.

The `send_draft` tool signature is exactly:

```
send_draft(draft_id: string, confirmation_token: string) → { sent, message_id }
```

There is **no phase-1 preview call for send** — the token is issued
by the preceding drafting tool. This is the crucial difference from
`delete_message`'s two-phase pattern: delete has no upstream tool to
issue the token, so it must issue its own. Send does have an
upstream, so we reuse it.

### D2. No one-shot `send_mail(to, subject, body)` signature

The following signatures are explicitly forbidden:

- `send_mail(to, subject, body, ...)` (no draft, no confirmation)
- `send_draft(draft_id)` without confirmation token
- Any convenience wrapper that composes draft-create → send in a
  single tool call

Rationale: the draft is the human-reviewable artefact. Skipping the
draft phase collapses the audit trail and denies the CEO the chance
to see the marked draft in Outlook's Drafts folder before it goes
out.

### D3. Full recipient surfacing including BCC

The `send_draft` response MUST return `{ to, cc, bcc }` extracted
from the sent message. The BCC field is load-bearing: the harness /
audit log MUST record BCC recipients even though they are invisible
in the delivered mail.

### D4. Recipient set is CEO-authored, never inbox-derived

Per Shield C3 (inbox-as-untrusted-data), the recipient set of any
draft that a send tool can consume MUST have been authored by the
CEO / caller — never lifted verbatim from an inbound message body.
The draft-authoring tools (`create_draft`, `create_reply_draft`,
`create_forward_draft`) enforce this at the input-schema description
level in v0.1 already; v0.3 will additionally require that the
`send_draft` handler asserts:

- The draft in question was created by *this server* (checked via the
  `X-Juvant-Agent-Author` header presence on the draft, per Shield C4).
- The draft's `toRecipients[]` are all present in an allowlist that
  the harness supplies (mechanism TBD — likely an env-configured
  domain allowlist plus per-run CEO overrides).

### D5. Harness deny in non-interactive contexts

Shield's harness pre-tool-use hook MUST deny `mcp__m365-mail__send_draft`
by default in non-interactive / batch execution contexts. Interactive
contexts (a human is watching the tool call) may allow it under CEO
policy. This is a harness invariant, not a server invariant — the
server just needs to be denied cleanly (no side effects on rejection).

### D6. Rate cap ~5/hour, per-send audit row

The server enforces a soft rate cap of **≤ 5 sends per rolling hour**
per subprocess. On breach, the tool throws with a structured error
(`{ error: "rate_capped", retry_after_seconds }`) and does not
attempt the send. Every successful send emits an audit record via a
harness-provided sink (mechanism TBD; likely a JSON line on stderr
tagged with a `MCP_AUDIT` prefix that the harness parses).

### D7. `Mail.Send.Shared` is out of scope

The v0.3 send capability applies to `/me/sendMail` only. Sending as
another user (`/users/{id}/sendMail` via delegate) requires the
`Mail.Send.Shared` scope which we do NOT request in v0.3.
Shared-mailbox delegation is a separate v0.4+ conversation that
requires its own Shield review.

### D8. CI-enforced invariant: no send without consumed token

A CI check in `.github/workflows/ci.yml` MUST fail the build if any
tool under `src/tools/` references `/sendMail`, `/send`,
`/messages/{id}/reply` (send-form), `/messages/{id}/replyAll`
(send-form), or `/messages/{id}/forward` (send-form) without also
importing `consumeConfirmation` from `src/auth/confirmation_tokens.js`.
The check parallels the existing `write_irreversible` enforcement.

## Consequences

**Positive**:

- v0.3 lands as a purely-additive PR: one new tool file, one new
  scope, one new CI grep pattern, one new CHANGELOG entry.
- Shield can pre-approve the shape of the v0.3 PR *now* because the
  contract is pinned.
- Failure modes at the moment of send are structured and observable
  (rate cap, expired token, spec mismatch, wrong tool).

**Negative**:

- The two-phase-across-tools pattern (create issues, send consumes)
  is unusual and needs clear docs — an agent reading only `send_draft`
  can't discover the token issuance without following the chain.
- The `create_draft` output shape needs to grow a `confirmation_token`
  field in v0.3, which is a minor schema evolution. v0.1 does not
  need to prepare for this — the addition is backward-compatible.

**Neutral**:

- Rate cap is intentionally per-subprocess, not per-user. If two
  concurrent m365-mail subprocesses run against the same tenant,
  the effective cap doubles. Multi-subprocess rate coordination is
  out of scope for v0.3.

## Alternatives considered

1. **One-shot `send_mail(to, subject, body)`** — rejected. Collapses
   audit trail and denies the CEO the draft-preview window. See D2.

2. **`send_draft(draft_id)` without token** — rejected. Turns every
   drafting tool into a latent send trigger; a compromised draft-id
   could be replayed. The token-per-draft binding is what makes the
   send authorisation session-scoped.

3. **Reuse the existing two-phase-in-one-tool pattern from
   `delete_message`** — rejected. Would require a duplicate
   confirmation-token issue at send time, which is redundant given
   the draft tools already issue one, and would create a "why is
   there both a create_draft token and a send preview token?"
   documentation surface.

## References

- Shield C3 (inbox-as-untrusted-data), C4 (agent-draft marker),
  C5 (README op note) from the Azure app PR review, 2026-07-02.
- ADR 0009 (juvantlabs handbook) — untrusted-data provenance.
- Manifesto Article 4 — human-in-the-loop for irreversible outward
  actions.
- `src/auth/confirmation_tokens.ts` — token infra reused verbatim.
- Existing `write_irreversible` CI enforcement in `.github/workflows/ci.yml`.
