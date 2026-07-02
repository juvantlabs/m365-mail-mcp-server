# ADR 0002 — v0.2 shared / delegate mailbox parameter

- **Status**: Accepted (implemented in v0.2)
- **Date**: 2026-07-02
- **Author**: m365-mail maintainer
- **Applies to**: `juvantlabs/m365-mail-mcp-server` from v0.2 onward
- **Tracking issue**: [#5](https://github.com/juvantlabs/m365-mail-mcp-server/issues/5)

## Context

v0.1 shipped 13 tools scoped to the caller's own mailbox (`/me/…`). The
Juvant finance workflow needs to read `finance@juvant.io` — a shared
mailbox `antonio@juvant.io` has "Full Access" on — to build the Q2 2026
passive-invoice registry. That mailbox is unreachable through v0.1: the
tools have no way to route the Graph call anywhere except `/me`.

v0.2 adds:

- The `Mail.Read.Shared` + `Mail.ReadWrite.Shared` delegated scopes.
- An OPTIONAL `shared_user` UPN parameter on every tool that routes the
  Graph call to `/users/{shared_user}/…` when present, `/me/…`
  otherwise.

## Decision

The v0.2 shared-mailbox capability MUST ship with the following
invariants. These resolve the six open design questions raised on #5.

### D1. `shared_user` is a UPN string; GUID user ids are not accepted

The parameter accepts a User Principal Name (e.g. `finance@juvant.io`)
and rejects everything else at the schema/handler boundary. Graph would
also accept a raw GUID user id, but agents reason about mailboxes by
UPN, and the syntactic UPN gate lets us reject obvious garbage before
it reaches Graph. A malformed value fails LOUD (thrown error), never
silently degrading to `/me` — that failure mode would be a dangerous
scope-change bug.

Rationale: UPN is what humans and agents actually use; GUIDs would be a
separate, later change if we ever grow a concrete need.

### D2. The delegated-vs-shared distinction is deliberately not exposed

Graph treats a shared mailbox (`finance@`) and a delegated mailbox (a
grant of "read from another user's mailbox") uniformly under
`Mail.*.Shared` + `/users/{id}/…`. The tool surface does not surface
the internal Exchange distinction — the caller sees a single parameter
"another mailbox I'm authorised on". If the caller is not authorised,
Graph returns 403 and the tool surfaces it as a normal error.

### D3. `shared_user` composes across all 13 tools via a single helper

There is exactly one composition point: `mailboxRoot(sharedUser)` in
`src/tools/_mailbox.ts`. Every tool calls it once and threads the
returned root (`/me` or `/users/{encoded-upn}`) into its Graph URL.
Hand-rolling `/me/…` string concatenation in a new tool is a review
smell — it silently skips the shared-mailbox path. A registry test
(`registry.test.ts` "v0.2 exposes `shared_user`") fails the build if a
tool lands without wiring the parameter through.

### D4. Shield C4 markers still apply on shared-mailbox drafts

A draft this server creates in `finance@`'s Drafts folder is still an
agent-authored draft, and the mailbox's other delegates should be able
to tell it apart from a human-authored one. Both markers apply:

- Subject prefix `[agent-draft] ` — unchanged.
- Custom `X-Juvant-Agent-Author` header — unchanged.

Nothing about `shared_user` alters the marker logic; the same
`_shared.ts` helpers do the work.

### D5. Threat-model boundary — still within Outlook Mail

`Mail.*.Shared` widens *access to Mail*, not access to the m365 stack.
This ADR does NOT authorise:

- `Files.*.Shared`, `Sites.*.Shared`, `Calendars.*.Shared` scopes (those
  belong to `m365-graph-mcp-server`).
- Application (app-only) permissions on any of the above — a full
  Shield review would be required.
- Cross-tenant access.

The env-var namespacing (`M365_MAIL_*`), keychain service name, and
download sandbox subdirectory remain distinct from the sibling
`m365-graph-mcp-server`. The isolation from ADR 0003 (handbook, MCP
scope boundaries) holds.

### D6. `delete_message` folds `shared_user` into the confirmation-token spec hash

The two-phase delete's canonical-JSON spec now includes `shared_user`
(null for own-mailbox). This closes a specific replay-across-mailboxes
hole:

- Phase-1 preview against own-mailbox message A issues a token bound
  to `{ message_id: "A", shared_user: null }`.
- Phase-2 with the same `message_id` but a different `shared_user` (or
  vice versa) MUST fail with `spec_mismatch`.

Graph message ids are already mailbox-scoped in practice, but hashing
the mailbox routing key into the spec makes the guarantee explicit
and survives any future Graph id-format change. Three new unit tests
in `tests/unit/delete_message.test.ts` pin the invariant:

1. Token issued for own-mailbox, replayed with `shared_user` → reject.
2. Token issued for shared, replayed without → reject.
3. Token issued for shared A, replayed against shared B → reject.

The download sandbox path in `deriveSafeLocalPath` also folds
`sharedUser` into the SHA-256 key for the same defense-in-depth reason
(prevents a filesystem collision where the same `message_id` +
`attachment_id` pair could route to different mailboxes).

## Consequences

**Positive**:

- Finance workflow can read `finance@juvant.io` via the connector once
  eng-platform lands the app-registration scope change.
- v0.2 is a purely-additive schema change from v0.1: every tool grows
  one optional field, no existing field changes shape, callers omitting
  `shared_user` see v0.1 semantics bit-for-bit.
- The composition point (`mailboxRoot`) is a single 2-line function.
  Adding a v0.3+ tool "just works" as long as it calls the helper.
- Confirmation-token spec-hash safety generalises to any future tool
  that adds a `write_irreversible` category: fold the new distinguishing
  parameter into the spec, and the canonical-JSON canonicalisation
  covers replay-across-scope hygiene automatically.

**Negative**:

- The Entra app registration must be updated in `juvant-shared-infra`
  Terraform (see `terraform/shared/agent-appregistrations/m365_mail_mcp.tf`)
  to add the two `.Shared` scopes and admin consent. This is an
  eng-platform action, not an in-repo change; the PR is code-complete
  before that lands and end-to-end verification against
  `finance@juvant.io` is BLOCKED until it does.
- The `deriveSafeLocalPath` hash changes shape (mailbox key is now
  part of the hashed input). This is not a semver break — the local
  sandbox is treated as an implementation detail — but tenants that
  re-run against previously-downloaded attachments will see the same
  bytes cached under a new filename. Not a data-loss risk; only a
  small extra disk usage.

**Neutral**:

- `Mail.Send.Shared` remains OUT of scope (ADR 0001 §D7). v0.3 will add
  own-mailbox send only. Sending as another mailbox is a separate,
  post-v0.3 conversation that requires its own Shield review.
- The `delegated vs shared` Exchange distinction is not exposed. If a
  caller ever needs to know whether they hold a "Full Access" grant vs
  a "Send on Behalf Of" delegation, they can inspect Graph directly via
  `m365-graph`.

## Alternatives considered

1. **Separate `read_shared_*` tools instead of parameter widening.**
   Rejected — doubles the tool surface (26 instead of 13), forces
   agents to learn two names for every operation, and makes the
   deny-list surface harder to reason about. The optional-parameter
   shape is the industry norm for this pattern.

2. **`shared_user` accepts either UPN or GUID.** Rejected for v0.2 —
   UPN is what callers know; GUID acceptance would erode the loud-
   failure semantics of a mistyped UPN. Can be added later without
   breaking anything.

3. **Expose the delegated-vs-shared distinction as a second flag.**
   Rejected — the internal Exchange distinction is Microsoft's, not
   the caller's. Wrapping it would leak an implementation detail
   that Graph itself abstracts away.

4. **Application (app-only) permissions.** Rejected — this server is
   built around the delegated identity of the running user. Moving to
   app-only would collapse the identity boundary the whole security
   model rests on. If we ever need it (e.g. a batch-cron reading a
   shared mailbox), that is a full re-architect + Shield review, not
   a v0.2 change.

## References

- Tracking issue: [#5](https://github.com/juvantlabs/m365-mail-mcp-server/issues/5).
- ADR 0001 §D7 — `Mail.Send.Shared` explicitly beyond v0.3.
- Handbook MCP-server spec § Isolation from sibling server.
- `src/tools/_mailbox.ts` — the composition helper.
- `src/tools/delete_message.ts` — the spec-hash change.
- `src/auth/msal.ts` — the widened `DELEGATED_SCOPES`.
- `tests/unit/registry.test.ts` — the invariant flip from
  "must NOT ship `shared_user`" (v0.1) to
  "MUST expose `shared_user` on every tool" (v0.2).

## Enablement prerequisite (eng-platform handoff)

End-to-end verification is BLOCKED until eng-platform lands the
following in `juvant-shared-infra`:

- Add `Mail.Read.Shared` (delegated) to
  `terraform/shared/agent-appregistrations/m365_mail_mcp.tf`.
- Add `Mail.ReadWrite.Shared` (delegated).
- `terraform apply` via the shared-infra workflow.
- Grant admin consent for the two new scopes (Azure portal, one-time
  per tenant).
- Re-run `npm run setup` in this server to refresh the cached refresh
  token with the widened scope set.
- Verify `antonio@juvant.io` has at least "Read" permission on
  `finance@juvant.io` in Exchange (Full Access is what "Full Access +
  Send As" delegations grant).

The code merges "implemented, not yet e2e-verified against a live
shared mailbox." The first successful read of `finance@juvant.io`
through this server closes the v0.2 milestone.
