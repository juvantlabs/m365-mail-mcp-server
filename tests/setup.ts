/**
 * Vitest global setup — one-shot env-var defaults for the unit + integration
 * test suites.
 *
 * v0.2.1 — server-side shared-mailbox allowlist
 *
 * `validateSharedUser` in `src/tools/_mailbox.ts` gates every non-empty
 * `shared_user` value on `M365_MAIL_ALLOWED_SHARED_USERS` (fail-closed
 * when unset). Most unit tests that exercise the 13-tool surface pass
 * `shared_user: "finance@juvant.io"` (or similar) to prove that the
 * routing composition works end-to-end — they are NOT testing the
 * policy layer, and forcing every one of them to seed the env var
 * would just repeat noise.
 *
 * We install a wildcard default here so those tests behave the way
 * they always did (v0.2 posture: any UPN-shaped value is accepted,
 * Exchange is the sole gate). Tests that specifically exercise the
 * allowlist (`tests/unit/mailbox.test.ts::validateSharedUser —
 * allowlist (v0.2.1)`) manage the env var explicitly per-case with
 * `beforeEach` / `afterEach` — they overwrite whatever this file
 * sets and restore it after.
 *
 * IMPORTANT: this default is a TEST-ONLY convenience. In production,
 * unset env means "reject all shared_user values" (fail-closed), and
 * that invariant is exercised by the dedicated allowlist tests.
 */
if (process.env.M365_MAIL_ALLOWED_SHARED_USERS === undefined) {
  process.env.M365_MAIL_ALLOWED_SHARED_USERS = "*";
}
