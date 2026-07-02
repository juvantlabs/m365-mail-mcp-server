/**
 * Token persistence via the OS keychain.
 *
 * Uses `@napi-rs/keyring` (NOT `keytar` — archived since 2022 per
 * handbook spec anti-pattern #10). Tokens are stored under a
 * (service, account) pair where the account is keyed by tenant ID
 * so multiple tenant configs don't collide.
 *
 * Service name is INTENTIONALLY distinct from the sibling
 * `juvantlabs-m365-graph-mcp-server` service so the two servers'
 * cached tokens live in separate keychain entries. A revocation or
 * setup rerun on one server MUST NOT touch the other's cache.
 *
 * Platform backends:
 *   macOS    → Keychain
 *   Linux    → Secret Service / GNOME keyring
 *   Windows  → Credential Manager
 */

import { Entry } from "@napi-rs/keyring";

const SERVICE = "juvantlabs-m365-mail-mcp-server";

export interface TokenStore {
  load(): string | null;
  save(serialized: string): void;
  clear(): void;
}

export function getTokenStore(tenantId: string): TokenStore {
  const entry = new Entry(SERVICE, `tenant:${tenantId}`);
  return {
    load(): string | null {
      try {
        return entry.getPassword();
      } catch {
        // No entry yet — first run before setup, or it was cleared.
        return null;
      }
    },
    save(serialized: string): void {
      entry.setPassword(serialized);
    },
    clear(): void {
      try {
        entry.deletePassword();
      } catch {
        // already absent
      }
    },
  };
}
