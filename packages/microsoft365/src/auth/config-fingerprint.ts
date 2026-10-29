import { createHash } from "node:crypto"

// The interactive modes request `.default`, meaning "whatever this app registration has
// been consented". That string never changes, so MSAL's cache matches before and after
// a permission is granted in Azure, and keeps serving the access token minted under the
// old consent until it expires — up to an hour, and across restarts now that the cache
// is on disk. The usual moment that bites is a configuration change made alongside an
// Azure one: MS365_ALLOWED_MAILBOXES gains a mailbox, Mail.ReadWrite.Shared is granted,
// and Graph still answers 403.
//
// So the server's configuration is hashed, and when the hash differs from the one the
// cache was last used under, the cached access token is dropped. The refresh token is
// kept: the next call redeems it silently for a token under the current consent, with no
// sign-in. Every MS365_* and GRAPH_* variable counts, not a curated list, so a setting
// added later is covered without anyone remembering to add it here. A change that turns
// out not to matter costs one silent refresh.

// Secrets never reach the hash, even one-way: the fingerprint is written to disk, and
// none of these change what an interactive token can do.
const SECRET = /SECRET|PASSWORD|_KEY$|ACCESS_TOKEN|UPLOAD_TOKEN/

const CONFIG_PREFIXES = ["MS365_", "GRAPH_"]

export const configFingerprint = (
  auth: Readonly<Record<string, unknown>>,
  env: NodeJS.ProcessEnv = process.env,
): string => {
  const settings = Object.entries(env)
    .filter(
      ([name, value]) => value !== undefined && CONFIG_PREFIXES.some((p) => name.startsWith(p)) && !SECRET.test(name),
    )
    .sort(([a], [b]) => a.localeCompare(b))
  const authInputs = Object.entries(auth)
    .filter(([, value]) => value !== undefined)
    .sort(([a], [b]) => a.localeCompare(b))
  return createHash("sha256")
    .update(JSON.stringify({ auth: authInputs, settings }))
    .digest("hex")
}
