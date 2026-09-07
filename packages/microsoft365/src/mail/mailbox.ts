// Which mailbox a mail request addresses.
//
// Every mail path is either rooted at /me (the signed-in user's own mailbox) or at
// /users/{address} (someone else's, reached by delegation or by an app-only identity).
// Callers pass an address; this module turns it into a path prefix and decides whether
// the deployment is allowed to touch that mailbox at all.
//
// Two distinct boundaries are at work, and they are not redundant:
//
//   Azure          an ApplicationAccessPolicy (app-only) or a delegated grant caps what
//                  the credential could ever reach. This is the security boundary.
//   MS365_ALLOWED_MAILBOXES
//                  narrows one deployment further, within that cap. This is what lets a
//                  single app registration serve several agents with different reach —
//                  a triage bot limited to one mailbox, a household assistant given two —
//                  without minting a registration per agent.
//
// The env var is configuration, not security: it fails closed and gives a clear error
// instead of a Graph 403, but it is Azure that actually enforces access.

import { UserError } from "fastmcp"
import { type Either, Left, Right } from "functype/either"

import { getAuthMode } from "../auth/auth-manager"
import type { AuthMode } from "../types"

/** Graph path prefix for a mailbox: "/me" for the signed-in user, else "/users/{address}". */
export type MailboxScope = {
  /** Undefined means the signed-in user's own mailbox. */
  readonly mailbox?: string
  readonly prefix: string
}

export const OWN_MAILBOX: MailboxScope = { prefix: "/me" }

// Addresses are compared case-insensitively — Exchange treats them that way, and an
// allowlist that rejected "Bel@..." against "bel@..." would be a confusing failure.
const normalize = (address: string): string => address.trim().toLowerCase()

/**
 * The mailboxes this deployment may address, from MS365_ALLOWED_MAILBOXES
 * (comma-separated). Unset or empty means no other mailbox is reachable: the server
 * can still work, but only against /me.
 */
export const resolveAllowedMailboxes = (
  raw: string | undefined = process.env.MS365_ALLOWED_MAILBOXES,
): ReadonlySet<string> =>
  new Set(
    (raw ?? "")
      .split(",")
      .map(normalize)
      .filter((entry) => entry.length > 0),
  )

// Client-secret and certificate tokens belong to an application, not a user, so Graph has
// no /me for them. client-token is left out: the token handed in may be either kind.
const APP_ONLY_MODES: ReadonlySet<AuthMode> = new Set(["client-secret", "certificate"])

export const isAppOnlyAuth = (): boolean =>
  getAuthMode().fold(
    () => false,
    (mode) => APP_ONLY_MODES.has(mode),
  )

const requiresMailbox = () =>
  new UserError(
    "This server signs in as an application (client secret or certificate), which has no mailbox of its own. " +
      "Pass mailbox with the address of the mailbox to act on.",
  )

/**
 * Resolve a caller-supplied mailbox into a scope, refusing anything the deployment has
 * not been configured to reach.
 *
 * Omitting the mailbox yields /me, so every existing delegated caller keeps its behaviour
 * and no configuration is needed to stay on the current path. App-only deployments have
 * no /me, so there omitting it is refused here instead of surfacing as a Graph error.
 */
export const resolveMailboxScope = (
  mailbox: string | undefined,
  allowed: ReadonlySet<string> = resolveAllowedMailboxes(),
  appOnly: boolean = isAppOnlyAuth(),
): Either<UserError, MailboxScope> => {
  if (mailbox === undefined) return appOnly ? Left(requiresMailbox()) : Right(OWN_MAILBOX)

  // Blank is refused rather than read as "my own mailbox": a caller that meant to name a
  // mailbox and sent an empty value would otherwise act on the wrong one without noticing.
  const normalized = normalize(mailbox)
  if (normalized.length === 0)
    return Left(
      appOnly
        ? requiresMailbox()
        : new UserError("mailbox is empty. Omit mailbox for your own mailbox, or pass an email address."),
    )

  if (!allowed.has(normalized)) {
    // The configured addresses are not listed: on a shared deployment, anyone who
    // mistyped an address would otherwise learn every mailbox the server can reach.
    // Naming the setting is enough for an operator, who can read it.
    return Left(
      new UserError(`Mailbox "${mailbox}" is not permitted by this server's configuration (MS365_ALLOWED_MAILBOXES).`),
    )
  }

  // Encoded because it lands in a URL path segment; Graph accepts the address as-is,
  // but an unencoded "+" in an address would otherwise be read as a space.
  return Right({ mailbox: normalized, prefix: `/users/${encodeURIComponent(normalized)}` })
}
