// Which mailbox a raw Graph URL reaches, for the tools that take one (read_document and
// graph_query). The mail tools build their own paths from a resolved scope; these two take
// whatever the caller passes, so without this they would step around MS365_ALLOWED_MAILBOXES.

import { GRAPH_API_BASE } from "@sapientsai/ms-graph-core"

import { safeDecode } from "../utils/safe-decode"

// Directory reads about a user, not their mailbox: the profile itself, photo, org chart,
// group membership, presence, and OneDrive. Everything else under /users/{x}/ is treated as
// mailbox data (messages, mailFolders, mailboxSettings, events, calendars, contacts, todo,
// outlook, ...), including anything not named here, so a resource Graph adds later is
// refused until someone decides it is safe.
const DIRECTORY_RESOURCES: ReadonlySet<string> = new Set([
  "photo",
  "photos",
  "manager",
  "directreports",
  "memberof",
  "transitivememberof",
  "presence",
  "drive",
  "drives",
])

// users('x'), and the alternate-key forms users(id='x') / users(userPrincipalName='x').
const USERS_KEY_FORM = /^users\((?:[a-z]+=)?'(.*)'\)$/i
const USERS_ANY_KEY = /^users\(/i

// Segment names are compared the way graph_query's send check compares them: case-folded,
// NFKC-normalised and trimmed, with a trailing "(...)" (a key or function call) dropped.
const fold = (segment: string): string =>
  segment
    .trim()
    .normalize("NFKC")
    .toLowerCase()
    .replace(/\(.*\)$/, "")

type UsersHit = { readonly address: string; readonly resource: string | undefined }

const usersSegmentAt = (segments: ReadonlyArray<string>, at: number): UsersHit | undefined => {
  const segment = (segments.at(at) ?? "").trim().normalize("NFKC")
  const keyed = USERS_KEY_FORM.exec(segment)
  if (keyed) return { address: keyed[1]!, resource: segments.at(at + 1) }
  // A key form we do not parse still names some user; returning the whole segment as the
  // address makes it fail the mailbox check rather than slip past it.
  if (USERS_ANY_KEY.test(segment)) return { address: segment, resource: segments.at(at + 1) }
  if (segment.toLowerCase() === "users" && at + 1 < segments.length)
    return { address: segments[at + 1]!, resource: segments.at(at + 2) }
  return undefined
}

/**
 * The mailbox address a Graph URL reaches under /users/{address}/... or /users('{address}')/...,
 * or undefined when it reaches none: /me, /drives, /sites, a directory read about a user.
 *
 * `url` is a full URL, or a path relative to `base` (the version root, as Graph resolves the
 * relative URLs in a $batch). It is read as fetch will send it: the URL parser resolves "."
 * and ".." segments, so "/me/../users/x/messages" reaches /users/x and is caught as such. An
 * encoded "/" still separates segments, and each segment is decoded on its own, so a stray
 * "%" in one cannot leave the others undecoded. The first two segments are checked because
 * ".." can climb out of the version root into another version.
 */
export const mailboxOfUrl = (url: string, base: string = `${GRAPH_API_BASE}/v1.0/`): string | undefined => {
  const pathname = URL.canParse(url, base) ? new URL(url, base).pathname : url.split(/[?#]/)[0]!
  const segments = pathname
    .split(/\/|%2f/i)
    .slice(1)
    .map(safeDecode)
  const hit = [0, 1].map((at) => usersSegmentAt(segments, at)).find((found) => found !== undefined)
  if (hit === undefined) return undefined
  const resource = fold(hit.resource ?? "")
  return resource === "" || DIRECTORY_RESOURCES.has(resource) ? undefined : hit.address
}
