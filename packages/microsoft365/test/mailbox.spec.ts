import { describe, expect, it } from "vitest"

import { OWN_MAILBOX, resolveAllowedMailboxes, resolveMailboxScope } from "../src/mail/mailbox"

const allowed = (raw: string | undefined) => resolveAllowedMailboxes(raw)

describe("resolveAllowedMailboxes", () => {
  it("is empty when unset", () => {
    expect(allowed(undefined).size).toBe(0)
  })

  it("is empty for a blank or comma-only value", () => {
    expect(allowed("").size).toBe(0)
    expect(allowed(" , , ").size).toBe(0)
  })

  it("trims and lowercases each entry", () => {
    expect([...allowed(" Bel@Example.com , household@example.com ")]).toEqual([
      "bel@example.com",
      "household@example.com",
    ])
  })
})

describe("resolveMailboxScope", () => {
  it("defaults to the signed-in user's mailbox when no mailbox is given", () => {
    const scope = resolveMailboxScope(undefined, allowed("bel@example.com"))
    expect(scope.isRight()).toBe(true)
    expect(scope.orThrow()).toEqual(OWN_MAILBOX)
    expect(scope.orThrow().prefix).toBe("/me")
  })

  // Existing deployments set nothing, so the default path must not require config.
  it("allows the own mailbox even when no allowlist is configured", () => {
    const scope = resolveMailboxScope(undefined, allowed(undefined))
    expect(scope.isRight()).toBe(true)
    expect(scope.orThrow().prefix).toBe("/me")
  })

  it("builds a /users/ prefix for an allowed mailbox", () => {
    const scope = resolveMailboxScope("bel@example.com", allowed("bel@example.com"))
    expect(scope.isRight()).toBe(true)
    expect(scope.orThrow().prefix).toBe("/users/bel%40example.com")
    expect(scope.orThrow().mailbox).toBe("bel@example.com")
  })

  it("matches the allowlist case-insensitively and ignores surrounding space", () => {
    const scope = resolveMailboxScope("  BEL@Example.com ", allowed("bel@example.com"))
    expect(scope.isRight()).toBe(true)
    expect(scope.orThrow().mailbox).toBe("bel@example.com")
  })

  // Fails closed: an unlisted mailbox is refused here rather than reaching Graph, so a
  // typo or a guessed address gets a legible answer instead of a 403.
  it("refuses a mailbox that is not on the allowlist", () => {
    const scope = resolveMailboxScope("someone@example.com", allowed("bel@example.com"))
    expect(scope.isLeft()).toBe(true)
    expect((scope.value as { message: string }).message).toContain("someone@example.com")
  })

  // A hosted server answers many callers; a refusal that listed the allowlist would
  // disclose every reachable mailbox to anyone who typed a wrong address.
  it("does not disclose the allowed mailboxes when refusing", () => {
    const scope = resolveMailboxScope("nope@example.com", allowed("bel@example.com,household@example.com"))
    const message = (scope.value as { message: string }).message
    expect(message).toContain("not permitted by this server's configuration")
    expect(message).not.toContain("bel@example.com")
    expect(message).not.toContain("household@example.com")
  })

  it("says how to configure one when the allowlist is empty", () => {
    const scope = resolveMailboxScope("bel@example.com", allowed(undefined))
    expect(scope.isLeft()).toBe(true)
    expect((scope.value as { message: string }).message).toContain("MS365_ALLOWED_MAILBOXES")
  })

  // An address is a path segment; "+" in particular is a valid local part and would
  // otherwise be read as a space once Graph decodes the URL.
  it("percent-encodes an address with a plus", () => {
    const scope = resolveMailboxScope("bel+bills@example.com", allowed("bel+bills@example.com"))
    expect(scope.orThrow().prefix).toBe("/users/bel%2Bbills%40example.com")
  })

  // A blank value is more likely a mailbox the caller failed to fill in than a request
  // for its own, so it is refused rather than quietly acting on /me.
  it.each(["", "   "])("refuses an empty mailbox %j and says to omit it instead", (mailbox) => {
    const scope = resolveMailboxScope(mailbox, allowed("bel@example.com"))
    expect(scope.isLeft()).toBe(true)
    expect((scope.value as { message: string }).message).toContain("Omit mailbox for your own mailbox")
  })
})

// Client-secret and certificate tokens have no /me. Without this, an omitted mailbox reached
// Graph as /me and came back as a raw Graph error instead of saying what to pass.
describe("resolveMailboxScope in app-only auth", () => {
  it("refuses an omitted mailbox and says to pass one", () => {
    const scope = resolveMailboxScope(undefined, allowed("bel@example.com"), true)
    expect(scope.isLeft()).toBe(true)
    expect((scope.value as { message: string }).message).toContain("Pass mailbox")
  })

  it("refuses a blank mailbox the same way, not with the own-mailbox hint", () => {
    const message = (resolveMailboxScope(" ", allowed("bel@example.com"), true).value as { message: string }).message
    expect(message).toContain("Pass mailbox")
    expect(message).not.toContain("your own mailbox")
  })

  it("resolves a named, allowed mailbox as usual", () => {
    const scope = resolveMailboxScope("bel@example.com", allowed("bel@example.com"), true)
    expect(scope.orThrow().prefix).toBe("/users/bel%40example.com")
  })

  // The allowlist still applies unchanged; app-only only removes the /me default.
  it("still refuses a mailbox that is not on the allowlist", () => {
    const scope = resolveMailboxScope("someone@example.com", allowed("bel@example.com"), true)
    expect((scope.value as { message: string }).message).toContain("not permitted")
  })
})
