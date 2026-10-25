// In oauth-proxy mode one process serves many users, each request carrying its own token in
// async context. These tests drive the real auth manager and Graph client (only fetch is
// stubbed) to show the own-address lookup goes out with the caller's token and that one
// user's addresses are never found under another user's token.

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"

import { getAccessToken } from "../src/auth/auth-manager"
import { withToken } from "../src/auth/token-context"
import { initializeGraphClient } from "../src/client/graph-client"
import { clearOwnAddressCache, OWN_ADDRESS_CACHE_LIMIT, resolveMailbox } from "../src/mail/mailbox"

const IDENTITIES: Record<string, { mail: string; userPrincipalName: string }> = {
  "token-jake": { mail: "jake@example.com", userPrincipalName: "jake@example.com" },
  "token-bel": { mail: "bel@example.com", userPrincipalName: "bel@example.com" },
}

const json = (body: unknown, status = 200) =>
  new Response(JSON.stringify(body), { status, headers: { "Content-Type": "application/json" } })

// Answers /me as whichever user the bearer token belongs to.
const fetchSpy = vi.fn((_url: string, init?: RequestInit) => {
  const token = String((init?.headers as Record<string, string>).Authorization).replace("Bearer ", "")
  const identity = IDENTITIES[token] ?? { mail: `${token}@example.com`, userPrincipalName: `${token}@example.com` }
  return Promise.resolve(json(identity))
})

const client = () => initializeGraphClient({ getAccessToken })
const as = <T>(token: string, fn: () => Promise<T>) => withToken(token, fn)
const prefix = async (token: string, mailbox: string) =>
  (await as(token, () => resolveMailbox(mailbox, client()))).fold(
    () => "refused",
    (scope) => scope.prefix,
  )

beforeEach(() => {
  vi.clearAllMocks()
  clearOwnAddressCache()
  vi.stubEnv("MS365_ALLOWED_MAILBOXES", "")
  vi.stubEnv("MS365_GRAPH_MAX_RETRIES", "0")
  vi.stubGlobal("fetch", fetchSpy)
})

afterEach(() => {
  vi.unstubAllGlobals()
  vi.unstubAllEnvs()
})

describe("own address in oauth-proxy mode", () => {
  it("looks /me up with the caller's own token", async () => {
    expect(await prefix("token-jake", "jake@example.com")).toBe("/me")

    expect(fetchSpy).toHaveBeenCalledOnce()
    const [url, init] = fetchSpy.mock.calls[0]!
    expect(new URL(url).pathname).toBe("/v1.0/me")
    expect((init?.headers as Record<string, string>).Authorization).toBe("Bearer token-jake")
  })

  it("gives each token its own addresses and never another's", async () => {
    expect(await prefix("token-jake", "jake@example.com")).toBe("/me")
    expect(await prefix("token-bel", "bel@example.com")).toBe("/me")

    expect(await prefix("token-jake", "bel@example.com")).toBe("refused")
    expect(await prefix("token-bel", "jake@example.com")).toBe("refused")
    // Both answers came from the cache the second time round.
    expect(fetchSpy).toHaveBeenCalledTimes(2)
  })

  it("keeps concurrent requests for different users apart", async () => {
    const results = await Promise.all([
      prefix("token-jake", "jake@example.com"),
      prefix("token-bel", "jake@example.com"),
      prefix("token-bel", "bel@example.com"),
      prefix("token-jake", "bel@example.com"),
    ])

    expect(results).toEqual(["/me", "refused", "/me", "refused"])
  })

  it("evicts the oldest token once the cache is full", async () => {
    await prefix("token-jake", "jake@example.com")
    for (let i = 0; i < OWN_ADDRESS_CACHE_LIMIT; i++) await prefix(`t${i}`, "jake@example.com")
    expect(fetchSpy).toHaveBeenCalledTimes(OWN_ADDRESS_CACHE_LIMIT + 1)

    // The newest entry is still cached; the first one was evicted and is looked up again.
    await prefix(`t${OWN_ADDRESS_CACHE_LIMIT - 1}`, "jake@example.com")
    expect(fetchSpy).toHaveBeenCalledTimes(OWN_ADDRESS_CACHE_LIMIT + 1)
    expect(await prefix("token-jake", "jake@example.com")).toBe("/me")
    expect(fetchSpy).toHaveBeenCalledTimes(OWN_ADDRESS_CACHE_LIMIT + 2)
  })

  it("falls back to the allowlist when the /me lookup fails", async () => {
    fetchSpy.mockResolvedValueOnce(json({ error: { message: "boom" } }, 500))
    expect(await prefix("token-jake", "jake@example.com")).toBe("refused")

    vi.stubEnv("MS365_ALLOWED_MAILBOXES", "jake@example.com")
    fetchSpy.mockResolvedValueOnce(json({ error: { message: "boom" } }, 500))
    expect(await prefix("token-jake", "jake@example.com")).toBe("/users/jake%40example.com")
  })

  it("does nothing without a context token when auth is not initialised", async () => {
    const scope = await resolveMailbox("jake@example.com", client())

    expect(scope.isLeft()).toBe(true)
    expect(fetchSpy).not.toHaveBeenCalled()
  })
})
