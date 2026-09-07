// The decorator around token acquisition: it must re-acquire past the MSAL cache when a
// token predates a permission change, and must not loop when the permission genuinely is
// not granted.

import type { AccessToken, TokenCredential } from "@azure/identity"
import { beforeEach, describe, expect, it, vi } from "vitest"

import { revalidateScopes } from "../src/auth/auth-modes"

const REQUIRED = ["Mail.ReadWrite.Shared"]

const accessToken = (token: string): AccessToken => ({ token, expiresOnTimestamp: Date.now() + 3_600_000 })

// Stands in for the JWT reader: the token string is its own scope list, so these tests
// exercise the decorator's control flow rather than jwt.decode.
const readScopes = (token: string): ReadonlyArray<string> => (token.length === 0 ? [] : token.split(" "))

let consoleError: ReturnType<typeof vi.spyOn>

beforeEach(() => {
  // Restored between tests: a spy that survives would carry the previous test's calls
  // into the "reported only once" count and make it pass or fail for the wrong reason.
  vi.restoreAllMocks()
  consoleError = vi.spyOn(console, "error").mockImplementation(() => {})
})

describe("revalidateScopes", () => {
  it("passes the credential through untouched when nothing is required", async () => {
    const inner = { getToken: vi.fn(async () => accessToken("Mail.ReadWrite")) }
    const credential = revalidateScopes(inner as unknown as TokenCredential, [], readScopes)

    expect(credential).toBe(inner)
  })

  it("returns the token unchanged when every required scope is present", async () => {
    const inner = { getToken: vi.fn(async () => accessToken("Mail.ReadWrite Mail.ReadWrite.Shared")) }
    const credential = revalidateScopes(inner as unknown as TokenCredential, REQUIRED, readScopes)

    const token = await credential.getToken("scope")

    expect(token?.token).toBe("Mail.ReadWrite Mail.ReadWrite.Shared")
    expect(inner.getToken).toHaveBeenCalledTimes(1)
  })

  // The point of the change: the first call is served from the MSAL cache and predates the
  // consent, the second bypasses it and carries the new permission.
  it("re-acquires past the cache when the cached token predates a permission change", async () => {
    const inner = {
      getToken: vi
        .fn()
        .mockResolvedValueOnce(accessToken("Mail.ReadWrite"))
        .mockResolvedValueOnce(accessToken("Mail.ReadWrite Mail.ReadWrite.Shared")),
    }
    const credential = revalidateScopes(inner as unknown as TokenCredential, REQUIRED, readScopes)

    const token = await credential.getToken("scope")

    expect(token?.token).toBe("Mail.ReadWrite Mail.ReadWrite.Shared")
    expect(inner.getToken).toHaveBeenCalledTimes(2)
  })

  it("sends a claims challenge on the retry, which is what defeats the cache", async () => {
    const inner = {
      getToken: vi
        .fn()
        .mockResolvedValueOnce(accessToken("Mail.ReadWrite"))
        .mockResolvedValueOnce(accessToken("Mail.ReadWrite Mail.ReadWrite.Shared")),
    }
    const credential = revalidateScopes(inner as unknown as TokenCredential, REQUIRED, readScopes)

    await credential.getToken("scope")

    const [, retryOptions] = inner.getToken.mock.calls[1] as [unknown, { claims?: string }]
    expect(retryOptions.claims).toBeTruthy()
  })

  // A missing consent cannot be fixed by asking again, so the retry happens once and then
  // the situation is reported. Retrying per call would be an infinite loop against Azure.
  it("retries once and then gives up when the permission is genuinely not granted", async () => {
    const inner = { getToken: vi.fn(async () => accessToken("Mail.ReadWrite")) }
    const credential = revalidateScopes(inner as unknown as TokenCredential, REQUIRED, readScopes)

    await credential.getToken("scope")

    expect(inner.getToken).toHaveBeenCalledTimes(2)
    expect(consoleError).toHaveBeenCalledWith(expect.stringContaining("Mail.ReadWrite.Shared"))
  })

  it("still returns a usable token when the scope is missing, rather than failing the call", async () => {
    const inner = { getToken: vi.fn(async () => accessToken("Mail.ReadWrite")) }
    const credential = revalidateScopes(inner as unknown as TokenCredential, REQUIRED, readScopes)

    const token = await credential.getToken("scope")

    // Graph refuses the calls that need the scope; everything else keeps working.
    expect(token?.token).toBe("Mail.ReadWrite")
  })

  // This runs on every token acquisition, so an unfixable shortfall must not print a line
  // per Graph call.
  it("reports an unfixable shortfall only once", async () => {
    const inner = { getToken: vi.fn(async () => accessToken("Mail.ReadWrite")) }
    const credential = revalidateScopes(inner as unknown as TokenCredential, REQUIRED, readScopes)

    await credential.getToken("scope")
    await credential.getToken("scope")
    await credential.getToken("scope")

    const drift = consoleError.mock.calls.filter(([line]) => String(line).includes("app registration"))
    expect(drift).toHaveLength(1)
  })

  it("leaves a null token alone", async () => {
    const inner = { getToken: vi.fn(async () => null) }
    const credential = revalidateScopes(inner as unknown as TokenCredential, REQUIRED, readScopes)

    expect(await credential.getToken("scope")).toBeNull()
    expect(inner.getToken).toHaveBeenCalledTimes(1)
  })

  // An opaque token reads as no scopes at all; treating that as drift would re-acquire on
  // every single call for deployments this cannot parse.
  it("does not re-acquire when the token's scopes cannot be read", async () => {
    const inner = { getToken: vi.fn(async () => accessToken("")) }
    const credential = revalidateScopes(inner as unknown as TokenCredential, REQUIRED, readScopes)

    await credential.getToken("scope")

    expect(inner.getToken).toHaveBeenCalledTimes(1)
  })
})
