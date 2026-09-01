import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"

import { createCredential, resetTokenCachePersistenceForTests } from "../src/auth/auth-modes"
import type { AuthConfig } from "../src/types"

const interactive = (): AuthConfig => ({ mode: "interactive", tenantId: "tenant", clientId: "client" })

const saved = process.env.MS365_TOKEN_CACHE

beforeEach(() => {
  resetTokenCachePersistenceForTests()
  delete process.env.MS365_TOKEN_CACHE
  // restoreAllMocks first: spyOn an already-spied console.error returns the same spy,
  // so without this the previous test's calls leak into the next one's assertions.
  vi.restoreAllMocks()
  vi.spyOn(console, "error").mockImplementation(() => {})
})

afterEach(() => {
  if (saved === undefined) delete process.env.MS365_TOKEN_CACHE
  else process.env.MS365_TOKEN_CACHE = saved
})

describe("interactive token cache persistence", () => {
  // The whole point of the optional dependency: a host that cannot build keytar must still
  // get a working credential, just without persistence across restarts.
  it("should still create a credential when the persistence plugin cannot load", () => {
    const result = createCredential(interactive())
    expect(result.isRight()).toBe(true)
  })

  it("should not attempt registration when MS365_TOKEN_CACHE=false", () => {
    process.env.MS365_TOKEN_CACHE = "false"
    const spy = vi.spyOn(console, "error")
    const result = createCredential(interactive())
    expect(result.isRight()).toBe(true)
    // The "unavailable" warning is only emitted by a registration attempt.
    expect(spy).not.toHaveBeenCalled()
  })

  it("should only attempt plugin registration once across credentials", () => {
    const spy = vi.spyOn(console, "error")
    createCredential(interactive())
    const afterFirst = spy.mock.calls.length
    createCredential(interactive())
    createCredential(interactive())
    expect(spy.mock.calls.length).toBe(afterFirst)
  })

  it("should still require a client ID", () => {
    const result = createCredential({ mode: "interactive", tenantId: "tenant", clientId: "" })
    expect(result.isLeft()).toBe(true)
    expect((result.value as { message: string }).message).toContain("MS365_CLIENT_ID")
  })
})
