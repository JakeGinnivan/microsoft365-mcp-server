import { mkdtempSync, rmSync } from "node:fs"
import { tmpdir } from "node:os"
import { join } from "node:path"

import { DeviceCodeCredential, InteractiveBrowserCredential, type TokenCredential } from "@azure/identity"
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"

import { createCredential } from "../src/auth/auth-modes"

const config = { mode: "interactive" as const, tenantId: "common", clientId: "11111111-1111-1111-1111-111111111111" }
const token = { token: "test-token", expiresOnTimestamp: 1_000 }

// Keep the persisted token cache and authentication record out of the real config directory.
const cacheDir = { path: "" }

beforeEach(() => {
  cacheDir.path = mkdtempSync(join(tmpdir(), "ms365-auth-modes-"))
  vi.stubEnv("MS365_TOKEN_CACHE_PATH", cacheDir.path)
})

afterEach(() => {
  vi.restoreAllMocks()
  vi.unstubAllEnvs()
  rmSync(cacheDir.path, { recursive: true, force: true })
})

describe("interactive authentication mode", () => {
  it("uses the browser by default", async () => {
    const browserToken = vi.spyOn(InteractiveBrowserCredential.prototype, "getToken").mockResolvedValue(token)
    const result = createCredential(config)
    expect(result.isRight()).toBe(true)

    // With the token cache on (the default) the browser credential is wrapped so the
    // authentication record can be saved after sign-in; it still does the signing in.
    await expect((result.value as TokenCredential).getToken("scope")).resolves.toBe(token)
    expect(browserToken).toHaveBeenCalledTimes(1)
  })

  it("hands back the bare browser credential when the token cache is off", () => {
    vi.stubEnv("MS365_TOKEN_CACHE", "false")
    const result = createCredential(config)
    expect(result.isRight()).toBe(true)
    expect(result.value).toBeInstanceOf(InteractiveBrowserCredential)
  })

  it("uses one device-code credential for repeated token requests when explicitly enabled", async () => {
    const deviceToken = vi.spyOn(DeviceCodeCredential.prototype, "getToken").mockResolvedValue(token)
    const browserToken = vi.spyOn(InteractiveBrowserCredential.prototype, "getToken")
    const result = createCredential({ ...config, useDeviceCode: true })
    expect(result.isRight()).toBe(true)
    const credential = result.value as TokenCredential
    expect(credential).toBeInstanceOf(DeviceCodeCredential)

    const scopes = ["https://graph.microsoft.com/.default"]
    await expect(credential.getToken(scopes)).resolves.toBe(token)
    await expect(credential.getToken(scopes)).resolves.toBe(token)
    expect(deviceToken).toHaveBeenCalledTimes(2)
    expect(deviceToken.mock.contexts).toEqual([credential, credential])
    expect(browserToken).not.toHaveBeenCalled()
  })

  it("propagates browser errors without starting another authentication flow", async () => {
    const error = new Error("Unable to open browser")
    vi.spyOn(InteractiveBrowserCredential.prototype, "getToken").mockRejectedValue(error)
    const deviceToken = vi.spyOn(DeviceCodeCredential.prototype, "getToken").mockResolvedValue(token)
    const result = createCredential(config)
    const credential = result.value as TokenCredential

    await expect(credential.getToken("scope")).rejects.toBe(error)
    expect(deviceToken).not.toHaveBeenCalled()
  })
})
