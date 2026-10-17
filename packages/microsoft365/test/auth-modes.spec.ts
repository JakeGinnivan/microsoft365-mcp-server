import { DeviceCodeCredential, InteractiveBrowserCredential, type TokenCredential } from "@azure/identity"
import { afterEach, describe, expect, it, vi } from "vitest"

import { createCredential } from "../src/auth/auth-modes"

const config = { mode: "interactive" as const, tenantId: "common", clientId: "11111111-1111-1111-1111-111111111111" }
const token = { token: "test-token", expiresOnTimestamp: 1_000 }

afterEach(() => vi.restoreAllMocks())

describe("interactive authentication mode", () => {
  it("uses the browser by default", () => {
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
