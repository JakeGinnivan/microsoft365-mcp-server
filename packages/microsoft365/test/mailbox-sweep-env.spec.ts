import { afterEach, describe, expect, it, vi } from "vitest"

import { resolveAuthConfig } from "../src/auth/config"
import { runMailboxSweep } from "../src/scripts/mailbox-sweep"

const exit = () =>
  vi.spyOn(process, "exit").mockImplementation(() => {
    throw new Error("exit")
  })

afterEach(() => {
  vi.restoreAllMocks()
  vi.unstubAllEnvs()
})

describe("mailbox-sweep environment controls", () => {
  it("refuses --apply in read-only mode before opening rules, writing reports or authenticating", async () => {
    vi.stubEnv("MS365_READ_ONLY", "true")
    vi.spyOn(process, "argv", "get").mockReturnValue(["node", "sweep", "--apply", "--rules", "/does-not-exist"])
    const error = vi.spyOn(console, "error").mockImplementation(() => {})
    exit()
    await expect(runMailboxSweep()).rejects.toThrow("exit")
    expect(error).toHaveBeenCalledWith(expect.stringContaining("MS365_READ_ONLY"))
  })

  it("refuses OAuth proxy mode, which has no session in a standalone script", async () => {
    vi.stubEnv("MS365_READ_ONLY", "false")
    vi.stubEnv("MS365_AUTH_MODE", "oauth-proxy")
    vi.spyOn(process, "argv", "get").mockReturnValue(["node", "sweep"])
    const error = vi.spyOn(console, "error").mockImplementation(() => {})
    exit()
    await expect(runMailboxSweep()).rejects.toThrow("exit")
    expect(error).toHaveBeenCalledWith(expect.stringContaining("oauth-proxy"))
  })

  it("uses device-code configuration for an interactive sweep", () => {
    vi.stubEnv("MS365_AUTH_MODE", "interactive")
    vi.stubEnv("MS365_USE_DEVICE_CODE", "true")
    expect(resolveAuthConfig()).toMatchObject({ mode: "interactive", useDeviceCode: true })
  })

  it.each([
    ["client-secret", "MS365_CLIENT_SECRET", "clientSecret"],
    ["certificate", "MS365_CERT_PATH", "certPath"],
    ["client-token", "MS365_ACCESS_TOKEN", "accessToken"],
  ])("uses %s auth configuration for a sweep", (mode, variable, field) => {
    vi.stubEnv("MS365_AUTH_MODE", mode)
    vi.stubEnv(variable, "configured-value")
    expect(resolveAuthConfig()).toMatchObject({ mode, [field]: "configured-value" })
  })
})
