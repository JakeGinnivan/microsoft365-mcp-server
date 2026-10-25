// Naming your own address should work without listing it, since it is the same mailbox as
// /me. These tests pin when that shortcut applies and, more importantly, that every way the
// lookup can fail or be inapplicable leaves the caller to the allowlist instead of opening up.

import { None, Some } from "functype"
import { Left, Right } from "functype/either"
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"

vi.mock("../src/client/graph-client", () => ({ getGraphClient: vi.fn() }))
vi.mock("../src/auth/auth-manager", () => ({ getAuthMode: vi.fn(), getAccessToken: vi.fn() }))

import { getAccessToken, getAuthMode } from "../src/auth/auth-manager"
import { getGraphClient } from "../src/client/graph-client"
import { clearOwnAddressCache, ownAddressesOf, resolveMailbox } from "../src/mail/mailbox"
import { listMessages } from "../src/tools/mail-tools"
import { clearMessageRefs } from "../src/utils/message-refs"

const ME = {
  mail: "jake@example.com",
  userPrincipalName: "jake@corp.example.com",
  proxyAddresses: ["SMTP:jake@example.com", "smtp:jg@example.com", "SIP:jake@example.com", "X500:/o=ExchangeLabs"],
}

const client = { request: vi.fn() }
const mockClient = { ...client, listMessages: vi.fn() }

beforeEach(() => {
  vi.clearAllMocks()
  clearOwnAddressCache()
  clearMessageRefs()
  vi.stubEnv("MS365_ALLOWED_MAILBOXES", "")
  vi.mocked(getAuthMode).mockReturnValue(Some("interactive"))
  vi.mocked(getAccessToken).mockResolvedValue(Right("token-1"))
  client.request.mockResolvedValue(Right(ME))
  mockClient.request = client.request
  mockClient.listMessages.mockResolvedValue(Right({ value: [] }))
  vi.mocked(getGraphClient).mockReturnValue(Some(mockClient as never))
})

afterEach(() => {
  vi.unstubAllEnvs()
})

const message = (result: { value: unknown }) => (result.value as { message: string }).message

describe("ownAddressesOf", () => {
  it("takes mail, UPN and SMTP aliases, lowercased, and ignores other proxy types", () => {
    expect([...ownAddressesOf(ME)].sort()).toEqual(["jake@corp.example.com", "jake@example.com", "jg@example.com"])
  })
})

describe("naming your own mailbox", () => {
  it.each(["jake@example.com", "Jake@Corp.Example.com", "JG@EXAMPLE.COM"])(
    "routes %s to /me with an empty allowlist",
    async (mailbox) => {
      const scope = await resolveMailbox(mailbox, client)

      expect(scope.orThrow()).toEqual({ prefix: "/me" })
      expect(client.request).toHaveBeenCalledWith("GET", "/me", {
        odataParams: { $select: ["mail", "userPrincipalName", "proxyAddresses"] },
      })
    },
  )

  it("routes a tool call naming your own alias to /me", async () => {
    const result = await listMessages({ mailbox: "jg@example.com" })

    expect(result.isRight()).toBe(true)
    expect(mockClient.listMessages.mock.calls.at(-1)?.at(-1)).toBe("/me")
  })

  it("works in client-token mode", async () => {
    vi.mocked(getAuthMode).mockReturnValue(Some("client-token"))
    expect((await resolveMailbox("jake@example.com", client)).orThrow().prefix).toBe("/me")
  })

  it("still refuses someone else's unlisted address", async () => {
    const scope = await resolveMailbox("partner@example.com", client)

    expect(scope.isLeft()).toBe(true)
    expect(message(scope)).toContain("not permitted by this server's configuration")
  })

  it("does not look anything up when mailbox is omitted", async () => {
    expect((await resolveMailbox(undefined, client)).orThrow().prefix).toBe("/me")
    expect(client.request).not.toHaveBeenCalled()
  })
})

describe("own-address cache", () => {
  it("looks the identity up once per token", async () => {
    await resolveMailbox("jake@example.com", client)
    await resolveMailbox("jg@example.com", client)

    expect(client.request).toHaveBeenCalledOnce()
  })

  // set_access_token can hand a client-token deployment someone else's token.
  it("looks it up again when the token changes", async () => {
    vi.mocked(getAuthMode).mockReturnValue(Some("client-token"))
    await resolveMailbox("jake@example.com", client)

    vi.mocked(getAccessToken).mockResolvedValue(Right("token-2"))
    client.request.mockResolvedValue(Right({ mail: "bel@example.com", userPrincipalName: "bel@example.com" }))

    expect((await resolveMailbox("jake@example.com", client)).isLeft()).toBe(true)
    expect((await resolveMailbox("bel@example.com", client)).orThrow().prefix).toBe("/me")
    expect(client.request).toHaveBeenCalledTimes(2)
  })
})

describe("falling back to the allowlist", () => {
  it("refuses an unlisted address when the /me lookup fails", async () => {
    client.request.mockResolvedValue(Left({ type: "api", message: "boom" }))

    expect((await resolveMailbox("jake@example.com", client)).isLeft()).toBe(true)
  })

  it("routes a listed address to /users when the /me lookup fails", async () => {
    vi.stubEnv("MS365_ALLOWED_MAILBOXES", "jake@example.com")
    client.request.mockResolvedValue(Left({ type: "api", message: "boom" }))

    expect((await resolveMailbox("jake@example.com", client)).orThrow().prefix).toBe("/users/jake%40example.com")
  })

  it("does not cache a failed lookup", async () => {
    client.request.mockResolvedValueOnce(Left({ type: "api", message: "boom" }))
    expect((await resolveMailbox("jake@example.com", client)).isLeft()).toBe(true)
    expect((await resolveMailbox("jake@example.com", client)).orThrow().prefix).toBe("/me")
  })

  it("refuses when no token can be had", async () => {
    vi.mocked(getAccessToken).mockResolvedValue(Left({ type: "token", message: "no token" }))

    expect((await resolveMailbox("jake@example.com", client)).isLeft()).toBe(true)
    expect(client.request).not.toHaveBeenCalled()
  })

  it("refuses when auth is not initialised", async () => {
    vi.mocked(getAuthMode).mockReturnValue(None())

    expect((await resolveMailbox("jake@example.com", client)).isLeft()).toBe(true)
    expect(client.request).not.toHaveBeenCalled()
  })
})

describe("modes without the shortcut", () => {
  it.each(["client-secret", "certificate"] as const)("leaves %s app-only mode unchanged", async (mode) => {
    vi.mocked(getAuthMode).mockReturnValue(Some(mode))

    const scope = await resolveMailbox("jake@example.com", client)

    expect(scope.isLeft()).toBe(true)
    expect(message(scope)).toContain("not permitted")
    expect(client.request).not.toHaveBeenCalled()
  })

  it("routes a listed address to /users in app-only mode, as before", async () => {
    vi.mocked(getAuthMode).mockReturnValue(Some("client-secret"))
    vi.stubEnv("MS365_ALLOWED_MAILBOXES", "jake@example.com")

    expect((await resolveMailbox("jake@example.com", client)).orThrow().prefix).toBe("/users/jake%40example.com")
  })
})
