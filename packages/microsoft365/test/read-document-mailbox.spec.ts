// read_document takes a raw Graph path, so without its own check it would read attachments
// from any mailbox the token can open, whatever MS365_ALLOWED_MAILBOXES says. These tests pin
// that it applies the mail tools' mailbox check to /users/... paths, in every spelling, before
// any request, and leaves /me, drive and site paths alone.

import { Some } from "functype"
import { Right } from "functype/either"
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"

vi.mock("../src/client/graph-client", () => ({ getGraphClient: vi.fn() }))
vi.mock("../src/auth", () => ({ getAccessToken: vi.fn() }))
vi.mock("../src/auth/auth-manager", () => ({ getAuthMode: vi.fn(), getAccessToken: vi.fn() }))

import { getAccessToken } from "../src/auth"
import { getAccessToken as getManagerToken, getAuthMode } from "../src/auth/auth-manager"
import { getGraphClient } from "../src/client/graph-client"
import { clearOwnAddressCache } from "../src/mail/mailbox"
import { mailboxOfUrl } from "../src/mail/mailbox-path"
import { readDocument } from "../src/tools/read-document-tools"

const ATTACHMENT = "messages/msg-1/attachments/att-1/$value"

const mockClient = { graphQuery: vi.fn(), request: vi.fn() }
const fetchSpy = vi.fn()

beforeEach(() => {
  vi.clearAllMocks()
  clearOwnAddressCache()
  vi.stubEnv("MS365_ALLOWED_MAILBOXES", "bel@example.com")
  vi.mocked(getGraphClient).mockReturnValue(Some(mockClient) as never)
  vi.mocked(getAccessToken).mockResolvedValue(Right("token-abc") as never)
  vi.mocked(getManagerToken).mockResolvedValue(Right("token-abc") as never)
  vi.mocked(getAuthMode).mockReturnValue(Some("interactive"))
  mockClient.request.mockResolvedValue(Right({ mail: "jake@example.com", userPrincipalName: "jake@example.com" }))
  mockClient.graphQuery.mockResolvedValue(Right({ id: "att-1", name: "a.txt", contentType: "text/plain", size: 5 }))
  fetchSpy.mockResolvedValue({
    ok: true,
    headers: new Headers({ "content-type": "text/plain" }),
    arrayBuffer: () => Promise.resolve(new TextEncoder().encode("hello").buffer),
  })
  vi.stubGlobal("fetch", fetchSpy)
})

afterEach(() => {
  vi.unstubAllGlobals()
  vi.unstubAllEnvs()
})

describe("read_document mailbox check", () => {
  it.each([`/users/bel%40example.com/${ATTACHMENT}`, `/users/Bel@Example.com/${ATTACHMENT}`])(
    "reads an attachment in an allowed mailbox: %s",
    async (path) => {
      const result = await readDocument({ path })

      expect(result.isRight()).toBe(true)
      expect(result.value).toContain("hello")
    },
  )

  it("reads an attachment in your own mailbox addressed by address, without listing it", async () => {
    const result = await readDocument({ path: `/users/jake%40example.com/${ATTACHMENT}` })

    expect(result.isRight()).toBe(true)
  })

  it.each([
    `/users/other@example.com/${ATTACHMENT}`,
    `/users/other%40example.com/${ATTACHMENT}`,
    `/users('other@example.com')/${ATTACHMENT}`,
    `/users('other%40example.com')/${ATTACHMENT}`,
    `/Users/other@example.com/${ATTACHMENT}`,
    `/users%2Fother@example.com/${ATTACHMENT}`,
    `/me/../users/other@example.com/${ATTACHMENT}`,
    `/../beta/users/other@example.com/${ATTACHMENT}`,
    `/users/other@example.com/mailFolders/inbox/${ATTACHMENT}`,
    `/users//${ATTACHMENT}`,
    `/users/other@example.com/events/e1/attachments/att-1/$value`,
    `/users(userPrincipalName='other@example.com')/${ATTACHMENT}`,
    `/users%2Fother%40example.com/messages/x%zz/attachments/att-1/$value`,
  ])("refuses a mailbox that is not allowed, before any request: %s", async (path) => {
    const result = await readDocument({ path })

    expect(result.isLeft()).toBe(true)
    const message = (result.value as Error).message
    expect(message).not.toContain("bel@example.com")
    expect(mockClient.graphQuery).not.toHaveBeenCalled()
    expect(fetchSpy).not.toHaveBeenCalled()
  })

  it("uses the mail tools' non-enumerating refusal", async () => {
    const result = await readDocument({ path: `/users('other%40example.com')/${ATTACHMENT}` })

    expect((result.value as Error).message).toBe(
      `Mailbox "other@example.com" is not permitted by this server's configuration (MS365_ALLOWED_MAILBOXES).`,
    )
  })

  it.each([
    `/me/${ATTACHMENT}`,
    "/me/drive/items/1/content",
    "/drives/d1/items/1/content",
    "/sites/s1/drive/items/1/content",
    "/users/other@example.com/drive/items/1/content",
    "/users/other@example.com/drives/d1/items/1/content",
  ])("leaves a non-mailbox path alone: %s", async (path) => {
    if (path.endsWith("/content"))
      mockClient.graphQuery.mockResolvedValue(
        Right({ id: "1", name: "a.txt", size: 5, file: { mimeType: "text/plain" } }),
      )

    const result = await readDocument({ path })

    expect(result.isRight()).toBe(true)
    expect(mockClient.request).not.toHaveBeenCalled()
  })
})

describe("mailboxOfUrl", () => {
  it.each([
    [`/users/bel%40example.com/${ATTACHMENT}`, "bel@example.com"],
    [`/users('bel%40example.com')/${ATTACHMENT}`, "bel@example.com"],
    [`/users/bel%zz/${ATTACHMENT}`, "bel%zz"],
    [`/me/${ATTACHMENT}`, undefined],
    ["/users/bel@example.com/drive/items/1/content", undefined],
  ])("%s -> %s", (path, expected) => {
    expect(mailboxOfUrl(`https://graph.microsoft.com/v1.0${path}`)).toBe(expected)
  })
})
