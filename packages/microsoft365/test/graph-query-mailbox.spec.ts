// graph_query takes a raw Graph path, so it is the easiest way around MS365_ALLOWED_MAILBOXES.
// These tests pin that every /users/{x} mailbox path it would reach, in every spelling and
// inside a $batch, goes through the mail tools' mailbox check before anything is sent, while
// directory reads about a user and /me stay as they were.

import { Some } from "functype"
import { Right } from "functype/either"
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"

vi.mock("../src/client/graph-client", () => ({ getGraphClient: vi.fn() }))
vi.mock("../src/auth/auth-manager", () => ({ getAuthMode: vi.fn(), getAccessToken: vi.fn() }))

import { getAccessToken, getAuthMode } from "../src/auth/auth-manager"
import { getGraphClient } from "../src/client/graph-client"
import { clearOwnAddressCache } from "../src/mail/mailbox"
import { mailboxOfUrl } from "../src/mail/mailbox-path"
import { graphQuery, mailboxesReached } from "../src/tools/graph-query-tools"

const REFUSAL = `is not permitted by this server's configuration (MS365_ALLOWED_MAILBOXES).`

const mockClient = { graphQuery: vi.fn(), request: vi.fn() }

beforeEach(() => {
  vi.clearAllMocks()
  clearOwnAddressCache()
  vi.stubEnv("MS365_ALLOWED_MAILBOXES", "bel@example.com")
  vi.mocked(getGraphClient).mockReturnValue(Some(mockClient as never))
  vi.mocked(getAuthMode).mockReturnValue(Some("interactive"))
  vi.mocked(getAccessToken).mockResolvedValue(Right("token-jake"))
  mockClient.request.mockResolvedValue(Right({ mail: "jake@example.com", userPrincipalName: "jake@example.com" }))
  mockClient.graphQuery.mockResolvedValue(Right({ ok: true }))
})

afterEach(() => {
  vi.unstubAllEnvs()
})

const get = (path: string) => graphQuery({ method: "GET", path })

describe("graph_query mailbox check", () => {
  it.each([
    "/users/bel@example.com/messages?$top=5",
    "/users/bel%40example.com/mailFolders/inbox/messages",
    "/users('bel@example.com')/mailboxSettings",
  ])("allows a listed mailbox: %s", async (path) => {
    expect((await get(path)).isRight()).toBe(true)
    expect(mockClient.graphQuery).toHaveBeenCalledOnce()
  })

  it("allows your own address without listing it", async () => {
    expect((await get("/users/jake@example.com/messages")).isRight()).toBe(true)
    expect(mockClient.graphQuery).toHaveBeenCalledOnce()
  })

  it.each([
    "/users/other@example.com/messages",
    "/users/other%40example.com/messages?$select=id",
    "/users('other@example.com')/messages",
    "/users('other%40example.com')/events",
    "/users(userPrincipalName='other@example.com')/calendar/events",
    "/users(id='abc')/messages",
    "/Users/other@example.com/contacts",
    "/users%2Fother@example.com/todo/lists",
    "/me/../users/other@example.com/messages",
    "/../beta/users/other@example.com/messages",
    "/users/other@example.com/mailboxSettings",
    "/users/other@example.com/calendarView?startDateTime=2026-01-01&endDateTime=2026-02-01",
    "/users/other@example.com/inferenceClassification/overrides",
    "/users/other@example.com/somethingGraphAddsLater",
    "/users/abc-guid-1234/messages",
    // Graph reads an empty segment as nothing, so "//users/x" reaches /users/x.
    "//users/other@example.com/messages",
    "///users/other@example.com/messages",
    "/me/..//users/other@example.com/messages",
  ])("refuses a mailbox that is not allowed, before any request: %s", async (path) => {
    const result = await get(path)

    expect(result.isLeft()).toBe(true)
    expect((result.value as Error).message).toContain(REFUSAL)
    expect((result.value as Error).message).not.toContain("bel@example.com")
    expect(mockClient.graphQuery).not.toHaveBeenCalled()
  })

  it.each([
    "/users/other@example.com",
    "/users/other@example.com?$select=displayName,mail",
    "/users/other@example.com/photo/$value",
    "/users/other@example.com/photos('48x48')/$value",
    "/users/other@example.com/manager",
    "/users/other@example.com/directReports",
    "/users/other@example.com/memberOf",
    "/users/other@example.com/transitiveMemberOf",
    "/users/other@example.com/presence",
    "/users/other@example.com/drive/root/children",
    "/users/other@example.com/drives",
    "/users?$filter=startswith(mail,'o')",
    "/me/messages",
    "/me/mailFolders/inbox/messages",
    "/drives/d1/items/1",
    "/sites/s1/lists",
  ])("leaves a directory read or non-mailbox path alone: %s", async (path) => {
    expect((await get(path)).isRight()).toBe(true)
    expect(mockClient.graphQuery).toHaveBeenCalledOnce()
    expect(mockClient.request).not.toHaveBeenCalled()
  })

  it("refuses a whole $batch when one sub-request reaches a mailbox that is not allowed", async () => {
    const body = JSON.stringify({
      requests: [
        { id: "1", method: "GET", url: "/users/bel@example.com/messages" },
        { id: "2", method: "GET", url: "/me/messages" },
        { id: "3", method: "GET", url: "users('other%40example.com')/messages" },
      ],
    })

    const result = await graphQuery({ method: "POST", path: "/$batch", body })

    expect(result.isLeft()).toBe(true)
    expect((result.value as Error).message).toContain(REFUSAL)
    expect(mockClient.graphQuery).not.toHaveBeenCalled()
  })

  // A relative URL starting "//" would parse as protocol-relative, with "users" as the host.
  // Graph treats it as a path, so it has to be checked as one.
  it.each([
    "//users/other@example.com/messages",
    "\\\\users\\other@example.com\\messages",
    "///users/other@example.com/events",
  ])("refuses a $batch sub-request whose URL starts with repeated slashes: %s", async (url) => {
    const body = JSON.stringify({ requests: [{ id: "1", method: "GET", url }] })

    const result = await graphQuery({ method: "POST", path: "/$batch", body })

    expect(result.isLeft()).toBe(true)
    expect((result.value as Error).message).toContain(REFUSAL)
    expect(mockClient.graphQuery).not.toHaveBeenCalled()
  })

  it("sends a $batch whose sub-requests are all allowed", async () => {
    const body = JSON.stringify({
      requests: [
        { id: "1", method: "GET", url: "/users/bel@example.com/messages" },
        { id: "2", method: "GET", url: "/users/jake@example.com/events" },
        { id: "3", method: "GET", url: "/users/other@example.com/manager" },
      ],
    })

    expect((await graphQuery({ method: "POST", path: "/$batch", body })).isRight()).toBe(true)
    expect(mockClient.graphQuery).toHaveBeenCalledOnce()
  })
})

describe("mailboxesReached", () => {
  it("reads absolute and relative batch URLs, and the request's own", () => {
    const body = {
      Requests: [
        { URL: "https://graph.microsoft.com/v1.0/users/a@example.com/messages" },
        { url: "users/b@example.com/messages" },
        { url: "/me/messages" },
      ],
    }
    expect(mailboxesReached("https://graph.microsoft.com/v1.0/$batch", body)).toEqual([
      "a@example.com",
      "b@example.com",
    ])
  })
})

describe("mailboxOfUrl", () => {
  it.each([
    ["https://graph.microsoft.com/v1.0/users/a%40example.com/messages", "a@example.com"],
    ["https://graph.microsoft.com/beta/users('a@example.com')/messages?$top=1", "a@example.com"],
    ["/users/a@example.com/messages", "a@example.com"],
    ["users/a@example.com/events", "a@example.com"],
    ["https://graph.microsoft.com/v1.0/users/a@example.com/Manager", undefined],
    ["https://graph.microsoft.com/v1.0/me/messages", undefined],
    ["https://graph.microsoft.com/v1.0//users/a@example.com/messages", "a@example.com"],
    ["//users/a@example.com/messages", "a@example.com"],
    ["///users/a@example.com/messages", "a@example.com"],
    ["\\\\users\\a@example.com\\messages", "a@example.com"],
    ["//me/messages", undefined],
  ])("%s -> %s", (url, expected) => {
    expect(mailboxOfUrl(url)).toBe(expected)
  })
})
