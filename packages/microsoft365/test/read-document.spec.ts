import { Some } from "functype"
import { Left, Right } from "functype/either"
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"

vi.mock("../src/client/graph-client", () => ({
  getGraphClient: vi.fn(),
}))

vi.mock("../src/auth", () => ({
  getAccessToken: vi.fn(),
}))

import { getAccessToken } from "../src/auth"
import { getGraphClient } from "../src/client/graph-client"
import { readDocument } from "../src/tools/read-document-tools"

const PDF = "application/pdf"
const XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
const MB = 1024 * 1024

const mockClient = { graphQuery: vi.fn() }

/** Arrange a drive item for the metadata read. */
const givenItem = (item: { name?: string; size?: number; mimeType?: string }) => {
  mockClient.graphQuery.mockResolvedValue(
    Right({ id: "1", name: item.name ?? "doc.pdf", size: item.size, file: { mimeType: item.mimeType ?? PDF } }),
  )
}

const givenContent = (body: string, contentType = "text/plain") =>
  vi.fn().mockResolvedValue({
    ok: true,
    headers: new Headers({ "content-type": contentType }),
    arrayBuffer: () => Promise.resolve(new TextEncoder().encode(body).buffer),
  })

beforeEach(() => {
  vi.clearAllMocks()
  delete process.env.MS365_MAX_EXTRACT_BYTES
  vi.mocked(getGraphClient).mockReturnValue(Some(mockClient) as never)
  vi.mocked(getAccessToken).mockResolvedValue(Right("token-abc") as never)
})

afterEach(() => {
  vi.unstubAllGlobals()
  delete process.env.MS365_MAX_EXTRACT_BYTES
})

describe("readDocument — size caps", () => {
  // The whole point of the metadata-first ordering: an over-cap file must be rejected without ever
  // transferring the body. Asserting on the error message alone would pass even if the old
  // buffer-then-check ordering came back, so assert that fetch was never called.
  it("rejects an over-cap file from metadata, without fetching the body", async () => {
    givenItem({ name: "huge.xlsx", size: 30 * MB, mimeType: XLSX })
    const fetchSpy = givenContent("never read")
    vi.stubGlobal("fetch", fetchSpy)

    const result = await readDocument({ path: "/me/drive/items/1/content" })

    expect(result.isLeft()).toBe(true)
    expect((result.value as Error).message).toContain("extraction cap")
    expect(fetchSpy).not.toHaveBeenCalled()
  })

  it("allows a 60 MB PDF, which the old flat 10 MB cap would have rejected", async () => {
    givenItem({ name: "report.pdf", size: 60 * MB, mimeType: PDF })
    const fetchSpy = givenContent("plain text stand-in", "text/plain")
    vi.stubGlobal("fetch", fetchSpy)

    const result = await readDocument({ path: "/me/drive/items/1/content" })

    expect(result.isRight()).toBe(true)
    expect(fetchSpy).toHaveBeenCalledOnce()
  })

  it("applies per-format caps: 30 MB passes as PDF but fails as XLSX", async () => {
    vi.stubGlobal("fetch", givenContent("ok"))

    givenItem({ size: 30 * MB, mimeType: PDF })
    expect((await readDocument({ path: "/me/drive/items/1/content" })).isRight()).toBe(true)

    givenItem({ size: 30 * MB, mimeType: XLSX })
    expect((await readDocument({ path: "/me/drive/items/1/content" })).isLeft()).toBe(true)
  })

  it("MS365_MAX_EXTRACT_BYTES tightens every format", async () => {
    process.env.MS365_MAX_EXTRACT_BYTES = String(5 * MB)
    givenItem({ size: 6 * MB, mimeType: PDF })
    const fetchSpy = givenContent("never read")
    vi.stubGlobal("fetch", fetchSpy)

    const result = await readDocument({ path: "/me/drive/items/1/content" })

    expect(result.isLeft()).toBe(true)
    expect(fetchSpy).not.toHaveBeenCalled()
  })

  it("is a ceiling, never a raise — it cannot lift a format above its default", async () => {
    process.env.MS365_MAX_EXTRACT_BYTES = String(500 * MB)
    givenItem({ size: 30 * MB, mimeType: XLSX })
    vi.stubGlobal("fetch", givenContent("never read"))

    expect((await readDocument({ path: "/me/drive/items/1/content" })).isLeft()).toBe(true)
  })
})

describe("readDocument — request shape", () => {
  it("reads metadata from the item path, not the /content endpoint", async () => {
    givenItem({ size: 10 })
    vi.stubGlobal("fetch", givenContent("hi"))

    await readDocument({ path: "/drives/d1/items/i1/content" })

    expect(mockClient.graphQuery).toHaveBeenCalledWith(
      "GET",
      "/drives/d1/items/i1?$select=id,name,size,file",
      undefined,
      "v1.0",
    )
  })

  it("sends the resolved access token as a bearer", async () => {
    givenItem({ size: 10 })
    const fetchSpy = givenContent("hi")
    vi.stubGlobal("fetch", fetchSpy)

    await readDocument({ path: "/me/drive/items/1/content" })

    expect(fetchSpy).toHaveBeenCalledWith(expect.stringContaining("/v1.0/me/drive/items/1/content"), {
      headers: { Authorization: "Bearer token-abc" },
    })
  })

  it("surfaces an auth failure rather than sending 'Bearer [object Object]'", async () => {
    vi.mocked(getAccessToken).mockResolvedValue(Left({ type: "token", message: "no token" }) as never)
    const fetchSpy = givenContent("never read")
    vi.stubGlobal("fetch", fetchSpy)

    const result = await readDocument({ path: "/me/drive/items/1/content" })

    expect(result.isLeft()).toBe(true)
    expect((result.value as Error).message).toBe("no token")
    expect(fetchSpy).not.toHaveBeenCalled()
  })
})

describe("readDocument — output", () => {
  it("returns extracted text with the filename and size header", async () => {
    givenItem({ name: "notes.txt", size: 5, mimeType: "text/plain" })
    vi.stubGlobal("fetch", givenContent("hello", "text/plain"))

    const result = await readDocument({ path: "/me/drive/items/1/content" })

    expect(result.value as string).toContain("File: notes.txt")
    expect(result.value as string).toContain("hello")
  })

  it("truncates at max_chars with an explicit marker", async () => {
    givenItem({ name: "long.txt", size: 100, mimeType: "text/plain" })
    vi.stubGlobal("fetch", givenContent("x".repeat(5000), "text/plain"))

    const result = await readDocument({ path: "/me/drive/items/1/content", max_chars: 1000 })

    expect(result.value as string).toContain("[truncated at 1,000 chars")
    expect(result.value as string).toContain("full document is 5,000 chars")
  })

  // "unsupported" is the arm that means download_file is the right fallback, so the message says so.
  it("points at download_file for an unsupported content type", async () => {
    givenItem({ name: "photo.png", size: 10, mimeType: "image/png" })
    vi.stubGlobal("fetch", givenContent("\x89PNG", "image/png"))

    const result = await readDocument({ path: "/me/drive/items/1/content" })

    expect(result.isLeft()).toBe(true)
    expect((result.value as Error).message).toContain("Unsupported content type")
    expect((result.value as Error).message).toContain("download_file")
  })
})
