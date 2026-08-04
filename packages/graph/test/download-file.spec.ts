import type { GraphApiError, GraphRequest } from "@sapientsai/ms-graph-core"
import { Left, Right } from "functype/either"
import { describe, expect, it, vi } from "vitest"

import { resolveServerRuntimeConfig } from "../src/config"
import { buildDownloadFileTool } from "../src/tools/download-file"

const DOWNLOAD_URL = "https://tenant.sharepoint.com/pre-authed/blob?tempauth=xyz"

const graphReturning = (item: Record<string, unknown>): GraphRequest =>
  ({
    request: vi.fn(() => Promise.resolve(Right(item))),
    requestPaginated: vi.fn(),
  }) as unknown as GraphRequest

const graphFailing = (message: string): GraphRequest =>
  ({
    request: vi.fn(() => Promise.resolve(Left<GraphApiError, unknown>({ type: "api", message }))),
    requestPaginated: vi.fn(),
  }) as unknown as GraphRequest

const binaryResponse = (body: string, ok = true, status = 200) =>
  ({
    ok,
    status,
    statusText: ok ? "OK" : "Not Found",
    arrayBuffer: () => Promise.resolve(Buffer.from(body)),
  }) as Response

const item = (over: Record<string, unknown> = {}) => ({
  id: "01ABC",
  name: "contract.pdf",
  size: 1024,
  file: { mimeType: "application/pdf" },
  "@microsoft.graph.downloadUrl": DOWNLOAD_URL,
  ...over,
})

const args = { path: "/drives/d1/items/01ABC", api_version: "v1.0" as const, inline: true }

describe("download_file tool", () => {
  it("returns metadata and the pre-authenticated download URL", async () => {
    const tool = buildDownloadFileTool(
      graphReturning(item()),
      vi.fn(() => Promise.resolve(binaryResponse("PDF"))),
    )

    const out = await tool.execute(args)

    expect(out).toContain("Name: contract.pdf")
    expect(out).toContain("Size: 1.0 KB")
    expect(out).toContain(DOWNLOAD_URL)
  })

  it("inlines base64 under the 256 KB limit", async () => {
    const tool = buildDownloadFileTool(
      graphReturning(item()),
      vi.fn(() => Promise.resolve(binaryResponse("hello"))),
    )

    const out = await tool.execute(args)

    expect(out).toContain("## Content (base64)")
    expect(out).toContain(Buffer.from("hello").toString("base64"))
  })

  // The point of returning a URL rather than bytes: a large file must not be transferred at all.
  it("does not fetch the body for a file over the inline limit", async () => {
    const fetchImpl = vi.fn(() => Promise.resolve(binaryResponse("never read")))
    const tool = buildDownloadFileTool(graphReturning(item({ size: 5 * 1024 * 1024 })), fetchImpl)

    const out = await tool.execute(args)

    expect(fetchImpl).not.toHaveBeenCalled()
    expect(out).toContain(DOWNLOAD_URL)
    expect(out).not.toContain("## Content (base64)")
  })

  it("skips the body when inline is false, even for a small file", async () => {
    const fetchImpl = vi.fn(() => Promise.resolve(binaryResponse("small")))
    const tool = buildDownloadFileTool(graphReturning(item()), fetchImpl)

    await tool.execute({ ...args, inline: false })

    expect(fetchImpl).not.toHaveBeenCalled()
  })

  // downloadUrl is pre-authenticated; sending a bearer with it makes Graph reject the request.
  it("fetches the download URL without an Authorization header", async () => {
    const fetchImpl = vi.fn(() => Promise.resolve(binaryResponse("x")))
    const tool = buildDownloadFileTool(graphReturning(item()), fetchImpl)

    await tool.execute(args)

    expect(fetchImpl).toHaveBeenCalledWith(DOWNLOAD_URL, { method: "GET" })
  })

  it("tolerates a trailing /content on the path", async () => {
    const graph = graphReturning(item())
    const tool = buildDownloadFileTool(
      graph,
      vi.fn(() => Promise.resolve(binaryResponse("x"))),
    )

    await tool.execute({ ...args, path: "/drives/d1/items/01ABC/content" })

    expect(graph.request).toHaveBeenCalledWith("GET", "/drives/d1/items/01ABC", { version: "v1.0" })
  })

  it("degrades to the URL when the inline fetch fails", async () => {
    const tool = buildDownloadFileTool(
      graphReturning(item()),
      vi.fn(() => Promise.resolve(binaryResponse("", false, 404))),
    )

    const out = await tool.execute(args)

    expect(out).toContain("inline fetch failed: HTTP 404")
    expect(out).toContain(DOWNLOAD_URL)
  })

  it("reports a missing download URL rather than pretending it has one", async () => {
    const tool = buildDownloadFileTool(
      graphReturning(item({ "@microsoft.graph.downloadUrl": undefined })),
      vi.fn(() => Promise.resolve(binaryResponse("x"))),
    )

    expect(await tool.execute(args)).toContain("not returned")
  })

  it("throws the Graph error message at the tool boundary", async () => {
    const tool = buildDownloadFileTool(graphFailing("Item not found"))

    await expect(tool.execute(args)).rejects.toThrow("Item not found")
  })
})

describe("GRAPH_ENABLE_DOWNLOAD_FILE", () => {
  const baseEnv = {
    MS_GRAPH_TENANT_ID: "tenant",
    MS_GRAPH_CLIENT_ID: "client",
    MS_GRAPH_CLIENT_SECRET: "secret",
  }
  const resolve = (over: Record<string, string> = {}) =>
    resolveServerRuntimeConfig({ ...baseEnv, ...over } as NodeJS.ProcessEnv).value as { enableDownloadFile: boolean }

  // Defaults ON: the deployed containers already expose the tool and must not lose it on re-pull.
  it("defaults on when unset", () => {
    expect(resolve().enableDownloadFile).toBe(true)
  })

  it('turns off only for an explicit "false"', () => {
    expect(resolve({ GRAPH_ENABLE_DOWNLOAD_FILE: "false" }).enableDownloadFile).toBe(false)
    expect(resolve({ GRAPH_ENABLE_DOWNLOAD_FILE: "FALSE" }).enableDownloadFile).toBe(false)
    expect(resolve({ GRAPH_ENABLE_DOWNLOAD_FILE: "true" }).enableDownloadFile).toBe(true)
    expect(resolve({ GRAPH_ENABLE_DOWNLOAD_FILE: "" }).enableDownloadFile).toBe(true)
  })
})
