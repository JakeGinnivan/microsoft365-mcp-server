import { Right } from "functype/either"
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"

const server = vi.hoisted(() => ({
  addTool: vi.fn(),
  start: vi.fn(),
  post: vi.fn(),
  put: vi.fn(),
  instructions: "",
}))
vi.mock("fastmcp", async (original) => ({
  ...(await original<typeof import("fastmcp")>()),
  FastMCP: class {
    constructor(options: { instructions: string }) {
      server.instructions = options.instructions
    }
    addTool = server.addTool
    start = server.start
    getApp = () => ({ post: server.post, put: server.put })
  },
}))
vi.mock("../src/auth", async (original) => ({
  ...(await original<typeof import("../src/auth")>()),
  initializeAuth: vi.fn(async () => Right(undefined)),
  getAccessToken: vi.fn(async () => Right("TOKEN")),
}))
vi.mock("../src/auth/oauth-provider", () => ({ createAzureAuthProvider: vi.fn() }))

beforeEach(() => {
  vi.resetModules()
  vi.clearAllMocks()
  for (const key of [
    "MS365_READ_ONLY",
    "MS365_PRESETS",
    "MS365_ENABLED_TOOLS",
    "MS365_REQUIRE_DRAFT",
    "MS365_ORG_MODE",
  ])
    vi.stubEnv(key, undefined)
  vi.stubEnv("TRANSPORT_TYPE", "httpStream")
})
afterEach(() => {
  vi.unstubAllEnvs()
  vi.unstubAllGlobals()
})

describe.each(["client-token", "oauth-proxy"])("%s server environment controls", (mode) => {
  it.each([
    ["MS365_READ_ONLY", "true"],
    ["MS365_PRESETS", "productivity"],
    ["MS365_ENABLED_TOOLS", "^list_"],
  ])("does not expose /upload when %s=%s", async (key, value) => {
    vi.stubEnv("MS365_AUTH_MODE", mode)
    vi.stubEnv(key, value)
    await import("../src/index")
    await vi.waitFor(() => expect(server.start).toHaveBeenCalled())
    expect(server.post).not.toHaveBeenCalled()
    expect(server.put).not.toHaveBeenCalled()
    expect(server.instructions).not.toContain("Upload workflows:")
    expect(server.addTool.mock.calls.map(([tool]) => tool.name)).not.toContain("get_upload_config")
  })

  it.each([undefined, "v1.0"])("uses the configured Graph version for /upload, with override %s", async (override) => {
    vi.stubEnv("MS365_AUTH_MODE", mode)
    vi.stubEnv("MS365_GRAPH_VERSION", "beta")
    vi.stubEnv("MS365_UPLOAD_TOKEN", "TOKEN")
    const fetchSpy = vi.fn().mockResolvedValue(new Response(JSON.stringify({ id: "uploaded" })))
    vi.stubGlobal("fetch", fetchSpy)
    await import("../src/index")
    await vi.waitFor(() => expect(server.start).toHaveBeenCalled())
    const handler = server.post.mock.calls[0]?.[1]
    const query: Record<string, string | undefined> = {
      path: "/me/drive/root:/test.txt:/content",
      apiVersion: override,
    }
    await handler({
      req: {
        query: (key: string) => query[key],
        header: () => "Bearer TOKEN",
        arrayBuffer: async () => new TextEncoder().encode("text").buffer,
      },
      json: vi.fn(),
    })
    expect(fetchSpy.mock.calls[0]?.[0]).toContain(`/${override ?? "beta"}/me/drive/`)
  })

  it("keeps /upload when get_upload_config is enabled", async () => {
    vi.stubEnv("MS365_AUTH_MODE", mode)
    vi.stubEnv("MS365_ENABLED_TOOLS", "^get_upload_config$")
    await import("../src/index")
    await vi.waitFor(() => expect(server.start).toHaveBeenCalled())
    expect(server.post).toHaveBeenCalledWith("/upload", expect.any(Function))
    expect(server.put).toHaveBeenCalledWith("/upload", expect.any(Function))
  })
})
