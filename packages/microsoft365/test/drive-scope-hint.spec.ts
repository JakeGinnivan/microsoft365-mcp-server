import { Some } from "functype"
import { Left, Right } from "functype/either"
import { beforeEach, describe, expect, it, vi } from "vitest"

vi.mock("../src/client/graph-client", () => ({ getGraphClient: vi.fn() }))
vi.mock("../src/auth", async (importOriginal) => ({
  ...(await importOriginal<typeof import("../src/auth")>()),
  getAuthStatus: vi.fn(),
}))

import { getAuthStatus } from "../src/auth"
import { getGraphClient } from "../src/client/graph-client"
import { listDriveItems, searchFiles } from "../src/tools/files-tools"

const mockClient = { listDriveItems: vi.fn(), searchFiles: vi.fn() }
const notFound = Left({ type: "not_found", status: 404, message: "Item not found" })
const withScopes = (scopes: ReadonlyArray<string>) =>
  vi.mocked(getAuthStatus).mockResolvedValue(Right({ mode: "interactive", authenticated: true, scopes: [...scopes] }))

const message = (result: { value: unknown }) => (result.value as { message: string }).message

beforeEach(() => {
  vi.clearAllMocks()
  vi.mocked(getGraphClient).mockReturnValue(Some(mockClient) as never)
})

// OneDrive answers a token without Files permission with 404 "Item not found", which reads as an empty or
// missing drive rather than a permission problem.
describe("drive errors when the token has no Files permission", () => {
  it("names the missing permission when listing", async () => {
    withScopes(["Mail.ReadWrite", "User.Read"])
    mockClient.listDriveItems.mockResolvedValue(notFound)

    const result = await listDriveItems({})

    expect(result.isLeft()).toBe(true)
    expect(message(result)).toContain("Item not found")
    expect(message(result)).toContain("no Files permission")
  })

  it("names the missing permission when searching", async () => {
    withScopes(["Mail.ReadWrite"])
    mockClient.searchFiles.mockResolvedValue(notFound)

    expect(message(await searchFiles({ query: "receipt" }))).toContain("no Files permission")
  })

  it("leaves a genuine not-found alone when the token can read files", async () => {
    withScopes(["Files.Read", "User.Read"])
    mockClient.listDriveItems.mockResolvedValue(notFound)

    expect(message(await listDriveItems({ folder_id: "gone" }))).toBe("Failed to list drive items: Item not found")
  })

  it("treats Sites permission as able to read OneDrive", async () => {
    withScopes(["Sites.Read.All", "Mail.ReadWrite"])
    mockClient.listDriveItems.mockResolvedValue(notFound)

    expect(message(await listDriveItems({ folder_id: "gone" }))).toBe("Failed to list drive items: Item not found")
  })

  // A throttled or failed request says nothing about permissions, and reading the scopes on an auth
  // failure would ask the credential for a token a second time.
  it("leaves other failures alone without reading the scopes", async () => {
    withScopes(["Mail.ReadWrite"])
    mockClient.listDriveItems.mockResolvedValue(Left({ type: "throttle", status: 429, message: "Too many requests" }))

    expect(message(await listDriveItems({}))).toBe("Failed to list drive items: Too many requests")
    expect(getAuthStatus).not.toHaveBeenCalled()
  })

  it("adds no hint when the token's scopes cannot be read", async () => {
    withScopes([])
    mockClient.listDriveItems.mockResolvedValue(notFound)

    expect(message(await listDriveItems({}))).not.toContain("Files permission")
  })
})
