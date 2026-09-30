import { readFile, rm } from "node:fs/promises"
import { tmpdir } from "node:os"
import { join } from "node:path"

import { Some } from "functype"
import { Left, Right } from "functype/either"
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"

vi.mock("../src/client/graph-client", () => ({ getGraphClient: vi.fn() }))

import { getGraphClient } from "../src/client/graph-client"
import { exportContacts } from "../src/tools/contacts-tools"

const mockClient = { requestPaginated: vi.fn() }

const pages: Record<string, ReadonlyArray<Record<string, unknown>>> = {
  "/me/contactFolders": [{ id: "F1", displayName: "Family" }],
  "/me/contactFolders/F1/childFolders": [{ id: "F2", displayName: "Cousins" }],
  "/me/contactFolders/F2/childFolders": [],
  "/me/contacts": [{ id: "C1", displayName: "Default Person", categories: ["Friends"] }],
  "/me/contactFolders/F1/contacts": [{ id: "C2", displayName: "Family Person" }],
  "/me/contactFolders/F2/contacts": [{ id: "C3", displayName: "Cousin Person" }],
}

let outDir: string

beforeEach(() => {
  vi.clearAllMocks()
  vi.mocked(getGraphClient).mockReturnValue(Some(mockClient as never))
  mockClient.requestPaginated.mockImplementation(async (path: string) => Right(pages[path] ?? []))
  outDir = join(tmpdir(), `export-contacts-test-${Date.now()}-${Math.random().toString(36).slice(2)}`)
})

afterEach(async () => {
  await rm(outDir, { recursive: true, force: true })
})

const readExport = async (message: string) => {
  const path = /Saved: (.+)/.exec(message)?.[1]
  expect(path).toBeDefined()
  return JSON.parse(await readFile(path as string, "utf-8")) as {
    folders: ReadonlyArray<{ path: string; count: number }>
    contacts: ReadonlyArray<{ id: string; folderPath: string }>
  }
}

describe("exportContacts", () => {
  it("exports the default folder and every nested contact folder", async () => {
    const result = await exportContacts({ out_dir: outDir })

    expect(result.isRight()).toBe(true)
    expect(result.value).toContain("Contacts: 3")
    const exported = await readExport(result.value as string)
    expect(exported.contacts.map((c) => [c.id, c.folderPath])).toEqual([
      ["C1", "Contacts"],
      ["C2", "Contacts/Family"],
      ["C3", "Contacts/Family/Cousins"],
    ])
    expect(exported.folders.map((f) => f.count)).toEqual([1, 1, 1])
  })

  // The default page size would hit requestPaginated's 50-page cap at 500 contacts and
  // truncate without an error, so every contact listing must ask for full pages.
  it("requests full pages with the complete field set", async () => {
    await exportContacts({ out_dir: outDir })

    const contactCalls = mockClient.requestPaginated.mock.calls.filter(([path]) => String(path).endsWith("/contacts"))
    expect(contactCalls).toHaveLength(3)
    for (const [, options] of contactCalls) {
      expect(options.odataParams.$top).toBe(999)
      expect(options.odataParams.$select).toContain("categories")
      expect(options.odataParams.$select).toContain("emailAddresses")
    }
  })

  it("exports a folder reached twice only once", async () => {
    mockClient.requestPaginated.mockImplementation(async (path: string) =>
      Right(
        path === "/me/contactFolders"
          ? [
              { id: "F1", displayName: "Family" },
              { id: "F2", displayName: "Cousins" },
            ]
          : (pages[path] ?? []),
      ),
    )

    const result = await exportContacts({ out_dir: outDir })

    const exported = await readExport(result.value as string)
    expect(exported.contacts.filter((c) => c.id === "C3")).toHaveLength(1)
  })

  it("fails rather than writing a partial export when a folder cannot be read", async () => {
    mockClient.requestPaginated.mockImplementation(async (path: string) =>
      path === "/me/contactFolders/F1/contacts"
        ? Left({ type: "http", message: "Access is denied" })
        : Right(pages[path] ?? []),
    )

    const result = await exportContacts({ out_dir: outDir })

    expect(result.isLeft()).toBe(true)
    expect((result.value as Error).message).toContain("Contacts/Family")
  })
})
