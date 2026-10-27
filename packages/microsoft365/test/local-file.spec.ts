import { mkdir, readdir, readFile, rm, stat, writeFile } from "node:fs/promises"
import { tmpdir } from "node:os"
import { join } from "node:path"
import { afterEach, beforeEach, describe, expect, it } from "vitest"
import { writeNewFile } from "../src/utils/local-file"

let dir: string

beforeEach(() => {
  dir = join(tmpdir(), `local-file-test-${Date.now()}-${Math.random().toString(36).slice(2)}`)
})

afterEach(async () => {
  await rm(dir, { recursive: true, force: true })
})

describe("writeNewFile", () => {
  it("creates the folder and writes a private file", async () => {
    const result = await writeNewFile(dir, "a.txt", "hello")

    expect(result.isRight()).toBe(true)
    expect(result.value).toBe(join(dir, "a.txt"))
    expect(await readFile(join(dir, "a.txt"), "utf-8")).toBe("hello")
    expect((await stat(join(dir, "a.txt"))).mode & 0o777).toBe(0o600)
    expect(await readdir(dir)).toEqual(["a.txt"])
  })

  it("refuses a taken name by default and leaves the existing file alone", async () => {
    await mkdir(dir, { recursive: true })
    await writeFile(join(dir, "a.txt"), "earlier")

    const result = await writeNewFile(dir, "a.txt", "later")

    expect(result.isLeft()).toBe(true)
    expect((result.value as Error).message).toContain("already exists")
    expect(await readFile(join(dir, "a.txt"), "utf-8")).toBe("earlier")
    expect(await readdir(dir)).toEqual(["a.txt"])
  })

  it("falls back to a numbered name when allowed", async () => {
    await mkdir(dir, { recursive: true })
    await writeFile(join(dir, "a.txt"), "first")
    await writeFile(join(dir, "a (1).txt"), "second")

    const result = await writeNewFile(dir, "a.txt", "third", 100)

    expect(result.value).toBe(join(dir, "a (2).txt"))
    expect(await readFile(join(dir, "a (2).txt"), "utf-8")).toBe("third")
    expect(await readFile(join(dir, "a.txt"), "utf-8")).toBe("first")
  })

  it("gives up once every allowed name is taken", async () => {
    await mkdir(dir, { recursive: true })
    await writeFile(join(dir, "a.txt"), "first")
    await writeFile(join(dir, "a (1).txt"), "second")

    const result = await writeNewFile(dir, "a.txt", "third", 2)

    expect(result.isLeft()).toBe(true)
    expect((result.value as Error).message).toContain("2 files named like a.txt already exist")
    expect((await readdir(dir)).sort()).toEqual(["a (1).txt", "a.txt"])
  })

  it("refuses a name that resolves outside the folder", async () => {
    const result = await writeNewFile(dir, "../escape.txt", "x")

    expect(result.isLeft()).toBe(true)
    expect((result.value as Error).message).toContain("Refusing to write outside")
  })
})
