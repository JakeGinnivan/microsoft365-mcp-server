import { createHash } from "node:crypto"
import {
  chmodSync,
  existsSync,
  mkdirSync,
  mkdtempSync,
  readdirSync,
  readFileSync,
  rmSync,
  statSync,
  symlinkSync,
  writeFileSync,
} from "node:fs"
import { tmpdir } from "node:os"
import { join } from "node:path"

import type { TokenCachePersistenceOptions } from "@azure/identity"
import { type ICachePlugin, TokenCacheContext } from "@azure/msal-node"
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"

import { createFileCachePlugin, fileCachePersistencePlugin, resolveCacheDirectory } from "../src/auth/token-cache"

const cacheContext = (serialized: string, hasChanged = true) => {
  const deserialize = vi.fn()
  return {
    context: new TokenCacheContext({ deserialize, serialize: () => serialized }, hasChanged),
    deserialize,
  }
}

describe("token cache", () => {
  let directory: string

  beforeEach(() => {
    directory = mkdtempSync(join(tmpdir(), "token-cache-"))
  })

  afterEach(() => {
    rmSync(directory, { recursive: true, force: true })
    vi.restoreAllMocks()
  })

  describe("resolveCacheDirectory", () => {
    it("prefers an explicit MS365_TOKEN_CACHE_PATH", () => {
      expect(resolveCacheDirectory({ MS365_TOKEN_CACHE_PATH: "/custom/path" })).toBe("/custom/path")
    })

    it("falls back to XDG_CONFIG_HOME", () => {
      expect(resolveCacheDirectory({ XDG_CONFIG_HOME: "/xdg" })).toBe("/xdg/microsoft365-mcp-server/token-cache")
    })

    // TOKEN_STORAGE_PATH belongs to the oauth-proxy provider. Reading it here would
    // couple two unrelated features together.
    it("ignores TOKEN_STORAGE_PATH", () => {
      expect(resolveCacheDirectory({ TOKEN_STORAGE_PATH: "/proxy", XDG_CONFIG_HOME: "/xdg" })).not.toContain("/proxy")
    })
  })

  describe("round trip", () => {
    it("writes on change and reads back", async () => {
      const plugin = createFileCachePlugin("test", directory)
      await plugin.afterCacheAccess(cacheContext('{"token":"abc"}').context)

      const reader = cacheContext("")
      await plugin.beforeCacheAccess(reader.context)

      expect(reader.deserialize).toHaveBeenCalledWith('{"token":"abc"}')
    })

    it("does not write when the cache is unchanged", async () => {
      const plugin = createFileCachePlugin("test", directory)
      await plugin.afterCacheAccess(cacheContext("{}", false).context)

      const reader = cacheContext("")
      await plugin.beforeCacheAccess(reader.context)

      expect(reader.deserialize).not.toHaveBeenCalled()
    })

    it("treats a missing cache as empty rather than failing", async () => {
      const plugin = createFileCachePlugin("absent", join(directory, "does-not-exist"))
      const reader = cacheContext("")

      await expect(plugin.beforeCacheAccess(reader.context)).resolves.toBeUndefined()
      expect(reader.deserialize).not.toHaveBeenCalled()
    })

    it("keeps separate caches for different names", async () => {
      const a = createFileCachePlugin("cache-a", directory)
      const b = createFileCachePlugin("cache-b", directory)

      await a.afterCacheAccess(cacheContext('"from-a"').context)
      await b.afterCacheAccess(cacheContext('"from-b"').context)

      const reader = cacheContext("")
      await a.beforeCacheAccess(reader.context)
      expect(reader.deserialize).toHaveBeenCalledWith('"from-a"')
    })

    // A name that sanitises away entirely must still land in its own file.
    it("distinguishes names that sanitise to the same string", async () => {
      const a = createFileCachePlugin("!!!", directory)
      const b = createFileCachePlugin("???", directory)

      await a.afterCacheAccess(cacheContext('"a"').context)
      await b.afterCacheAccess(cacheContext('"b"').context)

      const reader = cacheContext("")
      await a.beforeCacheAccess(reader.context)
      expect(reader.deserialize).toHaveBeenCalledWith('"a"')
    })

    it("keeps a path separator in the name from escaping the directory", async () => {
      const plugin = createFileCachePlugin("../../escape", directory)
      await plugin.afterCacheAccess(cacheContext('"contained"').context)

      const reader = cacheContext("")
      await plugin.beforeCacheAccess(reader.context)
      expect(reader.deserialize).toHaveBeenCalledWith('"contained"')
    })
  })

  // The token is a bearer credential: file permissions are the security model.
  describe("permissions", () => {
    it("writes the cache readable only by its owner", async () => {
      const plugin = createFileCachePlugin("perms", directory)
      await plugin.afterCacheAccess(cacheContext('{"token":"abc"}').context)

      const file = readFileSync(join(directory, findCacheFile(directory)), "utf8")
      expect(file).toBe('{"token":"abc"}')
      expect(statSync(join(directory, findCacheFile(directory))).mode & 0o777).toBe(0o600)
    })

    it("discards a cache that others can read", async () => {
      const plugin = createFileCachePlugin("loose", directory)
      await plugin.afterCacheAccess(cacheContext('{"token":"abc"}').context)

      const path = join(directory, findCacheFile(directory))
      chmodSync(path, 0o644)

      const reader = cacheContext("")
      vi.spyOn(console, "error").mockImplementation(() => undefined)
      await plugin.beforeCacheAccess(reader.context)

      expect(reader.deserialize).not.toHaveBeenCalled()
    })

    it("accepts a cache its owner has made read-only", async () => {
      const plugin = createFileCachePlugin("readonly", directory)
      await plugin.afterCacheAccess(cacheContext('{"token":"abc"}').context)
      chmodSync(join(directory, findCacheFile(directory)), 0o400)

      const reader = cacheContext("")
      await plugin.beforeCacheAccess(reader.context)

      expect(reader.deserialize).toHaveBeenCalledWith('{"token":"abc"}')
    })

    // Node on Windows reports 0666 or 0444 for every file, so a POSIX mode check there
    // would discard the cache on every start.
    it("does not judge the cache by its mode on Windows", async () => {
      const plugin = createFileCachePlugin("windows", directory)
      await plugin.afterCacheAccess(cacheContext('{"token":"abc"}').context)
      const path = join(directory, findCacheFile(directory))
      chmodSync(path, 0o666)

      const reader = cacheContext("")
      const log = vi.spyOn(console, "error").mockImplementation(() => undefined)
      await withPlatform("win32", () => plugin.beforeCacheAccess(reader.context))

      expect(reader.deserialize).toHaveBeenCalledWith('{"token":"abc"}')
      expect(existsSync(path)).toBe(true)
      expect(log).not.toHaveBeenCalled()
    })

    it("does not reuse a leftover temp file's permissions", async () => {
      const plugin = createFileCachePlugin("stale", directory)
      const path = join(directory, cacheFileNameFor("stale"))
      mkdirSync(directory, { recursive: true })
      writeFileSync(`${path}.${process.pid}.tmp`, "stale", { mode: 0o666 })

      await plugin.afterCacheAccess(cacheContext('{"token":"fresh"}').context)

      expect(statSync(path).mode & 0o777).toBe(0o600)
    })
  })

  describe("temporary file", () => {
    it("does not write through a symlink planted at the temp path", async () => {
      const plugin = createFileCachePlugin("planted", directory)
      const path = join(directory, cacheFileNameFor("planted"))
      const target = join(directory, "attacker-owned")
      writeFileSync(target, "untouched")
      symlinkSync(target, `${path}.${process.pid}.tmp`)

      await plugin.afterCacheAccess(cacheContext('{"token":"secret"}').context)

      expect(readFileSync(target, "utf8")).toBe("untouched")
      expect(readFileSync(path, "utf8")).toBe('{"token":"secret"}')
    })

    it("replaces a leftover temp file from an earlier crash", async () => {
      const plugin = createFileCachePlugin("leftover", directory)
      const path = join(directory, cacheFileNameFor("leftover"))
      writeFileSync(`${path}.${process.pid}.tmp`, "stale")

      await plugin.afterCacheAccess(cacheContext('{"token":"fresh"}').context)

      expect(readFileSync(path, "utf8")).toBe('{"token":"fresh"}')
      expect(existsSync(`${path}.${process.pid}.tmp`)).toBe(false)
    })
  })

  describe("failure handling", () => {
    it("discards a cache that cannot be deserialised", async () => {
      const plugin = createFileCachePlugin("corrupt", directory)
      await plugin.afterCacheAccess(cacheContext("{not json").context)
      const path = join(directory, findCacheFile(directory))

      const reader = cacheContext("")
      reader.deserialize.mockImplementation(() => {
        throw new SyntaxError("Unexpected token n in JSON")
      })
      const log = vi.spyOn(console, "error").mockImplementation(() => undefined)

      await expect(plugin.beforeCacheAccess(reader.context)).resolves.toBeUndefined()
      expect(existsSync(path)).toBe(false)
      expect(log).toHaveBeenCalledWith(expect.stringContaining("Discarding unreadable token cache"))
    })

    it("degrades to in-memory when the cache cannot be written", async () => {
      // A file where the directory should be makes every write fail.
      const blocked = join(directory, "blocked")
      writeFileSync(blocked, "not a directory")

      const plugin = createFileCachePlugin("test", blocked)
      vi.spyOn(console, "error").mockImplementation(() => undefined)

      await expect(plugin.afterCacheAccess(cacheContext("{}").context)).resolves.toBeUndefined()
    })
  })

  describe("fileCachePersistencePlugin", () => {
    it("registers a provider that yields a working plugin", async () => {
      const setPersistence =
        vi.fn<(factory: (options?: TokenCachePersistenceOptions) => Promise<ICachePlugin>) => void>()
      fileCachePersistencePlugin({ cachePluginControl: { setPersistence } })

      expect(setPersistence).toHaveBeenCalledOnce()

      const factory = setPersistence.mock.calls[0]![0]
      const plugin = await factory({ name: "from-provider" })

      expect(typeof plugin.beforeCacheAccess).toBe("function")
    })
  })
})

const withPlatform = async <T>(platform: NodeJS.Platform, fn: () => Promise<T>): Promise<T> => {
  const original = Object.getOwnPropertyDescriptor(process, "platform")!
  Object.defineProperty(process, "platform", { ...original, value: platform })
  try {
    return await fn()
  } finally {
    Object.defineProperty(process, "platform", original)
  }
}

// The plugin owns its naming scheme; tests locate the file rather than duplicating it.
const findCacheFile = (directory: string): string => {
  const found = readdirSync(directory).find((f) => f.endsWith(".json"))
  if (!found) throw new Error(`No cache file written in ${directory}`)
  return found
}

const cacheFileNameFor = (name: string): string => {
  const safe = name.replace(/[^A-Za-z0-9._-]/g, "_").slice(0, 64)
  return `${safe}.${createHash("sha256").update(name).digest("hex").slice(0, 8)}.json`
}
