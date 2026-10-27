import { createRequire } from "node:module"

import { mkdtempSync, readdirSync, readFileSync, rmSync, statSync } from "node:fs"
import { tmpdir } from "node:os"
import { dirname, join } from "node:path"

import { type TokenCachePersistenceOptions, useIdentityPlugin } from "@azure/identity"
import { type ICachePlugin, PublicClientApplication, TokenCacheContext } from "@azure/msal-node"
import { afterEach, beforeEach, describe, expect, it } from "vitest"

import { createFileCachePlugin, fileCachePersistencePlugin } from "../src/auth/token-cache"

// Unit tests exercise our own plugin object. This drives the real @azure/identity
// plugin machinery instead — the path that silently did nothing under the keytar
// implementation, where the plugin failed to load and persistence quietly degraded.
describe("token cache integration with @azure/identity", () => {
  let directory: string

  beforeEach(() => {
    directory = mkdtempSync(join(tmpdir(), "cache-integration-"))
    process.env.MS365_TOKEN_CACHE_PATH = directory
  })

  afterEach(() => {
    delete process.env.MS365_TOKEN_CACHE_PATH
    rmSync(directory, { recursive: true, force: true })
  })

  it("registers with useIdentityPlugin and persists through the real config path", async () => {
    // msalClient reaches the plugin through this control object, which
    // useIdentityPlugin is responsible for populating. Asserting on it proves
    // registration actually took effect — the step that silently no-opped under the
    // keytar implementation.
    // This module is internal to identity and has no published types, so the import is
    // typed by hand — with identity's and MSAL's own published types for the factory.
    type PersistenceFactory = (options?: TokenCachePersistenceOptions) => Promise<ICachePlugin>
    const require_ = createRequire(import.meta.url)
    const identityRoot = dirname(require_.resolve("@azure/identity/package.json"))
    const { msalNodeFlowCacheControl } = (await import(
      join(identityRoot, "dist/esm/msal/nodeFlows/msalPlugins.js")
    )) as { msalNodeFlowCacheControl: { setPersistence: (factory: PersistenceFactory) => void } }

    // Capture the provider identity registered, then exercise it exactly as
    // generatePluginConfiguration does internally.
    const captured: Array<PersistenceFactory> = []
    msalNodeFlowCacheControl.setPersistence = (factory) => captured.push(factory)

    useIdentityPlugin(fileCachePersistencePlugin)
    expect(captured, "useIdentityPlugin did not register a persistence provider").toHaveLength(1)

    const config = { cache: { cachePlugin: captured[0]!({ name: "integration" }) } }

    const plugin = await config.cache.cachePlugin
    expect(plugin, "identity produced no cache plugin").toBeDefined()

    await plugin.afterCacheAccess(
      new TokenCacheContext({ serialize: () => '{"secret":"xyz"}', deserialize: () => undefined }, true),
    )

    const files = readdirSync(directory)
    expect(files, "nothing was persisted to disk").toHaveLength(1)
    expect(statSync(join(directory, files[0]!)).mode & 0o777).toBe(0o600)

    const seen: string[] = []
    await plugin.beforeCacheAccess(
      new TokenCacheContext({ serialize: () => "", deserialize: (d: string) => void seen.push(d) }, false),
    )
    expect(seen).toEqual(['{"secret":"xyz"}'])
  })

  // A stub cache cannot catch this one: MSAL's real TokenCache keeps the rejected text as
  // its in-memory snapshot, so serialize() throws the same SyntaxError until the cache is
  // reset. Unreset, the first save after a sign-in throws and takes the server down.
  it("recovers from a corrupt cache file without poisoning MSAL's in-memory cache", async () => {
    const plugin = createFileCachePlugin("corrupt", directory)
    await plugin.afterCacheAccess(
      new TokenCacheContext({ serialize: () => "{not json", deserialize: () => undefined }, true),
    )
    expect(readdirSync(directory)).toHaveLength(1)

    const pca = new PublicClientApplication({
      auth: { clientId: "00000000-0000-0000-0000-000000000000" },
      cache: { cachePlugin: plugin },
    })
    const cache = pca.getTokenCache()

    // Loading the cache reads the corrupt file and discards it.
    await expect(cache.getAllAccounts()).resolves.toEqual([])
    expect(readdirSync(directory)).toEqual([])

    // What a sign-in does next: hand the changed cache back to be saved.
    await expect(plugin.afterCacheAccess(new TokenCacheContext(cache, true))).resolves.toBeUndefined()
    const files = readdirSync(directory)
    expect(files).toHaveLength(1)
    expect(JSON.parse(readFileSync(join(directory, files[0]!), "utf8"))).toHaveProperty("RefreshToken")
  })

  it("keeps running when the cache cannot be serialized", async () => {
    const plugin = createFileCachePlugin("unserializable", directory)
    const unserializable = {
      serialize: (): string => {
        throw new SyntaxError("Unexpected token")
      },
      deserialize: () => undefined,
    }

    await expect(plugin.afterCacheAccess(new TokenCacheContext(unserializable, true))).resolves.toBeUndefined()
    expect(readdirSync(directory)).toEqual([])
  })
})
