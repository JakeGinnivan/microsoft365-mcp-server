import { mkdtempSync, readdirSync, readFileSync, rmSync, statSync } from "node:fs"
import { tmpdir } from "node:os"
import { join } from "node:path"

import { type INetworkModule, type NetworkResponse, PublicClientApplication } from "@azure/msal-node"
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"

import { configFingerprint } from "../src/auth/config-fingerprint"
import { createFileCachePlugin, dropAccessTokensOnConfigChange } from "../src/auth/token-cache"

const AUTH = { mode: "interactive", tenantId: "tenant-1", clientId: "client-1" }

describe("configFingerprint", () => {
  it("is stable across the order variables are set in", () => {
    expect(configFingerprint(AUTH, { MS365_READ_ONLY: "true", MS365_ALLOWED_MAILBOXES: "a@x.com" })).toBe(
      configFingerprint(AUTH, { MS365_ALLOWED_MAILBOXES: "a@x.com", MS365_READ_ONLY: "true" }),
    )
  })

  it("changes when a setting changes", () => {
    expect(configFingerprint(AUTH, { MS365_ALLOWED_MAILBOXES: "a@x.com" })).not.toBe(
      configFingerprint(AUTH, { MS365_ALLOWED_MAILBOXES: "a@x.com,b@x.com" }),
    )
    expect(configFingerprint(AUTH, { GRAPH_API_VERSIONS: "v1.0" })).not.toBe(configFingerprint(AUTH, {}))
  })

  it("changes when an auth input changes", () => {
    expect(configFingerprint(AUTH, {})).not.toBe(configFingerprint({ ...AUTH, clientId: "client-2" }, {}))
  })

  it("ignores variables that are not the server's", () => {
    expect(configFingerprint(AUTH, { PATH: "/bin", HOME: "/home/a" })).toBe(configFingerprint(AUTH, {}))
  })

  it("leaves secrets out", () => {
    for (const name of [
      "MS365_CLIENT_SECRET",
      "MS365_CERT_PASSWORD",
      "MS365_ACCESS_TOKEN",
      "MS365_UPLOAD_TOKEN",
      "MS365_TOKEN_ENCRYPTION_KEY",
      "MS365_JWT_SIGNING_KEY",
    ]) {
      expect(configFingerprint(AUTH, { [name]: "s3cret" }), name).toBe(configFingerprint(AUTH, {}))
    }
  })
})

// The four sections MSAL writes, as in a real cache file.
const msalCache = () =>
  JSON.stringify({
    Account: { "uid.utid-login.microsoftonline.com-utid": { username: "a@x.com" } },
    IdToken: { "id-key": { secret: "id" } },
    AccessToken: { "at-key": { secret: "old-access-token" } },
    RefreshToken: { "rt-key": { secret: "refresh-token" } },
    AppMetadata: {},
  })

describe("dropAccessTokensOnConfigChange", () => {
  let directory: string

  beforeEach(() => {
    directory = mkdtempSync(join(tmpdir(), "fingerprint-"))
    vi.spyOn(console, "error").mockImplementation(() => undefined)
  })

  afterEach(() => {
    rmSync(directory, { recursive: true, force: true })
    vi.restoreAllMocks()
  })

  const seed = async (contents: string) =>
    createFileCachePlugin("cache", directory).afterCacheAccess({
      cacheHasChanged: true,
      tokenCache: { serialize: () => contents, deserialize: () => undefined },
    } as never)
  const cacheFile = () =>
    join(
      directory,
      readdirSync(directory).find((f) => f.startsWith("cache."))!,
    )
  const sections = () => JSON.parse(readFileSync(cacheFile(), "utf8")) as Record<string, Record<string, unknown>>

  it("leaves the cache alone when the configuration is unchanged", async () => {
    await seed(msalCache())
    dropAccessTokensOnConfigChange("fp-a", "cache", directory)
    await seed(msalCache())
    expect(dropAccessTokensOnConfigChange("fp-a", "cache", directory)).toBe(false)
    expect(Object.keys(sections().AccessToken!)).toEqual(["at-key"])
  })

  it("drops only the access tokens when the configuration changes", async () => {
    await seed(msalCache())
    dropAccessTokensOnConfigChange("fp-a", "cache", directory)
    await seed(msalCache())
    expect(dropAccessTokensOnConfigChange("fp-b", "cache", directory)).toBe(true)
    const after = sections()
    expect(after.AccessToken).toEqual({})
    expect(Object.keys(after.RefreshToken!)).toEqual(["rt-key"])
    expect(Object.keys(after.Account!)).toHaveLength(1)
    expect(statSync(cacheFile()).mode & 0o777).toBe(0o600)
    expect(statSync(join(directory, "config-fingerprint")).mode & 0o777).toBe(0o600)
  })

  // A cache written by a version without the fingerprint may hold a token minted under
  // any earlier consent, so it is treated as changed once.
  it("treats a cache with no recorded fingerprint as changed", async () => {
    await seed(msalCache())
    expect(dropAccessTokensOnConfigChange("fp-a", "cache", directory)).toBe(true)
    expect(sections().AccessToken).toEqual({})
  })

  it("records the fingerprint when there is no cache yet", () => {
    expect(dropAccessTokensOnConfigChange("fp-a", "cache", directory)).toBe(true)
    expect(dropAccessTokensOnConfigChange("fp-a", "cache", directory)).toBe(false)
  })

  it("leaves an unparseable cache for the plugin to discard", async () => {
    await seed("{not json")
    expect(dropAccessTokensOnConfigChange("fp-a", "cache", directory)).toBe(true)
    expect(readFileSync(cacheFile(), "utf8")).toBe("{not json")
  })
})

// Drives real MSAL against a fake token endpoint: sign in once, restart, change the
// configuration, restart again. What matters is the last step — the dropped access token
// must be replaced through the refresh-token grant, silently, and never by a sign-in.
describe("dropAccessTokensOnConfigChange with MSAL", () => {
  let directory: string

  beforeEach(() => {
    directory = mkdtempSync(join(tmpdir(), "fingerprint-msal-"))
    vi.spyOn(console, "error").mockImplementation(() => undefined)
  })

  afterEach(() => {
    rmSync(directory, { recursive: true, force: true })
    vi.restoreAllMocks()
  })

  const TENANT = "11111111-1111-1111-1111-111111111111"
  const CLIENT = "22222222-2222-2222-2222-222222222222"
  const AUTHORITY = `https://login.microsoftonline.com/${TENANT}`
  const SCOPES = ["https://graph.microsoft.com/.default"]
  const b64url = (value: object) => Buffer.from(JSON.stringify(value)).toString("base64url")

  const fakeEntra = () => {
    const grants: Array<URLSearchParams> = []
    const issued = { count: 0 }
    const respond = <T>(body: unknown): NetworkResponse<T> => ({ headers: {}, body: body as T, status: 200 })
    const network: INetworkModule = {
      sendGetRequestAsync: async <T>(url: string) =>
        url.includes("openid-configuration")
          ? respond<T>({
              token_endpoint: `${AUTHORITY}/oauth2/v2.0/token`,
              authorization_endpoint: `${AUTHORITY}/oauth2/v2.0/authorize`,
              end_session_endpoint: `${AUTHORITY}/oauth2/v2.0/logout`,
              issuer: `https://login.microsoftonline.com/${TENANT}/v2.0`,
            })
          : respond<T>({
              tenant_discovery_endpoint: `${AUTHORITY}/v2.0/.well-known/openid-configuration`,
              metadata: [
                {
                  preferred_network: "login.microsoftonline.com",
                  preferred_cache: "login.windows.net",
                  aliases: ["login.microsoftonline.com", "login.windows.net"],
                },
              ],
            }),
      sendPostRequestAsync: async <T>(_url: string, options?: { body?: string }) => {
        grants.push(new URLSearchParams(options?.body ?? ""))
        issued.count += 1
        return respond<T>({
          token_type: "Bearer",
          // As a real cache records it: the granted scopes plus the .default requested.
          scope:
            "profile openid email https://graph.microsoft.com/Mail.ReadWrite https://graph.microsoft.com/User.Read " +
            "https://graph.microsoft.com/.default",
          expires_in: 3600,
          ext_expires_in: 3600,
          access_token: `access-token-${issued.count}`,
          refresh_token: `refresh-token-${issued.count}`,
          id_token: `${b64url({ alg: "none" })}.${b64url({
            aud: CLIENT,
            iss: `https://login.microsoftonline.com/${TENANT}/v2.0`,
            tid: TENANT,
            oid: "user-oid",
            preferred_username: "a@x.com",
            iat: Math.floor(Date.now() / 1000),
            exp: Math.floor(Date.now() / 1000) + 3600,
          })}.`,
          client_info: b64url({ uid: "user-oid", utid: TENANT }),
        })
      },
    }
    return { grants, network }
  }

  // A new PublicClientApplication per "process", sharing only the files on disk.
  const start = (network: INetworkModule) =>
    new PublicClientApplication({
      auth: { clientId: CLIENT, authority: AUTHORITY },
      cache: { cachePlugin: createFileCachePlugin("msal", directory) },
      system: { networkClient: network },
    })

  const silent = async (network: INetworkModule) => {
    const app = start(network)
    const [account] = await app.getTokenCache().getAllAccounts()
    return app.acquireTokenSilent({ account: account!, scopes: SCOPES })
  }

  it("refreshes silently with the refresh token after a configuration change", async () => {
    const entra = fakeEntra()

    dropAccessTokensOnConfigChange("config-a", "msal", directory)
    await start(entra.network).acquireTokenByCode({
      code: "auth-code",
      scopes: SCOPES,
      redirectUri: "http://localhost",
    })
    expect(entra.grants.map((g) => g.get("grant_type"))).toEqual(["authorization_code"])

    // Restart with the same configuration: the cached access token is used as is.
    expect(dropAccessTokensOnConfigChange("config-a", "msal", directory)).toBe(false)
    expect((await silent(entra.network)).accessToken).toBe("access-token-1")
    expect(entra.grants).toHaveLength(1)

    // Restart with a changed configuration: a new access token, via the refresh token.
    expect(dropAccessTokensOnConfigChange("config-b", "msal", directory)).toBe(true)
    expect((await silent(entra.network)).accessToken).toBe("access-token-2")
    expect(entra.grants.map((g) => g.get("grant_type"))).toEqual(["authorization_code", "refresh_token"])
    expect(entra.grants[1]!.get("refresh_token")).toBe("refresh-token-1")
  })
})
