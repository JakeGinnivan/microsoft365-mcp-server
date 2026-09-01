import type { AccessToken as AzureAccessToken, TokenCredential } from "@azure/identity"
import {
  ClientCertificateCredential,
  ClientSecretCredential,
  DeviceCodeCredential,
  InteractiveBrowserCredential,
  useIdentityPlugin,
} from "@azure/identity"
import { Ref } from "functype"
import { type Either, Left, Right } from "functype/either"
import { Try } from "functype/try"

import type { AuthConfig, AuthError } from "../types"
import { GRAPH_DEFAULT_SCOPE } from "./scopes"

const ONE_HOUR_MS = 60 * 60 * 1000

const DEFAULT_REDIRECT_URI = "http://localhost:3000"

const TOKEN_CACHE_NAME = "microsoft365-mcp-server"

// Without a persistence plugin, Azure Identity keeps tokens in memory only, so every
// process restart forces a fresh interactive sign-in. The plugin is an optional
// dependency because it pulls a native keychain binding (keytar) that will not build
// everywhere — so registration is attempted once, and its absence degrades to the old
// in-memory behaviour rather than failing startup.
// Ref keeps the one-shot registration state without a mutable binding, matching the
// functional style enforced across this package.
const cachePersistence = Ref<"unattempted" | "enabled" | "unavailable">("unattempted")

const registerCachePersistencePlugin = (): "enabled" | "unavailable" => {
  try {
    // require rather than import: this must not become a hard dependency of the bundle.
    // eslint-disable-next-line @typescript-eslint/no-require-imports -- optional native dep, resolved at runtime
    const plugin = require("@azure/identity-cache-persistence") as { cachePersistencePlugin: unknown }
    useIdentityPlugin(plugin.cachePersistencePlugin as Parameters<typeof useIdentityPlugin>[0])
    return "enabled"
  } catch {
    console.error(
      "[Auth] Token cache persistence unavailable (@azure/identity-cache-persistence could not be loaded). " +
        "Sign-in will be required again after restart.",
    )
    return "unavailable"
  }
}

const enableTokenCachePersistence = (): boolean => {
  const current = cachePersistence.get()
  if (current !== "unattempted") return current === "enabled"

  const outcome = registerCachePersistencePlugin()
  cachePersistence.set(outcome)
  return outcome === "enabled"
}

/** Test seam: reset the one-shot plugin registration state. */
export const resetTokenCachePersistenceForTests = (): void => cachePersistence.set("unattempted")

const tryCredential = (fn: () => TokenCredential, label: string): Either<AuthError, TokenCredential> =>
  Try(fn).fold(
    (e): Either<AuthError, TokenCredential> =>
      Left<AuthError, TokenCredential>({
        type: "credential",
        message: `Failed to create ${label} credential: ${String(e)}`,
      }),
    (cred): Either<AuthError, TokenCredential> => Right(cred),
  )

export const createCredential = (config: AuthConfig): Either<AuthError, TokenCredential> => {
  switch (config.mode) {
    case "interactive":
      return createInteractiveCredential(config)
    case "certificate":
      return createCertificateCredential(config)
    case "client-secret":
      return createClientSecretCredential(config)
    case "client-token":
      return createClientProvidedTokenCredential(config)
    case "oauth-proxy":
      return Left<AuthError, TokenCredential>({
        type: "config",
        message: "OAuth proxy mode uses AzureProvider, not credential-based auth",
      })
  }
}

const createInteractiveCredential = (
  config: Extract<AuthConfig, { mode: "interactive" }>,
): Either<AuthError, TokenCredential> => {
  const { tenantId, clientId } = config

  if (!clientId) {
    return Left<AuthError, TokenCredential>({ type: "config", message: "Interactive mode requires MS365_CLIENT_ID" })
  }

  // Opt out with MS365_TOKEN_CACHE=false, e.g. on a shared machine where a persisted
  // token should not outlive the session.
  const persist = process.env.MS365_TOKEN_CACHE !== "false" && enableTokenCachePersistence()
  const tokenCachePersistenceOptions = persist ? { enabled: true, name: TOKEN_CACHE_NAME } : undefined

  return tryCredential(() => {
    try {
      return new InteractiveBrowserCredential({
        tenantId,
        clientId,
        redirectUri: config.redirectUri ?? DEFAULT_REDIRECT_URI,
        tokenCachePersistenceOptions,
      }) as TokenCredential
    } catch {
      // Fallback to device code flow for headless environments
      return new DeviceCodeCredential({
        tenantId,
        clientId,
        tokenCachePersistenceOptions,
        userPromptCallback: (info) => {
          console.error(`\nAuthentication Required:`)
          console.error(`Please visit: ${info.verificationUri}`)
          console.error(`And enter code: ${info.userCode}\n`)
        },
      }) as TokenCredential
    }
  }, "interactive")
}

const createCertificateCredential = (
  config: Extract<AuthConfig, { mode: "certificate" }>,
): Either<AuthError, TokenCredential> => {
  if (!config.tenantId || !config.clientId || !config.certPath) {
    return Left<AuthError, TokenCredential>({
      type: "config",
      message: "Certificate mode requires MS365_TENANT_ID, MS365_CLIENT_ID, and MS365_CERT_PATH",
    })
  }

  return tryCredential(
    () =>
      new ClientCertificateCredential(config.tenantId, config.clientId, {
        certificatePath: config.certPath,
        certificatePassword: config.certPassword,
      }) as TokenCredential,
    "certificate",
  )
}

const createClientSecretCredential = (
  config: Extract<AuthConfig, { mode: "client-secret" }>,
): Either<AuthError, TokenCredential> => {
  if (!config.tenantId || !config.clientId || !config.clientSecret) {
    return Left<AuthError, TokenCredential>({
      type: "config",
      message: "Client secret mode requires MS365_TENANT_ID, MS365_CLIENT_ID, and MS365_CLIENT_SECRET",
    })
  }

  return tryCredential(
    () => new ClientSecretCredential(config.tenantId, config.clientId, config.clientSecret) as TokenCredential,
    "client-secret",
  )
}

export class ClientProvidedTokenCredential implements TokenCredential {
  private _accessToken: string | undefined
  private _expiresOn: Date

  constructor(accessToken?: string, expiresOn?: Date) {
    this._accessToken = accessToken
    this._expiresOn = expiresOn ?? (accessToken ? new Date(Date.now() + ONE_HOUR_MS) : new Date(0))
  }

  // eslint-disable-next-line @typescript-eslint/require-await -- TokenCredential interface requires async
  async getToken(_scopes: string | string[]): Promise<AzureAccessToken | null> {
    if (!this._accessToken || this._expiresOn <= new Date()) {
      return null
    }
    return { token: this._accessToken, expiresOnTimestamp: this._expiresOn.getTime() }
  }

  updateToken(token: string, expiresOn?: Date): void {
    this._accessToken = token
    this._expiresOn = expiresOn ?? new Date(Date.now() + ONE_HOUR_MS)
  }

  isExpired(): boolean {
    return !this._accessToken || this._expiresOn <= new Date()
  }

  getExpirationTime(): Date {
    return this._expiresOn
  }

  getAccessTokenValue(): string | undefined {
    return this._accessToken
  }
}

const createClientProvidedTokenCredential = (
  config: Extract<AuthConfig, { mode: "client-token" }>,
): Either<AuthError, TokenCredential> =>
  Right(new ClientProvidedTokenCredential(config.accessToken, config.expiresOn) as TokenCredential)

export const isClientProvidedToken = (credential: TokenCredential): credential is ClientProvidedTokenCredential =>
  credential instanceof ClientProvidedTokenCredential

export const testCredential = async (credential: TokenCredential): Promise<Either<AuthError, true>> => {
  // Skip test for client-provided token without initial token
  if (isClientProvidedToken(credential) && credential.isExpired()) {
    return Right(true as const)
  }

  try {
    const token = await credential.getToken(GRAPH_DEFAULT_SCOPE)
    if (!token) {
      return Left<AuthError, true>({ type: "token", message: "Failed to acquire token during credential test" })
    }
    return Right(true as const)
  } catch (e) {
    return Left<AuthError, true>({ type: "token", message: `Authentication test failed: ${String(e)}` })
  }
}
