import type { AccessToken as AzureAccessToken, TokenCachePersistenceOptions, TokenCredential } from "@azure/identity"
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
import {
  type AuthenticationRecordLike,
  fileCachePersistencePlugin,
  readAuthenticationRecord,
  writeAuthenticationRecord,
} from "./token-cache"

const ONE_HOUR_MS = 60 * 60 * 1000

const DEFAULT_REDIRECT_URI = "http://localhost:3000"

const TOKEN_CACHE_NAME = "microsoft365-mcp-server"

// useIdentityPlugin mutates module-level state in @azure/identity, so it must run once
// and before any credential is constructed. Ref holds the one-shot flag without a
// mutable binding, matching the functional style enforced across this package.
const cacheRegistered = Ref(false)

const enableTokenCachePersistence = (): boolean => {
  // Opt out with MS365_TOKEN_CACHE=false — e.g. a shared machine, or a CI run where a
  // persisted token would outlive the job.
  if (process.env.MS365_TOKEN_CACHE === "false") return false

  if (!cacheRegistered.get()) {
    useIdentityPlugin(fileCachePersistencePlugin)
    cacheRegistered.set(true)
  }
  return true
}

/** Test seam: reset the one-shot plugin registration. */
export const resetTokenCacheRegistrationForTests = (): void => cacheRegistered.set(false)

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

const createDeviceCodeCredential = (
  tenantId: string | undefined,
  clientId: string,
  tokenCachePersistenceOptions: TokenCachePersistenceOptions | undefined,
): TokenCredential =>
  new DeviceCodeCredential({
    tenantId,
    clientId,
    tokenCachePersistenceOptions,
    userPromptCallback: (info) => {
      console.error(`\nAuthentication Required:`)
      console.error(`Please visit: ${info.verificationUri}`)
      console.error(`And enter code: ${info.userCode}\n`)
    },
  }) as TokenCredential

// An AuthenticationRecord only exists once a sign-in has completed, so it is captured
// on the way past rather than up front, and written once — it does not change, and
// rewriting it per token acquisition would be pointless disk traffic.
//
// The record is reached through the credential's internal msalClient. That is private
// API, so it is probed defensively: if the shape ever changes, the record simply is not
// saved and the next start signs in again, which is the current behaviour anyway. A
// missing convenience must not become a failed authentication.
const readActiveAccount = (credential: TokenCredential): AuthenticationRecordLike | undefined => {
  const client = (credential as { msalClient?: { getActiveAccount?: () => unknown } }).msalClient
  const account = client?.getActiveAccount?.()
  const record = account as Partial<AuthenticationRecordLike> | undefined
  return record?.homeAccountId && record.username && record.clientId ? (record as AuthenticationRecordLike) : undefined
}

const rememberAccount = (credential: TokenCredential): TokenCredential => {
  const saved = Ref(false)
  return {
    getToken: async (scopes, options) => {
      const token = await credential.getToken(scopes, options)
      if (!saved.get()) {
        const record = readActiveAccount(credential)
        if (record) {
          writeAuthenticationRecord(record)
          saved.set(true)
        }
      }
      return token
    },
  }
}

const createInteractiveCredential = (
  config: Extract<AuthConfig, { mode: "interactive" }>,
): Either<AuthError, TokenCredential> => {
  const { tenantId, clientId } = config

  if (!clientId) {
    return Left<AuthError, TokenCredential>({ type: "config", message: "Interactive mode requires MS365_CLIENT_ID" })
  }

  // Both flows below share this: whichever one is used, the resulting token survives
  // a restart, which is the entire point of the interactive mode being usable at all.
  const tokenCachePersistenceOptions = enableTokenCachePersistence()
    ? { enabled: true, name: TOKEN_CACHE_NAME }
    : undefined

  // Opt out of the browser entirely: headless hosts, SSH sessions, or a machine whose
  // default browser cannot be launched on demand.
  if (config.useDeviceCode) {
    return tryCredential(
      () => createDeviceCodeCredential(tenantId, clientId, tokenCachePersistenceOptions),
      "device code",
    )
  }

  // Without this the persisted cache is unreachable: the credential raises
  // AuthenticationRequiredError before consulting it, because it does not know which
  // account to look for.
  const authenticationRecord = tokenCachePersistenceOptions ? readAuthenticationRecord() : undefined

  return tryCredential(() => {
    const browser = new InteractiveBrowserCredential({
      tenantId,
      clientId,
      redirectUri: config.redirectUri ?? DEFAULT_REDIRECT_URI,
      tokenCachePersistenceOptions,
      authenticationRecord,
    }) as TokenCredential

    return tokenCachePersistenceOptions ? rememberAccount(browser) : browser
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
