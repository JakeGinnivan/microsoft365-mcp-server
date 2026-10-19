import type { AuthConfig } from "../types"

export const resolveAuthConfig = (): AuthConfig => {
  const mode = process.env.MS365_AUTH_MODE ?? "interactive"
  const tenantId = process.env.MS365_TENANT_ID ?? "common"
  const clientId = process.env.MS365_CLIENT_ID ?? ""

  switch (mode) {
    case "certificate":
      return {
        mode: "certificate",
        tenantId,
        clientId,
        certPath: process.env.MS365_CERT_PATH ?? "",
        certPassword: process.env.MS365_CERT_PASSWORD,
      }
    case "client-secret":
      return {
        mode: "client-secret",
        tenantId,
        clientId,
        clientSecret: process.env.MS365_CLIENT_SECRET ?? "",
      }
    case "client-token":
      return {
        mode: "client-token",
        accessToken: process.env.MS365_ACCESS_TOKEN,
      }
    case "oauth-proxy":
      return {
        mode: "oauth-proxy",
        tenantId,
        clientId,
        clientSecret: process.env.MS365_CLIENT_SECRET ?? "",
        baseUrl: process.env.MS365_OAUTH_BASE_URL ?? "http://localhost:3000",
      }
    default:
      return {
        mode: "interactive",
        tenantId,
        clientId,
        redirectUri: process.env.MS365_REDIRECT_URI,
        useDeviceCode: process.env.MS365_USE_DEVICE_CODE === "true",
      }
  }
}
