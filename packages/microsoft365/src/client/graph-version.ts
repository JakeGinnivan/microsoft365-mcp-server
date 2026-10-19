import type { GraphApiVersion } from "../types"

export const defaultVersion = (): GraphApiVersion => (process.env.MS365_GRAPH_VERSION === "beta" ? "beta" : "v1.0")
