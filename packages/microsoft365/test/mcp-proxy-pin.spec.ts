import { readFileSync } from "node:fs"
import { dirname, join } from "node:path"
import { fileURLToPath } from "node:url"

import { describe, expect, it } from "vitest"

// Tripwire for the mcp-proxy override in pnpm-workspace.yaml (see the comment there for why it
// exists). The override holds a *transitive* dependency down, which means nothing else in this
// repo can notice it: dependabot only watches direct dependencies, `pnpm outdated` is masked by the
// override, and the suite passes on either side of the pin. A transitive pin is silent by
// construction.
//
// The failure this guards is not "we kept the pin too long" — that is merely stale. It is the pin
// turning *harmful*: if fastmcp moves to a range that excludes 6.6.x, our override silently
// DOWNGRADES it, and the next person debugs a fresh mystery with no hint that a pin caused it.
// This fails loudly at that moment instead.

const HERE = dirname(fileURLToPath(import.meta.url))
const WORKSPACE_YAML = join(HERE, "..", "..", "..", "pnpm-workspace.yaml")
const FASTMCP_PKG = join(HERE, "..", "node_modules", "fastmcp", "package.json")

type Version = readonly [number, number, number]

const parseVersion = (raw: string): Version => {
  const parts = raw.replace(/^[^~]/, (c) => (/\d/.test(c) ? c : "")).match(/(\d+)\.(\d+)\.(\d+)/)
  if (!parts) throw new Error(`Unparseable version: ${raw}`)
  return [Number(parts[1]), Number(parts[2]), Number(parts[3])]
}

const compare = (a: Version, b: Version): number => a[0] - b[0] || a[1] - b[1] || a[2] - b[2]

/**
 * Whether `version` falls inside `range`, for the subset of npm range syntax that actually appears
 * here: `^x.y.z`, `~x.y.z`, and a bare version. Deliberately not a semver library — this file must
 * not add a dependency to assert something about dependencies.
 */
export const satisfies = (version: string, range: string): boolean => {
  const v = parseVersion(version)
  const floor = parseVersion(range)

  if (compare(v, floor) < 0) return false
  if (range.startsWith("^")) return v[0] === floor[0]
  if (range.startsWith("~")) return v[0] === floor[0] && v[1] === floor[1]
  return compare(v, floor) === 0
}

const overriddenRange = (): string => {
  const yaml = readFileSync(WORKSPACE_YAML, "utf-8")
  const match = yaml.match(/^\s+mcp-proxy:\s*"([^"]+)"/m)
  if (!match?.[1]) {
    throw new Error(
      "No mcp-proxy override found in pnpm-workspace.yaml. If it was removed on purpose because " +
        "punkpeye/mcp-proxy#96 is fixed, delete this spec too.",
    )
  }
  return match[1]
}

const fastmcpDeclaredRange = (): { version: string; range: string } => {
  // Read from disk rather than require(): fastmcp's exports map does not expose ./package.json.
  const pkg = JSON.parse(readFileSync(FASTMCP_PKG, "utf-8")) as {
    version: string
    dependencies: Record<string, string>
  }
  return { version: pkg.version, range: pkg.dependencies["mcp-proxy"] ?? "" }
}

describe("satisfies", () => {
  it.each([
    ["6.6.0", "^6.4.6", true],
    ["6.6.0", "^6.6.0", true],
    ["6.6.0", "^6.8.0", false], // fastmcp moved past our pin
    ["6.6.0", "^7.0.0", false], // fastmcp moved to a new major
    ["6.6.0", "~6.6.0", true],
    ["6.7.0", "~6.6.0", false],
    ["6.6.0", "6.6.0", true],
  ])("%s vs %s -> %s", (version, range, expected) => {
    expect(satisfies(version, range)).toBe(expected)
  })
})

describe("mcp-proxy pin stays compatible with fastmcp", () => {
  it("pins a version fastmcp still accepts", () => {
    const pin = overriddenRange()
    const { version, range } = fastmcpDeclaredRange()

    expect(range, "fastmcp no longer declares an mcp-proxy dependency").not.toBe("")

    // The pin's floor is the lowest version the override can resolve to. If fastmcp's own range
    // excludes it, the override is now forcing fastmcp onto something it does not support.
    expect(
      satisfies(pin.replace(/^[~^]/, ""), range),
      `The mcp-proxy override (${pin}) is no longer inside the range fastmcp ${version} declares ` +
        `(${range}). The override is now DOWNGRADING fastmcp rather than protecting it.\n\n` +
        `Check punkpeye/mcp-proxy#96: if the crash is fixed, drop the override from ` +
        `pnpm-workspace.yaml and delete this spec. If it is not, the pin needs to move to a ` +
        `version that satisfies both.`,
    ).toBe(true)
  })
})
