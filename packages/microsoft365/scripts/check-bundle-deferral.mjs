#!/usr/bin/env node
// Post-build half of the guard in test/lazy-extraction.spec.ts.
//
// That spec asserts the *sources* keep document extraction behind a dynamic import. This script
// asserts the *bundler honored it*: the emitted bundle must still reach the extract package
// through `import(...)` and must not have inlined mammoth/unpdf/exceljs.
//
// It lives in the build step rather than the test suite on purpose. `ts-builds validate` runs
// test before build, so a vitest assertion about dist/ can only ever read a bundle from some
// earlier build — it skips on a clean checkout (dead in CI) and asserts against a stale artifact
// locally. Running right after `tsdown --clean` is the only point where the bundle on disk is
// known to correspond to the sources just checked.
//
// Wired two ways, both of which must stay in place:
//   - `ts-builds.config.json` appends `check:bundle` to the validate chain, so CI covers it.
//     That chain restates ts-builds' default steps because there is no append mechanism; if the
//     toolchain's default chain gains a step, add it there too.
//   - the `build` script in package.json, so an explicit `pnpm build` verifies what it emitted.

import { readFileSync } from "node:fs"
import { dirname, join, relative } from "node:path"
import { fileURLToPath } from "node:url"

const PACKAGE_ROOT = join(dirname(fileURLToPath(import.meta.url)), "..")
const BUNDLES = ["dist/index.js", "dist/bin.js"]

const EXTRACT_PACKAGE = "@sapientsai/document-extract"
const PARSERS = ["mammoth", "unpdf", "exceljs"]

// Comments are stripped so that a specifier named in a preserved comment or a sourcemap URL is not
// mistaken for an import. The `[^:]` guard keeps `https://` from being read as a line comment.
const stripComments = (code) => code.replace(/\/\*[\s\S]*?\*\//g, "").replace(/(^|[^:])\/\/.*$/gm, "$1")

const escape = (specifier) => specifier.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")

const countAll = (code, specifier) => code.match(new RegExp(`["']${escape(specifier)}["']`, "g"))?.length ?? 0

const countDynamic = (code, specifier) =>
  code.match(new RegExp(`\\bimport\\s*\\(\\s*["']${escape(specifier)}["']`, "g"))?.length ?? 0

const failures = []

for (const bundle of BUNDLES) {
  const path = join(PACKAGE_ROOT, bundle)

  let raw
  try {
    raw = readFileSync(path, "utf-8")
  } catch {
    failures.push(`${bundle} is missing — the build did not emit it, so the deferral is unverifiable.`)
    continue
  }

  const code = stripComments(raw)

  // Every reference that is not a dynamic `import(...)` is a static one. Counting rather than
  // matching import syntax stays robust to `export ... from`, `require(...)`, and minified forms.
  for (const parser of PARSERS) {
    const statics = countAll(code, parser) - countDynamic(code, parser)
    if (statics > 0) {
      failures.push(
        `${bundle} references "${parser}" outside a dynamic import (${statics}x) — the parser is on the startup path.`,
      )
    }
  }

  const dynamic = countDynamic(code, EXTRACT_PACKAGE)
  const statics = countAll(code, EXTRACT_PACKAGE) - dynamic

  if (statics > 0) {
    failures.push(
      `${bundle} references "${EXTRACT_PACKAGE}" outside a dynamic import (${statics}x) — extraction loads at startup.`,
    )
  }

  // Only index.js is expected to carry the import; bin.js must stay clear of it entirely.
  if (bundle === "dist/index.js" && dynamic === 0) {
    failures.push(
      `${bundle} contains no \`import("${EXTRACT_PACKAGE}")\` — the dependency was inlined by the bundler, ` +
        `which turns the dynamic import into a no-op. Keep it a runtime dependency, not a devDependency.`,
    )
  }
}

if (failures.length > 0) {
  console.error(`\n✗ bundle deferral check failed (${relative(process.cwd(), PACKAGE_ROOT) || "."}):\n`)
  for (const failure of failures) console.error(`  - ${failure}`)
  console.error("\nSee test/lazy-extraction.spec.ts for the source-level half of this guard.\n")
  process.exit(1)
}

console.log(`✔ bundle deferral check passed (${BUNDLES.join(", ")})`)
