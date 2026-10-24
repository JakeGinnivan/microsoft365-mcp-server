# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a TypeScript library template designed to be cloned/forked for creating new npm packages. It uses the `ts-builds` toolchain for standardized build scripts with ESM output.

**Template Usage**: See STANDARDIZATION_GUIDE.md for applying this pattern to other TypeScript projects.

## Environment variable compatibility

When adding or changing tools, routes, scripts, authentication, or Graph requests, follow the [environment variable compatibility rules in AGENTS.md](AGENTS.md#environment-variable-compatibility).

## Development Commands

All commands delegate to `ts-builds` for consistency across projects:

```bash
pnpm validate        # Main command: format + lint + test + build (use before commits)

pnpm format          # Format code with Prettier
pnpm format:check    # Check formatting only

pnpm lint            # Fix ESLint issues
pnpm lint:check      # Check ESLint issues only

pnpm test            # Run tests once
pnpm test:watch      # Run tests in watch mode
pnpm test:coverage   # Run tests with coverage

pnpm build           # Production build (outputs to dist/)
pnpm dev             # Development build with watch mode

pnpm typecheck       # Check TypeScript types
```

### Running a Single Test

```bash
pnpm test -- --testNamePattern="pattern"    # Filter by test name
pnpm test -- test/specific.spec.ts          # Run specific file
```

## Architecture

### Build System: ts-builds + tsdown

- **ts-builds**: Centralized toolchain package providing all build scripts
- **tsdown**: Underlying bundler configured via `ts-builds/tsdown`
- **Configuration**: `tsdown.config.ts` imports default config from ts-builds
- **TypeScript**: `tsconfig.json` extends `ts-builds/tsconfig`
- **Prettier**: Uses `ts-builds/prettier` shared config

### Output Format

- **dist/**: Production builds containing:
  - `index.js` - ES module format
  - `index.d.ts` - TypeScript declarations
- **lib/**: Development builds (also published)

### Package Exports

```json
{
  "main": "./dist/index.js",
  "module": "./dist/index.js",
  "types": "./dist/index.d.ts",
  "exports": {
    ".": {
      "types": "./dist/index.d.ts",
      "import": "./dist/index.js",
      "default": "./dist/index.js"
    }
  }
}
```

### Testing: Vitest

- Tests located in `test/*.spec.ts`
- Uses Vitest with configuration from ts-builds
- Coverage via v8 provider

## Key Files

- `src/index.ts` - Main library entry point
- `test/*.spec.ts` - Test files
- `tsdown.config.ts` - Build config (imports from ts-builds)
- `tsconfig.json` - TypeScript config (extends ts-builds)
- `.claude/skills/ts-builds-template/` - Claude Code skill for bootstrapping libraries from this template

## Publishing

Bump `packages/microsoft365/package.json`, re-sync `server.json`, then commit and tag:

```bash
cd packages/microsoft365
npm version patch|minor|major --no-git-tag-version
pnpm sync:server-json
cd ../..
git commit -am "chore(release): microsoft365-mcp-server v<new>"
git tag v<new> && git push --follow-tags
```

`npm version` cannot commit or tag from here — `packages/microsoft365` is not a git
root, so npm does not see a repository and skips both. That is why the bump is split
from the commit, and why `sync:server-json` is a step rather than a `version` lifecycle
hook: the hook would never fire.

The tag drives two workflows: `publish.yml` publishes to npm, and `publish-registry.yml`
publishes `server.json` to the MCP Registry. The `prepublishOnly` hook runs `pnpm validate`
and `pnpm check:server-json` before npm accepts the package, and CI runs the same
`check:server-json` on every PR so drift fails there instead of on the tag.
