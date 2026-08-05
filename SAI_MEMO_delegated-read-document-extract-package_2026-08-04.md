# SAI_MEMO — Delegated `read_document` + shared extract package (build handoff)

**Repo:** `sapientsai/microsoft365-mcp-server`
**Reference commits:** `bdaaf77` (survey), `fd00bec` (the `unpdf` fix — see §4.2a), `30d9670` (current HEAD). All three are pushed and reachable on origin.
**Date:** 2026-08-04
**Author:** Jordan Burke
**Status:** Ready to build. One wave, four changes, landed together, single coordinated redeploy afterward.

---

## 1. Why

A Finance user asked Claude to read a PDF in `OncalaBio-Finance_Private` and it could not, through any route. Root cause is a capability split, not a permission problem and not a PDF regression:

- The **delegated** server (`packages/microsoft365`, `ms365.civala.ai`) can see private sites, because it runs as the requesting user. It has no binary-to-text path at all. `download_file` is its only file-content tool and it returns metadata plus a URL for anything that is not text under 100 KB.
- The **app-only** servers (`packages/graph`, the seven `ms-mcp-*` connectors plus `ms-mcp-admin`) have `read_document` with real PDF/DOCX/XLSX extraction, but are scoped by `Sites.Selected` and cannot see private sites without a per-site grant.

Granting app-only into private Finance content would let any consumer of that connector read finance documents regardless of their own SharePoint permissions, which cuts against the delegated-auth-for-audit-trail rule. So the fix belongs on the delegated server.

**Verified empirically before writing this:**

| Check | Result |
| --- | --- |
| `read_document` on app-only against a SharePoint PDF | Works. Extracted full text of a 2-page, 39.3 KB contract. |
| App-only Graph resolve `/sites/civala.sharepoint.com:/sites/OncalaBio` | 200 |
| App-only Graph resolve `/sites/civala.sharepoint.com:/sites/OncalaBio-Finance_Private` | `Access denied` |
| Delegated `download_file` with a SharePoint drive item ID | `Failed to get file info: The resource could not be found.` (it is `/me/drive`-rooted) |
| Delegated tool list contains `read_document` | No |

---

## 2. Corrections to earlier planning

Two things in prior docs are wrong and will mislead whoever builds this.

**2.1 `SAI_PLAN_ms-graph-monorepo_2026-06-20.md` says the port is "cheap, since the capabilities already live in `core`."** True for `microsoft_graph_batch` and Azure AI Search. **Not true for `read_document`.** Extraction lives only in `packages/graph/src/extract/extract.ts`, and `mammoth` / `unpdf` / `exceljs` are declared only in `packages/graph/package.json`. `packages/core` has one runtime dependency, `functype`, no `extract/` directory, and no extraction exports. Confirmed three ways: core's barrel export, its `src/` tree, and a zero-hit code search for the parser names under `packages/core`.

**2.2 `buildReadDocumentTool` cannot be dropped into the delegated server.** The two packages use incompatible conventions:

| | `packages/graph` | `packages/microsoft365` |
| --- | --- | --- |
| Shell | somamcp | FastMCP |
| Tool shape | `buildXTool(auth \| graph)` factory, `server.addTool(...)` | module-level `const toolDefinitions: ReadonlyArray<ToolDefinition>` |
| Auth reach | injected `AuthStrategy` | module-global `requireClient()` returning the client or null |
| Return | `Promise<string>` | `Either<UserError, string>`, unwrapped by `unwrapResult` |

So `read_document` on the delegated side is a **new native tool following local convention**, sharing only the extraction module.

---

## 3. Decisions

**3.1 Extraction moves to a new `packages/extract`, decoupled from core's error type.** Options weighed:

| Option | Verdict |
| --- | --- |
| Add `mammoth`/`unpdf`/`exceljs` to `packages/microsoft365` | Rejected. Duplicates the module and contradicts the stated intent in `extract.ts`'s own header comment. |
| Leave it in `packages/graph`, depend on that from `microsoft365` | Rejected. Makes the delegated server depend on a *server* package, pulling in the somamcp shell to get a text extractor. Layering inversion. |
| `packages/core`, with a `./extract` subpath deliberately absent from the barrel | Rejected. Makes the workspace's most-imported package expensive to import. See below. |
| **New `packages/extract` (chosen)** | Keeps core cheap to import. See below. |

One argument that does **not** justify avoiding core, recorded so nobody re-litigates it:

- "It pollutes core's dependencies." Dead. `@sapientsai/ms-graph-core` is `"private": true` at version `0.0.0` and is never published. Its consumers are exactly the two server packages, and after this change both need extraction. `pnpm-workspace.yaml` sets `injectWorkspacePackages: true`, so workspace deps hard-link into the deployed `node_modules` and the parsers land on disk in the delegated image regardless of which package declares them. Lazy loading controls startup and bundle, never disk.

The argument that **does** decide it:

**Import weight.** Both servers import `@sapientsai/ms-graph-core` eagerly, at module load. Today that costs nothing — core's only dependency is `functype`. Put extraction inside core, even behind a subpath the barrel does not export, and "importing core is free" stops being a fact and becomes something each contributor has to check: *which part of core am I pulling, and does it drag three parsers into startup?* That is a permanent tax on the most-imported package in the repo, paid to avoid one `package.json`.

The parsers also churn on their own schedule — `unpdf` shipped a breaking change mid-wave (§4.2a). Behind their own boundary, a parser bump touches one manifest and cannot perturb core. And "core is eager, extract is lazy" is a fact about packages, which survives contributor turnover better than "this subpath of core must never enter the barrel," which is a convention someone has to remember.

Cohesion points the same way — every current member of core is Graph-specific (`auth-strategy`, `graph-request`, `constants`, `types`, `upload/`, `odata-helpers`, `pagination`, `upload-helpers`), and a document parser would be the lone exception in a package described as "Shared Microsoft Graph plumbing." But treat that as corroboration, not the reason: a package description can always be rewritten.

**One argument previously listed here has been withdrawn.** An earlier draft claimed the separate package is "safe by construction," because nobody can widen a barrel that does not exist. That is weaker than it sounds. A separate package does not prevent the actual failure — a top-level `import { extractTextFromBuffer } from "@sapientsai/document-extract"` in the delegated server breaks lazy loading exactly as well, and no package boundary catches it. What catches it is the startup assertion in §9.7, which is required either way. Do not skip §9.7 on the theory that the package split already bought you that guarantee. It did not.

**Corollary — drop the `GraphApiError` return type.** `extractTextFromBuffer` currently returns `Either<GraphApiError, string>`, which is an artifact of the original port, not a real coupling: a document parser has no business returning a Graph error. Give it its own error type so the package depends on nothing but `functype` and has no relationship to core at all.

**On the name.** The package is `@sapientsai/document-extract`, not `ms-graph-extract`. Once the `GraphApiError` corollary lands, the module's imports are `node:path`, `exceljs`, `functype`, `mammoth`, `unpdf` — nothing Microsoft. A name asserting a Graph coupling would invite exactly the misreading this section exists to prevent. `packages/core` keeps its `ms-graph-` prefix because it has earned it.

**3.2 `download_file` is ported to `packages/graph`, not dropped.** `packages/graph/README.md` currently says it is "not ported — out of scope for this server's document-RAG purpose." That was a defensible design call, but two facts override it:

1. All eight deployed app-only containers currently expose `download_file`. They are running the pre-consolidation image. The moment they re-pull `:main`, the tool disappears from all eight simultaneously. Caller audit was attempted and could not be completed: GitHub code search returns zero results across `civala-ai` even for control queries, so the private-repo index is unusable. `vectorizer` and `vectorizer-ocr-rescue` are the suspected callers.
2. It is the correct fallback when extraction fails. `read_document` has no OCR (`unpdf` reads embedded text only), enforces per-format size caps (§3.4), and returns `Left` on unsupported content types. Scanned financial statements hit all three. Raw bytes, or a URL to them, are the escape hatch.

Update the README section to reflect the reversal and the reasoning.

**3.3 Lazy import is load-bearing, not cosmetic.** In `resolveFilterConfig`, `presets: process.env.MS365_PRESETS?.split(",")`. When `MS365_PRESETS` is unset, `filterTools` computes `allowedDomains` as `undefined` and every tool passes the preset check. So a new `rag` domain does **not** keep the parsers out of a default deployment. Only a dynamic `import()` inside `execute` does.

**3.4 Size limits are replaced, and the existing check is in the wrong place.** `read-document.ts` currently does this:

```ts
const buffer = Buffer.from(await response.arrayBuffer())

if (buffer.length > MAX_FILE_SIZE) {
  throw new Error(`File too large (${formatBytes(buffer.length)}). Maximum is ${formatBytes(MAX_FILE_SIZE)}.`)
}
```

The whole file is buffered into memory **before** the size check runs. For the exact scenario the cap exists to prevent, the `arrayBuffer()` call is where the process dies, so the check never executes. It is decorative for its own purpose. It also spends the full network transfer before rejecting.

Fix: read `size` and `file.mimeType` from the drive item metadata first, decide, then fetch content only if it passes. One extra Graph round trip, and it makes the cap actually load-bearing.

The `10 * 1024 * 1024` constant itself is inherited from the pre-consolidation `microsoft-mcp-server` (see the "Ported from" comment in the file header) with no recorded rationale. It is too low for this tenant: IND submissions and clinical study reports routinely run 50 to 200 MB. Replace with per-format caps, because memory cost is not uniform in file size:

| Format | Cap | Reason |
| --- | --- | --- |
| PDF | 100 MB | `unpdf` text extraction is roughly linear in input. |
| DOCX | 50 MB | A 50 MB Word file is almost entirely embedded images, which mammoth ignores for raw text. |
| XLSX | 25 MB | `exceljs` builds a full workbook object model, 10x to 20x file size in heap. This is the real bomb, not PDF. |
| text/* | 25 MB | Bounded by `max_chars` on output anyway. |

Single env override, `MS365_MAX_EXTRACT_BYTES` / `GRAPH_MAX_EXTRACT_BYTES`, applied as a ceiling over the per-format defaults so one variable can tighten everything in a constrained container.

Note that output size is already bounded independently by `max_chars` (50k default, 200k max) with an explicit truncation marker. The input cap protects container heap only. It has nothing to do with context window, and the two should not be tuned as if they were one number.

Also set `--max-old-space-size` explicitly in both Dockerfiles. Node's default old-space limit is derived from available memory, and `civala-ai-vm-01` is a 16 GB host running roughly 25 services, so the effective limit today is unpredictable. An explicit value turns a bad file into a catchable allocation failure instead of an OOM kill that drops every other session on that connector.

---

## 4. Change A — `packages/extract` (`@sapientsai/document-extract`)

Move `packages/graph/src/extract/` into the new package, carrying the **current** file contents (see §4.2a — do not copy from an older revision). Preserve the current exports: `extractTextFromBuffer`, `CONTENT_TYPE_MAP`, `EXTRACTABLE_TYPES`, `isTextContent`, `resolveContentType`. Move the extraction specs from `packages/graph/test/` across with the module. `packages/*` is already the workspace glob, so no `pnpm-workspace.yaml` change is needed.

There are exactly two intended deltas from the surveyed file: the error type in §4.1, and the already-committed `loadingTask.destroy()` form in §4.2a. Anything else differing is a mistake.

### 4.1 Signature change: own error type

Today:

```ts
): Promise<Either<GraphApiError, string>>
```

Becomes:

```ts
export type ExtractError = { readonly type: "parse" | "unsupported"; readonly message: string }

export const extractTextFromBuffer = async (
  buffer: Buffer,
  contentType: string,
  filename: string,
): Promise<Either<ExtractError, string>>
```

The internal `parseError` helper keeps its shape, just typed as `ExtractError`. The package then has no dependency on core at all.

**This costs nothing at the existing call site.** `packages/graph/src/tools/read-document.ts` currently does:

```ts
const fullText = extracted.fold(
  (error) => { throw new Error(error.message) },
  (text) => text,
)
```

It only reads `.message`, which `ExtractError` still has. So graph needs no change beyond the import path.

### 4.2 `package.json`

```json
{
  "name": "@sapientsai/document-extract",
  "version": "0.0.0",
  "private": true,
  "type": "module",
  "dependencies": {
    "exceljs": "^4.4.0",
    "functype": "1.4.4",
    "mammoth": "^1.12.0",
    "unpdf": "^1.8.0"
  },
  "devDependencies": {
    "@types/node": "^24.13.3",
    "ts-builds": "^3.2.3",
    "tsdown": "^0.22.14"
  }
}
```

Mirror core's `main`/`module`/`types`/`exports`/`files`/`scripts`/`engines` block. Keep `functype` at the exact `1.4.4` pin: per the note in `pnpm-workspace.yaml`, two copies of functype's structurally recursive `Either` blow the TS instantiation budget (TS2589), and the workspace `overrides` entry is what prevents that. A fourth package asking for a range instead of the pin is the way this reintroduces itself.

**Do not downgrade these pins to match an older draft of this memo.** They were re-read from `packages/graph/package.json` at `30d9670` and are unchanged since `fd00bec`; read the current file before writing the new one rather than trusting the numbers above. Note `fd00bec` also bumped `fastmcp` 4.12.2, `functype-log` 1.8.0, `somamcp` 1.1.1 and pnpm 11.20.0 elsewhere in the workspace, and `30d9670` is a lockfile-only follow-up; none of that touches the new package's dependencies.

### 4.2a `unpdf` 1.8 breaking change — the moved file is not the surveyed file

Between the survey commit (`bdaaf77`) and now, `unpdf` went 1.6.2 → 1.8.0 and the pdf.js it bundles **removed `PDFDocumentProxy.destroy()`**. Verified at runtime: `doc.destroy` is `undefined` in 1.8.0, `doc.loadingTask.destroy` is a function. Commit `fd00bec` fixed `extract.ts` accordingly:

```ts
} finally {
  await pdf.loadingTask.destroy()   // was: await pdf.destroy()
}
```

This was not a cosmetic type fix. Under 1.8.0 the old call throws a `TypeError` from the `finally` block on **every** PDF extraction, discarding the successfully extracted text.

Consequence for this wave: moving `extract.ts` "verbatim" from any pre-`fd00bec` checkout, or pinning `unpdf ^1.6.2` in the new package, resurrects the bug. Take the file at current HEAD and keep the pin at `^1.8.0`. `loadingTask.destroy()` exists in both versions, so the current form is safe either way.

This is also the template for the failure mode: it typechecked clean and broke only at runtime, which is why §9.1c asserts on extracted content rather than on a passing build.

### 4.3 `packages/graph`

Depend on `@sapientsai/document-extract`, change `read-document.ts`'s import from `../extract/extract` to the package, and **remove `exceljs`, `mammoth`, `unpdf` from `packages/graph/package.json`**. Graph imports it eagerly; it is the RAG server and always needs it.

---

## 5. Change B — port `download_file` to `packages/graph`

Follow the local factory convention and register in `buildServer` alongside its siblings.

**Do not stream bytes through the response by default.** Graph returns `@microsoft.graph.downloadUrl` on the drive item: short-lived (roughly one hour), pre-authenticated, needs no `Authorization` header. Return metadata plus that URL always, and inline base64 only under a small threshold. Base64 inflates by 4/3 and lands in model context, so a multi-megabyte inline return is the one place a size cap genuinely protects the caller rather than the container. The delegated `download_file` already works this way, so this also aligns the two servers.

```ts
export const buildDownloadFileTool = (graph: GraphRequest) => ({
  name: "download_file",
  description:
    "Get a file's metadata and a short-lived pre-authenticated download URL from SharePoint or " +
    "OneDrive. Small files are returned inline as base64. Use read_document first for text; use " +
    "this when extraction fails (scanned PDFs with no embedded text, unsupported content types, " +
    "files over the extraction cap) or when raw bytes are needed. The URL expires in about an " +
    "hour, so do not persist it.",
  parameters: z.object({
    path: z.string().describe("Graph path to the drive item (no trailing /content)"),
    api_version: z.enum(["v1.0", "beta"]).default("v1.0"),
    inline: z.boolean().default(true).describe("Inline small files as base64"),
  }),
  // execute: GET the item selecting id,name,size,file,@microsoft.graph.downloadUrl
  //   -> always return metadata + downloadUrl
  //   -> if inline && size <= INLINE_LIMIT (256 KB), fetch content and append base64
})
```

Note it takes `graph` (core's request helper) rather than `auth`, unlike `buildReadDocumentTool`. The metadata read goes through the normal JSON layer; only the optional inline fetch needs a raw request, and at 256 KB that is cheap.

Gate it: `GRAPH_ENABLE_DOWNLOAD_FILE`, defaulting **on**, so the Civala deployment keeps its current surface without extra env wiring.

Note from the graph README, still accurate and relevant: the somamcp-side blocker on content-array and image returns is resolved (`wrapTool` passes them through), so returning non-text content is viable.

---

## 6. Change C — `read_document` on `packages/microsoft365`

### 6.1 Registry — `src/tools/tool-registry.ts`

```diff
 export type ToolDomain =
   | "auth"
   ...
   | "todo"
   | "query"
+  | "rag"
```

```diff
 export const PRESETS: Record<string, ReadonlyArray<ToolDomain>> = {
   personal: [...],
   collaboration: [...],
   productivity: [...],
+  rag: ["rag", "files", "query"],
   all: [
     ...
     "todo",
     "query",
+    "rag",
   ],
 }
```

```diff
   // Query
   { name: "graph_query", domain: "query", readOnly: false, orgOnly: false },
+  // RAG
+  { name: "read_document", domain: "rag", readOnly: true, orgOnly: false },
 ]
```

`orgOnly: false` because the tool works against `/me/drive` as well as sites. `filterTools` always unions in `"auth"`, so no change needed there.

### 6.2 Implementation — new `src/tools/read-document-tools.ts`

Sketch, not verbatim. Two API facts below were verified against HEAD and are load-bearing; both were wrong in the previous draft of this memo.

**Fact 1 — `getAccessToken` returns an `Either`, not a nullable string.** `src/auth/auth-manager.ts:147`:

```ts
export const getAccessToken = async (): Promise<Either<AuthError, string>> => { ... }
```

So a truthiness guard is always dead code — an `Either` is an object and always truthy — and interpolating the result into a header yields `Bearer [object Object]`. Use the `isLeft()` shape that `src/client/graph-client.ts:151` already uses.

**Fact 2 — the client comes from `requireClient()`, not from thin air.** `src/tools/files-tools.ts:98` is the local pattern. Raw bytes still need `fetch` with an explicit token, because the JSON request layer will not return a buffer; that is the same reason `packages/graph` does it this way. So this tool needs *both* the client (metadata via `graphQuery`) and a token (content via `fetch`).

`client.graphQuery`'s real signature, `src/client/graph-client.ts:419`:

```ts
graphQuery<T>(method: string, path: string, body?: Record<string, unknown>, version?: GraphApiVersion, headers?: Record<string, string>)
```

**It also returns an `Either`.** The sketch below elides that unwrap for brevity, but do not carry the elision into the code: reading `.size` or `.file` straight off the return value is the identical mistake Fact 1 describes. Unwrap it the way every other call site in `files-tools.ts` does.

```ts
import { UserError } from "fastmcp"
import { type Either, Left, Right } from "functype/either"

import { getAccessToken } from "../auth"
import { GRAPH_API_BASE } from "../auth/scopes"
import { requireClient } from "./<wherever requireClient lives — see files-tools.ts>"

export const readDocument = async (params: {
  path: string
  api_version?: "v1.0" | "beta"
  format?: string
  max_chars?: number
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const tokenResult = await getAccessToken()
  if (tokenResult.isLeft()) {
    return Left(new UserError((tokenResult.value as { message: string }).message))
  }
  const token = tokenResult.value as string

  const version = params.api_version ?? "v1.0"

  // Metadata FIRST — see §3.4. Checking size after arrayBuffer() is too late.
  // Derive the item path by stripping a trailing /content from params.path.
  // NB: graphQuery returns Either — unwrap before touching .size / .file.
  const meta = await client.graphQuery("GET", itemPathFrom(params.path), undefined, version)
  //   -> select id,name,size,file
  const cap = capForMimeType(meta.file?.mimeType) // §3.4 table, ceiling from env
  if ((meta.size ?? 0) > cap) {
    return Left(new UserError(
      `File is ${formatBytes(meta.size)}, over the ${formatBytes(cap)} extraction cap for this ` +
      `format. Use download_file to get a download URL instead.`,
    ))
  }

  const query = params.format ? `?format=${params.format}` : ""
  const response = await fetch(`${GRAPH_API_BASE}/${version}${params.path}${query}`, {
    headers: { Authorization: `Bearer ${token}` },
  })
  if (!response.ok) return Left(new UserError(/* mirror read-document.ts error branching */))

  const buffer = Buffer.from(await response.arrayBuffer())

  // Lazy: keeps mammoth/unpdf/exceljs out of the startup path. See §3.3.
  const { extractTextFromBuffer } = await import("@sapientsai/document-extract")
  const extracted = await extractTextFromBuffer(buffer, contentType, filename)
  // ... truncate to max_chars, return Right(`File: ${filename} (${formatBytes(...)})\n\n${text}`)
}
```

**Critical correctness requirement.** In oauth-proxy mode the access token arrives per-request through the session and is placed in an async token context by `withToken(context.session?.accessToken, ...)` in `wrapExecute`. The token **must** be obtained via `getAccessToken` from `./auth` so it resolves from that context — confirmed at `auth-manager.ts:149`, which checks `getContextToken()` before falling back to the credential path. Do not import a credential-based token path or reach for `@azure/identity` directly. That would appear to work in standard mode and silently break oauth-proxy, which is how `ms365.civala.ai` runs.

**Reuse the prior art.** `files-tools.ts` already has `MAX_INLINE_SIZE` and `isTextFile`. Extend those rather than defining parallel constants with different names.

### 6.3 Definition — `src/index.ts`

Add to `toolDefinitions`, next to the Files block, and export `readDocument` from the `./tools` barrel:

```ts
  {
    name: "read_document",
    description:
      "Download a file from SharePoint or OneDrive and return its readable text content. Supports " +
      "DOCX, PDF, XLSX, and text-based files. Use instead of download_file when you need document " +
      "contents. Pair with search_site_files (SharePoint) or search_files (OneDrive) to get IDs, " +
      "then pass /drives/{driveId}/items/{itemId}/content or /me/drive/items/{id}/content. " +
      "Text extraction only, no OCR: scanned PDFs return no text.",
    parameters: z.object({
      path: z.string().describe("Graph path to the file content endpoint, ending in /content"),
      api_version: z.enum(["v1.0", "beta"]).optional(),
      format: z.string().optional(),
      max_chars: z.number().int().min(1000).max(200000).optional(),
    }),
    execute: async (params) => unwrapResult(await readDocument(params)),
    domain: "rag",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
```

The "pair with `search_site_files`" clause matters. Discovery on this server is split: `search_files` is `/me/drive`-rooted and `search_site_files` takes a `site_id`. Without the hint, callers guess and fail.

---

## 7. Change D — widen delegated `download_file`

Current shape takes only `item_id` and resolves against `/me/drive` — `graph-client.ts:148` is literally `` request(`GET`, `/me/drive/items/${id}`) `` — which is the error the Finance request produced. Add an optional `drive_id`, and when present resolve `/drives/{drive_id}/items/{item_id}` instead. Keep the existing single-argument behavior working so nothing breaks.

Also worth adjusting: the tool's description promises "content inline for text files under 100KB, otherwise returns metadata and download URL" without saying what to do next. Point it at `read_document`.

---

## 8. Non-goals

- OCR. Out of scope. `vectorizer-ocr-rescue` exists on the cluster and is the right owner if this becomes a real need.
- Switching XLSX to `exceljs`'s streaming `WorkbookReader`. Separate ticket. It would let the 25 MB spreadsheet cap rise substantially, since the object-model expansion is the only reason that cap is the tightest of the four.
- Porting `microsoft_graph_batch` or Azure AI Search to the delegated server. Separate tickets, genuinely cheap since they sit on core's request layer.
- Any `Sites.Selected` grant for private FedCo sites. Explicitly rejected in §1.

---

## 9. Test plan

1. Extraction unit tests pass after the move into `packages/extract`, unchanged apart from the error type (they were written against `buildReadDocumentTool` with an injected `fetchImpl`, so keep that seam).
1b. `pnpm typecheck` across the workspace, specifically to confirm the new package's `functype` pin did not reintroduce TS2589. That failure mode presents as an instantiation-depth error in an unrelated file, so do not attribute it to whatever file it points at.
1c. PDF extraction returns text end to end, not just a passing typecheck. §4.2a's bug typechecked clean under `unpdf` 1.6.2 and failed only at runtime, so assert on extracted content.
2. `packages/graph`: `read_document` still extracts; `download_file` returns metadata and a working download URL, with base64 only under the inline threshold; with `GRAPH_ENABLE_DOWNLOAD_FILE=false` the tool is absent from the listing.
3. `packages/microsoft365` in **stdio/credential** mode: `read_document` extracts a PDF from OneDrive.
4. `packages/microsoft365` in **oauth-proxy** mode: same, against a SharePoint site path, confirming the per-request session token resolves. This is the case that catches a wrong token source.
5. Negative cases: a file over its per-format cap is rejected **from metadata, without transferring the body** (verify by network trace, not just by the error message, since the whole point of §3.4 is that the old check ran too late); a scanned PDF returns a clear no-embedded-text message rather than an empty string; an unsupported content type surfaces `extractTextFromBuffer`'s `Left`.
5b. Size policy: confirm a 60 MB PDF extracts (previously rejected at the 10 MB cap), a 30 MB XLSX is rejected against the 25 MB format cap, and `MS365_MAX_EXTRACT_BYTES=5242880` tightens all formats to 5 MB.
5c. `download_file` returns a working `@microsoft.graph.downloadUrl` for a large file without transferring it, and inlines base64 only under 256 KB.
6. Preset behavior: with `MS365_PRESETS` unset, `read_document` is present. With `MS365_PRESETS=personal`, absent. With `MS365_PRESETS=rag`, present.
6b. Also assert `MS365_PRESETS=""`. It currently splits to `[""]`, length 1, so `allowedDomains` collapses to `{"auth"}` and every non-auth tool vanishes. Pre-existing behavior, not introduced by this wave, but the unset and `personal` cases above both miss it.
7. Startup check: confirm the three parsers are not loaded until first `read_document` call.

---

## 10. Deploy sequencing

Land all four changes before any redeploy. Then one coordinated pass, which is currently blocked on a separate issue and should stay blocked until this lands, to avoid churning the graph image twice.

1. ~~Regenerate the infrastructure index first.~~ **Done 2026-08-05** (`civ-devops@0fe11ba`). It found more than a stale version row — see §10.1 below. Re-read it before planning anything here.
2. **Bump the `ms365` pinned tag.** This is the only part of the rollout that is ready. `ms365` tracks an explicit version tag, is currently `1.0.31`, and runs the delegated server this wave actually changed.
3. **The eight app-only containers are deferred to a dedicated window.** Not part of this rollout. See §10.2.
4. Parity test, scoped to what shipped: `ms365` lists `read_document`, and `download_file` there accepts `drive_id`. The app-only `download_file` parity check moves to whenever step 3 happens.

### 10.1 What regenerating the index actually found

The Dokploy panel had moved from `cluster.civala.ai` to `cluster.civalaos.com` with no redirect — the old host returns a bare Traefik 404, which reads like a bad path or key rather than a moved service. Beyond that: `ms365` was already at `1.0.31` (not the recorded `1.0.21`), the Data project's Postgres is gone so Azure `civala-pdos-db` is the only managed database, `civala-api` moved to `api.civalaos.com`, and `vectorizer-ocr-rescue` is idle.

### 10.2 Why the app-only redeploy is not a redeploy

The instruction this memo originally carried — "redeploy the eight app-only containers to `:main`" — was wrong twice over, and acting on it would have been worse than doing nothing.

**It would not have worked.** All eight services reference an immutable image digest, `microsoft-mcp-server@sha256:8886eef2…318b`, not a tag. A digest reference reproduces identical bytes on every pull, so redeploying is a no-op that reports success. This is stronger than the "static images do not auto-pull" caveat already in the index.

**And if it had worked, it would have been a codebase swap.** `docker-graph.yml` publishes `packages/graph` to `ghcr.io/sapientsai/microsoft-mcp-server` — the same image name those eight containers use — on `v*` tags. So that name now carries two unrelated codebases: the archived pre-consolidation repo (the digest in production) and the consolidated monorepo rewrite (the semver tags). Pointing the eight at any current tag moves them from a known-good old build to a rewrite that **has never been tested as a drop-in replacement**, across all eight subsidiary connectors simultaneously.

Only `ms365` is current. It is a different image name, tracks a version tag, and is the server this wave changed.

So the app-only migration is its own project with its own window, not a step at the end of this one. It needs at minimum: one connector cut over and exercised before the rest, a rollback path (the outgoing digest is the rollback target — record it), and a decision about the shared image name, which is the underlying hazard. Two codebases behind one name will cause this again.

---

## 11. Known limits to document in both READMEs

1. No OCR. Scanned PDFs yield nothing from `read_document`; fall back to `download_file`.
2. Per-format extraction caps (§3.4): 100 MB PDF, 50 MB DOCX, 25 MB XLSX, 25 MB text, with an `MS365_MAX_EXTRACT_BYTES` ceiling. Over the cap, `download_file` returns a download URL instead.
3. Delegated discovery is split between `search_files` (OneDrive) and `search_site_files` (per-site). Neither is a tenant-wide KQL search. `sharepoint_search` on the app-only server is the only Graph Search API path in the monorepo.
4. `@sapientsai/ms-graph-core` is a `devDependency` in `packages/microsoft365` despite runtime imports, presumably relying on tsdown bundling. Pre-existing oddity, worth a look but not part of this wave. Note that `@sapientsai/document-extract` cannot follow that pattern, because it is loaded through a runtime `import()` rather than bundled at build time, so it belongs in `dependencies`.

---

## Appendix — verification log

Structural claims were re-checked against the tree rather than carried over from the survey. Rows marked `fd00bec` were verified by Claude Code against a working checkout; rows marked `origin` were independently re-verified against the pushed remote.

**On commit provenance.** `fd00bec` is pushed and reachable on origin, as is `bdaaf77`. It is no longer the tip, though: origin/main is now `30d9670` ("bump"), a lockfile-only commit touching `pnpm-lock.yaml` alone (+4/−26). No source file, dependency range, or line number cited below moved between `fd00bec` and `30d9670`, so every row still resolves at current HEAD. The §4.2 pins were re-read at `30d9670` and are unchanged.

| § | Claim | Verified |
| --- | --- | --- |
| 2.1 | Extraction only in `packages/graph`; core has no parsers | `packages/graph/src/extract/extract.ts` is the sole file; core `src/` is 10 Graph-specific files; zero grep hits for `mammoth\|unpdf\|exceljs` under `packages/core`; core's only dep is `functype` (`fd00bec` + `origin`) |
| 2.2 | Convention mismatch between the packages | Confirmed both directions (`fd00bec`) |
| 3.1 | `injectWorkspacePackages`, functype pin, TS2589 rationale | All three present verbatim in `pnpm-workspace.yaml` (`fd00bec` + `origin`) |
| 3.1 | Core is `"private": true`, version `0.0.0`, single barrel export | `packages/core/package.json` (`origin`) |
| 3.3 | Lazy import load-bearing | `tool-registry.ts:158` gates on `config.presets && config.presets.length > 0`; unset → `allowedDomains` undefined → all tools pass (`fd00bec` + `origin`) |
| 3.4 | Size check runs after buffering | `read-document.ts:54` buffers, `:56` checks (`fd00bec`) |
| 4.1 | Call site reads only `.message` | `read-document.ts:63` (`fd00bec` + `origin`) |
| 4.2 | Dependency pins current | Re-read from `packages/graph/package.json` at `30d9670`: `unpdf ^1.8.0`, `mammoth ^1.12.0`, `exceljs ^4.4.0`, `functype 1.4.4`, `ts-builds ^3.2.3`, `tsdown ^0.22.14`, `@types/node ^24.13.3` |
| 4.2a | `unpdf` 1.8 removed `PDFDocumentProxy.destroy()` | Runtime check: `doc.destroy === undefined`, `doc.loadingTask.destroy` is a function (`fd00bec`). Commit message on `fd00bec3b5d7403d55ea1c47d727589f47352c71` confirms the diagnosis and the `+1/-1` to `extract.ts`; commit is pushed and reachable on origin. |
| 6.2 | `getAccessToken` returns `Either<AuthError, string>` | `auth-manager.ts:147`; context token checked first at `:149` (`fd00bec` + `origin`). Note this signature also holds at the survey commit `bdaaf77`, so the earlier draft's nullable-string sketch was wrong when written, not invalidated by a later change. |
| 6.2 | `graphQuery` signature | `graph-client.ts:419` (`fd00bec`) |
| 7 | Delegated `download_file` is `/me/drive`-rooted, `item_id` only | `graph-client.ts:148`; params at `index.ts:626` (`fd00bec`) |
