import { formatBytes, type GraphRequest } from "@sapientsai/ms-graph-core"
import { z } from "zod"

const API_VERSIONS = ["v1.0", "beta"] as const

// Base64 inflates by 4/3 and lands in model context, so inlining is the one place a size cap protects
// the caller rather than the container. Above this, the pre-authenticated URL is the whole answer.
const INLINE_LIMIT = 256 * 1024

type DriveItem = {
  readonly id?: string
  readonly name?: string
  readonly size?: number
  readonly file?: { readonly mimeType?: string }
  readonly lastModifiedDateTime?: string
  readonly webUrl?: string
  readonly "@microsoft.graph.downloadUrl"?: string
}

const describeItem = (item: DriveItem): string =>
  [
    `Name: ${item.name ?? "(unnamed)"}`,
    `ID: ${item.id ?? "(unknown)"}`,
    `Size: ${formatBytes(item.size ?? 0)}`,
    `Type: ${item.file?.mimeType ?? "(unknown)"}`,
    item.lastModifiedDateTime ? `Modified: ${item.lastModifiedDateTime}` : undefined,
    item.webUrl ? `Web URL: ${item.webUrl}` : undefined,
  ]
    .filter((line) => line !== undefined)
    .join("\n")

// download_file: metadata plus a short-lived pre-authenticated download URL, with small files
// inlined as base64. The counterpart to read_document — that one returns text, this one returns
// the bytes (or a way to get them) when extraction cannot help.
//
// Takes `graph` rather than `auth`, unlike buildReadDocumentTool: the metadata read goes through the
// normal JSON layer, and the optional inline fetch uses @microsoft.graph.downloadUrl, which is
// pre-authenticated and must NOT carry an Authorization header — adding one makes Graph reject it.
export const buildDownloadFileTool = (graph: GraphRequest, fetchImpl: typeof fetch = fetch) => ({
  name: "download_file",
  description:
    "Get a file's metadata and a short-lived pre-authenticated download URL from SharePoint or OneDrive. " +
    "Small files are returned inline as base64. Use read_document first for text; use this when extraction " +
    "fails (scanned PDFs with no embedded text, unsupported content types, files over the extraction cap) or " +
    "when raw bytes are needed. The URL expires in about an hour, so do not persist it.",
  parameters: z.object({
    path: z.string().describe("Graph path to the drive item, e.g. /drives/{driveId}/items/{itemId} (no /content)"),
    api_version: z.enum(API_VERSIONS).default("v1.0").describe("Graph API version"),
    inline: z.boolean().default(true).describe("Inline the file as base64 when it is under 256 KB"),
  }),
  execute: async (args: {
    path: string
    api_version: (typeof API_VERSIONS)[number]
    inline: boolean
  }): Promise<string> => {
    // No $select here on purpose: @microsoft.graph.downloadUrl is an OData annotation, not a plain
    // property, and selecting it is inconsistent across API versions. A driveItem GET returns it by
    // default and the payload is small.
    const path = args.path.replace(/\/content\/?$/, "")
    const result = await graph.request<DriveItem>("GET", path, { version: args.api_version })

    const item = result.fold(
      (error) => {
        throw new Error(error.message)
      },
      (value) => value,
    )

    const downloadUrl = item["@microsoft.graph.downloadUrl"]
    const detail = describeItem(item)
    const header = downloadUrl
      ? `${detail}\nDownload URL (expires in ~1 hour): ${downloadUrl}`
      : `${detail}\nDownload URL: (not returned — the item may be a folder or have no content)`

    const size = item.size ?? 0
    if (!args.inline || !downloadUrl || size > INLINE_LIMIT) return header

    // Pre-authenticated: deliberately no Authorization header.
    const response = await fetchImpl(downloadUrl, { method: "GET" })
    if (!response.ok) {
      return `${header}\n\n[inline fetch failed: HTTP ${response.status} ${response.statusText} — use the URL above]`
    }
    const base64 = Buffer.from(await response.arrayBuffer()).toString("base64")
    return `${header}\n\n## Content (base64)\n\n${base64}`
  },
})
