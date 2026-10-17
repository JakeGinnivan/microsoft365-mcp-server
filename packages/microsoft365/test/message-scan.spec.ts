import { beforeEach, describe, expect, it } from "vitest"

import type { GraphMessage } from "../src/types"
import { formatMessageScan, formatMessageScanRow } from "../src/utils/formatters"
import {
  clearMessageRefs,
  messageRefCount,
  rememberMessageId,
  resolveMessageIdOrRef,
  resolveMessageRef,
} from "../src/utils/message-refs"

const message = (overrides: Partial<GraphMessage> = {}): GraphMessage => ({
  id: "AAMkAGI0YjA3OTNhLWY2MDEtNGZlYy1hNzU2LTE4NDFiODg5ZjliMgBGAAAAAABtR",
  subject: "Test Subject",
  from: { emailAddress: { name: "John", address: "john@example.com" } },
  receivedDateTime: "2026-01-15T10:30:45Z",
  isRead: true,
  hasAttachments: false,
  ...overrides,
})

describe("message refs", () => {
  beforeEach(() => clearMessageRefs())

  it("assigns stable refs, reusing the same ref for the same id", () => {
    const first = rememberMessageId("id-a")
    const second = rememberMessageId("id-b")

    expect(rememberMessageId("id-a")).toBe(first)
    expect(second).not.toBe(first)
    expect(messageRefCount()).toBe(2)
  })

  it("starts refs at 1 so a ref is never falsy", () => {
    expect(rememberMessageId("id-a")).toBe(1)
  })

  it("round-trips a ref back to its message id", () => {
    const ref = rememberMessageId("id-a")
    expect(resolveMessageRef(ref)).toBe("id-a")
  })

  it("passes a full Graph id straight through", () => {
    const graphId = "AAMkAGI0YjA3OTNhLWY2MDEtNGZlYy1hNzU2LTE4NDFiODg5ZjliMg=="
    expect(resolveMessageIdOrRef(graphId)).toBe(graphId)
  })

  it("resolves a numeric string as a ref", () => {
    rememberMessageId("id-a")
    expect(resolveMessageIdOrRef("1")).toBe("id-a")
  })

  // The dangerous failure is resolving to the wrong message rather than to nothing,
  // so an unknown ref must come back undefined for the caller to turn into an error.
  it("returns undefined for a ref that was never issued", () => {
    expect(resolveMessageIdOrRef("999")).toBeUndefined()
  })
})

describe("formatMessageScanRow", () => {
  it("emits ref, minute-precision date, sender, subject and flags", () => {
    expect(formatMessageScanRow(message(), 7)).toBe("7|2026-01-15T10:30|John|Test Subject|")
  })

  it("marks unread and attachments", () => {
    const row = formatMessageScanRow(message({ isRead: false, hasAttachments: true }), 1)
    expect(row.endsWith("|UA")).toBe(true)
  })

  it("falls back to the address when the sender has no display name", () => {
    const row = formatMessageScanRow(message({ from: { emailAddress: { address: "a@b.com" } } }), 1)
    expect(row).toContain("|a@b.com|")
  })

  // A subject containing a pipe or newline would otherwise forge extra columns or rows.
  it("strips delimiters out of the subject", () => {
    const row = formatMessageScanRow(message({ subject: "a|b\nc" }), 1)
    expect(row.split("|")).toHaveLength(5)
  })

  it("truncates a runaway subject", () => {
    const row = formatMessageScanRow(message({ subject: "x".repeat(500) }), 1)
    expect(row.length).toBeLessThan(200)
  })

  it("handles a message with no subject", () => {
    expect(formatMessageScanRow(message({ subject: undefined }), 1)).toContain("(No Subject)")
  })
})

describe("formatMessageScan", () => {
  it("reports an empty result plainly", () => {
    expect(formatMessageScan([], [], {})).toContain("No messages found on this page.")
  })

  it("names the scanned folder and explains the ref column", () => {
    const out = formatMessageScan([message()], [1], { folder: "archive" })

    expect(out).toContain("1 in archive")
    expect(out).toContain("ref|received|from|subject|flags")
    expect(out).not.toContain("More available")
  })

  it("tells the caller how to fetch the next page", () => {
    const out = formatMessageScan([message()], [1], { nextCursor: "next-page" })
    expect(out).toContain('cursor: "next-page"')
  })

  // The whole point of the format is that scanning stays affordable in bulk.
  it("stays far cheaper per row than the markdown list format", () => {
    const many = Array.from({ length: 100 }, () => message())
    const refs = many.map((_, i) => i + 1)

    const perRow = formatMessageScan(many, refs, {}).length / 100
    expect(perRow).toBeLessThan(100)
  })
})

describe("scan coverage honesty", () => {
  it("preserves the continuation on an empty page", () => {
    const out = formatMessageScan([], [], { nextCursor: "next-page" })
    expect(out).toContain("INCOMPLETE")
    expect(out).toContain('cursor: "next-page"')
  })

  it("warns about the search cap even on the final page", () => {
    const out = formatMessageScan([message()], [1], { searched: true })
    expect(out).toContain("1,000")
    expect(out).toContain("does not establish mailbox coverage")
    expect(out).toContain("filter")
    expect(out).not.toContain("cannot be paged")
  })

  it("offers the same cursor continuation for search", () => {
    const out = formatMessageScan([message()], [1], { searched: true, nextCursor: "next-page" })
    expect(out).toContain("INCOMPLETE")
    expect(out).toContain('cursor: "next-page"')
    expect(out).not.toContain("skip:")
  })

  it("marks the end only when there is no continuation", () => {
    const out = formatMessageScan([message()], [1], {})
    expect(out).not.toContain("INCOMPLETE")
    expect(out).toContain("End of results for this query")
  })
})
