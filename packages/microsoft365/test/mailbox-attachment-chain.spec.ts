// scan_messages -> list_attachments -> read_document is the documented way to read an
// attachment. Each step hands the next a string it printed, so a mailbox lost at any hop
// sends the last request to the caller's own mailbox, where the IDs do not exist. This
// drives the real Graph client end to end and asserts every request stayed in one mailbox.

import { Right } from "functype/either"
import { afterEach, describe, expect, it, vi } from "vitest"

vi.mock("../src/auth", async (importOriginal) => ({
  ...(await importOriginal<typeof import("../src/auth")>()),
  getAccessToken: vi.fn(() => Promise.resolve(Right("test-token"))),
}))

import { initializeGraphClient } from "../src/client/graph-client"
import { listAttachments, scanMessages } from "../src/tools/mail-tools"
import { readDocument } from "../src/tools/read-document-tools"
import { clearMessageRefs } from "../src/utils/message-refs"

const BEL = "bel@example.com"
const BEL_ROOT = "/v1.0/users/bel%40example.com"

const json = (body: unknown) =>
  new Response(JSON.stringify(body), { status: 200, headers: { "Content-Type": "application/json" } })

// Answers by path, so the test fails on a request to the wrong mailbox instead of
// handing it the next canned response.
const graph = (url: string): Response => {
  const { pathname } = new URL(url)
  switch (pathname) {
    case `${BEL_ROOT}/messages`:
      return json({ value: [{ id: "msg-1", subject: "Invoice", hasAttachments: true }] })
    case `${BEL_ROOT}/messages/msg-1/attachments`:
      return json({
        value: [
          {
            "@odata.type": "#microsoft.graph.fileAttachment",
            id: "att-1",
            name: "invoice.txt",
            contentType: "text/plain",
            size: 12,
          },
        ],
      })
    case `${BEL_ROOT}/messages/msg-1/attachments/att-1`:
      return json({ id: "att-1", name: "invoice.txt", contentType: "text/plain", size: 12 })
    case `${BEL_ROOT}/messages/msg-1/attachments/att-1/$value`:
      return new Response("amount: $42", { status: 200, headers: { "Content-Type": "text/plain" } })
    default:
      return new Response(JSON.stringify({ error: { message: `unexpected ${pathname}` } }), {
        status: 404,
        headers: { "Content-Type": "application/json" },
      })
  }
}

afterEach(() => {
  vi.unstubAllGlobals()
  vi.unstubAllEnvs()
  clearMessageRefs()
})

describe("reading an attachment in another mailbox", () => {
  it("keeps scan, list_attachments and read_document in the addressed mailbox", async () => {
    vi.stubEnv("MS365_ALLOWED_MAILBOXES", BEL)
    const fetchSpy = vi.fn((url: string) => Promise.resolve(graph(url)))
    vi.stubGlobal("fetch", fetchSpy)
    initializeGraphClient({ getAccessToken: () => Promise.resolve(Right("test-token")) })

    const scan = await scanMessages({ mailbox: BEL })
    const ref = /\n([a-z0-9]{4}-\d+)\|/.exec(String(scan.value))?.[1]
    if (!ref) throw new Error(`expected a scan row in: ${String(scan.value)}`)

    const listed = await listAttachments({ message_id: ref, mailbox: BEL })
    expect(listed.isRight()).toBe(true)
    const path = /read_document path: (\S+)/.exec(String(listed.value))?.[1]
    expect(path).toBe("/users/bel%40example.com/messages/msg-1/attachments/att-1/$value")

    const read = await readDocument({ path: path! })
    expect(read.isRight()).toBe(true)
    expect(read.value).toContain("amount: $42")

    expect(fetchSpy).toHaveBeenCalledTimes(4)
    for (const [url] of fetchSpy.mock.calls) expect(new URL(url).pathname.startsWith(`${BEL_ROOT}/`)).toBe(true)
  })

  it("still prints a /me path for the signed-in user's own mailbox", async () => {
    const fetchSpy = vi.fn(() =>
      Promise.resolve(
        json({
          value: [
            { "@odata.type": "#microsoft.graph.fileAttachment", id: "att-1", name: "a.txt", contentType: "text/plain" },
          ],
        }),
      ),
    )
    vi.stubGlobal("fetch", fetchSpy)
    initializeGraphClient({ getAccessToken: () => Promise.resolve(Right("test-token")) })

    const listed = await listAttachments({ message_id: "msg-1" })

    expect(listed.value).toContain("read_document path: /me/messages/msg-1/attachments/att-1/$value")
  })
})
