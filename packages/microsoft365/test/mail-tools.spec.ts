import { Some } from "functype"
import { Right } from "functype/either"
import { beforeEach, describe, expect, it, vi } from "vitest"

import type { GraphMessage } from "../src/types"

vi.mock("../src/client/graph-client", () => ({
  getGraphClient: vi.fn(),
}))

import { getGraphClient } from "../src/client/graph-client"
import {
  createDraft,
  createForwardDraft,
  createReplyAllDraft,
  createReplyDraft,
  listAttachments,
  listMailFolders,
  moveMessage,
  sendDraft,
  sendForward,
  sendMessage,
  sendReply,
  sendReplyAll,
} from "../src/tools/mail-tools"

const mockClient = {
  sendMessage: vi.fn(),
  createDraft: vi.fn(),
  sendDraft: vi.fn(),
  sendReply: vi.fn(),
  sendReplyAll: vi.fn(),
  sendForward: vi.fn(),
  createReplyDraft: vi.fn(),
  createReplyAllDraft: vi.fn(),
  createForwardDraft: vi.fn(),
  listAttachments: vi.fn(),
  listMailFolders: vi.fn(),
  moveMessage: vi.fn(),
  requestPaginated: vi.fn(),
}

beforeEach(() => {
  vi.clearAllMocks()
  vi.mocked(getGraphClient).mockReturnValue(Some(mockClient as never))
})

describe("mail-tools", () => {
  describe("sendMessage", () => {
    it("should send a message with default content type", async () => {
      mockClient.sendMessage.mockResolvedValue(Right({}))
      const result = await sendMessage({ to: "alice@example.com", subject: "Hi", body: "Hello" })
      expect(result.isRight()).toBe(true)
      expect(result.value).toContain("alice@example.com")
      expect(mockClient.sendMessage).toHaveBeenCalledWith({
        message: {
          subject: "Hi",
          body: { contentType: "Text", content: "Hello" },
          toRecipients: [{ emailAddress: { address: "alice@example.com" } }],
        },
      })
    })

    it("should send a message with HTML content type", async () => {
      mockClient.sendMessage.mockResolvedValue(Right({}))
      await sendMessage({ to: "bob@example.com", subject: "Hi", body: "<b>Bold</b>", content_type: "HTML" })
      expect(mockClient.sendMessage).toHaveBeenCalledWith({
        message: {
          subject: "Hi",
          body: { contentType: "HTML", content: "<b>Bold</b>" },
          toRecipients: [{ emailAddress: { address: "bob@example.com" } }],
        },
      })
    })

    it("should split comma-separated 'to' into multiple toRecipients", async () => {
      mockClient.sendMessage.mockResolvedValue(Right({}))
      await sendMessage({
        to: "alice@example.com, bob@example.com,carol@example.com",
        subject: "Hi",
        body: "Hello",
      })
      const callArg = mockClient.sendMessage.mock.calls[0][0] as { message: Record<string, unknown> }
      expect(callArg.message.toRecipients).toEqual([
        { emailAddress: { address: "alice@example.com" } },
        { emailAddress: { address: "bob@example.com" } },
        { emailAddress: { address: "carol@example.com" } },
      ])
    })

    it("should reject empty 'to' field", async () => {
      const result = await sendMessage({ to: "", subject: "Hi", body: "Hello" })
      expect(result.isLeft()).toBe(true)
      expect((result.value as Error).message).toContain("recipient is required")
      expect(mockClient.sendMessage).not.toHaveBeenCalled()
    })

    it("should reject 'to' containing only whitespace and commas", async () => {
      const result = await sendMessage({ to: " , , ", subject: "Hi", body: "Hello" })
      expect(result.isLeft()).toBe(true)
      expect(mockClient.sendMessage).not.toHaveBeenCalled()
    })
  })

  describe("createDraft", () => {
    const draftResponse: Partial<GraphMessage> = { id: "draft-123", subject: "Test Draft" }

    it("should create a draft with basic params", async () => {
      mockClient.createDraft.mockResolvedValue(Right(draftResponse))
      const result = await createDraft({ to: "alice@example.com", subject: "Draft", body: "Content" })
      expect(result.isRight()).toBe(true)
      expect(result.value).toContain("draft-123")
      expect(mockClient.createDraft).toHaveBeenCalledWith({
        subject: "Draft",
        body: { contentType: "Text", content: "Content" },
        toRecipients: [{ emailAddress: { address: "alice@example.com" } }],
      })
    })

    it("should create a draft with HTML content type", async () => {
      mockClient.createDraft.mockResolvedValue(Right(draftResponse))
      await createDraft({ to: "alice@example.com", subject: "Draft", body: "<p>Hi</p>", content_type: "HTML" })
      expect(mockClient.createDraft).toHaveBeenCalledWith(
        expect.objectContaining({
          body: { contentType: "HTML", content: "<p>Hi</p>" },
        }),
      )
    })

    it("should create a draft with cc recipients", async () => {
      mockClient.createDraft.mockResolvedValue(Right(draftResponse))
      await createDraft({
        to: "alice@example.com",
        subject: "Draft",
        body: "Hi",
        cc: "bob@example.com,carol@example.com",
      })
      expect(mockClient.createDraft).toHaveBeenCalledWith(
        expect.objectContaining({
          ccRecipients: [
            { emailAddress: { address: "bob@example.com" } },
            { emailAddress: { address: "carol@example.com" } },
          ],
        }),
      )
    })

    it("should create a draft with bcc recipients", async () => {
      mockClient.createDraft.mockResolvedValue(Right(draftResponse))
      await createDraft({ to: "alice@example.com", subject: "Draft", body: "Hi", bcc: "secret@example.com" })
      expect(mockClient.createDraft).toHaveBeenCalledWith(
        expect.objectContaining({
          bccRecipients: [{ emailAddress: { address: "secret@example.com" } }],
        }),
      )
    })

    it("should handle cc with whitespace and empty entries", async () => {
      mockClient.createDraft.mockResolvedValue(Right(draftResponse))
      await createDraft({
        to: "alice@example.com",
        subject: "Draft",
        body: "Hi",
        cc: " bob@example.com , , carol@example.com ",
      })
      expect(mockClient.createDraft).toHaveBeenCalledWith(
        expect.objectContaining({
          ccRecipients: [
            { emailAddress: { address: "bob@example.com" } },
            { emailAddress: { address: "carol@example.com" } },
          ],
        }),
      )
    })

    it("should omit cc when empty string", async () => {
      mockClient.createDraft.mockResolvedValue(Right(draftResponse))
      await createDraft({ to: "alice@example.com", subject: "Draft", body: "Hi", cc: "" })
      const callArg = mockClient.createDraft.mock.calls[0][0] as Record<string, unknown>
      expect(callArg).not.toHaveProperty("ccRecipients")
    })

    it("should split comma-separated 'to' into multiple toRecipients", async () => {
      mockClient.createDraft.mockResolvedValue(Right(draftResponse))
      await createDraft({
        to: "alice@example.com, bob@example.com, carol@example.com",
        subject: "Draft",
        body: "Hi",
      })
      expect(mockClient.createDraft).toHaveBeenCalledWith(
        expect.objectContaining({
          toRecipients: [
            { emailAddress: { address: "alice@example.com" } },
            { emailAddress: { address: "bob@example.com" } },
            { emailAddress: { address: "carol@example.com" } },
          ],
        }),
      )
    })

    it("should trim whitespace and drop empty entries in 'to'", async () => {
      mockClient.createDraft.mockResolvedValue(Right(draftResponse))
      await createDraft({
        to: " alice@example.com , , bob@example.com ",
        subject: "Draft",
        body: "Hi",
      })
      expect(mockClient.createDraft).toHaveBeenCalledWith(
        expect.objectContaining({
          toRecipients: [
            { emailAddress: { address: "alice@example.com" } },
            { emailAddress: { address: "bob@example.com" } },
          ],
        }),
      )
    })

    it("should reject empty 'to' field", async () => {
      const result = await createDraft({ to: "", subject: "Draft", body: "Hi" })
      expect(result.isLeft()).toBe(true)
      expect((result.value as Error).message).toContain("recipient is required")
      expect(mockClient.createDraft).not.toHaveBeenCalled()
    })
  })

  describe("sendDraft", () => {
    it("should send a draft by ID", async () => {
      mockClient.sendDraft.mockResolvedValue(Right({}))
      const result = await sendDraft({ message_id: "draft-123" })
      expect(result.isRight()).toBe(true)
      expect(result.value).toContain("Draft sent successfully")
      expect(mockClient.sendDraft).toHaveBeenCalledWith("draft-123")
    })
  })

  describe("sendReply", () => {
    it("should send a reply by message ID", async () => {
      mockClient.sendReply.mockResolvedValue(Right({}))
      const result = await sendReply({ message_id: "msg-1", comment: "Thanks!" })
      expect(result.isRight()).toBe(true)
      expect(result.value).toContain("Reply sent successfully")
      expect(mockClient.sendReply).toHaveBeenCalledWith("msg-1", "Thanks!")
    })
  })

  describe("sendReplyAll", () => {
    it("should send a reply-all by message ID", async () => {
      mockClient.sendReplyAll.mockResolvedValue(Right({}))
      const result = await sendReplyAll({ message_id: "msg-1", comment: "Thanks all!" })
      expect(result.isRight()).toBe(true)
      expect(result.value).toContain("Reply-all sent successfully")
      expect(mockClient.sendReplyAll).toHaveBeenCalledWith("msg-1", "Thanks all!")
    })
  })

  describe("sendForward", () => {
    it("should forward with recipients and an optional comment", async () => {
      mockClient.sendForward.mockResolvedValue(Right({}))
      const result = await sendForward({ message_id: "msg-1", to: "alice@example.com", comment: "FYI" })
      expect(result.isRight()).toBe(true)
      expect(result.value).toContain("alice@example.com")
      expect(mockClient.sendForward).toHaveBeenCalledWith("msg-1", "FYI", [
        { emailAddress: { address: "alice@example.com" } },
      ])
    })

    it("should default an omitted comment to an empty string", async () => {
      mockClient.sendForward.mockResolvedValue(Right({}))
      await sendForward({ message_id: "msg-1", to: "alice@example.com" })
      expect(mockClient.sendForward).toHaveBeenCalledWith("msg-1", "", [
        { emailAddress: { address: "alice@example.com" } },
      ])
    })

    it("should reject an empty 'to' field", async () => {
      const result = await sendForward({ message_id: "msg-1", to: "" })
      expect(result.isLeft()).toBe(true)
      expect((result.value as Error).message).toContain("recipient is required")
      expect(mockClient.sendForward).not.toHaveBeenCalled()
    })
  })

  describe("createReplyDraft", () => {
    it("should create a threaded reply draft and return its ID", async () => {
      mockClient.createReplyDraft.mockResolvedValue(Right({ id: "draft-r1" }))
      const result = await createReplyDraft({ message_id: "msg-1", comment: "Will do" })
      expect(result.isRight()).toBe(true)
      expect(result.value).toContain("draft-r1")
      expect(result.value).toContain("send_draft")
      expect(mockClient.createReplyDraft).toHaveBeenCalledWith("msg-1", "Will do")
    })
  })

  describe("createReplyAllDraft", () => {
    it("should create a threaded reply-all draft and return its ID", async () => {
      mockClient.createReplyAllDraft.mockResolvedValue(Right({ id: "draft-ra1" }))
      const result = await createReplyAllDraft({ message_id: "msg-1", comment: "Will do" })
      expect(result.isRight()).toBe(true)
      expect(result.value).toContain("draft-ra1")
      expect(mockClient.createReplyAllDraft).toHaveBeenCalledWith("msg-1", "Will do")
    })
  })

  describe("createForwardDraft", () => {
    it("should create a forward draft with recipients and return its ID", async () => {
      mockClient.createForwardDraft.mockResolvedValue(Right({ id: "draft-f1" }))
      const result = await createForwardDraft({ message_id: "msg-1", to: "alice@example.com", comment: "FYI" })
      expect(result.isRight()).toBe(true)
      expect(result.value).toContain("draft-f1")
      expect(mockClient.createForwardDraft).toHaveBeenCalledWith("msg-1", "FYI", [
        { emailAddress: { address: "alice@example.com" } },
      ])
    })

    it("should reject an empty 'to' field", async () => {
      const result = await createForwardDraft({ message_id: "msg-1", to: "" })
      expect(result.isLeft()).toBe(true)
      expect((result.value as Error).message).toContain("recipient is required")
      expect(mockClient.createForwardDraft).not.toHaveBeenCalled()
    })
  })

  describe("listMailFolders", () => {
    it("should list folders with their counts", async () => {
      mockClient.listMailFolders.mockResolvedValue(
        Right({ value: [{ id: "f1", displayName: "Archive", totalItemCount: 12, unreadItemCount: 3 }] }),
      )
      const result = await listMailFolders()
      expect(result.isRight()).toBe(true)
      expect(result.value).toContain("Archive")
      expect(result.value).toContain("12 items, 3 unread")
      expect(mockClient.listMailFolders).toHaveBeenCalledWith({ $top: 100 })
    })

    it("should page through all folders when asked", async () => {
      mockClient.requestPaginated.mockResolvedValue(Right([{ id: "f1", displayName: "Archive" }]))
      const result = await listMailFolders({ fetch_all_pages: true })
      expect(result.isRight()).toBe(true)
      expect(mockClient.requestPaginated).toHaveBeenCalledWith("/me/mailFolders")
      expect(mockClient.listMailFolders).not.toHaveBeenCalled()
    })
  })

  describe("moveMessage", () => {
    it("should pass a well-known folder name straight through", async () => {
      mockClient.moveMessage.mockResolvedValue(Right({ id: "msg-1", subject: "Receipt" }))
      const result = await moveMessage({ message_id: "msg-1", destination: "archive" })
      expect(result.isRight()).toBe(true)
      expect(mockClient.moveMessage).toHaveBeenCalledWith("msg-1", "archive")
      expect(mockClient.listMailFolders).not.toHaveBeenCalled()
    })

    it("should map a well-known alias and ignore case", async () => {
      mockClient.moveMessage.mockResolvedValue(Right({ id: "msg-1" }))
      await moveMessage({ message_id: "msg-1", destination: "Deleted Items" })
      expect(mockClient.moveMessage).toHaveBeenCalledWith("msg-1", "deleteditems")
    })

    it("should resolve a folder display name to its ID", async () => {
      mockClient.listMailFolders.mockResolvedValue(Right({ value: [{ id: "f-receipts", displayName: "Receipts" }] }))
      mockClient.moveMessage.mockResolvedValue(Right({ id: "msg-1" }))
      await moveMessage({ message_id: "msg-1", destination: "Receipts" })
      expect(mockClient.moveMessage).toHaveBeenCalledWith("msg-1", "f-receipts")
    })

    it("should error rather than guess when a display name is ambiguous", async () => {
      mockClient.listMailFolders.mockResolvedValue(
        Right({
          value: [
            { id: "f-a", displayName: "Receipts" },
            { id: "f-b", displayName: "Receipts" },
          ],
        }),
      )
      const result = await moveMessage({ message_id: "msg-1", destination: "Receipts" })
      expect(result.isLeft()).toBe(true)
      expect((result.value as Error).message).toContain("Multiple folders")
      expect((result.value as Error).message).toContain("f-a")
      expect(mockClient.moveMessage).not.toHaveBeenCalled()
    })

    it("should fall through to Graph when nothing matches, assuming a folder ID", async () => {
      mockClient.listMailFolders.mockResolvedValue(Right({ value: [{ id: "f1", displayName: "Archive" }] }))
      mockClient.moveMessage.mockResolvedValue(Right({ id: "msg-1" }))
      await moveMessage({ message_id: "msg-1", destination: "AAMkAGI0-opaque-id" })
      expect(mockClient.moveMessage).toHaveBeenCalledWith("msg-1", "AAMkAGI0-opaque-id")
    })
  })

  describe("listAttachments", () => {
    it("should list attachments with a read_document path for each", async () => {
      mockClient.listAttachments.mockResolvedValue(
        Right({
          value: [{ id: "att-1", name: "invoice.pdf", contentType: "application/pdf", size: 20480 }],
        }),
      )
      const result = await listAttachments({ message_id: "msg-1" })
      expect(result.isRight()).toBe(true)
      expect(result.value).toContain("invoice.pdf")
      expect(result.value).toContain("application/pdf")
      expect(result.value).toContain("20.0 KB")
      expect(result.value).toContain("/me/messages/msg-1/attachments/att-1/$value")
    })

    it("should mark inline attachments so signature images are recognisable", async () => {
      mockClient.listAttachments.mockResolvedValue(
        Right({ value: [{ id: "att-2", name: "logo.png", contentType: "image/png", size: 900, isInline: true }] }),
      )
      const result = await listAttachments({ message_id: "msg-1" })
      expect(result.value).toContain("[inline]")
      expect(result.value).toContain("900 B")
    })

    it("should report no attachments rather than an empty list", async () => {
      mockClient.listAttachments.mockResolvedValue(Right({ value: [] }))
      const result = await listAttachments({ message_id: "msg-1" })
      expect(result.isRight()).toBe(true)
      expect(result.value).toContain("No attachments found")
    })

    it("should surface a failure as a UserError", async () => {
      const { Left: L } = await import("functype/either")
      mockClient.listAttachments.mockResolvedValue(L({ message: "not found" }))
      const result = await listAttachments({ message_id: "bad" })
      expect(result.isLeft()).toBe(true)
      expect((result.value as Error).message).toContain("Failed to list attachments")
    })
  })
})
