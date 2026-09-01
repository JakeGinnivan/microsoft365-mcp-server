// Mail tool definitions.

import { z } from "zod"

import {
  batchMoveMessages,
  createDraft,
  createForwardDraft,
  createReplyAllDraft,
  createReplyDraft,
  getMessage,
  listAttachments,
  listMailFolders,
  listMessages,
  moveMessage,
  searchMessages,
  sendDraft,
  sendForward,
  sendMessage,
  sendReply,
  sendReplyAll,
} from ".."
import type { ToolDefinition } from "../tool-definitions"
import { FETCH_ALL_PAGES_PARAM, unwrapResult } from "./shared"

export const mailTools: ReadonlyArray<ToolDefinition> = [
  {
    name: "list_messages",
    description: "List email messages from your inbox",
    parameters: z.object({
      top: z.number().optional().describe("Number of messages to return (default: 25)"),
      filter: z.string().optional().describe("OData filter expression"),
      fetch_all_pages: FETCH_ALL_PAGES_PARAM,
    }),
    execute: async (params) => unwrapResult(await listMessages(params)),
    domain: "mail",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "get_message",
    description:
      "Get a specific email message with full body content. Pass body_format:'text' for marketing or newsletter mail — Graph converts server-side, avoiding tens of thousands of characters of HTML and CSS.",
    parameters: z.object({
      message_id: z.string().describe("The message ID"),
      body_format: z
        .enum(["text", "html"])
        .optional()
        .describe("Body format to request. 'text' strips HTML/CSS server-side. Default: the message's own format"),
    }),
    execute: async (params) => unwrapResult(await getMessage(params)),
    domain: "mail",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "list_mail_folders",
    description: "List mail folders with item and unread counts, for resolving move destinations",
    parameters: z.object({
      fetch_all_pages: FETCH_ALL_PAGES_PARAM,
    }),
    execute: async (params) => unwrapResult(await listMailFolders(params)),
    domain: "mail",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "move_message",
    description:
      "Move a message to another mail folder. Destination accepts a well-known name (archive, deleteditems, inbox, junkemail), a folder display name, or a folder ID. Moving to deleteditems is recoverable; use list_mail_folders to see what exists.",
    parameters: z.object({
      message_id: z.string().describe("The message ID to move"),
      destination: z
        .string()
        .describe("Destination folder: well-known name (e.g. archive), display name, or folder ID"),
    }),
    execute: async (params) => unwrapResult(await moveMessage(params)),
    domain: "mail",
    readOnly: false,
    annotations: { destructiveHint: true },
  },
  {
    name: "batch_move_messages",
    description:
      "Move several messages to the same folder in one call. Resolves the destination once and returns a single summary instead of one result per message. Reports any failures individually.",
    parameters: z.object({
      message_ids: z.array(z.string()).describe("Message IDs to move (max 50)"),
      destination: z
        .string()
        .describe("Destination folder: well-known name (e.g. archive), display name, or folder ID"),
    }),
    execute: async (params) => unwrapResult(await batchMoveMessages(params)),
    domain: "mail",
    readOnly: false,
    annotations: { destructiveHint: true },
  },
  {
    name: "list_attachments",
    description:
      "List a message's attachments with name, content type and size. Returns a read_document path per attachment for extracting its text (PDF, Office, etc).",
    parameters: z.object({
      message_id: z.string().describe("The message ID whose attachments to list"),
    }),
    execute: async (params) => unwrapResult(await listAttachments(params)),
    domain: "mail",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "send_message",
    description: "Send a new email message",
    parameters: z.object({
      to: z.string().describe("Recipient email address(es), comma-separated for multiple"),
      subject: z.string().describe("Email subject"),
      body: z.string().describe("Email body content"),
      content_type: z.string().optional().describe("Body content type: Text or HTML (default: Text)"),
    }),
    execute: async (params) => unwrapResult(await sendMessage(params)),
    domain: "mail",
    readOnly: false,
    annotations: { destructiveHint: true },
  },
  {
    name: "send_reply",
    description:
      "Reply to the sender of an email and send immediately. Threads into the conversation and quotes the original. Use create_reply_draft to review before sending.",
    parameters: z.object({
      message_id: z.string().describe("The message ID to reply to"),
      comment: z.string().describe("Reply content (added above the quoted original)"),
    }),
    execute: async (params) => unwrapResult(await sendReply(params)),
    domain: "mail",
    readOnly: false,
    annotations: { destructiveHint: true },
  },
  {
    name: "send_reply_all",
    description:
      "Reply to all recipients of an email and send immediately. Threads into the conversation and quotes the original. Use create_reply_all_draft to review before sending.",
    parameters: z.object({
      message_id: z.string().describe("The message ID to reply to"),
      comment: z.string().describe("Reply content (added above the quoted original)"),
    }),
    execute: async (params) => unwrapResult(await sendReplyAll(params)),
    domain: "mail",
    readOnly: false,
    annotations: { destructiveHint: true },
  },
  {
    name: "send_forward",
    description:
      "Forward an email to new recipients and send immediately. Quotes the original. Use create_forward_draft to review before sending.",
    parameters: z.object({
      message_id: z.string().describe("The message ID to forward"),
      to: z.string().describe("Recipient email address(es), comma-separated for multiple"),
      comment: z.string().optional().describe("Optional note added above the quoted original"),
    }),
    execute: async (params) => unwrapResult(await sendForward(params)),
    domain: "mail",
    readOnly: false,
    annotations: { destructiveHint: true },
  },
  {
    name: "create_reply_draft",
    description:
      "Create a reply draft (to the sender) in the Drafts folder. Threads into the conversation with the original quoted underneath. Review, then send with send_draft.",
    parameters: z.object({
      message_id: z.string().describe("The message ID to reply to"),
      comment: z.string().describe("Reply content (added above the quoted original)"),
    }),
    execute: async (params) => unwrapResult(await createReplyDraft(params)),
    domain: "mail",
    readOnly: false,
  },
  {
    name: "create_reply_all_draft",
    description:
      "Create a reply-all draft (to all recipients) in the Drafts folder. Threads into the conversation with the original quoted underneath. Review, then send with send_draft.",
    parameters: z.object({
      message_id: z.string().describe("The message ID to reply to"),
      comment: z.string().describe("Reply content (added above the quoted original)"),
    }),
    execute: async (params) => unwrapResult(await createReplyAllDraft(params)),
    domain: "mail",
    readOnly: false,
  },
  {
    name: "create_forward_draft",
    description:
      "Create a forward draft in the Drafts folder with the original quoted underneath. Review, then send with send_draft.",
    parameters: z.object({
      message_id: z.string().describe("The message ID to forward"),
      to: z.string().describe("Recipient email address(es), comma-separated for multiple"),
      comment: z.string().optional().describe("Optional note added above the quoted original"),
    }),
    execute: async (params) => unwrapResult(await createForwardDraft(params)),
    domain: "mail",
    readOnly: false,
  },
  {
    name: "search_messages",
    description: "Search email messages",
    parameters: z.object({
      query: z.string().describe("Search query string"),
      top: z.number().optional().describe("Number of results to return (default: 25)"),
    }),
    execute: async (params) => unwrapResult(await searchMessages(params)),
    domain: "mail",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "create_draft",
    description: "Create a new email draft in the Drafts folder",
    parameters: z.object({
      to: z.string().describe("Recipient email address(es), comma-separated for multiple"),
      subject: z.string().describe("Email subject"),
      body: z.string().describe("Email body content"),
      content_type: z.string().optional().describe("Body content type: Text or HTML (default: Text)"),
      cc: z.string().optional().describe("CC recipients (comma-separated email addresses)"),
      bcc: z.string().optional().describe("BCC recipients (comma-separated email addresses)"),
    }),
    execute: async (params) => unwrapResult(await createDraft(params)),
    domain: "mail",
    readOnly: false,
  },
  {
    name: "send_draft",
    description: "Send an existing email draft",
    parameters: z.object({
      message_id: z.string().describe("The draft message ID to send"),
    }),
    execute: async (params) => unwrapResult(await sendDraft(params)),
    domain: "mail",
    readOnly: false,
    annotations: { destructiveHint: true },
  },
]
