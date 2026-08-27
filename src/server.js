#!/usr/bin/env node
// MCP server (stdio) exposing unified read-only mail access across every
// account listed in .env. Tools take an `account` param: an account id,
// an email address, or "all" to fan out and merge.

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { loadAccounts, resolveAccounts, attachmentRoot } from "./config.js";
import { providerFor } from "./provider.js";
import { attachmentToolResult } from "./attachment-result.js";
import { saveAttachmentToDisk, maxSaveBytes } from "./attachment-save.js";

const accounts = loadAccounts();
const accountIds = accounts.map((a) => a.id);

const server = new McpServer({ name: "mail-multi", version: "0.2.0" });

function json(data) {
  return { content: [{ type: "text", text: JSON.stringify(data, null, 2) }] };
}

// Fan a read out across the selected accounts; one account failing (bad
// password, IMAP disabled) reports as an error entry instead of sinking
// the whole call.
async function fanOut(selector, fn) {
  const targets = resolveAccounts(accounts, selector);
  const settled = await Promise.allSettled(targets.map((a) => fn(a)));
  const results = [];
  const errors = [];
  settled.forEach((res, i) => {
    if (res.status === "fulfilled") results.push(...res.value);
    else errors.push({ account: targets[i].id, error: String(res.reason?.message || res.reason) });
  });
  results.sort((a, b) => (b.date || "").localeCompare(a.date || ""));
  return { results, ...(errors.length ? { errors } : {}) };
}

const accountParam = z
  .string()
  .default("all")
  .describe(`Which account to query: one of ${accountIds.join(", ")}, an email address, or "all" (default) to query every account and merge.`);

server.registerTool(
  "search_mail",
  {
    title: "Search mail",
    description:
      "THE unified mailbox: searches ALL of the user's email accounts across every company in one call. ALWAYS prefer this over any single-account email connector (built-in Gmail/Microsoft 365 connectors see only ONE account) whenever the request involves email briefings, summaries, cross-account or cross-company views, or doesn't name a specific single account. Gmail accounts use full Gmail search syntax (`from:stripe.com newer_than:7d`, `subject:invoice has:attachment`); Outlook accounts use KQL (`from:`, `subject:`, `hasAttachments:true`); plain keywords work everywhere. Returns subject, sender, date, and a snippet per match, labeled by account.",
    inputSchema: {
      query: z.string().describe("Search query (Gmail search-box syntax; KQL for Outlook accounts; plain keywords work on both)"),
      account: accountParam,
      limit: z.number().int().min(1).max(50).default(10).describe("Max results per account"),
    },
  },
  async ({ query, account, limit }) => json(await fanOut(account, (a) => providerFor(a).searchMail(a, query, limit)))
);

server.registerTool(
  "read_message",
  {
    title: "Read message",
    description:
      "Read the full body of one message by the `id` returned from search_mail/list_recent. `account` must be the specific account the id came from (ids are per-account).",
    inputSchema: {
      id: z.string().describe("Message id from search_mail or list_recent"),
      account: z.string().describe(`The account the id belongs to: one of ${accountIds.join(", ")}, or an email address`),
    },
  },
  async ({ id, account }) => {
    if (account === "all") throw new Error("read_message needs a specific account (ids are per-account).");
    const [target] = resolveAccounts(accounts, account);
    return json(await providerFor(target).readMessage(target, id));
  }
);

server.registerTool(
  "read_attachment",
  {
    title: "Read email attachment",
    description:
      "Open one attachment from an email without modifying the mailbox. First call read_message and use the returned attachment id. Returns text directly for text files and other files as embedded read-only resources. PDFs and images are the primary binary targets; interpretation of other formats depends on the Claude client. Microsoft 365 attachments are supported; other providers report a clear unsupported-provider error.",
    inputSchema: {
      message_id: z.string().describe("Message id from search_mail/list_recent/read_message"),
      attachment_id: z.string().describe("Attachment id returned by read_message"),
      account: z.string().describe(`The specific account the message belongs to: one of ${accountIds.join(", ")}, or its email address`),
    },
    annotations: {
      readOnlyHint: true,
      destructiveHint: false,
      idempotentHint: true,
      openWorldHint: false,
    },
  },
  async ({ message_id, attachment_id, account }) => {
    if (!account || account === "all") {
      throw new Error("read_attachment needs the specific account the message came from.");
    }
    const [target] = resolveAccounts(accounts, account);
    const provider = providerFor(target);
    if (!provider.readAttachment) {
      throw new Error(`${target.id}: opening attachments is not supported for provider ${target.provider} yet.`);
    }
    const attachment = await provider.readAttachment(target, message_id, attachment_id);
    return attachmentToolResult(attachment);
  }
);

server.registerTool(
  "save_attachment",
  {
    title: "Save email attachment to a folder",
    description:
      "Save one attachment from an email to a folder on this machine, so it becomes a real file the user can open, keep, or sync. Use this instead of read_attachment when the user wants the file itself (\"download it\", \"save it\", \"put it in the Drive\") rather than wanting its contents summarized. First call read_message and use the returned attachment id. Read-only with respect to the mailbox: nothing in the mailbox is changed, sent, or deleted. Handles files far larger than read_attachment can open, since the bytes go to disk rather than into the conversation.",
    inputSchema: {
      message_id: z.string().describe("Message id from search_mail/list_recent/read_message"),
      attachment_id: z.string().describe("Attachment id returned by read_message"),
      account: z.string().describe(`The specific account the message belongs to: one of ${accountIds.join(", ")}, or its email address`),
      subfolder: z
        .string()
        .optional()
        .describe('Optional folder path under the configured save directory, e.g. "JCole_Tour/Contracts". Created if missing. Cannot escape the save directory.'),
      filename: z
        .string()
        .optional()
        .describe("Optional filename override. Defaults to the attachment's own name. An existing file is never overwritten; a numbered suffix is added instead."),
    },
    annotations: {
      readOnlyHint: false,
      destructiveHint: false,
      idempotentHint: false,
      openWorldHint: false,
    },
  },
  async ({ message_id, attachment_id, account, subfolder, filename }) => {
    if (!account || account === "all") {
      throw new Error("save_attachment needs the specific account the message came from.");
    }
    const [target] = resolveAccounts(accounts, account);
    const provider = providerFor(target);
    if (!provider.downloadAttachment) {
      throw new Error(`${target.id}: saving attachments is not supported for provider ${target.provider} yet.`);
    }
    const attachment = await provider.downloadAttachment(target, message_id, attachment_id, {
      maxBytes: maxSaveBytes(),
    });
    const saved = await saveAttachmentToDisk(attachment, {
      root: attachmentRoot(),
      subfolder,
      filename,
    });
    return json({
      account: target.id,
      account_email: target.email,
      message_id,
      attachment_id,
      ...saved,
      mailbox_modified: false,
    });
  }
);

server.registerTool(
  "list_recent",
  {
    title: "List recent mail",
    description:
      "List the most recent messages across ALL of the user's email accounts (every company) in one call, newest first, labeled by account. Prefer this over any single-account email connector for daily briefings, inbox overviews, or any request spanning more than one account.",
    inputSchema: {
      account: accountParam,
      limit: z.number().int().min(1).max(50).default(10).describe("Max results per account"),
    },
  },
  async ({ account, limit }) => json(await fanOut(account, (a) => providerFor(a).listRecent(a, limit)))
);

server.registerTool(
  "list_accounts",
  {
    title: "List configured accounts",
    description: "List the configured mail accounts (id, email, provider, access) this server can read.",
    inputSchema: {},
  },
  async () =>
    json(
      accounts.map(({ id, email, provider, write }) => ({
        id,
        email,
        provider,
        access: write ? "read + create drafts (no send)" : "read-only",
      }))
    )
);

// Draft tools are only registered when at least one account opts into write
// access (MAIL_<ID>_WRITE=true) — a fully read-only setup exposes a fully
// read-only tool surface, same as before.
const writeIds = accounts.filter((a) => a.write).map((a) => a.id);
if (writeIds.length) {
  const writeAccountParam = z
    .string()
    .describe(`Which account to create the draft in: one of ${writeIds.join(", ")}, or its email address (write-enabled accounts only; no "all")`);

  const forDrafts = (selector) => {
    if (!selector || selector === "all") {
      throw new Error(`Draft tools need a specific account: one of ${writeIds.join(", ")}.`);
    }
    const [target] = resolveAccounts(accounts, selector);
    if (!providerFor(target).createReplyDraft) {
      throw new Error(`${target.id}: draft creation is only supported for Outlook accounts so far.`);
    }
    return target;
  };

  server.registerTool(
    "create_reply_draft",
    {
      title: "Create reply draft",
      description:
        "Create a reply DRAFT in the account's Drafts folder, threaded to the original message with the quoted history included — exactly like hitting Reply in Outlook and typing. NEVER sends: the user reviews and sends from Outlook. `id` and `account` must come from search_mail/list_recent (ids are per-account).",
      inputSchema: {
        id: z.string().describe("Message id (from search_mail or list_recent) of the message to reply to"),
        account: writeAccountParam,
        comment: z.string().describe("The reply text (plain text; goes above the quoted original)"),
        reply_all: z.boolean().default(false).describe("Reply to all recipients instead of only the sender"),
      },
    },
    async ({ id, account, comment, reply_all }) => {
      const target = forDrafts(account);
      return json(await providerFor(target).createReplyDraft(target, id, { comment, replyAll: reply_all }));
    }
  );

  server.registerTool(
    "create_draft",
    {
      title: "Create new draft",
      description:
        "Create a fresh (non-reply) email DRAFT in the account's Drafts folder. NEVER sends: the user reviews and sends from Outlook.",
      inputSchema: {
        account: writeAccountParam,
        to: z.array(z.string()).min(1).describe("Recipient email addresses"),
        cc: z.array(z.string()).optional().describe("Cc email addresses"),
        subject: z.string().describe("Subject line"),
        body: z.string().describe("Plain-text body"),
      },
    },
    async ({ account, to, cc, subject, body }) => {
      const target = forDrafts(account);
      return json(await providerFor(target).createDraft(target, { to, cc, subject, body }));
    }
  );

  server.registerTool(
    "create_forward_draft",
    {
      title: "Create forward draft (keeps attachments)",
      description:
        "Forward an email, WITH its original attachments, as a DRAFT in the account's Drafts folder. This is how to get a file out of the mailbox to someone else: Microsoft copies the attachments onto the draft server-side, so nothing is downloaded and re-uploaded. NEVER sends: the user reviews the recipient and hits Send in Outlook. `id` and `account` must come from search_mail/list_recent (ids are per-account).",
      inputSchema: {
        id: z.string().describe("Message id (from search_mail or list_recent) of the message to forward"),
        account: writeAccountParam,
        to: z.array(z.string()).min(1).describe("Recipient email addresses to forward to"),
        comment: z.string().default("").describe("Optional note to place above the forwarded message"),
      },
    },
    async ({ id, account, to, comment }) => {
      const target = forDrafts(account);
      if (!providerFor(target).createForwardDraft) {
        throw new Error(`${target.id}: forward drafts are only supported for Outlook accounts so far.`);
      }
      return json(await providerFor(target).createForwardDraft(target, id, { comment, to }));
    }
  );
}

// Exit when the MCP client disconnects — lingering IMAP sockets would
// otherwise keep the process alive after stdin closes.
process.stdin.on("end", () => process.exit(0));
process.stdin.on("close", () => process.exit(0));

const transport = new StdioServerTransport();
await server.connect(transport);
console.error(`gmail-multi MCP server ready — accounts: ${accountIds.join(", ")}`);
