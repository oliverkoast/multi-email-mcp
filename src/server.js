#!/usr/bin/env node
// MCP server (stdio) exposing unified read-only Gmail access across every
// account listed in .env. Tools take an `account` param: an account id,
// an email address, or "all" to fan out and merge.

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { loadAccounts, resolveAccounts } from "./config.js";
import { providerFor } from "./provider.js";

const accounts = loadAccounts();
const accountIds = accounts.map((a) => a.id);

const server = new McpServer({ name: "gmail-multi", version: "0.1.0" });

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
}

// Exit when the MCP client disconnects — lingering IMAP sockets would
// otherwise keep the process alive after stdin closes.
process.stdin.on("end", () => process.exit(0));
process.stdin.on("close", () => process.exit(0));

const transport = new StdioServerTransport();
await server.connect(transport);
console.error(`gmail-multi MCP server ready — accounts: ${accountIds.join(", ")}`);
