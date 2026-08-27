# multi-email-mcp — multi-account mail MCP server

A small local MCP server (stdio) that gives Claude unified **read-only**
access to any number of email accounts at once — Gmail and Microsoft 365 —
each authenticating independently, all exposed through one connection.

> **Setting this up for the first time? Start here → [CLIENT-ONEPAGER.md](CLIENT-ONEPAGER.md)**
> — a plain-English walkthrough (install, connect your accounts, done).
> The rest of this README is the technical reference.
>
> **Docs map:** [CLIENT-ONEPAGER.md](CLIENT-ONEPAGER.md) (start here) ·
> [CLIENT-SETUP.md](CLIENT-SETUP.md) (per-account steps + IT request templates) ·
> [DAILY-BRIEFING.md](DAILY-BRIEFING.md) (automate a daily briefing) ·
> [CLIENT-NOTE.md](CLIENT-NOTE.md) (what it is, in plain terms).

## Providers (per account, set in `.env`)

| Provider | Backend | Auth | When |
|---|---|---|---|
| `gmail` (default) | IMAP (`X-GM-RAW` = full Gmail search syntax) | App password | Google accounts that allow app passwords — the 5-minute path |
| `gmail-api` | Gmail REST API | OAuth, `gmail.readonly` scope | Google orgs where IT disabled IMAP/app passwords |
| `outlook` | Microsoft Graph | OAuth device-code, `Mail.Read` scope | Microsoft 365 / outlook.com accounts |

**Why app passwords first?** For accounts you control they need no Google
Cloud project, no consent screen, no token storage — just 2-Step Verification
plus one generated password. The tradeoffs: an app password is a full-access
mailbox credential, and hardened orgs disable them — that's what the two
OAuth providers are for. Microsoft has no app-password lane at all (basic
IMAP auth was retired in 2022–2024), so Outlook accounts always use Graph.

## Tools

| Tool | Params | Returns |
|---|---|---|
| `search_mail` | `query`, `account` (id / email / `"all"`), `limit` | subject, sender, date, snippet per match, labeled by account |
| `read_message` | `id`, `account` (specific — ids are per-account) | full body + headers + attachment list, including attachment ids |
| `read_attachment` | `message_id`, `attachment_id`, `account` | opens a Microsoft 365 attachment as an embedded read-only resource |
| `save_attachment` | `message_id`, `attachment_id`, `account`, `subfolder?`, `filename?` | writes the attachment to a folder on disk and returns the path |
| `list_recent` | `account`, `limit` | newest messages first |
| `list_accounts` | — | configured account ids + emails + providers |

`account: "all"` fans out to every account in parallel and merges results
newest-first; a failing account comes back as an `errors` entry instead of
failing the call. Result shape is identical across providers.

### Reading attachments

For a Microsoft 365 message, call `read_message` first, then pass the returned
message id, attachment id, and account to `read_attachment`. Text attachments
are returned as text; other files are returned as MCP embedded resources.
PDFs and images are the primary binary targets; whether Claude can interpret
other formats depends on the Claude client. The tool uses only
Microsoft Graph `Mail.Read`, so opening an attachment does not grant send,
delete, or modify access. Binary attachments over 7 MB and text attachments over 1 MB are rejected to protect the
model context. Gmail attachment retrieval is not implemented yet.

### Saving attachments to a folder

`read_attachment` opens a file *in the conversation*; `save_attachment` writes
it *to disk* and returns the path. Use the second when the user wants the file
itself. Both stay inside delegated `Mail.Read` and neither changes the mailbox.

Set `MAIL_ATTACHMENT_DIR` to the save root (defaults to `~/Downloads`). Point
it at a locally-synced Drive folder and saved files land in the Drive with no
Drive API involved. The optional `subfolder` argument organises below that root
and is resolved and bounds-checked, so a caller cannot write outside it, and an
existing file is never overwritten (a numbered suffix is added instead).

Because saved bytes never enter the conversation, saving is not bound by the
7 MB inline-read cap. Its own ceiling is `MAIL_MAX_ATTACHMENT_SAVE_MB`
(default 50), which is what makes large scanned contracts workable.

### Draft tools (write-enabled Outlook accounts only)

Registered only when an account sets `MAIL_<ID>_WRITE=true`, which is also the
only case where `Mail.ReadWrite` is requested instead of `Mail.Read`. A
read-only setup exposes a read-only tool surface.

| Tool | Params | Returns |
|---|---|---|
| `create_reply_draft` | `id`, `account`, `comment`, `reply_all` | threaded reply draft in Drafts |
| `create_forward_draft` | `id`, `account`, `to`, `comment?` | forward draft in Drafts, **original attachments carried over**, plus the list of files that rode along |
| `create_draft` | `account`, `to`, `cc?`, `subject`, `body` | fresh draft in Drafts |

`create_forward_draft` is the supported way to get a file out of the mailbox to
someone else: Microsoft Graph's `createForward` copies the attachments onto the
draft server-side, so nothing is downloaded and re-uploaded.

**There is no send path.** `Mail.Send` is never requested and no tool sends.
Every draft lands in Drafts for a human to review and send from Outlook.

## Setup

1. `npm install`
2. `cp .env.example .env`, list your accounts (see comments in the file)
3. Per account: paste an app password, or `npm run auth <id>` for the OAuth
   providers (one-time interactive sign-in; tokens cached in `tokens/`,
   git-ignored). Step-by-step per lane: **[CLIENT-SETUP.md](CLIENT-SETUP.md)**
   — including the IT request templates for locked-down orgs.
4. `npm run check` — verifies every account end-to-end and prints counts.

### Adding an account

Append an id to `MAIL_ACCOUNTS`, add its `MAIL_<ID>_*` vars, auth if OAuth.
No code changes.

## Wiring into Claude

Claude Code:

```sh
claude mcp add --scope user mail -- node /path/to/multi-email-mcp/src/server.js
```

Claude Desktop (`claude_desktop_config.json`):

```json
{ "mcpServers": { "mail": { "command": "node", "args": ["/path/to/multi-email-mcp/src/server.js"] } } }
```

`.env` and `tokens/` are resolved relative to this folder, so launch cwd
doesn't matter.

## Security notes

- `.env` and `tokens/` are git-ignored; nothing secret is ever committed.
- Read-only by construction: IMAP mailboxes open with a read-only lock, and
  the OAuth scopes requested (`gmail.readonly`, `Mail.Read`) cannot send,
  modify, or delete even if the code tried.
- The MCP server is headless — it never opens an interactive auth flow; if a
  token is missing or revoked it returns an error telling you to run
  `npm run auth <id>`.
- Revocation per lane is documented in [CLIENT-SETUP.md](CLIENT-SETUP.md).
