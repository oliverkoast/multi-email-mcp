# Client setup runbook — unified mail access for Claude

One local MCP server gives Claude read-only search across all of your email
accounts at once. Nothing is sent anywhere: the server runs on your machine,
credentials stay in a local file, and the code physically contains no
send/delete/modify operations.

## What gets installed

- Node.js (LTS) if not already present
- This folder (`multi-email-mcp`), with your account credentials in a local `.env`
  file and OAuth tokens (if any) in `tokens/` — both excluded from version
  control

## Per-account setup — pick the lane per account

### Lane A — Gmail with an app password (fastest, ~5 min, no IT needed)

Works when your Google account allows app passwords (most do by default).

1. Signed in as the account, open <https://myaccount.google.com/apppasswords>
   - If it asks you to enable **2-Step Verification** first, do that
     (Security → 2-Step Verification), then come back.
   - If the page says app passwords aren't available for your account, your
     org has disabled them — use **Lane B** for this account.
2. App name: `claude-mail` → **Create** → copy the 16-character password.
3. Add the account to `.env` (spaces in the password are fine):
   ```
   MAIL_WORK_EMAIL=you@yourcompany.com
   MAIL_WORK_APP_PASSWORD=xxxx xxxx xxxx xxxx
   ```
   and append `work` to `MAIL_ACCOUNTS`.

### Lane B — Gmail via Google OAuth (when IT disabled IMAP/app passwords)

Uses the Gmail API with the **read-only** scope. One-time browser sign-in.

1. In `.env`: set `MAIL_<ID>_PROVIDER=gmail-api` and the account email
   (the `GOOGLE_OAUTH_CLIENT_ID/SECRET` app settings are already filled in
   by your consultant).
2. Run `npm run auth <id>` → a browser opens → sign in as that account →
   approve the read-only Gmail permission.
3. If Google blocks the sign-in with an admin-approval message, your org
   restricts third-party apps — send IT the **Google IT request** below and
   re-run after they approve.

### Lane C — Microsoft 365 / Outlook (via Microsoft Graph)

One-time browser sign-in with the **Mail.Read** (read-only) permission.

1. In `.env`: set `MAIL_<ID>_PROVIDER=outlook` and the account email.
2. Run `npm run auth <id>` → your browser opens → sign in with the account
   (MFA works normally) → **Accept** the read-only mail consent. The tab
   redirects to a local "Authenticated" page and the token is cached.
3. If you see "Need admin approval", your org requires admin consent — send
   IT the **Microsoft IT request** below and re-run after they approve.

### Lane C write option — draft creation for a Microsoft 365 account

By default every account is read-only. An Outlook account can additionally
opt into **draft creation**: Claude can then drop reply drafts (threaded,
with quoted history), **forward drafts that keep the original attachments**,
or fresh drafts into the Drafts folder. Sending stays impossible — the code
never requests Mail.Send and contains no send path, so a human still has to
hit Send on every draft.

Forwarding is the supported way to pass a file from the mailbox to someone
else. Graph's `createForward` copies the original attachments onto the draft
server-side, so the file is never downloaded and re-uploaded, and the draft
waits in Drafts until a person checks the recipient and sends it.

To pull files *out* to a folder instead, no write access is needed at all:
`save_attachment` works on a plain read-only account (see
`MAIL_ATTACHMENT_DIR` in `.env.example`).

Requirements:

1. An app registration that declares **Mail.ReadWrite (delegated)**. Two
   ways to get one:
   - **Add the scope to the existing registration** (API permissions → Add
     → Microsoft Graph → Delegated → Mail.ReadWrite). Simplest when the
     registration serves one client. Read-only accounts are unaffected:
     the code only ever *requests* Mail.ReadWrite for accounts with the
     write flag, so their sign-ins and tokens stay Mail.Read.
   - **A client-owned registration in the client's own tenant** — their IT
     creates it (same steps as the consultant prerequisites below, with
     Mail.ReadWrite instead of Mail.Read, single-tenant is fine) and owns
     it outright: full audit, kill switch, no external app in the trust
     chain. Point the account at it with
     `MAIL_<ID>_MS_CLIENT_ID=<their client ID>` and
     `MAIL_<ID>_MS_TENANT=<their tenant>`; other accounts keep using the
     global `MS_CLIENT_ID`.
2. In `.env`, on the account: `MAIL_<ID>_WRITE=true`.
3. If the tenant requires admin consent, IT approves the write scope
   (template below).
4. Re-run `npm run auth <id>` — the cached token was issued for Mail.Read
   and must be re-issued with the new scope.

**Microsoft IT request (write lane):**

> Subject: Approve draft creation for the claude-mail app for <user@org.com>
>
> Please grant admin consent for the app registration `<CLIENT_ID>`
> ("claude-mail") for the delegated Microsoft Graph permission
> `Mail.ReadWrite` (read the signed-in user's own mailbox and create/edit
> drafts in it, including forward drafts that retain the original
> attachments; no ability to send — that would be Mail.Send, which the
> app does not request). To limit the grant to specific users, set
> Enterprise applications → the app → Properties → "Assignment required"
> = Yes and assign only them; delegated permissions only ever reach
> mailboxes of users who sign in. Access is revocable anytime under the
> user's My Apps or by removing the enterprise application. If you would
> rather own the app outright, an alternative is to create the
> registration in your own tenant (public client, delegated Mail.ReadWrite,
> redirect URI http://localhost:3000, "allow public client flows" on) and
> send back its client ID — everything else works the same.

## Verify + wire into Claude

```sh
npm run check          # one ✅ line per account with a message count
```

**Claude Desktop** — add to `claude_desktop_config.json`
(macOS: `~/Library/Application Support/Claude/claude_desktop_config.json`):

```json
{
  "mcpServers": {
    "mail": {
      "command": "node",
      "args": ["/ABSOLUTE/PATH/TO/multi-email-mcp/src/server.js"]
    }
  }
}
```

then fully quit and reopen Claude Desktop.

**Claude Code:**

```sh
claude mcp add --scope user mail -- node /ABSOLUTE/PATH/TO/multi-email-mcp/src/server.js
```

Smoke test in Claude: *"Search all my accounts for `<keyword>` and show which
account each result came from."*

Attachment test for a Microsoft 365 account:

> Find a recent email with a PDF attachment. Read the message, open the PDF
> using `read_attachment`, and summarize it. Tell me the source account,
> message subject, attachment name, and whether anything could not be read.
> Do not create, modify, delete, or send anything.

Opening a Microsoft 365 attachment stays within the delegated `Mail.Read`
permission. The tool can open ordinary binary file attachments up to 7 MB and text attachments up to 1 MB. Attached
Outlook items and cloud-reference attachments are identified but not opened;
Gmail attachment retrieval is not implemented yet.

## IT request templates

**Google IT request** (org restricts third-party app access):

> Subject: Allowlist a read-only Gmail OAuth app for <user@org.com>
>
> Please allow the OAuth app with client ID `<GOOGLE_OAUTH_CLIENT_ID>` to
> access Google data for <user@org.com>, scope
> `https://www.googleapis.com/auth/gmail.readonly` (read-only mail; no send,
> modify, or delete). Admin console → Security → API controls → App access
> control → Configure new app. This is a locally-run mail search tool; access
> is revocable anytime from the user's account permissions page.

**Microsoft IT request** (tenant requires admin consent):

> Subject: Approve a read-only mail app for <user@org.com>
>
> Please grant admin consent for the app registration `<MS_CLIENT_ID>`
> ("claude-mail") requesting only the delegated Microsoft Graph permission
> `Mail.Read` (read the signed-in user's own mailbox; no send or modify).
> The consent request should already be pending under Entra ID → Enterprise
> applications → Admin consent requests. Access is revocable anytime under
> the user's My Apps or by removing the enterprise application.

## Consultant prerequisites (one-time, before the engagement)

These create the shared OAuth "apps" that lanes B and C sign in through.
Lane A needs nothing.

**Google OAuth client (lane B)** — in <https://console.cloud.google.com>:

1. Create a project (e.g. `claude-mail`) → APIs & Services → **Enable the
   Gmail API**.
2. OAuth consent screen: External, app name `claude-mail`, add each lane-B
   user's email as a **test user**.
3. Credentials → Create credentials → OAuth client ID → **Desktop app** →
   copy client ID + secret into `.env`.
4. ⚠️ **Testing-mode gotcha:** while the consent screen is in "Testing"
   status, refresh tokens **expire after 7 days** (re-run `npm run auth`).
   Publishing to Production for the `gmail.readonly` restricted scope
   requires Google's app verification — heavy. For a long-term client
   deployment on a **Workspace** account, the clean fix is to create this
   OAuth client in a GCP project under the *client's own* Workspace org with
   consent screen type **Internal**: no verification, no 7-day expiry, and
   their IT gets full ownership. For personal `@gmail.com` accounts, prefer
   lane A (app passwords) — lane B there is demo-grade only.

**Entra ID app registration (lane C)** — in <https://entra.microsoft.com>.
Needs a directory: a bare personal Microsoft account has none, so if
"App registrations" bounces you ("create applications outside of a directory
has been deprecated"), first create one — the reliable route is a free Azure
account at <https://azure.microsoft.com/free> (card for identity check, never
charged; app registrations are free), which provisions a "Default Directory".
Then:

1. App registrations → New registration → name `claude-mail` → supported
   account types: **"Accounts in any organizational directory and personal
   Microsoft accounts"**.
2. Authentication → **"Allow public client flows" = Yes**.
3. Authentication → **Add a platform → Mobile and desktop applications**, and
   register redirect URI **`http://localhost:3000`** (the exact loopback the
   auth CLI uses — see `MS_REDIRECT_PORT` in src/auth.js). Personal Microsoft
   accounts require an exact port match; work accounts are more lenient, but
   registering this one URI covers both.
4. API permissions → Add → Microsoft Graph → Delegated → **Mail.Read**.
5. Copy the Application (client) ID into `.env` as `MS_CLIENT_ID`.

No client secret is needed (public client), so the registration itself
contains nothing sensitive. One registration covers every Outlook account in
every engagement. Sign in per account with `npm run auth <id>`: it opens the
browser, you approve read-only mail, and it captures the result on
`http://localhost:3000` — no code to type.

## Revoking access (any time, ~30 seconds)

- **Lane A:** delete the app password — myaccount.google.com → Security →
  2-Step Verification → App passwords.
- **Lane B:** remove the app — <https://myaccount.google.com/permissions>.
- **Lane C:** remove the app — <https://myapps.microsoft.com> → app → Remove
  (or your admin removes the enterprise app).
- Locally: delete `.env` and `tokens/` and the server can read nothing.
