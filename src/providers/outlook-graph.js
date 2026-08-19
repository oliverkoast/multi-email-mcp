// Microsoft 365 / Outlook provider via Microsoft Graph (delegated Mail.Read,
// or Mail.ReadWrite for write-enabled accounts — write means creating drafts
// only; there is deliberately no send path anywhere in this module).
// Auth is a one-time browser sign-in (`npm run auth <id>`); MSAL's token
// cache is persisted to tokens/<id>.json and refreshed silently from then on.
// Search uses KQL (`from:`, `subject:`, `hasAttachments:true`, ...).

import { PublicClientApplication } from "@azure/msal-node";
import fs from "node:fs";
import { msOAuthConfig } from "../config.js";
import { maxAttachmentBytesForContentType } from "../attachment-result.js";
import {
  assertOutlookAttachmentSize,
  decodeOutlookAttachment,
  outlookAttachmentSummary,
} from "./outlook-attachment.js";

const GRAPH = "https://graph.microsoft.com/v1.0";
// Mail.ReadWrite is a superset of Mail.Read (read + create/edit drafts). It
// does NOT include sending — that would be Mail.Send, which we never request.
export const scopesFor = (account) => [account.write ? "Mail.ReadWrite" : "Mail.Read"];

export function cachePlugin(file) {
  return {
    beforeCacheAccess: async (ctx) => {
      if (fs.existsSync(file)) ctx.tokenCache.deserialize(fs.readFileSync(file, "utf8"));
    },
    afterCacheAccess: async (ctx) => {
      if (ctx.cacheHasChanged) fs.writeFileSync(file, ctx.tokenCache.serialize(), { mode: 0o600 });
    },
  };
}

export function makePca(account) {
  const { clientId, tenant } = msOAuthConfig(account);
  return new PublicClientApplication({
    auth: { clientId, authority: `https://login.microsoftonline.com/${tenant}` },
    cache: { cachePlugin: cachePlugin(account.tokenFile) },
  });
}

async function getToken(account) {
  if (!fs.existsSync(account.tokenFile)) {
    throw new Error(
      `${account.id} (${account.email}) is not authenticated yet — run: npm run auth ${account.id}`
    );
  }
  const pca = makePca(account);
  const cached = await pca.getTokenCache().getAllAccounts();
  if (!cached.length) {
    throw new Error(`${account.id}: token cache is empty — run: npm run auth ${account.id}`);
  }
  try {
    const result = await pca.acquireTokenSilent({ account: cached[0], scopes: scopesFor(account) });
    return result.accessToken;
  } catch {
    throw new Error(
      `${account.id}: token refresh failed (expired, revoked, or the cached token predates a ` +
        `scope change like enabling MAIL_${account.id.toUpperCase()}_WRITE) — run: npm run auth ${account.id}`
    );
  }
}

async function graph(token, path, { preferText = false } = {}) {
  const headers = { Authorization: `Bearer ${token}` };
  if (preferText) headers.Prefer = 'outlook.body-content-type="text"';
  const res = await fetch(`${GRAPH}${path}`, { headers });
  if (!res.ok) {
    const body = await res.text().catch(() => "");
    throw new Error(`Graph ${res.status}: ${body.slice(0, 300)}`);
  }
  return res.json();
}

function addr(recipient) {
  const e = recipient?.emailAddress;
  if (!e) return "";
  return e.name && e.name !== e.address ? `${e.name} <${e.address}>` : e.address || "";
}

function toSummary(account, msg) {
  return {
    account: account.id,
    account_email: account.email,
    id: msg.id,
    subject: msg.subject || "(no subject)",
    from: addr(msg.from),
    to: (msg.toRecipients || []).map(addr).join(", "),
    date: msg.receivedDateTime || null,
    snippet: (msg.bodyPreview || "").replace(/\s+/g, " ").trim().slice(0, 300),
  };
}

const SELECT = "id,subject,from,toRecipients,receivedDateTime,bodyPreview";

export async function searchMail(account, query, limit) {
  const token = await getToken(account);
  // $search takes a quoted KQL string and doesn't combine with $orderby.
  const search = encodeURIComponent(`"${query.replace(/"/g, '\\"')}"`);
  const data = await graph(token, `/me/messages?$search=${search}&$top=${limit}&$select=${SELECT}`);
  const results = (data.value || []).map((m) => toSummary(account, m));
  results.sort((a, b) => (b.date || "").localeCompare(a.date || ""));
  return results;
}

export async function listRecent(account, limit) {
  const token = await getToken(account);
  const data = await graph(
    token,
    `/me/messages?$top=${limit}&$orderby=receivedDateTime desc&$select=${SELECT}`
  );
  return (data.value || []).map((m) => toSummary(account, m));
}

export async function readMessage(account, id) {
  const token = await getToken(account);
  const msg = await graph(
    token,
    `/me/messages/${encodeURIComponent(id)}?$select=subject,from,toRecipients,ccRecipients,receivedDateTime,body,hasAttachments`,
    { preferText: true }
  );
  let attachments = [];
  if (msg.hasAttachments) {
    const att = await graph(
      token,
      `/me/messages/${encodeURIComponent(id)}/attachments?$select=id,name,contentType,size,isInline`
    );
    attachments = (att.value || []).map(outlookAttachmentSummary);
  }
  return {
    account: account.id,
    account_email: account.email,
    id,
    subject: msg.subject || "(no subject)",
    from: addr(msg.from),
    to: (msg.toRecipients || []).map(addr).join(", "),
    cc: (msg.ccRecipients || []).map(addr).join(", "),
    date: msg.receivedDateTime || null,
    body: msg.body?.content || "",
    attachments,
  };
}

export async function readAttachment(account, messageId, attachmentId) {
  const token = await getToken(account);
  const path = `/me/messages/${encodeURIComponent(messageId)}/attachments/${encodeURIComponent(attachmentId)}`;
  // Fetch metadata without contentBytes first. This prevents a large base64
  // payload from entering memory before the safety limit can be enforced.
  const metadata = await graph(token, `${path}?$select=id,name,contentType,size`);
  assertOutlookAttachmentSize(metadata, maxAttachmentBytesForContentType(metadata.contentType));
  const attachment = await graph(token, path);
  return {
    account: account.id,
    messageId,
    ...decodeOutlookAttachment(attachment),
  };
}

export async function checkAccount(account) {
  const token = await getToken(account);
  const inbox = await graph(token, "/me/mailFolders/inbox?$select=totalItemCount");
  const mode = account.write ? "read + drafts" : "read-only";
  return `${account.email} — ${inbox.totalItemCount} messages in Inbox (Microsoft Graph, ${mode})`;
}

// --- Draft creation (write-enabled accounts only; never sends) ------------

async function graphWrite(token, method, path, payload) {
  const res = await fetch(`${GRAPH}${path}`, {
    method,
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    body: payload === undefined ? undefined : JSON.stringify(payload),
  });
  if (!res.ok) {
    const body = await res.text().catch(() => "");
    throw new Error(`Graph ${res.status}: ${body.slice(0, 300)}`);
  }
  return res.status === 204 ? null : res.json();
}

function assertWritable(account) {
  if (!account.write) {
    throw new Error(
      `${account.id} (${account.email}) is read-only. To enable draft creation: set ` +
        `MAIL_${account.id.toUpperCase()}_WRITE=true in .env, make sure the app registration ` +
        `declares Mail.ReadWrite (delegated), and re-run: npm run auth ${account.id}`
    );
  }
}

// Graph's createReply `comment` lands inside an HTML body, so escape and
// convert newlines or plain-text formatting is lost.
function asHtmlComment(text) {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\n/g, "<br>\n");
}

export async function createReplyDraft(account, id, { comment, replyAll = false }) {
  assertWritable(account);
  const token = await getToken(account);
  // createReply/createReplyAll makes a draft in Drafts, threaded to the
  // original, with the quoted history below the comment — same as hitting
  // Reply in Outlook and typing. Nothing is sent (that would need Mail.Send).
  const action = replyAll ? "createReplyAll" : "createReply";
  const draft = await graphWrite(token, "POST", `/me/messages/${encodeURIComponent(id)}/${action}`, {
    comment: asHtmlComment(comment),
  });
  return {
    account: account.id,
    account_email: account.email,
    draft_id: draft.id,
    subject: draft.subject || "(no subject)",
    to: (draft.toRecipients || []).map(addr).join(", "),
    cc: (draft.ccRecipients || []).map(addr).join(", "),
    saved_to: "Drafts",
    sent: false,
    note: "Draft created in the Drafts folder, threaded to the original. Nothing was sent — review and send it from Outlook.",
  };
}

export async function createDraft(account, { to, cc, subject, body }) {
  assertWritable(account);
  const token = await getToken(account);
  const recipients = (list) => (list || []).map((address) => ({ emailAddress: { address } }));
  const draft = await graphWrite(token, "POST", "/me/messages", {
    subject,
    body: { contentType: "Text", content: body },
    toRecipients: recipients(to),
    ...(cc?.length ? { ccRecipients: recipients(cc) } : {}),
  });
  return {
    account: account.id,
    account_email: account.email,
    draft_id: draft.id,
    subject: draft.subject || "(no subject)",
    to: (draft.toRecipients || []).map(addr).join(", "),
    saved_to: "Drafts",
    sent: false,
    note: "Draft created in the Drafts folder. Nothing was sent — review and send it from Outlook.",
  };
}
