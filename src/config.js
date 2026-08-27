// Account list is driven entirely by env vars so adding an account never
// touches code:
//
//   MAIL_ACCOUNTS=arcforma,formai,personal,clientco
//
//   # provider "gmail" (default): IMAP + app password
//   MAIL_ARCFORMA_EMAIL=oliver@arcforma.ai
//   MAIL_ARCFORMA_APP_PASSWORD=xxxxxxxxxxxxxxxx
//
//   # provider "gmail-api": Gmail REST API + OAuth (hardened Google orgs)
//   MAIL_CLIENTCO_PROVIDER=gmail-api
//   MAIL_CLIENTCO_EMAIL=user@clientco.com
//
//   # provider "outlook": Microsoft Graph + device-code OAuth
//   MAIL_CLIENTMS_PROVIDER=outlook
//   MAIL_CLIENTMS_EMAIL=user@clientms.com
//
// OAuth providers authenticate once via `npm run auth <id>`; tokens are
// cached in tokens/<id>.json (git-ignored). The legacy GMAIL_* prefix from
// v0.1 still works.

import dotenv from "dotenv";
import { fileURLToPath } from "node:url";
import os from "node:os";
import path from "node:path";

export const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
export const tokensDir = path.join(projectRoot, "tokens");
dotenv.config({ path: path.join(projectRoot, ".env"), quiet: true });

export const PROVIDERS = ["gmail", "gmail-api", "outlook"];

function envFor(id, suffix) {
  const key = id.toUpperCase().replace(/[^A-Z0-9]/g, "_");
  return process.env[`MAIL_${key}_${suffix}`] ?? process.env[`GMAIL_${key}_${suffix}`];
}

export function loadAccounts() {
  const ids = (process.env.MAIL_ACCOUNTS || process.env.GMAIL_ACCOUNTS || "")
    .split(",")
    .map((s) => s.trim())
    .filter(Boolean);

  if (ids.length === 0) {
    throw new Error("No accounts configured. Set MAIL_ACCOUNTS in .env (see .env.example).");
  }

  const problems = [];
  const accounts = ids.map((id) => {
    const provider = (envFor(id, "PROVIDER") || "gmail").toLowerCase();
    const email = envFor(id, "EMAIL");
    const password = (envFor(id, "APP_PASSWORD") || "").replace(/\s+/g, "");
    const write = /^(1|true|yes)$/i.test(envFor(id, "WRITE") || "");

    if (!PROVIDERS.includes(provider)) {
      problems.push(`${id}: unknown provider "${provider}" (use ${PROVIDERS.join(" | ")})`);
    }
    if (!email) problems.push(`${id}: missing MAIL_${id.toUpperCase()}_EMAIL`);
    if (provider === "gmail" && !password) {
      problems.push(`${id}: missing MAIL_${id.toUpperCase()}_APP_PASSWORD (required for provider "gmail")`);
    }
    if (write && provider !== "outlook") {
      problems.push(`${id}: MAIL_${id.toUpperCase()}_WRITE is only supported for provider "outlook"`);
    }
    return {
      id,
      provider,
      email,
      password,
      write,
      // Per-account Entra app override, so one write-enabled account can use
      // a Mail.ReadWrite registration while the rest stay on the shared
      // read-only one.
      msClientId: envFor(id, "MS_CLIENT_ID"),
      msTenant: envFor(id, "MS_TENANT"),
      tokenFile: path.join(tokensDir, `${id}.json`),
    };
  });

  if (problems.length) {
    throw new Error(`Account config problems:\n  ${problems.join("\n  ")}`);
  }

  return accounts;
}

// Shared OAuth app settings (one Google OAuth client / one Entra app
// registration covers every account of that provider).
export function googleOAuthConfig() {
  const clientId = process.env.GOOGLE_OAUTH_CLIENT_ID;
  const clientSecret = process.env.GOOGLE_OAUTH_CLIENT_SECRET;
  if (!clientId || !clientSecret) {
    throw new Error(
      'Provider "gmail-api" needs GOOGLE_OAUTH_CLIENT_ID + GOOGLE_OAUTH_CLIENT_SECRET in .env ' +
        "(a Desktop-type OAuth client from any Google Cloud project — see CLIENT-SETUP.md)."
    );
  }
  return { clientId, clientSecret };
}

// Where save_attachment writes. Point this at a synced Drive folder and saved
// contracts land in the Drive with no Drive API involved; leave it unset and
// they land in ~/Downloads.
export function attachmentRoot(env = process.env) {
  const dir = (env.MAIL_ATTACHMENT_DIR || "").trim();
  if (!dir) return path.join(os.homedir(), "Downloads");
  return dir.startsWith("~") ? path.join(os.homedir(), dir.slice(1)) : path.resolve(dir);
}

export function msOAuthConfig(account = {}) {
  const clientId = account.msClientId || process.env.MS_CLIENT_ID;
  if (!clientId) {
    throw new Error(
      'Provider "outlook" needs MS_CLIENT_ID in .env (an Entra ID app registration ' +
        "with Mail.Read delegated permission — see CLIENT-SETUP.md)."
    );
  }
  return { clientId, tenant: account.msTenant || process.env.MS_TENANT || "common" };
}

export function resolveAccounts(accounts, selector) {
  if (!selector || selector === "all") return accounts;
  const found = accounts.find(
    (a) => a.id === selector || a.email.toLowerCase() === selector.toLowerCase()
  );
  if (!found) {
    const known = accounts.map((a) => `${a.id} (${a.email})`).join(", ");
    throw new Error(`Unknown account "${selector}". Known accounts: ${known}, or "all".`);
  }
  return [found];
}
