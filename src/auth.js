// One-time interactive auth for OAuth providers: `npm run auth <accountId>`.
// Stores tokens in tokens/<id>.json (git-ignored). The MCP server itself is
// headless and never triggers these flows — it only consumes the cached
// tokens and tells you to run this when they're missing or revoked.

import { OAuth2Client } from "google-auth-library";
import http from "node:http";
import fs from "node:fs";
import { spawn } from "node:child_process";
import { loadAccounts, tokensDir } from "./config.js";
import { GMAIL_SCOPE } from "./providers/gmail-api.js";
import { makePca, scopesFor } from "./providers/outlook-graph.js";
import { CryptoProvider } from "@azure/msal-node";

// Fixed loopback port. Personal Microsoft accounts require an EXACT redirect
// URI match (including port), unlike work accounts which allow any port on
// http://localhost. So we pin one and register http://localhost:<PORT>.
const MS_REDIRECT_PORT = 3000;
import { googleOAuthConfig } from "./config.js";

const id = process.argv[2];
if (!id) {
  console.error("Usage: npm run auth <accountId>");
  process.exit(1);
}

const account = loadAccounts().find((a) => a.id === id);
if (!account) {
  console.error(`Unknown account "${id}". Configured: ${loadAccounts().map((a) => a.id).join(", ")}`);
  process.exit(1);
}
fs.mkdirSync(tokensDir, { recursive: true });

function tryOpen(url) {
  const cmd = process.platform === "darwin" ? "open" : process.platform === "win32" ? "start" : "xdg-open";
  spawn(cmd, [url], { shell: process.platform === "win32", stdio: "ignore", detached: true }).unref();
}

if (account.provider === "gmail") {
  console.log(
    `${account.id} uses IMAP + app password — nothing to authenticate here.\n` +
      `Create an app password at https://myaccount.google.com/apppasswords and put it in .env.`
  );
  process.exit(0);
}

if (account.provider === "gmail-api") {
  const { clientId, clientSecret } = googleOAuthConfig();
  // Loopback (installed-app) flow: Desktop OAuth clients accept any
  // http://127.0.0.1:<port> redirect, so we bind an ephemeral port.
  const server = http.createServer();
  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  const redirectUri = `http://127.0.0.1:${server.address().port}`;
  const client = new OAuth2Client(clientId, clientSecret, redirectUri);
  const authUrl = client.generateAuthUrl({
    access_type: "offline",
    prompt: "consent",
    scope: [GMAIL_SCOPE],
    login_hint: account.email,
  });

  console.log(`\nSign in as ${account.email} (read-only Gmail scope):\n\n${authUrl}\n`);
  tryOpen(authUrl);

  const code = await new Promise((resolve, reject) => {
    server.on("request", (req, res) => {
      const url = new URL(req.url, redirectUri);
      const err = url.searchParams.get("error");
      const code = url.searchParams.get("code");
      res.end(
        err
          ? `Authorization failed: ${err}. You can close this tab.`
          : "Authenticated — you can close this tab and return to the terminal."
      );
      if (err) reject(new Error(err));
      else if (code) resolve(code);
    });
  });
  server.close();

  const { tokens } = await client.getToken(code);
  if (!tokens.refresh_token) {
    console.error("No refresh token returned — remove the app's access at https://myaccount.google.com/permissions and re-run.");
    process.exit(1);
  }
  fs.writeFileSync(account.tokenFile, JSON.stringify({ refresh_token: tokens.refresh_token }), { mode: 0o600 });
  console.log(`✅ ${account.id} authenticated — token saved to tokens/${account.id}.json`);
  process.exit(0);
}

if (account.provider === "outlook") {
  const pca = makePca(account);
  const scopes = scopesFor(account);
  const mode = account.write ? "read + create drafts (no send)" : "read-only";
  // Interactive loopback (auth-code + PKCE): opens the browser, captures the
  // result on http://localhost:<port>. More robust than device code for
  // personal Microsoft accounts (no code-entry race, no expiry window), and
  // a smoother experience for the client. Needs http://localhost registered
  // as a redirect URI on the app (Mobile & desktop platform).
  // Manual auth-code + PKCE loopback on a FIXED port (see MS_REDIRECT_PORT).
  const redirectUri = `http://localhost:${MS_REDIRECT_PORT}`;
  const crypto = new CryptoProvider();
  const { verifier, challenge } = await crypto.generatePkceCodes();

  const authUrl = await pca.getAuthCodeUrl({
    scopes,
    redirectUri,
    codeChallenge: challenge,
    codeChallengeMethod: "S256",
    loginHint: account.email,
  });

  const server = http.createServer();
  await new Promise((resolve, reject) => {
    server.once("error", (e) =>
      reject(e.code === "EADDRINUSE"
        ? new Error(`Port ${MS_REDIRECT_PORT} is in use. Free it (or change MS_REDIRECT_PORT) and retry.`)
        : e));
    server.listen(MS_REDIRECT_PORT, "127.0.0.1", resolve);
  });

  console.log(`\nSign in as ${account.email} (${mode} mail). Opening your browser...`);
  console.log(`\nIf it doesn't open, paste this URL into your browser:\n\n${authUrl}\n`);
  tryOpen(authUrl);

  const code = await new Promise((resolve, reject) => {
    server.on("request", (req, res) => {
      const url = new URL(req.url, redirectUri);
      const err = url.searchParams.get("error");
      const code = url.searchParams.get("code");
      res.end(err
        ? `Sign-in failed: ${err}. You can close this tab.`
        : "Authenticated. You can close this tab and return to the terminal.");
      if (err) reject(new Error(`${err}: ${url.searchParams.get("error_description") || ""}`));
      else if (code) resolve(code);
    });
  });
  server.close();

  const result = await pca.acquireTokenByCode({
    scopes,
    redirectUri,
    code,
    codeVerifier: verifier,
  });
  console.log(`✅ ${account.id} authenticated as ${result.account?.username} — token cache saved to tokens/${account.id}.json`);
  process.exit(0);
}
