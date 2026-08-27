// Provider dispatch: every provider module exposes the same read surface
// (searchMail, listRecent, readMessage, checkAccount) over the same
// normalized result shape, so the MCP tools don't care what's behind an
// account. Providers may additionally expose attachment reading and saving
// (readAttachment, downloadAttachment), and draft creation (createReplyDraft,
// createDraft, createForwardDraft) for write-enabled accounts — currently
// outlook only. No provider has a send path.

import * as gmailImap from "./providers/gmail-imap.js";
import * as gmailApi from "./providers/gmail-api.js";
import * as outlookGraph from "./providers/outlook-graph.js";

const providers = {
  gmail: gmailImap,
  "gmail-api": gmailApi,
  outlook: outlookGraph,
};

export function providerFor(account) {
  const provider = providers[account.provider];
  if (!provider) throw new Error(`${account.id}: unknown provider "${account.provider}"`);
  return provider;
}
