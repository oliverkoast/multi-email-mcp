import test from "node:test";
import assert from "node:assert/strict";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

test("advertises read_attachment as a read-only MCP tool", async () => {
  const transport = new StdioClientTransport({
    command: process.execPath,
    args: [path.join(root, "src/server.js")],
    cwd: root,
    env: {
      ...process.env,
      MAIL_ACCOUNTS: "test",
      MAIL_TEST_PROVIDER: "outlook",
      MAIL_TEST_EMAIL: "test@example.com",
      MS_CLIENT_ID: "00000000-0000-0000-0000-000000000000",
      MS_TENANT: "organizations",
    },
    stderr: "pipe",
  });
  const client = new Client({ name: "attachment-test", version: "1.0.0" });

  try {
    await client.connect(transport);
    const { tools } = await client.listTools();
    const attachmentTool = tools.find((tool) => tool.name === "read_attachment");
    assert.ok(attachmentTool, "read_attachment should be listed");
    assert.equal(attachmentTool.annotations?.readOnlyHint, true);
    assert.deepEqual(attachmentTool.inputSchema.required?.sort(), ["account", "attachment_id", "message_id"]);
  } finally {
    await client.close();
  }
});

async function listToolsWith(env) {
  const transport = new StdioClientTransport({
    command: process.execPath,
    args: [path.join(root, "src/server.js")],
    cwd: root,
    env: {
      ...process.env,
      MAIL_ACCOUNTS: "test",
      MAIL_TEST_PROVIDER: "outlook",
      MAIL_TEST_EMAIL: "test@example.com",
      MS_CLIENT_ID: "00000000-0000-0000-0000-000000000000",
      MS_TENANT: "organizations",
      ...env,
    },
    stderr: "pipe",
  });
  const client = new Client({ name: "tool-surface-test", version: "1.0.0" });
  try {
    await client.connect(transport);
    return (await client.listTools()).tools;
  } finally {
    await client.close();
  }
}

test("advertises save_attachment on a read-only account", async () => {
  const tools = await listToolsWith({});
  const save = tools.find((tool) => tool.name === "save_attachment");
  assert.ok(save, "save_attachment should be listed without any write flag");
  assert.equal(save.annotations?.destructiveHint, false);
  assert.deepEqual(save.inputSchema.required?.sort(), ["account", "attachment_id", "message_id"]);
});

// Forwarding writes to the mailbox, so it must stay behind the same write flag
// as the other draft tools rather than riding along with read-only access.
test("hides create_forward_draft until an account opts into write access", async () => {
  const readOnly = await listToolsWith({});
  assert.equal(readOnly.find((tool) => tool.name === "create_forward_draft"), undefined);
  assert.equal(readOnly.find((tool) => tool.name === "create_draft"), undefined);

  const writable = await listToolsWith({ MAIL_TEST_WRITE: "true" });
  const forward = writable.find((tool) => tool.name === "create_forward_draft");
  assert.ok(forward, "create_forward_draft should appear once MAIL_TEST_WRITE=true");
  assert.deepEqual(forward.inputSchema.required?.sort(), ["account", "id", "to"]);
});

// The whole point of the draft lane: no tool anywhere may claim to send.
test("exposes no send tool even with write access enabled", async () => {
  const tools = await listToolsWith({ MAIL_TEST_WRITE: "true" });
  const senders = tools.filter((tool) => /send|forward_mail|deliver/i.test(tool.name));
  assert.deepEqual(senders, []);
});
