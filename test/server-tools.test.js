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
