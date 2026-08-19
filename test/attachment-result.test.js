import test from "node:test";
import assert from "node:assert/strict";

import {
  attachmentToolResult,
  MAX_ATTACHMENT_BYTES,
  MAX_ATTACHMENT_OUTPUT_BYTES,
  MAX_TEXT_ATTACHMENT_BYTES,
} from "../src/attachment-result.js";

test("returns text attachments as an embedded text resource", () => {
  const result = attachmentToolResult({
    account: "polymateria",
    messageId: "message/1",
    id: "attachment 1",
    filename: "board-notes.txt",
    contentType: "text/plain",
    size: 11,
    content: Buffer.from("hello board"),
  });

  assert.equal(result.content[0].type, "text");
  assert.match(result.content[0].text, /board-notes\.txt/);
  assert.deepEqual(result.content[1], {
    type: "resource",
    resource: {
      uri: "mail-attachment://polymateria/message%2F1/attachment%201/board-notes.txt",
      mimeType: "text/plain",
      text: "hello board",
    },
  });
});

test("returns binary attachments as base64 embedded resources", () => {
  const bytes = Buffer.from([0x25, 0x50, 0x44, 0x46]);
  const result = attachmentToolResult({
    account: "polymateria",
    messageId: "m1",
    id: "a1",
    filename: "board.pdf",
    contentType: "application/pdf",
    size: bytes.length,
    content: bytes,
  });

  assert.deepEqual(result.content[1], {
    type: "resource",
    resource: {
      uri: "mail-attachment://polymateria/m1/a1/board.pdf",
      mimeType: "application/pdf",
      blob: bytes.toString("base64"),
    },
  });
});

test("rejects attachments larger than the safety limit", () => {
  assert.throws(
    () =>
      attachmentToolResult({
        account: "polymateria",
        messageId: "m1",
        id: "a1",
        filename: "huge.pdf",
        contentType: "application/pdf",
        size: MAX_ATTACHMENT_BYTES + 1,
        content: Buffer.alloc(0),
      }),
    /too large/i
  );
});


test("keeps the base64 MCP resource inside the output budget", () => {
  const bytes = Buffer.alloc(MAX_ATTACHMENT_BYTES);
  const result = attachmentToolResult({
    account: "polymateria",
    messageId: "m1",
    id: "a1",
    filename: "largest-supported.pdf",
    contentType: "application/pdf",
    size: bytes.length,
    content: bytes,
  });

  assert.ok(Buffer.byteLength(JSON.stringify(result)) <= MAX_ATTACHMENT_OUTPUT_BYTES);
});


test("rejects text attachments that could exceed the serialized output budget", () => {
  const bytes = Buffer.alloc(MAX_TEXT_ATTACHMENT_BYTES + 1, 0);
  assert.throws(
    () =>
      attachmentToolResult({
        account: "polymateria",
        messageId: "m1",
        id: "a1",
        filename: "hostile.txt",
        contentType: "text/plain",
        size: bytes.length,
        content: bytes,
      }),
    /too large/i
  );
});

test("keeps worst-case text escaping inside the serialized output budget", () => {
  const bytes = Buffer.alloc(MAX_TEXT_ATTACHMENT_BYTES, 0);
  const result = attachmentToolResult({
    account: "polymateria",
    messageId: "m1",
    id: "a1",
    filename: "bounded.txt",
    contentType: "text/plain",
    size: bytes.length,
    content: bytes,
  });
  assert.ok(Buffer.byteLength(JSON.stringify(result)) <= MAX_ATTACHMENT_OUTPUT_BYTES);
});
