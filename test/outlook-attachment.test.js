import test from "node:test";
import assert from "node:assert/strict";

import {
  outlookAttachmentSummary,
  decodeOutlookAttachment,
  assertOutlookAttachmentSize,
} from "../src/providers/outlook-attachment.js";

test("preserves the Microsoft attachment id needed for a later read", () => {
  assert.deepEqual(
    outlookAttachmentSummary({
      id: "AAMk-attachment-1",
      name: "board.pdf",
      contentType: "application/pdf",
      size: 4,
      isInline: false,
    }),
    {
      id: "AAMk-attachment-1",
      filename: "board.pdf",
      contentType: "application/pdf",
      size: 4,
      isInline: false,
    }
  );
});

test("decodes Microsoft file attachment content without mutating the mailbox", () => {
  const result = decodeOutlookAttachment({
    id: "a1",
    name: "board.pdf",
    contentType: "application/pdf",
    size: 4,
    contentBytes: Buffer.from("%PDF").toString("base64"),
  });

  assert.equal(result.id, "a1");
  assert.equal(result.filename, "board.pdf");
  assert.equal(result.contentType, "application/pdf");
  assert.deepEqual(result.content, Buffer.from("%PDF"));
});

test("rejects item attachments that do not contain readable file bytes", () => {
  assert.throws(
    () =>
      decodeOutlookAttachment({
        id: "a2",
        name: "attached-message.eml",
        contentType: "message/rfc822",
        size: 123,
      }),
    /does not expose file content/i
  );
});


test("rejects oversized Microsoft attachments before content is fetched", () => {
  assert.throws(
    () => assertOutlookAttachmentSize({ name: "huge.pdf", size: 8_000_000 }, 7_000_000),
    /too large/i
  );
});

test("rejects Microsoft attachments with no trustworthy preflight size", () => {
  assert.throws(
    () => assertOutlookAttachmentSize({ name: "unknown.pdf" }, 7_000_000),
    /size is unavailable/i
  );
});
