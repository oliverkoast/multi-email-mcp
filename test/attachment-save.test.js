import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import {
  DEFAULT_MAX_SAVE_BYTES,
  assertSaveSize,
  maxSaveBytes,
  resolveTargetDir,
  safeFilename,
  saveAttachmentToDisk,
  uniquePath,
} from "../src/attachment-save.js";

async function tmpdir() {
  return fs.mkdtemp(path.join(os.tmpdir(), "mail-save-test-"));
}

// Real contract filenames are full of spaces, digits, ampersands and dots.
// Over-aggressive sanitizing silently mangles every one of them, so this is
// pinned deliberately.
test("keeps ordinary contract filenames intact", () => {
  assert.equal(
    safeFilename("Marriott Tampa - J.Cole Tour (2026) & rider.pdf"),
    "Marriott Tampa - J.Cole Tour (2026) & rider.pdf"
  );
});

test("strips path separators and control characters from filenames", () => {
  assert.equal(safeFilename("../../etc/passwd"), "passwd");
  assert.equal(safeFilename("in\u0000voice\u001f.pdf"), "in_voice_.pdf");
  assert.equal(safeFilename('re:port<1>|2?.pdf'), "re_port_1__2_.pdf");
});

test("falls back rather than producing an empty or hidden filename", () => {
  assert.equal(safeFilename(".."), "attachment");
  assert.equal(safeFilename("/"), "attachment");
  assert.equal(safeFilename(""), "attachment");
  assert.equal(safeFilename(".hidden"), "hidden");
});

test("allows nested subfolders inside the save root", () => {
  const dir = resolveTargetDir("/tmp/mail", "JCole_Tour/Contracts");
  assert.equal(dir, path.resolve("/tmp/mail/JCole_Tour/Contracts"));
});

test("refuses subfolders that escape the save root", () => {
  assert.throws(() => resolveTargetDir("/tmp/mail", "../../etc"), /Refusing to save outside/);
  assert.throws(() => resolveTargetDir("/tmp/mail", "/etc"), /Refusing to save outside/);
  assert.throws(() => resolveTargetDir("/tmp/mail", "ok/../../.."), /Refusing to save outside/);
});

test("numbers around an existing file instead of overwriting it", async () => {
  const taken = new Set(["/d/contract.pdf", "/d/contract (2).pdf"]);
  const next = await uniquePath("/d", "contract.pdf", async (p) => taken.has(p));
  assert.equal(next, path.join("/d", "contract (3).pdf"));
});

test("rejects attachments over the save limit, and those with no known size", () => {
  assert.throws(() => assertSaveSize(999, "big.pdf", 100), /over the 100-byte save limit/);
  assert.throws(() => assertSaveSize(undefined, "odd.pdf", 100), /size is unavailable/);
  assert.doesNotThrow(() => assertSaveSize(99, "fine.pdf", 100));
});

// The context limit for opening an attachment inline is 7 MB. Saving must not
// inherit it, or the large scanned contracts this exists for would be rejected.
test("the save limit is far above the inline-read limit and is configurable", () => {
  assert.ok(DEFAULT_MAX_SAVE_BYTES > 7 * 1024 * 1024);
  assert.equal(maxSaveBytes({}), DEFAULT_MAX_SAVE_BYTES);
  assert.equal(maxSaveBytes({ MAIL_MAX_ATTACHMENT_SAVE_MB: "2" }), 2 * 1024 * 1024);
  assert.equal(maxSaveBytes({ MAIL_MAX_ATTACHMENT_SAVE_MB: "nonsense" }), DEFAULT_MAX_SAVE_BYTES);
});

test("writes the attachment bytes to the requested subfolder", async () => {
  const root = await tmpdir();
  try {
    const result = await saveAttachmentToDisk(
      {
        filename: "Hotel Contract.pdf",
        contentType: "application/pdf",
        content: Buffer.from("%PDF-1.4 contract"),
      },
      { root, subfolder: "JCole_Tour/Contracts" }
    );

    assert.equal(result.filename, "Hotel Contract.pdf");
    assert.equal(result.renamed, false);
    assert.equal(result.bytes, 17);
    assert.equal(
      await fs.readFile(path.join(root, "JCole_Tour/Contracts/Hotel Contract.pdf"), "utf8"),
      "%PDF-1.4 contract"
    );
  } finally {
    await fs.rm(root, { recursive: true, force: true });
  }
});

test("saving the same attachment twice keeps both files", async () => {
  const root = await tmpdir();
  try {
    const attachment = {
      filename: "contract.pdf",
      contentType: "application/pdf",
      content: Buffer.from("first"),
    };
    const one = await saveAttachmentToDisk(attachment, { root });
    const two = await saveAttachmentToDisk(
      { ...attachment, content: Buffer.from("second") },
      { root }
    );

    assert.equal(one.filename, "contract.pdf");
    assert.equal(two.filename, "contract (2).pdf");
    assert.equal(two.renamed, true);
    assert.equal(await fs.readFile(one.saved_to, "utf8"), "first");
    assert.equal(await fs.readFile(two.saved_to, "utf8"), "second");
  } finally {
    await fs.rm(root, { recursive: true, force: true });
  }
});

test("a traversing subfolder writes nothing to disk", async () => {
  const root = await tmpdir();
  try {
    await assert.rejects(
      saveAttachmentToDisk(
        { filename: "x.pdf", content: Buffer.from("x") },
        { root, subfolder: "../escaped" }
      ),
      /Refusing to save outside/
    );
    assert.deepEqual(await fs.readdir(root), []);
    await assert.rejects(fs.access(path.join(path.dirname(root), "escaped")));
  } finally {
    await fs.rm(root, { recursive: true, force: true });
  }
});
