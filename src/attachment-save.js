// Saving an attachment to disk is a different job from opening one in the
// model's context, so it gets its own size ceiling. The context limits in
// attachment-result.js exist because base64 in an MCP payload competes with
// the conversation; bytes written to a folder do not, so a 30 MB scanned
// contract saves fine even though it could never be opened inline.

import fs from "node:fs/promises";
import path from "node:path";

export const DEFAULT_MAX_SAVE_BYTES = 50 * 1024 * 1024;

export function maxSaveBytes(env = process.env) {
  const mb = Number(env.MAIL_MAX_ATTACHMENT_SAVE_MB);
  return Number.isFinite(mb) && mb > 0 ? Math.floor(mb * 1024 * 1024) : DEFAULT_MAX_SAVE_BYTES;
}

// Path separators, Windows-reserved characters, and control codes all become
// "_". A name that sanitizes to nothing ("..", "/") falls back rather than
// producing an empty or hidden filename.
export function safeFilename(name) {
  const base = path.basename(String(name || "").trim());
  const cleaned = base
    .replace(/[\u0000-\u001f\u007f<>:"/\\|?*]/g, "_")
    .replace(/^\.+/, "")
    .trim();
  return cleaned || "attachment";
}

// A subfolder is caller-supplied, so it is resolved and then proven to still
// sit inside the root. This rejects "..", absolute paths, and symlink-style
// escapes alike, instead of trying to pattern-match traversal.
export function resolveTargetDir(root, subfolder) {
  const base = path.resolve(root);
  if (!subfolder) return base;
  const target = path.resolve(base, subfolder);
  const rel = path.relative(base, target);
  if (rel.startsWith("..") || path.isAbsolute(rel)) {
    throw new Error(
      `Refusing to save outside the attachment folder: "${subfolder}" resolves to ${target}, which is not inside ${base}.`
    );
  }
  return target;
}

async function fileExists(p) {
  try {
    await fs.access(p);
    return true;
  } catch {
    return false;
  }
}

// Never silently overwrite: an existing "Hotel Contract.pdf" becomes
// "Hotel Contract (2).pdf" so a re-run cannot destroy the earlier file.
export async function uniquePath(dir, filename, exists = fileExists) {
  const ext = path.extname(filename);
  const stem = filename.slice(0, filename.length - ext.length);
  let candidate = path.join(dir, filename);
  for (let n = 2; await exists(candidate); n += 1) {
    if (n > 999) throw new Error(`Too many files named like ${filename} in ${dir}.`);
    candidate = path.join(dir, `${stem} (${n})${ext}`);
  }
  return candidate;
}

export function assertSaveSize(size, filename, limit) {
  if (!Number.isFinite(size)) {
    throw new Error(`${filename || "Attachment"} size is unavailable, so it cannot be saved safely.`);
  }
  if (size > limit) {
    throw new Error(
      `${filename || "Attachment"} is ${size} bytes, over the ${limit}-byte save limit. Raise MAIL_MAX_ATTACHMENT_SAVE_MB in .env if this is expected.`
    );
  }
}

export async function saveAttachmentToDisk(attachment, { root, subfolder, filename }) {
  const dir = resolveTargetDir(root, subfolder);
  const name = safeFilename(filename || attachment.filename);
  const content = Buffer.isBuffer(attachment.content)
    ? attachment.content
    : Buffer.from(attachment.content || "");

  await fs.mkdir(dir, { recursive: true });
  const target = await uniquePath(dir, name);
  await fs.writeFile(target, content);

  return {
    saved_to: target,
    filename: path.basename(target),
    content_type: attachment.contentType || "application/octet-stream",
    bytes: content.length,
    renamed: path.basename(target) !== name,
  };
}
