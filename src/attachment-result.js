// Binary files expand by roughly 4/3 when embedded as base64 in MCP.
// Cap source bytes at 7 MiB so the serialized resource stays under 10 MiB.
export const MAX_ATTACHMENT_OUTPUT_BYTES = 10 * 1024 * 1024;
export const MAX_ATTACHMENT_BYTES = 7 * 1024 * 1024;
export const MAX_TEXT_ATTACHMENT_BYTES = 1 * 1024 * 1024;

const TEXT_TYPES = new Set([
  "application/json",
  "application/xml",
  "application/javascript",
  "application/x-javascript",
  "application/csv",
]);

function isTextContentType(contentType) {
  const mime = (contentType || "application/octet-stream").split(";", 1)[0].trim().toLowerCase();
  return mime.startsWith("text/") || TEXT_TYPES.has(mime) || mime.endsWith("+json") || mime.endsWith("+xml");
}

export function maxAttachmentBytesForContentType(contentType) {
  return isTextContentType(contentType) ? MAX_TEXT_ATTACHMENT_BYTES : MAX_ATTACHMENT_BYTES;
}

function attachmentUri({ account, messageId, id, filename }) {
  return `mail-attachment://${encodeURIComponent(account)}/${encodeURIComponent(messageId)}/${encodeURIComponent(id)}/${encodeURIComponent(filename || "attachment")}`;
}

export function attachmentToolResult(attachment) {
  const content = Buffer.isBuffer(attachment.content)
    ? attachment.content
    : Buffer.from(attachment.content || "");
  const size = Number.isFinite(attachment.size) ? attachment.size : content.length;

  const mimeType = attachment.contentType || "application/octet-stream";
  const sourceLimit = maxAttachmentBytesForContentType(mimeType);
  if (size > sourceLimit || content.length > sourceLimit) {
    throw new Error(
      `${attachment.filename || "Attachment"} is too large to open safely (${size} bytes; limit ${sourceLimit} bytes for ${mimeType}).`
    );
  }

  const resource = {
    uri: attachmentUri(attachment),
    mimeType,
    ...(isTextContentType(mimeType)
      ? { text: content.toString("utf8") }
      : { blob: content.toString("base64") }),
  };

  const result = {
    content: [
      {
        type: "text",
        text: JSON.stringify(
          {
            account: attachment.account,
            message_id: attachment.messageId,
            attachment_id: attachment.id,
            filename: attachment.filename,
            content_type: mimeType,
            size,
            read_only: true,
          },
          null,
          2
        ),
      },
      { type: "resource", resource },
    ],
  };

  if (Buffer.byteLength(JSON.stringify(result)) > MAX_ATTACHMENT_OUTPUT_BYTES) {
    throw new Error(`${attachment.filename || "Attachment"} exceeds the safe MCP output limit.`);
  }
  return result;
}
