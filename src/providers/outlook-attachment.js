export function assertOutlookAttachmentSize(attachment, maxBytes) {
  if (!Number.isFinite(attachment.size)) {
    throw new Error(
      `${attachment.name || "Attachment"} size is unavailable, so it cannot be opened safely.`
    );
  }
  if (attachment.size > maxBytes) {
    throw new Error(
      `${attachment.name || "Attachment"} is too large to open safely (${attachment.size} bytes; limit ${maxBytes} bytes).`
    );
  }
}

export function outlookAttachmentSummary(attachment) {
  return {
    id: attachment.id,
    filename: attachment.name,
    contentType: attachment.contentType || "application/octet-stream",
    size: attachment.size,
    isInline: Boolean(attachment.isInline),
  };
}

export function decodeOutlookAttachment(attachment) {
  if (!attachment.contentBytes) {
    throw new Error(
      `${attachment.name || "This attachment"} does not expose file content. Attached messages and cloud-reference attachments are not supported yet.`
    );
  }

  return {
    id: attachment.id,
    filename: attachment.name,
    contentType: attachment.contentType || "application/octet-stream",
    size: attachment.size,
    content: Buffer.from(attachment.contentBytes, "base64"),
  };
}
