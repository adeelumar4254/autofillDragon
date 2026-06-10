
import { EmailDetails, AttachmentInfo } from "../types/email";

const MAX_SIZE = 5 * 1024 * 1024; // 5MB
const ALLOWED_TYPES = [
  "application/pdf",
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  "text/plain",
];

/**
 * Extracts all required data from the current Office item
 */
export const getActiveEmailDetails = (): EmailDetails | null => {
  if (typeof Office === "undefined" || !Office.context?.mailbox?.item) {
    return null;
  }

  const item = Office.context.mailbox.item;

 // Update the mapping logic in getActiveEmailDetails
const attachments: AttachmentInfo[] = (item.attachments || []).map((att: any) => {
  const isSizeValid = att.size <= MAX_SIZE;
  const isTypeValid = ALLOWED_TYPES.includes(att.contentType);

  return {
    id: att.id,
    name: att.name,
    size: att.size,
    contentType: att.contentType,
    attachmentType: att.attachmentType,
    isValid: isSizeValid && isTypeValid,
    // Add this line to identify WHY it failed
    errorReason: !isTypeValid ? "Unsupported Type" : !isSizeValid ? "Too Large (>5MB)" : null
  };
});
  return {
    subject: item.subject || "(No Subject)",
    senderName: item.from?.displayName || "Unknown",
    senderEmail: item.from?.emailAddress || "Unknown",
    userEmail: Office.context.mailbox.userProfile.emailAddress,
    attachments,
  };
};

/**
 * Format file size into readable string
 */
export const formatSize = (bytes: number): string => {
  if (bytes === 0) return "0 Bytes";
  const k = 1024;
  const sizes = ["Bytes", "KB", "MB"];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + " " + sizes[i];
};