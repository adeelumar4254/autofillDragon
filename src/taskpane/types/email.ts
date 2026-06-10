export interface AttachmentInfo {
  id: string;
  name: string;
  size: number;
  contentType: string;
  attachmentType: number;
  isValid: boolean;
}

export interface EmailDetails {
  subject: string;
  senderName: string;
  senderEmail: string;
  userEmail: string;
  attachments: AttachmentInfo[];
  bodyHtml?: string;
  bodyText?: string;
}

export interface ProcessResponse {
  uniqueId: string;
  url?: string;
  errorMessage?: string;
}