/**
 * Configuration for Azure authentication
 *
 * Light User Mode:
 * Provide only accessToken for temporary access (no auto-renewal)
 *
 * Medium User Mode:
 * Provide refreshToken + credentials for 90-day auto-renewal
 *
 * Super User Mode:
 * Provide tokenProvider + credentials for infinite auto-renewal
 */
export interface AzureConfig {
  // Light user: access token only (expires in ~1 hour)
  accessToken?: string;

  // Medium user: refresh token
  refreshToken?: string;

  // Super user: token provider
  tokenProvider?: () => Promise<string> | string;

  // Required credentials (can come from env variables)
  clientId?: string;
  clientSecret?: string;
  tenantId?: string;

  // Optional: Custom OAuth scopes
  // If not provided, uses DEFAULT_SCOPES (works in 99% of cases)
  // Example: ['User.Read', 'Mail.Send', 'Sites.ReadWrite.All']
  scopes?: string[];

  // Optional: Allow insecure SSL connections (for development/testing)
  // WARNING: Only use this in trusted development environments
  allowInsecure?: boolean;
}

/**
 * Stored credentials in cross-platform storage
 * NOTE: clientSecret is NOT stored (comes from env or config)
 */
export interface StoredCredentials {
  refreshToken: string;
  accessToken: string;
  expiresAt?: number; // Optional because we might not know for access-token-only mode
  clientId: string; // OK to store (public)
  tenantId: string; // OK to store (public)
  // NO clientSecret - must come from env or config
}

export interface Tag {
  id: string;
  displayName: string;
  type: string;
}

export interface Record {
  id: string;
  fields: {
    taskName: string;
    data: string;
    requestId: string;
    authToken: string;
    broadcaster: string;
  };
}

export interface Calendar {
  id: string;
  name: string;
}

export interface Holiday {
  name: string;
  date: string;
}

export interface Mail {
  from: {
    name: string;
    address: string;
  };
  subject: string;
  body: string;
  receivedDateTime: string;
}

/**
 * Email address with optional display name
 */
export interface EmailAddress {
  address: string;
  name?: string;
}

/**
 * Internet message header (e.g., for mentions)
 */
export interface InternetMessageHeader {
  name: string;
  value: string;
}

/**
 * Email recipient
 */
export interface Recipient {
  emailAddress: EmailAddress;
}

/**
 * Email attachment (Microsoft Graph format)
 */
export interface Attachment {
  "@odata.type": string;
  name: string;
  contentType: string;
  contentBytes: string;
}

/**
 * Follow-up flag for emails
 */
export interface FollowUpFlag {
  flagStatus: "notFlagged" | "complete" | "flagged";
  startDateTime?: { dateTime: string; timeZone: string };
  dueDateTime?: { dateTime: string; timeZone: string };
  completedDateTime?: { dateTime: string; timeZone: string };
  reminderDateTime?: { dateTime: string; timeZone: string };
}

/**
 * Complete email payload for Microsoft Graph API
 */
export interface MailPayload {
  message: {
    subject: string;
    internetMessageHeaders?: InternetMessageHeader[];
    bodyPreview?: string;
    body: {
      contentType: "HTML" | "Text";
      content: string;
    };
    uniqueBody?: {
      contentType: "HTML" | "Text";
      content: string;
    };
    from?: Recipient;
    replyTo?: Recipient[];
    toRecipients: Recipient[];
    ccRecipients?: Recipient[];
    bccRecipients?: Recipient[];
    attachments?: Attachment[];
    importance?: "low" | "normal" | "high";
    categories?: string[];
    isReadReceiptRequested?: boolean;
    flag?: FollowUpFlag;
  };
  saveToSentItems?: "true" | "false";
}
