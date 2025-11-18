import Axon, { AxonError } from "axios-fluent";
import dayjs from "dayjs";
import { JSDOM } from "jsdom";
import * as fs from "fs";
import * as path from "path";
import { AzureAuth } from "../core/auth";
import { AzureConfig, Mail, MailPayload } from "../types";

/**
 * Outlook/Mail service for Microsoft Graph API
 * Handles email operations (send, read, search)
 */
export class Outlook {
  private auth: AzureAuth;

  /**
   * Create a new Outlook service instance
   *
   * @param config - Optional config or AzureAuth instance
   *
   * @example
   * // Auto-load from env/storage
   * const outlook = new Outlook();
   *
   * @example
   * // Explicit config
   * const outlook = new Outlook({
   *   clientId: '...',
   *   clientSecret: '...',
   *   tenantId: '...',
   *   refreshToken: '...'
   * });
   *
   * @example
   * // Shared auth instance
   * const auth = new AzureAuth({ refreshToken: '...' });
   * const outlook = new Outlook(auth);
   */
  constructor(config?: AzureConfig | AzureAuth) {
    if (config instanceof AzureAuth) {
      this.auth = config;
    } else {
      this.auth = new AzureAuth(config);
    }
  }

  /**
   * Get current user's profile information
   *
   * @returns User profile data
   *
   * @example
   * const user = await outlook.getMe();
   * console.log(user.displayName, user.mail);
   */
  async getMe() {
    try {
      const token = await this.auth.getAccessToken();
      const url = `https://graph.microsoft.com/v1.0/me`;
      const res = await Axon.new().bearer(token).get(url);
      return res.data;
    } catch (error) {
      this.auth.handleApiError(error as AxonError);
    }
  }

  /**
   * Send an email
   *
   * @param payload - Email message payload (Microsoft Graph format)
   *
   * @example
   * await outlook.sendMail({
   *   message: {
   *     subject: 'Hello',
   *     body: {
   *       contentType: 'Text',
   *       content: 'World'
   *     },
   *     toRecipients: [
   *       { emailAddress: { address: 'user@example.com' } }
   *     ]
   *   }
   * });
   */
  async sendMail(payload: MailPayload) {
    try {
      const token = await this.auth.getAccessToken();
      const url = "https://graph.microsoft.com/v1.0/me/sendMail";
      return await Axon.new().bearer(token).post(url, payload);
    } catch (error) {
      this.auth.handleApiError(error as AxonError);
    }
  }

  /**
   * Get emails received on a specific date
   *
   * @param date - Date string (any format dayjs can parse)
   * @param subjectFilter - Optional subject filter
   * @returns Array of emails
   *
   * @example
   * const emails = await outlook.getMails('2024-01-15', 'invoice');
   */
  async getMails(date: string, subjectFilter?: string): Promise<Mail[]> {
    try {
      const token = await this.auth.getAccessToken();
      const url = "https://graph.microsoft.com/v1.0/me/messages";
      const formattedDate = dayjs(date).format("YYYY/MM/DD");

      const searchQuery = [
        `received:${formattedDate}`,
        ...(subjectFilter ? [`subject:${subjectFilter}`] : []),
      ].join(" AND ");

      const params = {
        $search: `"${searchQuery}"`,
        $select: "from,subject,body,receivedDateTime",
      };

      const fullResult = [];
      const res = await Axon.new()
        .bearer(token)
        .params(params)
        .get(url);
      fullResult.push(...res.data.value);

      let nextLink = res.data["@odata.nextLink"] || "";
      while (nextLink) {
        const nextRes = await Axon.new().bearer(token).get(nextLink);
        fullResult.push(...nextRes.data.value);
        nextLink = nextRes.data["@odata.nextLink"] || "";
      }

      return fullResult
        .filter((res) => res.subject)
        .map((res) => ({
          from: res.from.emailAddress,
          subject: res.subject,
          body: this.parseMailBody(res.body.content),
          receivedDateTime: res.receivedDateTime,
        }));
    } catch (error) {
      this.auth.handleApiError(error as AxonError);
    }
  }

  /**
   * Create a fluent email builder for composing emails
   *
   * @returns MailBuilder instance
   *
   * @example
   * await outlook.compose()
   *   .subject('Meeting Reminder')
   *   .body('Don\'t forget our meeting tomorrow!', 'Text')
   *   .to(['colleague@example.com'])
   *   .cc(['manager@example.com'])
   *   .importance('high')
   *   .send();
   *
   * @example
   * // With attachments
   * await outlook.compose()
   *   .subject('Report')
   *   .body('<h1>Monthly Report</h1>', 'HTML')
   *   .to(['boss@example.com'])
   *   .attachments(['./report.pdf', './charts.xlsx'])
   *   .send();
   */
  compose(): MailBuilder {
    return new MailBuilder(this);
  }

  /**
   * Parse HTML email body to plain text
   */
  private parseMailBody(html: string): string {
    const dom = new JSDOM(html);
    const text = dom.window.document.body.textContent;
    return text || "";
  }
}

/**
 * Fluent email builder for composing and sending emails
 * Use via outlook.compose()
 */
export class MailBuilder {
  private payload: MailPayload;
  private outlook: Outlook;

  constructor(outlook: Outlook) {
    this.outlook = outlook;
    this.payload = {
      message: {
        subject: "",
        body: {
          contentType: "Text",
          content: "",
        },
        toRecipients: [],
        attachments: [],
        importance: "normal",
        isReadReceiptRequested: false,
      },
      saveToSentItems: "true",
    };
  }

  /**
   * Set email subject
   */
  subject(subject: string): MailBuilder {
    this.payload.message.subject = subject;
    return this;
  }

  /**
   * Set email body preview (summary shown in inbox)
   */
  bodyPreview(bodyPreview: string): MailBuilder {
    this.payload.message.bodyPreview = bodyPreview;
    return this;
  }

  /**
   * Set email body content
   *
   * @param content - Email body content
   * @param contentType - 'Text' or 'HTML' (default: 'Text')
   */
  body(content: string, contentType: "HTML" | "Text" = "Text"): MailBuilder {
    this.payload.message.body = { content, contentType };
    return this;
  }

  /**
   * Set unique body (different from body for threading)
   */
  uniqueBody(content: string, contentType: "HTML" | "Text" = "Text"): MailBuilder {
    this.payload.message.uniqueBody = { content, contentType };
    return this;
  }

  /**
   * Set custom from address (requires Send As permission)
   */
  from(emailAddress: { address: string; name?: string }): MailBuilder {
    this.payload.message.from = { emailAddress };
    return this;
  }

  /**
   * Add mention headers (X-Mentions)
   */
  mention(recipients: string[]): MailBuilder {
    this.payload.message.internetMessageHeaders = recipients.map((mention) => ({
      name: "X-Mentions",
      value: mention,
    }));
    return this;
  }

  /**
   * Set reply-to addresses
   */
  replyTo(recipients: string[]): MailBuilder {
    this.payload.message.replyTo = recipients.map((replyTo) => ({
      emailAddress: { address: replyTo },
    }));
    return this;
  }

  /**
   * Set primary recipients (To field)
   */
  to(recipients: string[]): MailBuilder {
    this.payload.message.toRecipients = recipients.map((to) => ({
      emailAddress: { address: to },
    }));
    return this;
  }

  /**
   * Set carbon copy recipients (Cc field)
   */
  cc(recipients: string[]): MailBuilder {
    this.payload.message.ccRecipients = recipients.map((cc) => ({
      emailAddress: { address: cc },
    }));
    return this;
  }

  /**
   * Set blind carbon copy recipients (Bcc field)
   */
  bcc(recipients: string[]): MailBuilder {
    this.payload.message.bccRecipients = recipients.map((bcc) => ({
      emailAddress: { address: bcc },
    }));
    return this;
  }

  /**
   * Add file attachments
   *
   * @param paths - Array of file paths to attach
   * @param filesOrBuffers - Alternative: array of {name, content: Buffer, contentType?}
   *
   * @example
   * .attachments(['./report.pdf', './data.xlsx'])
   *
   * @example
   * .attachments([
   *   { name: 'data.json', content: Buffer.from('{}'), contentType: 'application/json' }
   * ])
   */
  attachments(
    filesOrBuffers: string[] | Array<{ name: string; content: Buffer; contentType?: string }>
  ): MailBuilder {
    if (!this.payload.message.attachments) {
      this.payload.message.attachments = [];
    }

    for (const item of filesOrBuffers) {
      try {
        if (typeof item === "string") {
          // File path
          const fileContent = fs.readFileSync(item);
          const fileName = path.basename(item);
          const contentType = this.getMimeType(fileName);

          this.payload.message.attachments.push({
            "@odata.type": "#microsoft.graph.fileAttachment",
            name: fileName,
            contentType,
            contentBytes: fileContent.toString("base64"),
          });
        } else {
          // Buffer object
          this.payload.message.attachments.push({
            "@odata.type": "#microsoft.graph.fileAttachment",
            name: item.name,
            contentType: item.contentType || this.getMimeType(item.name),
            contentBytes: item.content.toString("base64"),
          });
        }
      } catch (error) {
        console.error(`Error processing attachment:`, error);
        throw new Error(
          `Failed to process attachment: ${typeof item === "string" ? item : item.name}`
        );
      }
    }

    return this;
  }

  /**
   * Set whether to save email to Sent Items folder
   *
   * @param save - true to save, false to skip (default: true)
   */
  saveToSentItems(save: boolean): MailBuilder {
    this.payload.saveToSentItems = save ? "true" : "false";
    return this;
  }

  /**
   * Set email importance/priority
   */
  importance(importance: "low" | "normal" | "high"): MailBuilder {
    this.payload.message.importance = importance;
    return this;
  }

  /**
   * Add categories/tags to the email
   */
  categories(categories: string[]): MailBuilder {
    this.payload.message.categories = categories;
    return this;
  }

  /**
   * Request read receipt
   */
  requestReadReceipt(request: boolean = true): MailBuilder {
    this.payload.message.isReadReceiptRequested = request;
    return this;
  }

  /**
   * Flag the email for follow-up
   */
  flag(): MailBuilder {
    this.payload.message.flag = { flagStatus: "flagged" };
    return this;
  }

  /**
   * Send the email
   *
   * @returns Axios response from Microsoft Graph API
   */
  async send() {
    return await this.outlook.sendMail(this.payload);
  }

  /**
   * Get the current payload (useful for debugging)
   */
  getPayload(): MailPayload {
    return this.payload;
  }

  /**
   * Helper to infer MIME type from file extension
   */
  private getMimeType(fileName: string): string {
    const ext = path.extname(fileName).toLowerCase();
    const mimeTypes: Record<string, string> = {
      ".pdf": "application/pdf",
      ".doc": "application/msword",
      ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      ".xls": "application/vnd.ms-excel",
      ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      ".ppt": "application/vnd.ms-powerpoint",
      ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      ".jpg": "image/jpeg",
      ".jpeg": "image/jpeg",
      ".png": "image/png",
      ".gif": "image/gif",
      ".txt": "text/plain",
      ".csv": "text/csv",
      ".json": "application/json",
      ".xml": "application/xml",
      ".zip": "application/zip",
      ".mp4": "video/mp4",
      ".mp3": "audio/mpeg",
    };

    return mimeTypes[ext] || "application/octet-stream";
  }
}
