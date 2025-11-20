import { describe, it, expect, beforeEach, vi } from "vitest";
import { Outlook, MailBuilder } from "../src/services/Outlook";
import { AzureAuth } from "../src/core/auth";
import Axon from "axios-fluent";
import * as fs from "fs";

// Mock axios-fluent
vi.mock("axios-fluent");

// Mock fs
vi.mock("fs", () => ({
  readFileSync: vi.fn(),
}));

// Mock jsdom
vi.mock("jsdom", () => {
  return {
    JSDOM: class JSDOM {
      window: any;
      constructor(html: string) {
        this.window = {
          document: {
            body: {
              textContent: html.replace(/<[^>]*>/g, ""),
            },
          },
        };
      }
    },
  };
});

describe("Outlook Service", () => {
  let outlook: Outlook;
  let mockAuth: AzureAuth;

  beforeEach(() => {
    vi.clearAllMocks();

    mockAuth = new AzureAuth({
      clientId: "test-client",
      clientSecret: "test-secret",
      tenantId: "test-tenant",
      refreshToken: "test-refresh-token",
      accessToken: "mock-access-token",
    });

    vi.spyOn(mockAuth, "getAccessToken");
    vi.spyOn(mockAuth, "checkToken");
    vi.spyOn(mockAuth, "withRetry").mockImplementation(async (fn) => await fn());
    vi.spyOn(mockAuth, "getAxon").mockImplementation(() => Axon.new());
  });

  describe("Constructor", () => {
    it("should create instance with AzureAuth instance", () => {
      const ol = new Outlook(mockAuth);
      expect(ol).toBeInstanceOf(Outlook);
    });

    it("should create instance with config object", () => {
      const ol = new Outlook({
        clientId: "test-client",
        clientSecret: "test-secret",
        tenantId: "test-tenant",
        refreshToken: "test-token",
      });
      expect(ol).toBeInstanceOf(Outlook);
    });

    it("should create instance with no config", () => {
      const ol = new Outlook();
      expect(ol).toBeInstanceOf(Outlook);
    });
  });

  describe("getMe", () => {
    beforeEach(() => {
      outlook = new Outlook(mockAuth);
    });

    it("should fetch user profile", async () => {
      const mockUser = {
        id: "user-123",
        displayName: "John Doe",
        mail: "john@example.com",
      };

      const mockAxonInstance = {
        bearer: vi.fn().mockReturnThis(),
        get: vi.fn().mockResolvedValue({ data: mockUser }),
      };

      (Axon.new as any).mockReturnValue(mockAxonInstance);

      const result = await outlook.getMe();

      expect(mockAuth.getAccessToken).toHaveBeenCalled();
      expect(mockAxonInstance.bearer).toHaveBeenCalledWith("mock-access-token");
      expect(mockAxonInstance.get).toHaveBeenCalledWith(
        "https://graph.microsoft.com/v1.0/me"
      );
      expect(result).toEqual(mockUser);
    });

    it("should handle API errors", async () => {
      const mockError = new Error("Unauthorized");

      // Mock withRetry to throw the error
      vi.spyOn(mockAuth, "withRetry").mockRejectedValue(mockError);

      await expect(outlook.getMe()).rejects.toThrow("Unauthorized");
      expect(mockAuth.withRetry).toHaveBeenCalled();
    });
  });

  describe("sendMail", () => {
    beforeEach(() => {
      outlook = new Outlook(mockAuth);
    });

    it("should send email successfully", async () => {
      const payload = {
        message: {
          subject: "Test Email",
          body: {
            contentType: "Text",
            content: "Hello World",
          },
          toRecipients: [
            { emailAddress: { address: "recipient@example.com" } },
          ],
        },
      };

      const mockAxonInstance = {
        bearer: vi.fn().mockReturnThis(),
        post: vi.fn().mockResolvedValue({ status: 202 }),
      };

      (Axon.new as any).mockReturnValue(mockAxonInstance);

      const result = await outlook.sendMail(payload);

      expect(mockAxonInstance.bearer).toHaveBeenCalledWith("mock-access-token");
      expect(mockAxonInstance.post).toHaveBeenCalledWith(
        "https://graph.microsoft.com/v1.0/me/sendMail",
        payload
      );
      expect(result).toEqual({ status: 202 });
    });
  });

  describe("getMails", () => {
    beforeEach(() => {
      outlook = new Outlook(mockAuth);
    });

    it("should fetch emails with date filter", async () => {
      const mockEmails = [
        {
          from: { emailAddress: { address: "sender@example.com" } },
          subject: "Invoice #123",
          body: { content: "<p>Payment due</p>" },
          receivedDateTime: "2024-01-15T10:00:00Z",
        },
      ];

      const mockAxonInstance = {
        bearer: vi.fn().mockReturnThis(),
        params: vi.fn().mockReturnThis(),
        get: vi.fn().mockResolvedValue({
          data: { value: mockEmails },
        }),
      };

      (Axon.new as any).mockReturnValue(mockAxonInstance);

      const result = await outlook.getMails("2024-01-15");

      expect(mockAxonInstance.params).toHaveBeenCalledWith({
        $search: '"received:2024/01/15"',
        $select: "from,subject,body,receivedDateTime",
      });
      expect(result).toHaveLength(1);
      expect(result[0].subject).toBe("Invoice #123");
      expect(result[0].body).toBe("Payment due");
    });

    it("should fetch emails with date and subject filter", async () => {
      const mockEmails = [
        {
          from: { emailAddress: { address: "sender@example.com" } },
          subject: "Invoice #123",
          body: { content: "<p>Payment due</p>" },
          receivedDateTime: "2024-01-15T10:00:00Z",
        },
      ];

      const mockAxonInstance = {
        bearer: vi.fn().mockReturnThis(),
        params: vi.fn().mockReturnThis(),
        get: vi.fn().mockResolvedValue({
          data: { value: mockEmails },
        }),
      };

      (Axon.new as any).mockReturnValue(mockAxonInstance);

      await outlook.getMails("2024-01-15", "invoice");

      expect(mockAxonInstance.params).toHaveBeenCalledWith({
        $search: '"received:2024/01/15 AND subject:invoice"',
        $select: "from,subject,body,receivedDateTime",
      });
    });

    it("should handle pagination with @odata.nextLink", async () => {
      const mockEmailsPage1 = [
        {
          from: { emailAddress: { address: "sender1@example.com" } },
          subject: "Email 1",
          body: { content: "<p>Content 1</p>" },
          receivedDateTime: "2024-01-15T10:00:00Z",
        },
      ];

      const mockEmailsPage2 = [
        {
          from: { emailAddress: { address: "sender2@example.com" } },
          subject: "Email 2",
          body: { content: "<p>Content 2</p>" },
          receivedDateTime: "2024-01-15T11:00:00Z",
        },
      ];

      const mockAxonInstance = {
        bearer: vi.fn().mockReturnThis(),
        params: vi.fn().mockReturnThis(),
        get: vi.fn(),
      };

      // First page
      mockAxonInstance.get.mockResolvedValueOnce({
        data: {
          value: mockEmailsPage1,
          "@odata.nextLink":
            "https://graph.microsoft.com/v1.0/me/messages?$skip=10",
        },
      });

      // Second page
      mockAxonInstance.get.mockResolvedValueOnce({
        data: {
          value: mockEmailsPage2,
        },
      });

      (Axon.new as any).mockReturnValue(mockAxonInstance);

      const result = await outlook.getMails("2024-01-15");

      expect(result).toHaveLength(2);
      expect(result[0].subject).toBe("Email 1");
      expect(result[1].subject).toBe("Email 2");
    });

    it("should filter out emails without subject", async () => {
      const mockEmails = [
        {
          from: { emailAddress: { address: "sender@example.com" } },
          subject: "Email with subject",
          body: { content: "<p>Content</p>" },
          receivedDateTime: "2024-01-15T10:00:00Z",
        },
        {
          from: { emailAddress: { address: "sender2@example.com" } },
          subject: null,
          body: { content: "<p>No subject</p>" },
          receivedDateTime: "2024-01-15T11:00:00Z",
        },
      ];

      const mockAxonInstance = {
        bearer: vi.fn().mockReturnThis(),
        params: vi.fn().mockReturnThis(),
        get: vi.fn().mockResolvedValue({
          data: { value: mockEmails },
        }),
      };

      (Axon.new as any).mockReturnValue(mockAxonInstance);

      const result = await outlook.getMails("2024-01-15");

      expect(result).toHaveLength(1);
      expect(result[0].subject).toBe("Email with subject");
    });
  });

  describe("compose - MailBuilder", () => {
    beforeEach(() => {
      outlook = new Outlook(mockAuth);
    });

    it("should return MailBuilder instance", () => {
      const builder = outlook.compose();
      expect(builder).toBeInstanceOf(MailBuilder);
    });
  });
});

describe("MailBuilder", () => {
  let outlook: Outlook;
  let mockAuth: AzureAuth;
  let builder: MailBuilder;

  beforeEach(() => {
    vi.clearAllMocks();

    mockAuth = new AzureAuth({
      clientId: "test-client",
      clientSecret: "test-secret",
      tenantId: "test-tenant",
      refreshToken: "test-refresh-token",
      accessToken: "mock-access-token",
    });

    vi.spyOn(mockAuth, "getAccessToken");
    vi.spyOn(mockAuth, "handleApiError");

    outlook = new Outlook(mockAuth);
    builder = outlook.compose();
  });

  describe("Fluent API", () => {
    it("should build email with subject", () => {
      builder.subject("Test Subject");
      const payload = builder.getPayload();
      expect(payload.message.subject).toBe("Test Subject");
    });

    it("should build email with body", () => {
      builder.body("Test Body", "Text");
      const payload = builder.getPayload();
      expect(payload.message.body.content).toBe("Test Body");
      expect(payload.message.body.contentType).toBe("Text");
    });

    it("should build email with HTML body", () => {
      builder.body("<h1>HTML Body</h1>", "HTML");
      const payload = builder.getPayload();
      expect(payload.message.body.content).toBe("<h1>HTML Body</h1>");
      expect(payload.message.body.contentType).toBe("HTML");
    });

    it("should build email with bodyPreview", () => {
      builder.bodyPreview("Preview text");
      const payload = builder.getPayload();
      expect(payload.message.bodyPreview).toBe("Preview text");
    });

    it("should build email with uniqueBody", () => {
      builder.uniqueBody("Unique content", "HTML");
      const payload = builder.getPayload();
      expect(payload.message.uniqueBody).toEqual({
        content: "Unique content",
        contentType: "HTML",
      });
    });

    it("should build email with from address", () => {
      builder.from({ address: "custom@example.com", name: "Custom Sender" });
      const payload = builder.getPayload();
      expect(payload.message.from).toEqual({
        emailAddress: { address: "custom@example.com", name: "Custom Sender" },
      });
    });

    it("should build email with to recipients", () => {
      builder.to(["user1@example.com", "user2@example.com"]);
      const payload = builder.getPayload();
      expect(payload.message.toRecipients).toEqual([
        { emailAddress: { address: "user1@example.com" } },
        { emailAddress: { address: "user2@example.com" } },
      ]);
    });

    it("should build email with cc recipients", () => {
      builder.cc(["cc1@example.com", "cc2@example.com"]);
      const payload = builder.getPayload();
      expect(payload.message.ccRecipients).toEqual([
        { emailAddress: { address: "cc1@example.com" } },
        { emailAddress: { address: "cc2@example.com" } },
      ]);
    });

    it("should build email with bcc recipients", () => {
      builder.bcc(["bcc1@example.com"]);
      const payload = builder.getPayload();
      expect(payload.message.bccRecipients).toEqual([
        { emailAddress: { address: "bcc1@example.com" } },
      ]);
    });

    it("should build email with replyTo", () => {
      builder.replyTo(["replyto@example.com"]);
      const payload = builder.getPayload();
      expect(payload.message.replyTo).toEqual([
        { emailAddress: { address: "replyto@example.com" } },
      ]);
    });

    it("should build email with mentions", () => {
      builder.mention(["user1@example.com", "user2@example.com"]);
      const payload = builder.getPayload();
      expect(payload.message.internetMessageHeaders).toEqual([
        { name: "X-Mentions", value: "user1@example.com" },
        { name: "X-Mentions", value: "user2@example.com" },
      ]);
    });

    it("should set importance", () => {
      builder.importance("high");
      const payload = builder.getPayload();
      expect(payload.message.importance).toBe("high");
    });

    it("should set categories", () => {
      builder.categories(["Work", "Important"]);
      const payload = builder.getPayload();
      expect(payload.message.categories).toEqual(["Work", "Important"]);
    });

    it("should request read receipt", () => {
      builder.requestReadReceipt(true);
      const payload = builder.getPayload();
      expect(payload.message.isReadReceiptRequested).toBe(true);
    });

    it("should set flag", () => {
      builder.flag();
      const payload = builder.getPayload();
      expect(payload.message.flag).toEqual({ flagStatus: "flagged" });
    });

    it("should set saveToSentItems", () => {
      builder.saveToSentItems(false);
      const payload = builder.getPayload();
      expect(payload.saveToSentItems).toBe("false");
    });

    it("should chain multiple methods", () => {
      builder
        .subject("Chained Email")
        .body("Chained content")
        .to(["recipient@example.com"])
        .importance("high");

      const payload = builder.getPayload();
      expect(payload.message.subject).toBe("Chained Email");
      expect(payload.message.body.content).toBe("Chained content");
      expect(payload.message.toRecipients).toHaveLength(1);
      expect(payload.message.importance).toBe("high");
    });
  });

  describe("Attachments", () => {
    it("should add file attachments from buffer", () => {
      const mockBuffer = Buffer.from("test content");
      builder.attachments([
        {
          name: "test.txt",
          content: mockBuffer,
          contentType: "text/plain",
        },
      ]);

      const payload = builder.getPayload();
      expect(payload.message.attachments).toHaveLength(1);
      expect(payload.message.attachments![0]).toEqual({
        "@odata.type": "#microsoft.graph.fileAttachment",
        name: "test.txt",
        contentType: "text/plain",
        contentBytes: mockBuffer.toString("base64"),
      });
    });

    it("should infer MIME type from file extension for buffer", () => {
      const mockBuffer = Buffer.from("pdf content");
      builder.attachments([
        {
          name: "document.pdf",
          content: mockBuffer,
        },
      ]);

      const payload = builder.getPayload();
      expect(payload.message.attachments![0].contentType).toBe(
        "application/pdf"
      );
    });

    it("should add file attachments from file path", () => {
      const mockBuffer = Buffer.from("file content");
      (fs.readFileSync as any).mockReturnValue(mockBuffer);

      builder.attachments(["test.txt"]);

      const payload = builder.getPayload();
      expect(fs.readFileSync).toHaveBeenCalledWith("test.txt");
      expect(payload.message.attachments).toHaveLength(1);
      expect(payload.message.attachments![0].name).toBe("test.txt");
    });

    it("should handle file read errors", () => {
      (fs.readFileSync as any).mockImplementation(() => {
        throw new Error("File not found");
      });

      expect(() => {
        builder.attachments(["nonexistent.txt"]);
      }).toThrow("Failed to process attachment: nonexistent.txt");
    });

    it("should support multiple attachments", () => {
      const buffer1 = Buffer.from("content1");
      const buffer2 = Buffer.from("content2");

      builder.attachments([
        { name: "file1.txt", content: buffer1 },
        { name: "file2.pdf", content: buffer2 },
      ]);

      const payload = builder.getPayload();
      expect(payload.message.attachments).toHaveLength(2);
    });

    it("should infer MIME types correctly", () => {
      const testCases = [
        { ext: "file.pdf", expected: "application/pdf" },
        { ext: "file.docx", expected: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" },
        { ext: "file.xlsx", expected: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
        { ext: "file.jpg", expected: "image/jpeg" },
        { ext: "file.png", expected: "image/png" },
        { ext: "file.json", expected: "application/json" },
        { ext: "file.unknown", expected: "application/octet-stream" },
      ];

      testCases.forEach(({ ext, expected }) => {
        const b = outlook.compose();
        const buffer = Buffer.from("test");
        b.attachments([{ name: ext, content: buffer }]);
        const payload = b.getPayload();
        expect(payload.message.attachments![0].contentType).toBe(expected);
      });
    });
  });

  describe("send", () => {
    it("should send email successfully", async () => {
      const mockAxonInstance = {
        bearer: vi.fn().mockReturnThis(),
        post: vi.fn().mockResolvedValue({ status: 202 }),
      };

      (Axon.new as any).mockReturnValue(mockAxonInstance);

      builder.subject("Test").to(["user@example.com"]);

      const result = await builder.send();

      expect(mockAuth.getAccessToken).toHaveBeenCalled();
      expect(result).toEqual({ status: 202 });
    });

    it("should propagate errors from send", async () => {
      const mockError = new Error("Send failed");
      const mockAxonInstance = {
        bearer: vi.fn().mockReturnThis(),
        post: vi.fn().mockRejectedValue(mockError),
      };

      (Axon.new as any).mockReturnValue(mockAxonInstance);

      builder.subject("Test").to(["user@example.com"]);

      await expect(builder.send()).rejects.toThrow("Send failed");
    });
  });

  describe("getPayload", () => {
    it("should return current payload", () => {
      builder
        .subject("Test Subject")
        .body("Test Body")
        .to(["recipient@example.com"]);

      const payload = builder.getPayload();

      expect(payload.message.subject).toBe("Test Subject");
      expect(payload.message.body.content).toBe("Test Body");
      expect(payload.message.toRecipients).toHaveLength(1);
    });
  });
});
