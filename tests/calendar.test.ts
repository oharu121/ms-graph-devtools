import { describe, it, expect, beforeEach, vi } from "vitest";
import { Calendar } from "../src/services/Calendar";
import { AzureAuth } from "../src/core/auth";
import Axon from "axios-fluent";

// Mock axios-fluent
vi.mock("axios-fluent", () => {
  const mockGet = vi.fn();
  const mockBearer = vi.fn(() => ({ get: mockGet }));
  const mockParams = vi.fn(() => ({ get: mockGet }));
  const mockNew = vi.fn(() => ({
    bearer: mockBearer,
    params: mockParams,
  }));

  return {
    default: {
      new: mockNew,
    },
  };
});

describe("Calendar Service", () => {
  let calendar: Calendar;
  let mockAuth: AzureAuth;

  beforeEach(() => {
    vi.clearAllMocks();

    // Create mock auth instance with all required properties
    mockAuth = new AzureAuth({
      clientId: "test-client",
      clientSecret: "test-secret",
      tenantId: "test-tenant",
      refreshToken: "test-refresh-token",
      accessToken: "mock-access-token",
    });

    // Spy on methods
    vi.spyOn(mockAuth, "getAccessToken");
    vi.spyOn(mockAuth, "checkToken");
    vi.spyOn(mockAuth, "withRetry").mockImplementation(async (fn) => await fn());
    vi.spyOn(mockAuth, "getAxon").mockImplementation(() => Axon.new());
  });

  describe("Constructor", () => {
    it("should create instance with AzureAuth instance", () => {
      const cal = new Calendar(mockAuth);
      expect(cal).toBeInstanceOf(Calendar);
    });

    it("should create instance with config object", () => {
      const cal = new Calendar({
        clientId: "test-client",
        clientSecret: "test-secret",
        tenantId: "test-tenant",
        refreshToken: "test-token",
      });
      expect(cal).toBeInstanceOf(Calendar);
    });

    it("should create instance with no config", () => {
      const cal = new Calendar();
      expect(cal).toBeInstanceOf(Calendar);
    });
  });

  describe("getCalendars", () => {
    beforeEach(() => {
      calendar = new Calendar(mockAuth);
    });

    it("should fetch and return calendars", async () => {
      const mockCalendars = [
        { id: "cal-1", name: "My Calendar" },
        { id: "cal-2", name: "Work Calendar" },
      ];

      const mockAxonInstance = {
        bearer: vi.fn().mockReturnThis(),
        get: vi.fn().mockResolvedValue({
          data: { value: mockCalendars },
        }),
      };

      (Axon.new as any).mockReturnValue(mockAxonInstance);

      const result = await calendar.getCalendars();

      expect(mockAuth.getAccessToken).toHaveBeenCalled();
      expect(mockAxonInstance.bearer).toHaveBeenCalledWith("mock-access-token");
      expect(mockAxonInstance.get).toHaveBeenCalledWith(
        "https://graph.microsoft.com/v1.0/me/calendars"
      );
      expect(result).toEqual(mockCalendars);
    });

    it("should handle API errors", async () => {
      const mockError = new Error("API Error");

      // Mock withRetry to throw the error
      vi.spyOn(mockAuth, "withRetry").mockRejectedValue(mockError);

      await expect(calendar.getCalendars()).rejects.toThrow("API Error");
      expect(mockAuth.withRetry).toHaveBeenCalled();
    });
  });

  describe("getHolidaysByCalendarName", () => {
    beforeEach(() => {
      calendar = new Calendar(mockAuth);
    });

    it("should fetch holidays from specified calendar", async () => {
      const mockCalendars = [
        { id: "cal-1", name: "India holidays" },
        { id: "cal-2", name: "US Holidays" },
      ];

      const mockHolidays = [
        { subject: "Diwali", start: { dateTime: "2024-11-01T00:00:00Z" } },
        {
          subject: "Independence Day",
          start: { dateTime: "2024-08-15T00:00:00Z" },
        },
      ];

      const mockAxonInstance = {
        bearer: vi.fn().mockReturnThis(),
        params: vi.fn().mockReturnThis(),
        get: vi.fn(),
      };

      // First call for getCalendars
      mockAxonInstance.get.mockResolvedValueOnce({
        data: { value: mockCalendars },
      });

      // Second call for calendar view
      mockAxonInstance.get.mockResolvedValueOnce({
        data: { value: mockHolidays },
      });

      (Axon.new as any).mockReturnValue(mockAxonInstance);

      const result = await calendar.getHolidaysByCalendarName(
        "India holidays",
        "2024-01-01T00:00:00Z",
        "2024-12-31T23:59:59Z"
      );

      expect(result).toEqual([
        { name: "Diwali", date: "2024-11-01T00:00:00Z" },
        { name: "Independence Day", date: "2024-08-15T00:00:00Z" },
      ]);

      expect(mockAxonInstance.params).toHaveBeenCalledWith({
        startDateTime: "2024-01-01T00:00:00Z",
        endDateTime: "2024-12-31T23:59:59Z",
      });
    });

    it("should throw error when calendar not found", async () => {
      const mockCalendars = [{ id: "cal-1", name: "US Holidays" }];

      const mockAxonInstance = {
        bearer: vi.fn().mockReturnThis(),
        get: vi.fn().mockResolvedValue({
          data: { value: mockCalendars },
        }),
      };

      (Axon.new as any).mockReturnValue(mockAxonInstance);

      await expect(
        calendar.getHolidaysByCalendarName(
          "India holidays",
          "2024-01-01T00:00:00Z",
          "2024-12-31T23:59:59Z"
        )
      ).rejects.toThrow('Calendar "India holidays" not found');
    });
  });

  describe("getIndiaHolidays", () => {
    beforeEach(() => {
      calendar = new Calendar(mockAuth);
    });

    it("should fetch India holidays with default calendar name", async () => {
      const mockCalendars = [{ id: "cal-1", name: "India holidays" }];
      const mockHolidays = [
        { subject: "Diwali", start: { dateTime: "2024-11-01T00:00:00Z" } },
      ];

      const mockAxonInstance = {
        bearer: vi.fn().mockReturnThis(),
        params: vi.fn().mockReturnThis(),
        get: vi.fn(),
      };

      mockAxonInstance.get.mockResolvedValueOnce({
        data: { value: mockCalendars },
      });

      mockAxonInstance.get.mockResolvedValueOnce({
        data: { value: mockHolidays },
      });

      (Axon.new as any).mockReturnValue(mockAxonInstance);

      const result = await calendar.getIndiaHolidays(
        "2024-01-01T00:00:00Z",
        "2024-12-31T23:59:59Z"
      );

      expect(result).toEqual([
        { name: "Diwali", date: "2024-11-01T00:00:00Z" },
      ]);
    });

    it("should fetch India holidays with custom calendar name", async () => {
      const mockCalendars = [{ id: "cal-1", name: "भारत की छुट्टियाँ" }];
      const mockHolidays = [
        { subject: "होली", start: { dateTime: "2024-03-25T00:00:00Z" } },
      ];

      const mockAxonInstance = {
        bearer: vi.fn().mockReturnThis(),
        params: vi.fn().mockReturnThis(),
        get: vi.fn(),
      };

      mockAxonInstance.get.mockResolvedValueOnce({
        data: { value: mockCalendars },
      });

      mockAxonInstance.get.mockResolvedValueOnce({
        data: { value: mockHolidays },
      });

      (Axon.new as any).mockReturnValue(mockAxonInstance);

      const result = await calendar.getIndiaHolidays(
        "2024-01-01T00:00:00Z",
        "2024-12-31T23:59:59Z",
        "भारत की छुट्टियाँ"
      );

      expect(result).toEqual([
        { name: "होली", date: "2024-03-25T00:00:00Z" },
      ]);
    });
  });

  describe("getJapanHolidays", () => {
    beforeEach(() => {
      calendar = new Calendar(mockAuth);
    });

    it("should fetch Japan holidays with default calendar names", async () => {
      const mockCalendars = [{ id: "cal-1", name: "Japan holidays" }];
      const mockHolidays = [
        { subject: "New Year", start: { dateTime: "2024-01-01T00:00:00Z" } },
        {
          subject: "Golden Week",
          start: { dateTime: "2024-05-03T00:00:00Z" },
        },
      ];

      const mockAxonInstance = {
        bearer: vi.fn().mockReturnThis(),
        params: vi.fn().mockReturnThis(),
        get: vi.fn(),
      };

      mockAxonInstance.get.mockResolvedValueOnce({
        data: { value: mockCalendars },
      });

      mockAxonInstance.get.mockResolvedValueOnce({
        data: { value: mockHolidays },
      });

      (Axon.new as any).mockReturnValue(mockAxonInstance);

      const result = await calendar.getJapanHolidays(
        "2024-01-01T00:00:00Z",
        "2024-12-31T23:59:59Z"
      );

      expect(result).toEqual([
        { name: "New Year", date: "2024-01-01T00:00:00Z" },
        { name: "Golden Week", date: "2024-05-03T00:00:00Z" },
      ]);
    });

    it("should fetch Japan holidays with custom calendar names", async () => {
      const mockCalendars = [{ id: "cal-1", name: "日本の休日" }];
      const mockHolidays = [
        { subject: "正月", start: { dateTime: "2024-01-01T00:00:00Z" } },
      ];

      const mockAxonInstance = {
        bearer: vi.fn().mockReturnThis(),
        params: vi.fn().mockReturnThis(),
        get: vi.fn(),
      };

      mockAxonInstance.get.mockResolvedValueOnce({
        data: { value: mockCalendars },
      });

      mockAxonInstance.get.mockResolvedValueOnce({
        data: { value: mockHolidays },
      });

      (Axon.new as any).mockReturnValue(mockAxonInstance);

      const result = await calendar.getJapanHolidays(
        "2024-01-01T00:00:00Z",
        "2024-12-31T23:59:59Z",
        ["日本の休日"]
      );

      expect(result).toEqual([
        { name: "正月", date: "2024-01-01T00:00:00Z" },
      ]);
    });

    it("should throw error when Japan calendar not found", async () => {
      const mockCalendars = [{ id: "cal-1", name: "US Holidays" }];

      const mockAxonInstance = {
        bearer: vi.fn().mockReturnThis(),
        get: vi.fn().mockResolvedValue({
          data: { value: mockCalendars },
        }),
      };

      (Axon.new as any).mockReturnValue(mockAxonInstance);

      await expect(
        calendar.getJapanHolidays(
          "2024-01-01T00:00:00Z",
          "2024-12-31T23:59:59Z"
        )
      ).rejects.toThrow(
        "Japan holidays calendar not found. Searched for: Japan holidays, 日本 の休日"
      );
    });
  });
});
