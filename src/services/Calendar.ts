import Axon from "axios-fluent";
import { AzureAuth } from "../core/auth";
import { AzureConfig, Calendar as CalendarType, Holiday } from "../types";

/**
 * Calendar service for Microsoft Graph API
 * Handles calendar and holiday operations
 */
export class Calendar {
  private auth: AzureAuth;

  /**
   * Create a new Calendar service instance
   *
   * @param config - Optional config or AzureAuth instance
   *
   * @example
   * const calendar = new Calendar();
   * const calendars = await calendar.getCalendars();
   */
  constructor(config?: AzureConfig | AzureAuth) {
    if (config instanceof AzureAuth) {
      this.auth = config;
    } else {
      this.auth = new AzureAuth(config);
    }
  }

  /**
   * Get all calendars for the current user
   *
   * @returns Array of calendars
   *
   * @example
   * const calendars = await calendar.getCalendars();
   * console.log(calendars.map(c => c.name));
   */
  async getCalendars(): Promise<CalendarType[]> {
    try {
      const token = await this.auth.getAccessToken();
      const url = `https://graph.microsoft.com/v1.0/me/calendars`;
      const res = await Axon.new().bearer(token).get(url);
      return res.data.value as CalendarType[];
    } catch (error) {
      this.auth.handleApiError(error);
    }
  }

  /**
   * Get holidays from a specific calendar by name
   *
   * @param calendarName - Calendar name to search for (e.g., 'India holidays', 'US Holidays')
   * @param start - Start date (ISO format)
   * @param end - End date (ISO format)
   * @returns Array of holidays
   *
   * @example
   * const holidays = await calendar.getHolidaysByCalendarName(
   *   'India holidays',
   *   '2024-01-01T00:00:00Z',
   *   '2024-12-31T23:59:59Z'
   * );
   */
  async getHolidaysByCalendarName(
    calendarName: string,
    start: string,
    end: string
  ): Promise<Holiday[]> {
    try {
      const token = await this.auth.getAccessToken();
      const calendars = await this.getCalendars();
      const targetCalendar = calendars.find(
        (calendar) => calendar.name === calendarName
      );

      if (!targetCalendar) {
        throw new Error(`Calendar "${calendarName}" not found`);
      }

      const params = {
        startDateTime: start,
        endDateTime: end,
      };

      const url = `https://graph.microsoft.com/v1.0/me/calendars/${targetCalendar.id}/calendarView`;
      const res = await Axon.new()
        .bearer(token)
        .params(params)
        .get(url);

      type HolidayEvent = { subject: string; start: { dateTime: string } };

      return res.data.value.map((holiday: HolidayEvent) => ({
        name: holiday.subject,
        date: holiday.start.dateTime,
      }));
    } catch (error) {
      this.auth.handleApiError(error);
    }
  }

  /**
   * Get India holidays within a date range
   *
   * @param start - Start date (ISO format)
   * @param end - End date (ISO format)
   * @param calendarName - Optional calendar name (defaults to 'India holidays')
   * @returns Array of holidays
   *
   * @example
   * const holidays = await calendar.getIndiaHolidays(
   *   '2024-01-01T00:00:00Z',
   *   '2024-12-31T23:59:59Z'
   * );
   *
   * @example
   * // With custom calendar name
   * const holidays = await calendar.getIndiaHolidays(
   *   '2024-01-01T00:00:00Z',
   *   '2024-12-31T23:59:59Z',
   *   'भारत की छुट्टियाँ'
   * );
   */
  async getIndiaHolidays(
    start: string,
    end: string,
    calendarName: string = "India holidays"
  ): Promise<Holiday[]> {
    return this.getHolidaysByCalendarName(calendarName, start, end);
  }

  /**
   * Get Japan holidays within a date range
   *
   * @param start - Start date (ISO format)
   * @param end - End date (ISO format)
   * @param calendarNames - Optional calendar names to search for (defaults to common Japanese holiday calendar names)
   * @returns Array of holiday events
   *
   * @example
   * const holidays = await calendar.getJapanHolidays(
   *   '2024-01-01T00:00:00Z',
   *   '2024-12-31T23:59:59Z'
   * );
   *
   * @example
   * // With custom calendar name
   * const holidays = await calendar.getJapanHolidays(
   *   '2024-01-01T00:00:00Z',
   *   '2024-12-31T23:59:59Z',
   *   ['Japan holidays', '日本の休日', 'Japanese Holidays']
   * );
   */
  async getJapanHolidays(
    start: string,
    end: string,
    calendarNames: string[] = ["Japan holidays", "日本 の休日"]
  ): Promise<Holiday[]> {
    try {
      const token = await this.auth.getAccessToken();
      const calendars = await this.getCalendars();
      const japanCalendar = calendars.find((calendar) =>
        calendarNames.includes(calendar.name)
      );

      if (!japanCalendar) {
        throw new Error(
          `Japan holidays calendar not found. Searched for: ${calendarNames.join(", ")}`
        );
      }

      const params = {
        startDateTime: start,
        endDateTime: end,
      };

      const url = `https://graph.microsoft.com/v1.0/me/calendars/${japanCalendar.id}/calendarView`;
      const res = await Axon.new()
        .bearer(token)
        .params(params)
        .get(url);

      type HolidayEvent = { subject: string; start: { dateTime: string } };

      return res.data.value.map((holiday: HolidayEvent) => ({
        name: holiday.subject,
        date: holiday.start.dateTime,
      }));
    } catch (error) {
      this.auth.handleApiError(error);
    }
  }
}
