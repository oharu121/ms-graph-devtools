import { AzureAuth } from "./core/auth";
import { Outlook } from "./services/Outlook";
import { Calendar } from "./services/Calendar";
import { Teams } from "./services/Teams";
import { SharePoint } from "./services/SharePoint";
import { AzureConfig } from "./types";

/**
 * Main Azure utility class
 * Use Azure.config() to set global configuration, then access services directly
 *
 * @example
 * // Set global config once
 * Azure.config({
 *   clientId: '...',
 *   clientSecret: '...',
 *   tenantId: '...',
 *   refreshToken: '...'
 * });
 *
 * // Use services directly (no instantiation needed)
 * await Azure.outlook.sendMail({...});
 * await Azure.teams.postMessage({...});
 * await Azure.calendar.listEvents({...});
 *
 * @example
 * // With token provider (infinite renewal)
 * Azure.config({
 *   clientId: '...',
 *   clientSecret: '...',
 *   tenantId: '...',
 *   tokenProvider: async (callback) => await getAuthCode(callback)
 * });
 *
 * export default Azure; // Use globally in your app
 */
class Azure {
  private static _outlook?: Outlook;
  private static _calendar?: Calendar;
  private static _teams?: Teams;
  private static _sharePoint?: SharePoint;

  /**
   * Set global configuration for all service instances
   * Similar to AWS.config.update()
   *
   * @param config - Azure configuration
   *
   * @example
   * Azure.config({
   *   clientId: process.env.AZURE_CLIENT_ID,
   *   clientSecret: process.env.AZURE_CLIENT_SECRET,
   *   tenantId: process.env.AZURE_TENANT_ID,
   *   refreshToken: process.env.AZURE_REFRESH_TOKEN
   * });
   */
  static config(config: AzureConfig): void {
    AzureAuth.setGlobalConfig(config);
    // Reset service instances so they pick up new config
    this._outlook = undefined;
    this._calendar = undefined;
    this._teams = undefined;
    this._sharePoint = undefined;
  }

  /**
   * Reset global configuration (useful for testing)
   */
  static reset(): void {
    AzureAuth.reset();
    // Clear all service instances
    this._outlook = undefined;
    this._calendar = undefined;
    this._teams = undefined;
    this._sharePoint = undefined;
  }

  /**
   * Get Outlook service instance (lazy-loaded singleton)
   * Uses global config set via Azure.config()
   *
   * @example
   * await Azure.outlook.sendMail({...});
   */
  static get outlook(): Outlook {
    if (!this._outlook) {
      this._outlook = new Outlook();
    }
    return this._outlook;
  }

  /**
   * Get Calendar service instance (lazy-loaded singleton)
   * Uses global config set via Azure.config()
   *
   * @example
   * await Azure.calendar.listEvents({...});
   */
  static get calendar(): Calendar {
    if (!this._calendar) {
      this._calendar = new Calendar();
    }
    return this._calendar;
  }

  /**
   * Get Teams service instance (lazy-loaded singleton)
   * Uses global config set via Azure.config()
   *
   * @example
   * await Azure.teams.postMessage({...});
   */
  static get teams(): Teams {
    if (!this._teams) {
      this._teams = new Teams();
    }
    return this._teams;
  }

  /**
   * Get SharePoint service instance (lazy-loaded singleton)
   * Uses global config set via Azure.config()
   *
   * @example
   * await Azure.sharePoint.getList({...});
   */
  static get sharePoint(): SharePoint {
    if (!this._sharePoint) {
      this._sharePoint = new SharePoint();
    }
    return this._sharePoint;
  }

  /**
   * List all stored credentials
   *
   * @returns Array of stored credential metadata
   */
  static async listStoredCredentials() {
    return AzureAuth.listStoredCredentials();
  }

  /**
   * Clear stored credentials
   *
   * @param tenantId - Optional tenant ID to clear specific credentials
   * @param clientId - Optional client ID to clear specific credentials
   */
  static async clearStoredCredentials(tenantId?: string, clientId?: string) {
    return AzureAuth.clearStoredCredentials(tenantId, clientId);
  }
}

// Export main class
export default Azure;

// Export service classes
export { Outlook, Calendar, Teams, SharePoint };

// Export builder classes
export { MailBuilder } from "./services/Outlook";
export { AdaptiveCardBuilder } from "./services/Teams";

// Export auth class for advanced use
export { AzureAuth };

// Export types
export type { AzureConfig } from "./types";
