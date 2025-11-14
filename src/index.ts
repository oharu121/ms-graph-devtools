import { AzureAuth } from "./core/auth";
import { Outlook } from "./services/Outlook";
import { Calendar } from "./services/Calendar";
import { Teams } from "./services/Teams";
import { SharePoint } from "./services/SharePoint";
import { AzureConfig } from "./types";

/**
 * Main Azure utility class
 * Use Azure.config() to set global configuration for all services
 *
 * @example
 * // Set global config
 * Azure.config({
 *   clientId: '...',
 *   clientSecret: '...',
 *   tenantId: '...',
 *   refreshToken: '...'
 * });
 *
 * // Use services
 * const outlook = new Outlook();
 * await outlook.sendMail({...});
 *
 * @example
 * // Or use without global config (auto-loads from env/storage)
 * const outlook = new Outlook();
 * await outlook.sendMail({...});
 */
class Azure {
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
  }

  /**
   * Reset global configuration (useful for testing)
   */
  static reset(): void {
    AzureAuth.reset();
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
