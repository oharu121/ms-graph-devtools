import Axon, { AxonError } from "axios-fluent";
import os from "os";
import path from "path";
import fs from "fs/promises";
import { AzureConfig, StoredCredentials } from "../types";

const REDIRECT_URI = "https://oauth.pstmn.io/v1/callback";

/**
 * Default OAuth scopes that work in 99% of cases without admin consent
 *
 * Scope Explanations:
 * - openid: Required for OpenID Connect authentication
 * - profile: Access to user's basic profile information
 * - offline_access: Required to receive refresh tokens for long-term access
 * - User.Read: Read user's profile (name, email, etc.)
 * - Mail.Send: Send emails on behalf of the user
 * - Mail.Read: Read user's emails
 * - Calendars.ReadWrite: Read and write to user's calendars
 * - Calendars.ReadWrite.Shared: Read and write to calendars shared with the user
 * - ChannelMessage.Send: Send messages to Teams channels
 * - ChatMessage.Send: Send chat messages in Teams
 *
 * IMPORTANT - Admin Consent Required:
 * - Sites.ReadWrite.All: Requires admin consent in most organizations
 *   Only add this if you have admin approval or are an admin yourself
 *
 * Users can override these scopes by providing custom scopes in config
 */
const DEFAULT_SCOPES = [
  "openid",
  "profile",
  "offline_access",
  "User.Read",
  "Mail.Send",
  "Mail.Read",
  "Calendars.ReadWrite",
  "Calendars.ReadWrite.Shared",
  "ChannelMessage.Send",
  "ChatMessage.Send",
  // "Sites.ReadWrite.All", // ⚠️ Requires admin consent
];

/**
 * Core authentication module for Azure/Microsoft Graph API
 * Handles token management, refresh, and storage
 */
export class AzureAuth {
  private static globalInstance: AzureAuth | null = null;

  private expiredAt?: number;
  private refreshToken: string = "";
  private accessToken: string = "";
  private tokenRefreshPromise: Promise<void> | null = null;
  private storageLoadPromise: Promise<void> | null = null;
  private tokenProvider?: (callback: string) => Promise<string> | string;
  private storagePath: string;
  private clientId: string = "";
  private clientSecret: string = "";
  private tenantId: string = "";
  private isAccessTokenOnly: boolean = false;
  private scopes: string[] = DEFAULT_SCOPES;
  private scopesConfigured: boolean = false;
  private allowInsecure: boolean = false;
  private isRetrying: boolean = false;

  /**
   * Get configured Axon instance with appropriate security settings
   * Uses Axon.dev() when allowInsecure is true, Axon.new() otherwise
   */
  getAxon() {
    return this.allowInsecure ? Axon.dev() : Axon.new();
  }

  /**
   * Wrapper for API requests with automatic 401 retry
   * @param operation The API operation to execute
   * @returns The result of the operation
   */
  async withRetry<T>(operation: () => Promise<T>): Promise<T> {
    try {
      return await operation();
    } catch (error: unknown) {
      // Only retry on 401 and only once
      if (error instanceof AxonError && error.status === 401 && !this.isRetrying) {
        this.isRetrying = true;
        try {
          console.warn('Received 401, attempting to refresh token and retry...');
          await this.invalidateAndRefresh();

          // Retry the operation once
          return await operation();
        } finally {
          this.isRetrying = false;
        }
      }

      // For all other errors or if already retrying, throw with better message
      throw this.enhanceError(error);
    }
  }

  /**
   * Invalidate current tokens and refresh them
   * Called when we receive a 401 from the API despite having a token
   */
  private async invalidateAndRefresh(): Promise<void> {
    console.info('Token invalidated by 401 response, refreshing...');

    // Clear current access token
    this.accessToken = '';
    this.expiredAt = 0;

    // Try to refresh using the refresh token
    if (this.refreshToken) {
      try {
        await this.refreshAccessToken();
        return;
      } catch {
        console.warn('Failed to refresh with refresh token, will try provider');
        // If refresh fails, fall through to use provider
      }
    }

    // If no refresh token or refresh failed, use provider
    if (this.tokenProvider) {
      await this.forgeRefreshToken();
      await this.saveToStorage();
    } else {
      throw new Error(
        'Authentication failed and no token provider configured. ' +
          'Cannot recover from 401 error.'
      );
    }
  }

  /**
   * Enhance error messages to be more user-friendly
   */
  private enhanceError(error: unknown): Error {
    // If it's an AxonError, extract status and provide better messages
    if (error instanceof AxonError) {
      const status = error.status;
      const data = error.responseData;

      if (status === 401) {
        return new Error(
          'Authentication failed. Please check your credentials or re-authenticate.'
        );
      } else if (status === 404) {
        return new Error(
          `Resource not found: ${data?.message || data?.error?.message || 'The requested item does not exist'}`
        );
      } else if (status === 409) {
        return new Error(`Conflict: ${data?.message || data?.error?.message || 'An item with this name already exists'}`);
      } else if (status === 403) {
        return new Error(
          `Permission denied: ${data?.message || data?.error?.message || 'You do not have access to this resource'}`
        );
      } else if (status && status >= 500) {
        return new Error(
          `Microsoft server error (${status}): ${data?.message || data?.error?.message || 'Please try again later'}`
        );
      }
    }

    // If it's already an Error, return it
    if (error instanceof Error) {
      return error;
    }

    // Otherwise wrap it in an Error
    return new Error(`Unknown error: ${String(error)}`);
  }

  constructor(config?: AzureConfig | AzureAuth) {
    // If passed an AzureAuth instance, copy from it
    if (config instanceof AzureAuth) {
      this.copyFrom(config);
      this.storagePath = this.getDefaultStoragePath();
      return;
    }

    this.storagePath = this.getDefaultStoragePath();

    // Priority order:
    // 1. Explicit config
    // 2. Global instance
    // 3. Environment variables
    // 4. Storage file
    this.loadCredentials(config);
  }

  /**
   * Set global configuration for all service instances
   * Similar to AWS.config.update()
   */
  static setGlobalConfig(config: AzureConfig): void {
    AzureAuth.globalInstance = new AzureAuth(config);
  }

  /**
   * Get the global auth instance
   * Auto-creates from env/storage if not set
   */
  static getGlobalInstance(): AzureAuth {
    if (!AzureAuth.globalInstance) {
      AzureAuth.globalInstance = new AzureAuth();
    }
    return AzureAuth.globalInstance;
  }

  /**
   * Reset global instance (useful for testing)
   */
  static reset(): void {
    AzureAuth.globalInstance = null;
  }

  /**
   * Copy credentials from another AzureAuth instance
   */
  private copyFrom(other: AzureAuth): void {
    this.accessToken = other.accessToken;
    this.refreshToken = other.refreshToken;
    this.expiredAt = other.expiredAt;
    this.clientId = other.clientId;
    this.clientSecret = other.clientSecret;
    this.tenantId = other.tenantId;
    this.isAccessTokenOnly = other.isAccessTokenOnly;
    this.scopes = other.scopes;
    this.scopesConfigured = other.scopesConfigured;
    this.tokenProvider = other.tokenProvider;
    this.allowInsecure = other.allowInsecure;
  }

  /**
   * Load credentials from config, global instance, env, or storage
   */
  private loadCredentials(config?: AzureConfig): void {
    if (config) {
      // Use explicit config
      this.applyConfig(config);
    } else if (AzureAuth.globalInstance) {
      // Use global config
      this.copyFrom(AzureAuth.globalInstance);
    }
    // Else: will load from env/storage on first use via ensureRefreshToken()
  }

  /**
   * Apply configuration
   */
  private applyConfig(config: AzureConfig): void {
    if (config.accessToken) {
      this.accessToken = config.accessToken;
      this.isAccessTokenOnly = true;
    }

    if (config.refreshToken) {
      this.refreshToken = config.refreshToken;
    }

    if (config.tokenProvider) {
      this.tokenProvider = config.tokenProvider;
    }

    if (config.clientId) {
      this.clientId = config.clientId;
    }

    if (config.clientSecret) {
      this.clientSecret = config.clientSecret;
    }

    if (config.tenantId) {
      this.tenantId = config.tenantId;
    }

    if (config.scopes && !this.scopesConfigured) {
      this.scopes = config.scopes;
      this.scopesConfigured = true;
    }

    if (config.allowInsecure !== undefined) {
      this.allowInsecure = config.allowInsecure;
    }

    this.updateStoragePath();
  }

  /**
   * Get access token (auto-refreshes if needed)
   */
  async getAccessToken(): Promise<string> {
    await this.checkToken();
    return this.accessToken;
  }

  /**
   * Get the storage directory based on platform
   */
  private getStorageDirectory(): string {
    const homeDir = os.homedir();

    if (process.platform === "win32") {
      const localAppData =
        process.env.LOCALAPPDATA || path.join(homeDir, "AppData", "Local");
      return path.join(localAppData, "ms-graph-devtools");
    } else {
      const configHome =
        process.env.XDG_CONFIG_HOME || path.join(homeDir, ".config");
      return path.join(configHome, "ms-graph-devtools");
    }
  }

  /**
   * Get storage path for tokens
   */
  private getDefaultStoragePath(): string {
    const baseDir = this.getStorageDirectory();

    if (this.tenantId && this.clientId) {
      const filename = `tokens.${this.tenantId}.${this.clientId}.json`;
      return path.join(baseDir, filename);
    }

    return path.join(baseDir, "tokens.json");
  }

  /**
   * Update storage path based on current tenant/client
   */
  private updateStoragePath(): void {
    this.storagePath = this.getDefaultStoragePath();
  }

  /**
   * Save credentials to storage
   */
  private async saveToStorage(): Promise<void> {
    if (this.isAccessTokenOnly) {
      return;
    }

    if (!this.refreshToken) {
      return;
    }

    try {
      const credentials: StoredCredentials = {
        refreshToken: this.refreshToken,
        accessToken: this.accessToken,
        expiresAt: this.expiredAt,
        clientId: this.clientId,
        tenantId: this.tenantId,
      };

      const dir = path.dirname(this.storagePath);
      await fs.mkdir(dir, { recursive: true, mode: 0o700 });

      await fs.writeFile(
        this.storagePath,
        JSON.stringify(credentials, null, 2),
        { mode: 0o600 }
      );

      console.info(`Credentials saved to ${this.storagePath}`);
    } catch (error) {
      console.error("Failed to save credentials to storage:", error);
    }
  }

  /**
   * Load credentials from storage
   */
  private async loadFromStorage(): Promise<boolean> {
    try {
      const data = await fs.readFile(this.storagePath, "utf-8");
      const credentials: StoredCredentials = JSON.parse(data);

      this.refreshToken = credentials.refreshToken;
      this.accessToken = credentials.accessToken;
      this.expiredAt = credentials.expiresAt;
      this.clientId = credentials.clientId;
      this.tenantId = credentials.tenantId;

      this.updateStoragePath();

      console.info(`Credentials loaded from storage: ${this.storagePath}`);
      return true;
    } catch {
      return false;
    }
  }

  /**
   * Ensure we have required credentials
   */
  private ensureCredentials(): void {
    this.updateStoragePath();

    if (!this.clientId || !this.clientSecret || !this.tenantId) {
      throw new Error(
        "Missing required credentials. Please provide:\n" +
          (!this.clientId ? "  - clientId\n" : "") +
          (!this.clientSecret ? "  - clientSecret\n" : "") +
          (!this.tenantId ? "  - tenantId\n" : "") +
          "\nProvide via:\n" +
          "1. new Service({ clientId: '...', clientSecret: '...', tenantId: '...' })\n" +
          "2. Azure.config({ clientId: '...', clientSecret: '...', tenantId: '...' })\n"
      );
    }
  }

  /**
   * Ensure we have a refresh token
   */
  private async ensureRefreshToken(): Promise<void> {
    // Wait for any in-flight storage load or token provider call
    if (this.storageLoadPromise) {
      await this.storageLoadPromise;
      if (this.refreshToken) {
        return;
      }
    }

    if (this.refreshToken) {
      return;
    }

    // Check again after waiting - another concurrent call might have started loading
    if (this.storageLoadPromise) {
      await this.storageLoadPromise;
      if (this.refreshToken) {
        return;
      }
    }

    // ALWAYS try loading from storage first (regardless of tokenProvider)
    this.storageLoadPromise = (async () => {
      await this.loadFromStorage();
    })();

    try {
      await this.storageLoadPromise;
      if (this.refreshToken) {
        return;
      }
    } finally {
      this.storageLoadPromise = null;
    }

    // Check one more time before starting provider - storage might have been loaded by concurrent call
    if (this.refreshToken) {
      return;
    }

    // Check if another call already started the provider
    if (this.storageLoadPromise) {
      await this.storageLoadPromise;
      if (this.refreshToken) {
        return;
      }
    }

    // Only use tokenProvider if storage didn't have tokens
    if (this.tokenProvider) {
      this.storageLoadPromise = (async () => {
        await this.forgeRefreshToken();
        await this.saveToStorage();
        console.info("Obtained tokens from token provider");
      })();
      try {
        await this.storageLoadPromise;
        return;
      } finally {
        this.storageLoadPromise = null;
      }
    }

    throw new Error(
      "No refresh token available. Please provide one via:\n" +
        "1. new Service({ refreshToken: 'your-token' })\n" +
        "2. Azure.config({ refreshToken: 'your-token' })\n" +
        "3. Saved storage file at: " +
        this.storagePath +
        "\n" +
        "4. tokenProvider function\n\n" +
        "See documentation for how to obtain a refresh token."
    );
  }

  /**
   * Check if token needs refresh
   */
  async checkToken(): Promise<void> {
    if (this.isAccessTokenOnly) {
      return;
    }

    this.ensureCredentials();
    await this.ensureRefreshToken();

    if (this.tokenRefreshPromise) {
      await this.tokenRefreshPromise;
      if (this.refreshToken && this.expiredAt && Date.now() < this.expiredAt) {
        return;
      }
    }

    if (this.expiredAt && Date.now() >= this.expiredAt) {
      this.tokenRefreshPromise = this.refreshAccessToken();

      try {
        await this.tokenRefreshPromise;
        await this.saveToStorage();
      } finally {
        this.tokenRefreshPromise = null;
      }
    }
  }

  /**
   * Forge new refresh token via OAuth authorization code flow
   * Uses tokenProvider to get authorization code, then exchanges for tokens
   */
  private async forgeRefreshToken(): Promise<void> {
    if (!this.tokenProvider) {
      throw new Error(
        "No token provider configured. Please provide one via:\n" +
          "1. new Service({ tokenProvider: async (callback) => { ... } })\n" +
          "2. Provide refreshToken directly: new Service({ refreshToken: '...' })\n" +
          "\nExample with Playwright:\n" +
          "  new Outlook({\n" +
          "    tokenProvider: async (callback) => await Playwright.getAzureCode(callback)\n" +
          "  })\n"
      );
    }

    const callback =
      `https://login.microsoftonline.com/${this.tenantId}/oauth2/v2.0/authorize?` +
      [
        `client_id=${this.clientId}`,
        "response_type=code",
        `redirect_uri=${REDIRECT_URI}`,
        `scope=${this.scopes.join("%20")}`,
        "response_mode=query",
      ].join("&");

    const code = await this.tokenProvider(callback);

    const url = `https://login.microsoftonline.com/${this.tenantId}/oauth2/v2.0/token`;

    const reqTokenBody = {
      client_id: this.clientId,
      client_secret: this.clientSecret,
      code: code,
      redirect_uri: REDIRECT_URI,
      grant_type: "authorization_code",
      scope: this.scopes.join(" "),
    };

    try {
      const res = await this.getAxon()
        .encodeUrl()
        .post(url, reqTokenBody);

      if (res.status === 200) {
        this.accessToken = res.data.access_token;
        this.expiredAt = Date.now() + res.data.expires_in * 1000;
        this.refreshToken = res.data.refresh_token;
      } else {
        console.error(
          `Failed to forge refresh token: ${res.status} ${JSON.stringify(
            res.data
          )}`
        );
        throw new Error("Failed to forge refresh token");
      }
    } catch (error) {
      console.error("Error forging refresh token:", error);
      throw error;
    }
  }

  /**
   * Refresh access token
   */
  private async refreshAccessToken(): Promise<void> {
    const url = `https://login.microsoftonline.com/${this.tenantId}/oauth2/v2.0/token`;

    const reqTokenBody = {
      client_id: this.clientId,
      scope: this.scopes.join(" "),
      refresh_token: this.refreshToken,
      redirect_uri: REDIRECT_URI,
      grant_type: "refresh_token",
      client_secret: this.clientSecret,
    };

    try {
      const res = await this.getAxon()
        .encodeUrl()
        .post(url, reqTokenBody);

      if (res.status === 200) {
        this.accessToken = res.data.access_token;

        if (res.data.expires_in) {
          this.expiredAt = Date.now() + res.data.expires_in * 1000;
        }

        if (res.data.refresh_token) {
          this.refreshToken = res.data.refresh_token;
        }
      } else {
        console.error(
          `Failed to refresh access token: ${res.status} ${JSON.stringify(
            res.data
          )}`
        );
        throw new Error("Failed to refresh access token");
      }
    } catch (error) {
      console.error("Error refreshing access token:", error);
      throw new Error("Failed to refresh access token");
    }
  }

  /**
   * Handle API errors (especially 401 for light user mode)
   */
  handleApiError(error: AxonError): never {
    if (this.isAccessTokenOnly && error.status === 401) {
      throw new Error(
        "Access token is invalid or expired.\n\n" +
          "To continue:\n" +
          "  1. Provide a new access token: new Service({ accessToken: 'new-token' })\n" +
          "  2. For automatic renewal, see documentation on using refresh tokens\n"
      );
    }
    throw error;
  }

  /**
   * List all stored credentials
   */
  static async listStoredCredentials(): Promise<
    Array<{ tenantId?: string; clientId?: string; file: string }>
  > {
    const instance = new AzureAuth();
    const baseDir = instance.getStorageDirectory();

    try {
      const files = await fs.readdir(baseDir);
      return files
        .filter((f) => f.startsWith("tokens") && f.endsWith(".json"))
        .map((file) => {
          const parts = file.split(".");
          if (parts.length === 4 && parts[0] === "tokens") {
            return {
              tenantId: parts[1],
              clientId: parts[2],
              file: file,
            };
          }
          return { file: file };
        });
    } catch {
      return [];
    }
  }

  /**
   * Clear stored credentials
   */
  static async clearStoredCredentials(
    tenantId?: string,
    clientId?: string
  ): Promise<void> {
    const instance = new AzureAuth();
    const baseDir = instance.getStorageDirectory();

    try {
      if (tenantId && clientId) {
        const filename = `tokens.${tenantId}.${clientId}.json`;
        const filePath = path.join(baseDir, filename);
        await fs.unlink(filePath);
        console.info(
          `Cleared credentials for tenant=${tenantId}, client=${clientId}`
        );
      } else {
        const files = await fs.readdir(baseDir);
        const tokenFiles = files.filter(
          (f) => f.startsWith("tokens") && f.endsWith(".json")
        );

        await Promise.all(
          tokenFiles.map((file) => fs.unlink(path.join(baseDir, file)))
        );
        console.info(
          `Cleared all stored credentials (${tokenFiles.length} files)`
        );
      }
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code !== "ENOENT") {
        console.error("Failed to clear credentials:", error);
      }
    }
  }
}
