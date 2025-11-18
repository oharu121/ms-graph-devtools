import { describe, it, expect, beforeEach, afterEach, vi } from "vitest";
import { AzureAuth } from "../src/core/auth";
import fs from "fs/promises";
import path from "path";
import os from "os";
import type { AzureConfig } from "../src/types";

describe("AzureAuth - Comprehensive Tests", () => {
  let testStoragePath: string;
  let originalEnv: NodeJS.ProcessEnv;

  beforeEach(async () => {
    originalEnv = { ...process.env };
    const baseDir = getStorageDirectory();
    testStoragePath = path.join(
      baseDir,
      "tokens.test-tenant.test-client.json"
    );

    // Clean up any existing test files BEFORE each test
    try {
      await fs.unlink(testStoragePath);
    } catch {
      // Ignore
    }
  });

  afterEach(async () => {
    process.env = originalEnv;
    try {
      await fs.unlink(testStoragePath);
    } catch {
      // Ignore
    }
    AzureAuth.reset();
  });

  function getStorageDirectory(): string {
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

  describe("Constructor and Configuration", () => {
    it("should create instance with config object", () => {
      const config: AzureConfig = {
        clientId: "test-client",
        clientSecret: "test-secret",
        tenantId: "test-tenant",
        refreshToken: "test-token",
      };

      const auth = new AzureAuth(config);
      expect((auth as any).clientId).toBe("test-client");
      expect((auth as any).refreshToken).toBe("test-token");
    });

    it("should create instance with access token only", () => {
      const config: AzureConfig = {
        accessToken: "test-access-token",
      };

      const auth = new AzureAuth(config);
      expect((auth as any).accessToken).toBe("test-access-token");
      expect((auth as any).isAccessTokenOnly).toBe(true);
    });

    it("should copy from another AzureAuth instance", () => {
      const auth1 = new AzureAuth({
        clientId: "test-client",
        clientSecret: "test-secret",
        tenantId: "test-tenant",
        refreshToken: "test-token",
      });

      const auth2 = new AzureAuth(auth1);
      expect((auth2 as any).clientId).toBe("test-client");
      expect((auth2 as any).refreshToken).toBe("test-token");
    });

    it("should apply custom scopes", () => {
      const config: AzureConfig = {
        clientId: "test-client",
        clientSecret: "test-secret",
        tenantId: "test-tenant",
        scopes: ["custom.scope", "another.scope"],
      };

      const auth = new AzureAuth(config);
      expect((auth as any).scopes).toEqual(["custom.scope", "another.scope"]);
    });

    it("should set allowInsecure flag", () => {
      const auth = new AzureAuth({
        allowInsecure: true,
      });

      expect((auth as any).allowInsecure).toBe(true);
    });
  });

  describe("Global Instance Management", () => {
    it("should set and get global instance", () => {
      const config: AzureConfig = {
        clientId: "global-client",
        clientSecret: "global-secret",
        tenantId: "global-tenant",
        refreshToken: "global-token",
      };

      AzureAuth.setGlobalConfig(config);
      const instance = AzureAuth.getGlobalInstance();

      expect((instance as any).clientId).toBe("global-client");
    });

    it("should auto-create global instance if not set", () => {
      AzureAuth.reset();
      const instance = AzureAuth.getGlobalInstance();
      expect(instance).toBeDefined();
    });

    it("should reset global instance", () => {
      AzureAuth.setGlobalConfig({ refreshToken: "test" });
      AzureAuth.reset();
      expect((AzureAuth as any).globalInstance).toBeNull();
    });

    it("should use global instance when creating new auth without config", () => {
      AzureAuth.setGlobalConfig({
        clientId: "global-client",
        refreshToken: "global-token",
      });

      const auth = new AzureAuth();
      expect((auth as any).refreshToken).toBe("global-token");
    });
  });

  describe("Storage Management", () => {
    it("should generate correct storage path with tenant and client", () => {
      const auth = new AzureAuth({
        clientId: "my-client",
        tenantId: "my-tenant",
      });

      const storagePath = (auth as any).storagePath;
      expect(storagePath).toContain("tokens.my-tenant.my-client.json");
    });

    it("should use default storage path without tenant/client", () => {
      const auth = new AzureAuth({});
      const storagePath = (auth as any).storagePath;
      expect(storagePath).toContain("tokens.json");
    });

    it("should save credentials to storage", async () => {
      // Use unique IDs to avoid test interference
      const uniqueClient = `save-client-${Date.now()}`;
      const uniqueTenant = `save-tenant-${Date.now()}`;

      const config = {
        clientId: uniqueClient,
        clientSecret: "test-secret",
        tenantId: uniqueTenant,
        refreshToken: "test-token",
      };

      const auth = new AzureAuth(config);

      // Ensure storage path is set correctly
      (auth as any).updateStoragePath();
      const actualStoragePath = (auth as any).storagePath;

      await (auth as any).saveToStorage();

      const data = await fs.readFile(actualStoragePath, "utf-8");
      const credentials = JSON.parse(data);

      expect(credentials.refreshToken).toBe("test-token");
      expect(credentials.clientId).toBe(uniqueClient);

      // Cleanup
      await fs.unlink(actualStoragePath);
    });

    it("should load credentials from storage", async () => {
      const credentials = {
        refreshToken: "stored-token",
        accessToken: "stored-access",
        expiresAt: Date.now() + 3600000,
        clientId: "test-client",
        tenantId: "test-tenant",
      };

      const dir = path.dirname(testStoragePath);
      await fs.mkdir(dir, { recursive: true, mode: 0o700 });
      await fs.writeFile(
        testStoragePath,
        JSON.stringify(credentials, null, 2),
        { mode: 0o600 }
      );

      const auth = new AzureAuth();
      (auth as any).storagePath = testStoragePath;

      const loaded = await (auth as any).loadFromStorage();
      expect(loaded).toBe(true);
      expect((auth as any).refreshToken).toBe("stored-token");
    });

    it("should return false when loading from non-existent storage", async () => {
      const auth = new AzureAuth();
      (auth as any).storagePath = "/non/existent/path.json";

      const loaded = await (auth as any).loadFromStorage();
      expect(loaded).toBe(false);
    });

    it("should not save when using access token only", async () => {
      const auth = new AzureAuth({ accessToken: "test" });
      const saveSpy = vi.spyOn(fs, "writeFile");

      await (auth as any).saveToStorage();
      expect(saveSpy).not.toHaveBeenCalled();

      saveSpy.mockRestore();
    });

    it("should not save when using token provider", async () => {
      const auth = new AzureAuth({
        tokenProvider: async () => "test",
      });
      const saveSpy = vi.spyOn(fs, "writeFile");

      await (auth as any).saveToStorage();
      expect(saveSpy).not.toHaveBeenCalled();

      saveSpy.mockRestore();
    });
  });

  describe("Credential Loading Priority", () => {
    it("should use explicit config", () => {
      const auth = new AzureAuth({
        clientId: "config-client",
        clientSecret: "config-secret",
        tenantId: "config-tenant",
      });

      expect((auth as any).clientId).toBe("config-client");
      expect((auth as any).clientSecret).toBe("config-secret");
      expect((auth as any).tenantId).toBe("config-tenant");
    });

    it("should use global config when no local config provided", () => {
      AzureAuth.setGlobalConfig({
        clientId: "global-client",
        clientSecret: "global-secret",
        tenantId: "global-tenant",
      });

      const auth = new AzureAuth();

      expect((auth as any).clientId).toBe("global-client");
      expect((auth as any).clientSecret).toBe("global-secret");
      expect((auth as any).tenantId).toBe("global-tenant");

      AzureAuth.reset();
    });

    it("should prefer explicit config over global config", () => {
      AzureAuth.setGlobalConfig({
        clientId: "global-client",
      });

      const auth = new AzureAuth({
        clientId: "local-client",
      });

      expect((auth as any).clientId).toBe("local-client");

      AzureAuth.reset();
    });
  });

  describe("Token Provider", () => {
    it("should call token provider when no refresh token", async () => {
      const mockProvider = vi.fn(async () => "provider-token");

      const auth = new AzureAuth({
        clientId: "test-client",
        clientSecret: "test-secret",
        tenantId: "test-tenant",
        tokenProvider: mockProvider,
      });

      await (auth as any).ensureRefreshToken();

      expect(mockProvider).toHaveBeenCalled();
      expect((auth as any).refreshToken).toBe("provider-token");
    });

    it("should handle synchronous token provider", async () => {
      const mockProvider = vi.fn(() => "sync-token");

      const auth = new AzureAuth({
        clientId: "test-client",
        clientSecret: "test-secret",
        tenantId: "test-tenant",
        tokenProvider: mockProvider as any,
      });

      await (auth as any).ensureRefreshToken();

      expect((auth as any).refreshToken).toBe("sync-token");
    });
  });

  describe("Access Token Management", () => {
    it("should return access token directly when in access-token-only mode", async () => {
      const auth = new AzureAuth({
        accessToken: "my-access-token",
      });

      const token = await auth.getAccessToken();
      expect(token).toBe("my-access-token");
    });

    it("should throw error when no credentials available", async () => {
      const auth = new AzureAuth();

      await expect(auth.getAccessToken()).rejects.toThrow(
        "Missing required credentials"
      );
    });

    it("should throw error when no refresh token available", async () => {
      const auth = new AzureAuth({
        clientId: "test-client",
        clientSecret: "test-secret",
        tenantId: "test-tenant",
      });

      await expect((auth as any).ensureRefreshToken()).rejects.toThrow(
        "No refresh token available"
      );
    });
  });

  describe("Error Handling", () => {
    it("should handle API error in access-token-only mode", () => {
      const auth = new AzureAuth({ accessToken: "test" });

      const error = {
        response: { status: 401 },
      };

      expect(() => auth.handleApiError(error)).toThrow(
        "Access token is invalid or expired"
      );
    });

    it("should rethrow non-401 errors", () => {
      const auth = new AzureAuth({ accessToken: "test" });

      const error = new Error("Network error");

      expect(() => auth.handleApiError(error)).toThrow("Network error");
    });

    it("should handle storage save errors gracefully", async () => {
      const auth = new AzureAuth({
        clientId: "test-client",
        clientSecret: "test-secret",
        tenantId: "test-tenant",
        refreshToken: "test-token",
      });

      // Make storage path invalid
      (auth as any).storagePath = "/invalid/\0/path.json";

      // Should not throw
      await expect(
        (auth as any).saveToStorage()
      ).resolves.not.toThrow();
    });
  });

  describe("Multi-tenant Support", () => {
    it("should list stored credentials", async () => {
      // Create test credential files
      const baseDir = getStorageDirectory();
      await fs.mkdir(baseDir, { recursive: true, mode: 0o700 });

      const file1 = path.join(baseDir, "tokens.tenant1.client1.json");
      const file2 = path.join(baseDir, "tokens.tenant2.client2.json");

      await fs.writeFile(file1, JSON.stringify({}), { mode: 0o600 });
      await fs.writeFile(file2, JSON.stringify({}), { mode: 0o600 });

      const stored = await AzureAuth.listStoredCredentials();

      expect(stored.length).toBeGreaterThanOrEqual(2);
      expect(stored.some((c) => c.tenantId === "tenant1")).toBe(true);
      expect(stored.some((c) => c.tenantId === "tenant2")).toBe(true);

      // Cleanup
      await fs.unlink(file1);
      await fs.unlink(file2);
    });

    it("should clear specific tenant credentials", async () => {
      const config = {
        clientId: "clear-client",
        tenantId: "clear-tenant",
        clientSecret: "secret",
        refreshToken: "token",
      };

      const auth = new AzureAuth(config);
      await (auth as any).saveToStorage();

      await AzureAuth.clearStoredCredentials("clear-tenant", "clear-client");

      const baseDir = getStorageDirectory();
      const filePath = path.join(
        baseDir,
        "tokens.clear-tenant.clear-client.json"
      );

      await expect(fs.access(filePath)).rejects.toThrow();
    });

    it("should handle clear credentials for non-existent files", async () => {
      await expect(
        AzureAuth.clearStoredCredentials("non-existent", "tenant")
      ).resolves.not.toThrow();
    });

    it("should return empty array when storage directory doesn't exist", async () => {
      // Create temp auth to get directory
      const auth = new AzureAuth();
      const baseDir = (auth as any).getStorageDirectory();

      // Temporarily rename directory if it exists
      const tempDir = baseDir + ".tmp";
      try {
        await fs.rename(baseDir, tempDir);
      } catch {
        // Directory doesn't exist, which is what we want
      }

      const stored = await AzureAuth.listStoredCredentials();
      expect(stored).toEqual([]);

      // Restore directory
      try {
        await fs.rename(tempDir, baseDir);
      } catch {
        // Ignore
      }
    });
  });

  describe("Storage Path Updates", () => {
    it("should update storage path when tenant/client change", () => {
      const auth = new AzureAuth();
      const path1 = (auth as any).storagePath;

      (auth as any).clientId = "new-client";
      (auth as any).tenantId = "new-tenant";
      (auth as any).updateStoragePath();

      const path2 = (auth as any).storagePath;

      expect(path1).not.toBe(path2);
      expect(path2).toContain("new-tenant");
      expect(path2).toContain("new-client");
    });
  });

  describe("Configuration Edge Cases", () => {
    it("should handle empty config object", () => {
      const auth = new AzureAuth({});
      expect(auth).toBeDefined();
    });

    it("should handle partial config", () => {
      const auth = new AzureAuth({
        clientId: "only-client-id",
      });

      expect((auth as any).clientId).toBe("only-client-id");
    });

    it("should not override scopes once configured", () => {
      const auth = new AzureAuth({
        scopes: ["initial.scope"],
      });

      (auth as any).applyConfig({ scopes: ["new.scope"] });

      expect((auth as any).scopes).toEqual(["initial.scope"]);
    });
  });
});
