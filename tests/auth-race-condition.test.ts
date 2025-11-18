import { describe, it, expect, beforeEach, afterEach, vi } from "vitest";
import { AzureAuth } from "../src/core/auth";
import fs from "fs/promises";
import path from "path";
import os from "os";

describe("AzureAuth - Token Storage Race Condition Fix", () => {
  let testStoragePath: string;
  let originalEnv: NodeJS.ProcessEnv;

  beforeEach(async () => {
    // Save original env
    originalEnv = { ...process.env };

    // Set up test storage path
    const baseDir = getStorageDirectory();
    testStoragePath = path.join(
      baseDir,
      "tokens.test-tenant.test-client.json"
    );

    // Clean up any existing test files BEFORE each test
    try {
      await fs.unlink(testStoragePath);
    } catch {
      // Ignore if file doesn't exist
    }
  });

  afterEach(async () => {
    // Restore env
    process.env = originalEnv;

    // Clean up test files
    try {
      await fs.unlink(testStoragePath);
    } catch {
      // Ignore if file doesn't exist
    }

    // Reset global instance
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

  describe("Concurrent Storage Load Protection", () => {
    it("should call loadFromStorage only once despite concurrent requests", async () => {
      // Use unique IDs to avoid test interference
      const uniqueClient = `test-client-${Date.now()}`;
      const uniqueTenant = `test-tenant-${Date.now()}`;

      // Create test credentials in storage
      const testCredentials = {
        refreshToken: "test-refresh-token",
        accessToken: "test-access-token",
        expiresAt: Date.now() + 3600000,
        clientId: uniqueClient,
        tenantId: uniqueTenant,
      };

      const baseDir = getStorageDirectory();
      const uniqueStoragePath = path.join(
        baseDir,
        `tokens.${uniqueTenant}.${uniqueClient}.json`
      );

      const dir = path.dirname(uniqueStoragePath);
      await fs.mkdir(dir, { recursive: true, mode: 0o700 });
      await fs.writeFile(
        uniqueStoragePath,
        JSON.stringify(testCredentials, null, 2),
        { mode: 0o600 }
      );

      // Create auth with explicit credentials (no env vars)
      const auth = new AzureAuth({
        clientId: uniqueClient,
        clientSecret: "test-secret",
        tenantId: uniqueTenant,
      });

      // Clear the refresh token to force loading from storage
      (auth as any).refreshToken = "";

      // Spy on loadFromStorage to count calls
      const loadSpy = vi.spyOn(auth as any, "loadFromStorage");

      // Make 5 concurrent calls
      const promises = Array(5)
        .fill(null)
        .map(() => (auth as any).ensureRefreshToken());

      await Promise.all(promises);

      // Verify loadFromStorage was called only once
      expect(loadSpy).toHaveBeenCalledTimes(1);
      expect((auth as any).refreshToken).toBe("test-refresh-token");

      loadSpy.mockRestore();

      // Cleanup
      await fs.unlink(uniqueStoragePath);
    });

    it("should wait for in-flight storage load before making additional calls", async () => {
      // Use unique IDs to avoid test interference
      const uniqueClient = `test-client-wait-${Date.now()}`;
      const uniqueTenant = `test-tenant-wait-${Date.now()}`;

      // Create test credentials
      const testCredentials = {
        refreshToken: "test-refresh-token-2",
        accessToken: "test-access-token-2",
        expiresAt: Date.now() + 3600000,
        clientId: uniqueClient,
        tenantId: uniqueTenant,
      };

      const baseDir = getStorageDirectory();
      const uniqueStoragePath = path.join(
        baseDir,
        `tokens.${uniqueTenant}.${uniqueClient}.json`
      );

      const dir = path.dirname(uniqueStoragePath);
      await fs.mkdir(dir, { recursive: true, mode: 0o700 });
      await fs.writeFile(
        uniqueStoragePath,
        JSON.stringify(testCredentials, null, 2),
        { mode: 0o600 }
      );

      const auth = new AzureAuth({
        clientId: uniqueClient,
        clientSecret: "test-secret",
        tenantId: uniqueTenant,
      });

      // Clear the refresh token to force loading from storage
      (auth as any).refreshToken = "";

      // Track call order with actual timestamps
      const events: Array<{ type: string; time: number }> = [];
      const startTime = Date.now();

      // Spy on loadFromStorage to add delay
      const originalLoad = (auth as any).loadFromStorage.bind(auth);
      const loadSpy = vi.spyOn(auth as any, "loadFromStorage").mockImplementation(
        async function (this: any) {
          events.push({ type: "load-start", time: Date.now() - startTime });
          await new Promise((resolve) => setTimeout(resolve, 100));
          const result = await originalLoad.call(this);
          events.push({ type: "load-end", time: Date.now() - startTime });
          return result;
        }
      );

      // Start first call
      const promise1 = (auth as any).ensureRefreshToken();

      // Start second call shortly after (while first is still loading)
      await new Promise((resolve) => setTimeout(resolve, 20));
      const promise2 = (auth as any).ensureRefreshToken();

      await Promise.all([promise1, promise2]);

      // Verify storage was only loaded once (second call waited)
      expect(loadSpy).toHaveBeenCalledTimes(1);
      expect((auth as any).refreshToken).toBe("test-refresh-token-2");

      // Verify events show only one load cycle
      const loadStarts = events.filter((e) => e.type === "load-start");
      const loadEnds = events.filter((e) => e.type === "load-end");
      expect(loadStarts.length).toBe(1);
      expect(loadEnds.length).toBe(1);

      loadSpy.mockRestore();

      // Cleanup
      await fs.unlink(uniqueStoragePath);
    });
  });

  describe("Concurrent Token Provider Protection", () => {
    it("should call tokenProvider only once despite concurrent requests", async () => {
      let providerCallCount = 0;

      const mockProvider = vi.fn(async () => {
        providerCallCount++;
        await new Promise((resolve) => setTimeout(resolve, 50));
        return "provider-refresh-token";
      });

      const auth = new AzureAuth({
        clientId: "test-client",
        clientSecret: "test-secret",
        tenantId: "test-tenant",
        tokenProvider: mockProvider,
      });

      // Make 5 concurrent calls
      const promises = Array(5)
        .fill(null)
        .map(() => (auth as any).ensureRefreshToken());

      await Promise.all(promises);

      // Provider should be called only once
      expect(providerCallCount).toBe(1);
      expect(mockProvider).toHaveBeenCalledTimes(1);
      expect((auth as any).refreshToken).toBe("provider-refresh-token");
    });

    it("should skip storage load when tokenProvider is configured", async () => {
      const mockProvider = vi.fn(async () => "provider-token");

      const auth = new AzureAuth({
        clientId: "test-client",
        clientSecret: "test-secret",
        tenantId: "test-tenant",
        tokenProvider: mockProvider,
      });

      const loadSpy = vi.spyOn(auth as any, "loadFromStorage");

      await (auth as any).ensureRefreshToken();

      // Storage should not be accessed when tokenProvider is configured
      expect(loadSpy).not.toHaveBeenCalled();
      expect(mockProvider).toHaveBeenCalledTimes(1);

      loadSpy.mockRestore();
    });
  });

  describe("Mixed Concurrent Scenarios", () => {
    it("should handle concurrent calls with early return when token already exists", async () => {
      const auth = new AzureAuth({
        clientId: "test-client",
        clientSecret: "test-secret",
        tenantId: "test-tenant",
        refreshToken: "existing-token",
      });

      const loadSpy = vi.spyOn(auth as any, "loadFromStorage");

      // Make 10 concurrent calls
      const promises = Array(10)
        .fill(null)
        .map(() => (auth as any).ensureRefreshToken());

      await Promise.all(promises);

      // Neither storage nor provider should be called (token already exists)
      expect(loadSpy).not.toHaveBeenCalled();

      loadSpy.mockRestore();
    });

    it("should handle race between storage and token provider", async () => {
      const mockProvider = vi.fn(async () => "provider-token");

      const auth = new AzureAuth({
        clientId: "test-client",
        clientSecret: "test-secret",
        tenantId: "test-tenant",
        tokenProvider: mockProvider,
      });

      const loadSpy = vi.spyOn(auth as any, "loadFromStorage");

      // Make concurrent calls
      const promises = Array(5)
        .fill(null)
        .map(() => (auth as any).ensureRefreshToken());

      await Promise.all(promises);

      // Provider should be used (storage skipped when provider exists)
      expect(loadSpy).not.toHaveBeenCalled();
      expect(mockProvider).toHaveBeenCalledTimes(1);
      expect((auth as any).refreshToken).toBe("provider-token");

      loadSpy.mockRestore();
    });
  });

  describe("Promise Tracking Mechanism", () => {
    it("should set and clear storageLoadPromise correctly", async () => {
      const auth = new AzureAuth({
        clientId: "test-client",
        clientSecret: "test-secret",
        tenantId: "test-tenant",
        tokenProvider: async () => "test-token",
      });

      // Before call
      expect((auth as any).storageLoadPromise).toBeNull();

      // Start async call (don't await yet)
      const promise = (auth as any).ensureRefreshToken();

      // During call, promise should be set
      expect((auth as any).storageLoadPromise).not.toBeNull();

      // After call completes
      await promise;
      expect((auth as any).storageLoadPromise).toBeNull();
    });

    it("should reuse existing storageLoadPromise for concurrent calls", async () => {
      const mockProvider = vi.fn(async () => {
        await new Promise((resolve) => setTimeout(resolve, 100));
        return "test-token";
      });

      const auth = new AzureAuth({
        clientId: "test-client",
        clientSecret: "test-secret",
        tenantId: "test-tenant",
        tokenProvider: mockProvider,
      });

      // Start first call
      const promise1 = (auth as any).ensureRefreshToken();
      const firstPromise = (auth as any).storageLoadPromise;

      // Start second call while first is in flight
      const promise2 = (auth as any).ensureRefreshToken();
      const secondPromise = (auth as any).storageLoadPromise;

      // Both should reference the same promise instance
      expect(firstPromise).toBe(secondPromise);

      await Promise.all([promise1, promise2]);

      // Provider should only be called once
      expect(mockProvider).toHaveBeenCalledTimes(1);
    });
  });

  describe("Error Handling", () => {
    it("should clear storageLoadPromise even if loading fails", async () => {
      const mockProvider = vi.fn(async () => {
        throw new Error("Provider failed");
      });

      const auth = new AzureAuth({
        clientId: "test-client",
        clientSecret: "test-secret",
        tenantId: "test-tenant",
        tokenProvider: mockProvider,
      });

      await expect((auth as any).ensureRefreshToken()).rejects.toThrow(
        "Provider failed"
      );

      // Promise should be cleared even after failure
      expect((auth as any).storageLoadPromise).toBeNull();
    });

    it("should allow retry after failed provider call", async () => {
      let callCount = 0;
      const mockProvider = vi.fn(async () => {
        callCount++;
        if (callCount === 1) {
          throw new Error("First call failed");
        }
        return "success-token";
      });

      const auth = new AzureAuth({
        clientId: "test-client",
        clientSecret: "test-secret",
        tenantId: "test-tenant",
        tokenProvider: mockProvider,
      });

      // First call should fail
      await expect((auth as any).ensureRefreshToken()).rejects.toThrow(
        "First call failed"
      );

      // Second call should succeed
      await (auth as any).ensureRefreshToken();
      expect((auth as any).refreshToken).toBe("success-token");
      expect(mockProvider).toHaveBeenCalledTimes(2);
    });
  });
});
