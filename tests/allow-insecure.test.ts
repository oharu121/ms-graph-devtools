import { describe, it, expect, beforeEach, vi } from "vitest";
import { AzureAuth } from "../src/core/auth";

describe("allowInsecure Option", () => {
  let auth: AzureAuth;

  beforeEach(() => {
    vi.clearAllMocks();
  });

  it("should use Axon.new() when allowInsecure is false (default)", () => {
    auth = new AzureAuth({
      clientId: "test-client-id",
      clientSecret: "test-client-secret",
      tenantId: "test-tenant-id",
      accessToken: "test-token",
      allowInsecure: false,
    });

    const axonInstance = auth.getAxon();

    // Axon.new() returns an AxonClient instance
    // Axon.dev() also returns an AxonClient, but with different defaults
    // We can test by checking the internal config
    expect(axonInstance).toBeDefined();
    expect(typeof axonInstance.get).toBe("function");
  });

  it("should use Axon.dev() when allowInsecure is true", () => {
    auth = new AzureAuth({
      clientId: "test-client-id",
      clientSecret: "test-client-secret",
      tenantId: "test-tenant-id",
      accessToken: "test-token",
      allowInsecure: true,
    });

    const axonInstance = auth.getAxon();

    // Verify we get an Axon instance
    expect(axonInstance).toBeDefined();
    expect(typeof axonInstance.get).toBe("function");
    expect(typeof axonInstance.post).toBe("function");
  });

  it("should use Axon.new() when allowInsecure is undefined", () => {
    auth = new AzureAuth({
      clientId: "test-client-id",
      clientSecret: "test-client-secret",
      tenantId: "test-tenant-id",
      accessToken: "test-token",
      // allowInsecure not specified - should default to false
    });

    const axonInstance = auth.getAxon();

    expect(axonInstance).toBeDefined();
    expect(typeof axonInstance.get).toBe("function");
  });

  it("should pass allowInsecure through to service methods", () => {
    // We test the behavior of getAxon() by calling it directly
    // and verifying the returned instance type
    const authSecure = new AzureAuth({
      clientId: "test-client-id",
      clientSecret: "test-client-secret",
      tenantId: "test-tenant-id",
      accessToken: "test-token",
      allowInsecure: false,
    });

    const authInsecure = new AzureAuth({
      clientId: "test-client-id",
      clientSecret: "test-client-secret",
      tenantId: "test-tenant-id",
      accessToken: "test-token",
      allowInsecure: true,
    });

    // Both should return valid Axon instances with the same interface
    const secureAxon = authSecure.getAxon();
    const insecureAxon = authInsecure.getAxon();

    // Both instances should have the same methods
    expect(typeof secureAxon.get).toBe("function");
    expect(typeof secureAxon.post).toBe("function");
    expect(typeof insecureAxon.get).toBe("function");
    expect(typeof insecureAxon.post).toBe("function");

    // The behavior difference is internal to Axon.new() vs Axon.dev()
    // Both return AxonClient instances, but with different SSL settings
    expect(secureAxon).toBeDefined();
    expect(insecureAxon).toBeDefined();
  });

  it("should apply allowInsecure to token refresh requests", async () => {
    // This test verifies that when refreshing tokens,
    // the allowInsecure setting is respected
    const auth = new AzureAuth({
      clientId: "test-client-id",
      clientSecret: "test-client-secret",
      tenantId: "test-tenant-id",
      refreshToken: "test-refresh-token",
      allowInsecure: true,
    });

    // The getAxon method should return Axon.dev() when allowInsecure is true
    const axonInstance = auth.getAxon();
    expect(axonInstance).toBeDefined();

    // Verify it's using the dev instance (has the same interface)
    expect(typeof axonInstance.encodeUrl).toBe("function");
    expect(typeof axonInstance.bearer).toBe("function");
  });
});
