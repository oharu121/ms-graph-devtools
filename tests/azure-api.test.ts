import { describe, it, expect, beforeEach, afterEach } from "vitest";
import Azure from "../src/index";
import { Outlook } from "../src/services/Outlook";
import { Teams } from "../src/services/Teams";
import { Calendar } from "../src/services/Calendar";
import { SharePoint } from "../src/services/SharePoint";

describe("Azure Global Instance API", () => {
  beforeEach(() => {
    Azure.reset();
  });

  afterEach(() => {
    Azure.reset();
  });

  describe("Azure.config()", () => {
    it("should accept configuration", () => {
      expect(() => {
        Azure.config({
          accessToken: "fake-token-for-testing",
        });
      }).not.toThrow();
    });

    it("should accept all config options", () => {
      expect(() => {
        Azure.config({
          clientId: "test-client",
          clientSecret: "test-secret",
          tenantId: "test-tenant",
          refreshToken: "test-refresh-token",
        });
      }).not.toThrow();
    });

    it("should accept tokenProvider", () => {
      expect(() => {
        Azure.config({
          clientId: "test-client",
          clientSecret: "test-secret",
          tenantId: "test-tenant",
          tokenProvider: async () => "test-code",
        });
      }).not.toThrow();
    });

    it("should reset service instances when called", () => {
      Azure.config({ accessToken: "token1" });
      const outlook1 = Azure.outlook;

      Azure.config({ accessToken: "token2" });
      const outlook2 = Azure.outlook;

      // Should be different instances after config() is called again
      expect(outlook1).not.toBe(outlook2);
    });
  });

  describe("Azure.reset()", () => {
    it("should clear all service instances", () => {
      Azure.config({ accessToken: "test-token" });
      const outlook1 = Azure.outlook;

      Azure.reset();
      const outlook2 = Azure.outlook;

      // Should be different instances after reset
      expect(outlook1).not.toBe(outlook2);
    });
  });

  describe("Service Getters", () => {
    beforeEach(() => {
      Azure.config({
        accessToken: "test-token-for-getters",
      });
    });

    it("should return Outlook instance", () => {
      const outlook = Azure.outlook;
      expect(outlook).toBeInstanceOf(Outlook);
    });

    it("should return Teams instance", () => {
      const teams = Azure.teams;
      expect(teams).toBeInstanceOf(Teams);
    });

    it("should return Calendar instance", () => {
      const calendar = Azure.calendar;
      expect(calendar).toBeInstanceOf(Calendar);
    });

    it("should return SharePoint instance", () => {
      const sharePoint = Azure.sharePoint;
      expect(sharePoint).toBeInstanceOf(SharePoint);
    });

    it("should return same instance on multiple accesses (singleton)", () => {
      const outlook1 = Azure.outlook;
      const outlook2 = Azure.outlook;
      expect(outlook1).toBe(outlook2);

      const teams1 = Azure.teams;
      const teams2 = Azure.teams;
      expect(teams1).toBe(teams2);

      const calendar1 = Azure.calendar;
      const calendar2 = Azure.calendar;
      expect(calendar1).toBe(calendar2);

      const sharePoint1 = Azure.sharePoint;
      const sharePoint2 = Azure.sharePoint;
      expect(sharePoint1).toBe(sharePoint2);
    });

    it("should lazy-load services only when accessed", () => {
      Azure.reset();
      Azure.config({ accessToken: "test-token" });

      // Before accessing, instances should be undefined
      // We can't directly test this since they're private, but we can verify
      // that accessing them creates new instances
      const outlook = Azure.outlook;
      expect(outlook).toBeDefined();
      expect(outlook).toBeInstanceOf(Outlook);
    });
  });

  describe("Integration Pattern", () => {
    it("should support recommended usage pattern", () => {
      // Setup once
      Azure.config({
        clientId: "test-client",
        clientSecret: "test-secret",
        tenantId: "test-tenant",
        refreshToken: "test-refresh-token",
      });

      // Access services multiple times
      const outlook = Azure.outlook;
      const teams = Azure.teams;
      const calendar = Azure.calendar;
      const sharePoint = Azure.sharePoint;

      expect(outlook).toBeInstanceOf(Outlook);
      expect(teams).toBeInstanceOf(Teams);
      expect(calendar).toBeInstanceOf(Calendar);
      expect(sharePoint).toBeInstanceOf(SharePoint);

      // Should return same instances
      expect(Azure.outlook).toBe(outlook);
      expect(Azure.teams).toBe(teams);
      expect(Azure.calendar).toBe(calendar);
      expect(Azure.sharePoint).toBe(sharePoint);
    });

    it("should work with exported instance pattern", () => {
      // Simulate config file pattern
      Azure.config({
        accessToken: "test-token",
      });

      // Export would happen here in real code
      // import Azure from './config/azure'

      // Use in different modules
      const module1Outlook = Azure.outlook;
      const module2Outlook = Azure.outlook;

      // Should be same instance across "modules"
      expect(module1Outlook).toBe(module2Outlook);
    });
  });

  describe("Utility Methods", () => {
    it("should expose listStoredCredentials", async () => {
      expect(typeof Azure.listStoredCredentials).toBe("function");
    });

    it("should expose clearStoredCredentials", async () => {
      expect(typeof Azure.clearStoredCredentials).toBe("function");
    });
  });
});
