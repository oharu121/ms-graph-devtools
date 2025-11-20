import { describe, it, expect, beforeEach, vi } from "vitest";
import { Teams, AdaptiveCardBuilder } from "../src/services/Teams";
import { AzureAuth } from "../src/core/auth";
import Axon from "axios-fluent";

// Mock axios-fluent
vi.mock("axios-fluent");

describe("Teams Service", () => {
  let teams: Teams;
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
      const t = new Teams(mockAuth);
      expect(t).toBeInstanceOf(Teams);
    });

    it("should create instance with config object", () => {
      const t = new Teams({
        clientId: "test-client",
        clientSecret: "test-secret",
        tenantId: "test-tenant",
        refreshToken: "test-token",
      });
      expect(t).toBeInstanceOf(Teams);
    });

    it("should create instance with no config", () => {
      const t = new Teams();
      expect(t).toBeInstanceOf(Teams);
    });
  });

  describe("getTeams", () => {
    beforeEach(() => {
      teams = new Teams(mockAuth);
    });

    it("should fetch joined teams", async () => {
      const mockTeams = [
        {
          id: "team-1",
          displayName: "Engineering Team",
          description: "Engineering collaboration",
        },
        {
          id: "team-2",
          displayName: "Marketing Team",
          description: "Marketing campaigns",
        },
      ];

      const mockAxonInstance = {
        bearer: vi.fn().mockReturnThis(),
        get: vi.fn().mockResolvedValue({
          data: { value: mockTeams },
        }),
      };

      (Axon.new as any).mockReturnValue(mockAxonInstance);

      const result = await teams.getTeams();

      expect(mockAuth.getAccessToken).toHaveBeenCalled();
      expect(mockAxonInstance.bearer).toHaveBeenCalledWith("mock-access-token");
      expect(mockAxonInstance.get).toHaveBeenCalledWith(
        "https://graph.microsoft.com/v1.0/me/joinedTeams"
      );
      expect(result).toEqual(mockTeams);
    });

    it("should handle API errors", async () => {
      const mockError = new Error("Unauthorized");

      // Mock withRetry to throw the error
      vi.spyOn(mockAuth, "withRetry").mockRejectedValue(mockError);

      await expect(teams.getTeams()).rejects.toThrow("Unauthorized");
      expect(mockAuth.withRetry).toHaveBeenCalled();
    });
  });

  describe("getChannels", () => {
    beforeEach(() => {
      teams = new Teams(mockAuth);
    });

    it("should fetch channels for a team", async () => {
      const mockChannels = [
        {
          id: "channel-1",
          displayName: "General",
          description: "General channel",
          membershipType: "standard",
        },
        {
          id: "channel-2",
          displayName: "Development",
          description: "Development discussions",
          membershipType: "private",
        },
      ];

      const mockAxonInstance = {
        bearer: vi.fn().mockReturnThis(),
        get: vi.fn().mockResolvedValue({
          data: { value: mockChannels },
        }),
      };

      (Axon.new as any).mockReturnValue(mockAxonInstance);

      const result = await teams.getChannels("team-123");

      expect(mockAxonInstance.get).toHaveBeenCalledWith(
        "https://graph.microsoft.com/v1.0/teams/team-123/channels"
      );
      expect(result).toEqual(mockChannels);
    });
  });

  describe("getTags", () => {
    beforeEach(() => {
      teams = new Teams(mockAuth);
    });

    it("should fetch tags for a team", async () => {
      const mockTags = [
        {
          id: "tag-1",
          displayName: "Engineering",
          description: "Engineering team members",
          memberCount: 5,
        },
        {
          id: "tag-2",
          displayName: "Designers",
          description: "Design team members",
          memberCount: 3,
        },
      ];

      const mockAxonInstance = {
        bearer: vi.fn().mockReturnThis(),
        get: vi.fn().mockResolvedValue({
          data: { value: mockTags },
        }),
      };

      (Axon.new as any).mockReturnValue(mockAxonInstance);

      const result = await teams.getTags("team-123");

      expect(mockAxonInstance.get).toHaveBeenCalledWith(
        "https://graph.microsoft.com/v1.0/teams/team-123/tags"
      );
      expect(result).toEqual(mockTags);
    });
  });

  describe("postAdaptiveCard", () => {
    beforeEach(() => {
      teams = new Teams(mockAuth);
    });

    it("should post adaptive card without mentions", async () => {
      const card = {
        type: "AdaptiveCard",
        version: "1.4",
        body: [{ type: "TextBlock", text: "Hello Teams!" }],
      };

      const mockResponse = { id: "message-123" };

      const mockAxonInstance = {
        bearer: vi.fn().mockReturnThis(),
        post: vi.fn().mockResolvedValue({
          data: mockResponse,
        }),
      };

      (Axon.new as any).mockReturnValue(mockAxonInstance);

      const result = await teams.postAdaptiveCard("team-id", "channel-id", card);

      expect(mockAuth.getAccessToken).toHaveBeenCalled();
      expect(mockAxonInstance.post).toHaveBeenCalled();

      const [url, payload] = (mockAxonInstance.post as any).mock.calls[0];
      expect(url).toBe(
        "https://graph.microsoft.com/v1.0/teams/team-id/channels/channel-id/messages"
      );
      expect(payload.attachments).toHaveLength(1);
      expect(payload.attachments[0].contentType).toBe(
        "application/vnd.microsoft.card.adaptive"
      );
      expect(payload.mentions).toEqual([]);
      expect(result).toEqual(mockResponse);
    });

    it("should post adaptive card with team mention", async () => {
      const card = {
        type: "AdaptiveCard",
        body: [{ type: "TextBlock", text: "Notification" }],
      };

      const tags = [
        { id: "team-123", displayName: "Engineering Team", type: "team" as const },
      ];

      const mockAxonInstance = {
        bearer: vi.fn().mockReturnThis(),
        post: vi.fn().mockResolvedValue({
          data: { id: "msg-456" },
        }),
      };

      (Axon.new as any).mockReturnValue(mockAxonInstance);

      await teams.postAdaptiveCard("team-id", "channel-id", card, tags);

      const [, payload] = (mockAxonInstance.post as any).mock.calls[0];
      expect(payload.mentions).toHaveLength(1);
      expect(payload.mentions[0]).toEqual({
        id: 0,
        mentionText: "Engineering Team",
        mentioned: {
          conversation: {
            id: "team-123",
            displayName: "Engineering Team",
            conversationIdentityType: "team",
          },
        },
      });
      expect(payload.body.content).toContain('<at id="0">Engineering Team</at>');
    });

    it("should post adaptive card with tag mention", async () => {
      const card = { type: "AdaptiveCard", body: [] };
      const tags = [
        { id: "tag-789", displayName: "Developers", type: "tag" as const },
      ];

      const mockAxonInstance = {
        bearer: vi.fn().mockReturnThis(),
        post: vi.fn().mockResolvedValue({
          data: {},
        }),
      };

      (Axon.new as any).mockReturnValue(mockAxonInstance);

      await teams.postAdaptiveCard("team-id", "channel-id", card, tags);

      const [, payload] = (mockAxonInstance.post as any).mock.calls[0];
      expect(payload.mentions[0].mentioned).toEqual({
        tag: {
          id: "tag-789",
          displayName: "Developers",
          conversationIdentityType: "channel",
        },
      });
    });

    it("should post adaptive card with user mention", async () => {
      const card = { type: "AdaptiveCard", body: [] };
      const tags = [
        { id: "user-001", displayName: "John Doe", type: "user" as const },
      ];

      const mockAxonInstance = {
        bearer: vi.fn().mockReturnThis(),
        post: vi.fn().mockResolvedValue({
          data: {},
        }),
      };

      (Axon.new as any).mockReturnValue(mockAxonInstance);

      await teams.postAdaptiveCard("team-id", "channel-id", card, tags);

      const [, payload] = (mockAxonInstance.post as any).mock.calls[0];
      expect(payload.mentions[0].mentioned).toEqual({
        user: {
          id: "user-001",
          displayName: "John Doe",
          userIdentityType: "aadUser",
        },
      });
    });

    it("should post adaptive card with multiple mentions", async () => {
      const card = { type: "AdaptiveCard", body: [] };
      const tags = [
        { id: "team-1", displayName: "Team A", type: "team" as const },
        { id: "tag-1", displayName: "Tag B", type: "tag" as const },
        { id: "user-1", displayName: "User C", type: "user" as const },
      ];

      const mockAxonInstance = {
        bearer: vi.fn().mockReturnThis(),
        post: vi.fn().mockResolvedValue({
          data: {},
        }),
      };

      (Axon.new as any).mockReturnValue(mockAxonInstance);

      await teams.postAdaptiveCard("team-id", "channel-id", card, tags);

      const [, payload] = (mockAxonInstance.post as any).mock.calls[0];
      expect(payload.mentions).toHaveLength(3);
      expect(payload.body.content).toContain('<at id="0">Team A</at>');
      expect(payload.body.content).toContain('<at id="1">Tag B</at>');
      expect(payload.body.content).toContain('<at id="2">User C</at>');
    });

    it("should handle API errors", async () => {
      const mockError = new Error("Post failed");
      const mockAxonInstance = {
        bearer: vi.fn().mockReturnThis(),
        post: vi.fn().mockRejectedValue(mockError),
      };

      (Axon.new as any).mockReturnValue(mockAxonInstance);

      const card = { type: "AdaptiveCard", body: [] };

      await expect(
        teams.postAdaptiveCard("team-id", "channel-id", card)
      ).rejects.toThrow("Post failed");
    });
  });

  describe("createMentionBody", () => {
    beforeEach(() => {
      teams = new Teams(mockAuth);
    });

    it("should create mention body for different tag types", () => {
      const tags = [
        { id: "team-1", displayName: "Team", type: "team" as const },
        { id: "tag-1", displayName: "Tag", type: "tag" as const },
        { id: "user-1", displayName: "User", type: "user" as const },
      ];

      const result = teams.createMentionBody(tags);

      expect(result).toHaveLength(3);
      expect(result[0]).toEqual({
        id: 0,
        mentionText: "Team",
        mentioned: {
          conversation: {
            id: "team-1",
            displayName: "Team",
            conversationIdentityType: "team",
          },
        },
      });
    });
  });

  describe("compose - AdaptiveCardBuilder", () => {
    beforeEach(() => {
      teams = new Teams(mockAuth);
    });

    it("should return AdaptiveCardBuilder instance", () => {
      const builder = teams.compose();
      expect(builder).toBeInstanceOf(AdaptiveCardBuilder);
    });
  });
});

describe("AdaptiveCardBuilder", () => {
  let teams: Teams;
  let mockAuth: AzureAuth;
  let builder: AdaptiveCardBuilder;

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

    teams = new Teams(mockAuth);
    builder = teams.compose();
  });

  describe("Fluent API", () => {
    it("should set team ID", () => {
      builder.team("team-123");
      const config = builder.getConfig();
      expect(config.teamId).toBe("team-123");
    });

    it("should set channel ID", () => {
      builder.channel("channel-456");
      const config = builder.getConfig();
      expect(config.channelId).toBe("channel-456");
    });

    it("should set adaptive card", () => {
      const card = {
        type: "AdaptiveCard",
        version: "1.4",
        body: [{ type: "TextBlock", text: "Hello" }],
      };

      builder.card(card);
      const config = builder.getConfig();
      expect(config.card).toEqual(card);
    });

    it("should add team mention", () => {
      builder.mentionTeam("team-789", "Engineering Team");
      const config = builder.getConfig();
      expect(config.tags).toHaveLength(1);
      expect(config.tags[0]).toEqual({
        id: "team-789",
        displayName: "Engineering Team",
        type: "team",
      });
    });

    it("should add tag mention", () => {
      builder.mentionTag("tag-001", "Developers");
      const config = builder.getConfig();
      expect(config.tags).toHaveLength(1);
      expect(config.tags[0]).toEqual({
        id: "tag-001",
        displayName: "Developers",
        type: "tag",
      });
    });

    it("should add user mention", () => {
      builder.mentionUser("user-xyz", "John Doe");
      const config = builder.getConfig();
      expect(config.tags).toHaveLength(1);
      expect(config.tags[0]).toEqual({
        id: "user-xyz",
        displayName: "John Doe",
        type: "user",
      });
    });

    it("should add multiple mentions at once", () => {
      const tags = [
        { id: "team-1", displayName: "Team 1", type: "team" as const },
        { id: "user-1", displayName: "User 1", type: "user" as const },
      ];

      builder.mentions(tags);
      const config = builder.getConfig();
      expect(config.tags).toHaveLength(2);
    });

    it("should accumulate mentions from multiple calls", () => {
      builder
        .mentionTeam("team-1", "Team 1")
        .mentionUser("user-1", "User 1")
        .mentionTag("tag-1", "Tag 1");

      const config = builder.getConfig();
      expect(config.tags).toHaveLength(3);
    });

    it("should clear all mentions", () => {
      builder
        .mentionTeam("team-1", "Team 1")
        .mentionUser("user-1", "User 1")
        .clearMentions();

      const config = builder.getConfig();
      expect(config.tags).toHaveLength(0);
    });

    it("should chain all methods", () => {
      const card = { type: "AdaptiveCard", body: [] };

      builder
        .team("team-123")
        .channel("channel-456")
        .card(card)
        .mentionTeam("team-789", "Engineering")
        .mentionUser("user-001", "John");

      const config = builder.getConfig();
      expect(config.teamId).toBe("team-123");
      expect(config.channelId).toBe("channel-456");
      expect(config.card).toEqual(card);
      expect(config.tags).toHaveLength(2);
    });
  });

  describe("getConfig", () => {
    it("should return current configuration", () => {
      const card = { type: "AdaptiveCard", body: [] };

      builder
        .team("team-id")
        .channel("channel-id")
        .card(card)
        .mentionUser("user-1", "User");

      const config = builder.getConfig();

      expect(config).toEqual({
        teamId: "team-id",
        channelId: "channel-id",
        card: card,
        tags: [{ id: "user-1", displayName: "User", type: "user" }],
      });
    });
  });

  describe("send", () => {
    it("should send adaptive card successfully", async () => {
      const card = {
        type: "AdaptiveCard",
        version: "1.4",
        body: [{ type: "TextBlock", text: "Hello" }],
      };

      const mockResponse = { id: "message-123" };

      const mockAxonInstance = {
        bearer: vi.fn().mockReturnThis(),
        post: vi.fn().mockResolvedValue({
          data: mockResponse,
        }),
      };

      (Axon.new as any).mockReturnValue(mockAxonInstance);

      builder.team("team-id").channel("channel-id").card(card);

      const result = await builder.send();

      expect(mockAuth.getAccessToken).toHaveBeenCalled();
      expect(result).toEqual(mockResponse);
    });

    it("should send adaptive card with mentions", async () => {
      const card = { type: "AdaptiveCard", body: [] };

      const mockAxonInstance = {
        bearer: vi.fn().mockReturnThis(),
        post: vi.fn().mockResolvedValue({
          data: { id: "msg-456" },
        }),
      };

      (Axon.new as any).mockReturnValue(mockAxonInstance);

      builder
        .team("team-id")
        .channel("channel-id")
        .card(card)
        .mentionUser("user-1", "John");

      await builder.send();

      const [, payload] = (mockAxonInstance.post as any).mock.calls[0];
      expect(payload.mentions).toBeDefined();
      expect(payload.mentions).toHaveLength(1);
    });

    it("should throw error when team ID not set", async () => {
      const card = { type: "AdaptiveCard", body: [] };
      builder.channel("channel-id").card(card);

      await expect(builder.send()).rejects.toThrow(
        "Team ID is required. Use .team() to set it."
      );
    });

    it("should throw error when channel ID not set", async () => {
      const card = { type: "AdaptiveCard", body: [] };
      builder.team("team-id").card(card);

      await expect(builder.send()).rejects.toThrow(
        "Channel ID is required. Use .channel() to set it."
      );
    });

    it("should throw error when card not set", async () => {
      builder.team("team-id").channel("channel-id");

      await expect(builder.send()).rejects.toThrow(
        "Adaptive card is required. Use .card() to set it."
      );
    });

    it("should propagate API errors", async () => {
      const mockError = new Error("API Error");
      const mockAxonInstance = {
        bearer: vi.fn().mockReturnThis(),
        post: vi.fn().mockRejectedValue(mockError),
      };

      (Axon.new as any).mockReturnValue(mockAxonInstance);

      const card = { type: "AdaptiveCard", body: [] };
      builder.team("team-id").channel("channel-id").card(card);

      await expect(builder.send()).rejects.toThrow("API Error");
    });

    it("should not include tags if clearMentions was called", async () => {
      const card = { type: "AdaptiveCard", body: [] };

      const mockAxonInstance = {
        bearer: vi.fn().mockReturnThis(),
        post: vi.fn().mockResolvedValue({
          data: {},
        }),
      };

      (Axon.new as any).mockReturnValue(mockAxonInstance);

      builder
        .team("team-id")
        .channel("channel-id")
        .card(card)
        .mentionUser("user-1", "John")
        .clearMentions();

      await builder.send();

      const [, payload] = (mockAxonInstance.post as any).mock.calls[0];
      expect(payload.mentions).toEqual([]);
    });
  });
});
