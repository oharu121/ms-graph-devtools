import { AzureAuth } from "../core/auth";
import { AzureConfig, Tag } from "../types";

/**
 * Teams service for Microsoft Graph API
 * Handles Teams messaging, adaptive cards, and mentions
 */
export class Teams {
  private auth: AzureAuth;

  /**
   * Create a new Teams service instance
   *
   * @param config - Optional config or AzureAuth instance
   *
   * @example
   * const teams = new Teams();
   * await teams.postAdaptiveCard(teamId, channelId, card, tags);
   */
  constructor(config?: AzureConfig | AzureAuth) {
    if (config instanceof AzureAuth) {
      this.auth = config;
    } else {
      this.auth = new AzureAuth(config);
    }
  }

  /**
   * Get all teams for the current user
   *
   * @returns Array of teams with id and displayName
   *
   * @example
   * const teams = await teams.getTeams();
   * console.log(teams); // [{ id: '...', displayName: 'Engineering Team' }, ...]
   */
  async getTeams() {
    await this.auth.checkToken();
    return this.auth.withRetry(async () => {
      const token = await this.auth.getAccessToken();
      const url = "https://graph.microsoft.com/v1.0/me/joinedTeams";
      const res = await this.auth.getAxon().bearer(token).get(url);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      return res.data.value.map((team: any) => ({
        id: team.id,
        displayName: team.displayName,
        description: team.description,
      }));
    });
  }

  /**
   * Get all channels for a specific team
   *
   * @param teamId - Team ID
   * @returns Array of channels with id and displayName
   *
   * @example
   * const channels = await teams.getChannels('team-id');
   * console.log(channels); // [{ id: '...', displayName: 'General' }, ...]
   */
  async getChannels(teamId: string) {
    await this.auth.checkToken();
    return this.auth.withRetry(async () => {
      const token = await this.auth.getAccessToken();
      const url = `https://graph.microsoft.com/v1.0/teams/${teamId}/channels`;
      const res = await this.auth.getAxon().bearer(token).get(url);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      return res.data.value.map((channel: any) => ({
        id: channel.id,
        displayName: channel.displayName,
        description: channel.description,
        membershipType: channel.membershipType,
      }));
    });
  }

  /**
   * Get all tags for a specific team
   *
   * @param teamId - Team ID
   * @returns Array of tags with id and displayName
   *
   * @example
   * const tags = await teams.getTags('team-id');
   * console.log(tags); // [{ id: '...', displayName: 'Engineering', memberCount: 5 }, ...]
   */
  async getTags(teamId: string) {
    await this.auth.checkToken();
    return this.auth.withRetry(async () => {
      const token = await this.auth.getAccessToken();
      const url = `https://graph.microsoft.com/v1.0/teams/${teamId}/tags`;
      const res = await this.auth.getAxon().bearer(token).get(url);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      return res.data.value.map((tag: any) => ({
        id: tag.id,
        displayName: tag.displayName,
        description: tag.description,
        memberCount: tag.memberCount,
      }));
    });
  }

  /**
   * Post an adaptive card to a Teams channel
   *
   * @param teamId - Team ID
   * @param channelId - Channel ID
   * @param card - Adaptive card JSON
   * @param tags - Optional mention tags
   * @returns Response data
   *
   * @example
   * await teams.postAdaptiveCard(
   *   'a1b2c3d4-e5f6-7890-a1b2-c3d4e5f67890',
   *   '19:abc123def456ghi789jkl012mno345pqr678@thread.tacv2',
   *   { type: 'AdaptiveCard', body: [...] },
   *   [{ id: 'xyz-123', displayName: 'Team Name', type: 'team' }]
   * );
   */
  async postAdaptiveCard(
    teamId: string,
    channelId: string,
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    card: any,
    tags?: Tag[]
  ) {
    await this.auth.checkToken();
    return this.auth.withRetry(async () => {
      const token = await this.auth.getAccessToken();
      const tagText = tags ? this.createMentionTags(tags).join("<br>") : "";

      // Generate a random attachment ID
      const attachmentId = this.generateAttachmentId();

      const payload = {
        body: {
          contentType: "html",
          content: `<div><div>${tagText}<attachment id="${attachmentId}"></attachment></div></div>`,
        },
        attachments: [
          {
            id: attachmentId,
            contentType: "application/vnd.microsoft.card.adaptive",
            contentUrl: null,
            content: JSON.stringify(card),
            name: null,
            thumbnailUrl: null,
            teamsAppId: null,
          },
        ],
        mentions: tags ? this.createMentionBody(tags) : [],
      };

      const url = `https://graph.microsoft.com/v1.0/teams/${teamId}/channels/${channelId}/messages`;
      const res = await this.auth.getAxon().bearer(token).post(url, payload);
      return res.data;
    });
  }

  /**
   * Generate a random attachment ID for adaptive cards
   */
  private generateAttachmentId(): string {
    return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, (c) => {
      const r = (Math.random() * 16) | 0;
      const v = c === 'x' ? r : (r & 0x3) | 0x8;
      return v.toString(16);
    });
  }

  /**
   * Create mention tags for HTML content
   */
  private createMentionTags(tags: Tag[]): string[] {
    return tags.map((tag, i) => `<at id="${i}">${tag.displayName}</at>`);
  }

  /**
   * Create mention body for Teams API
   */
  createMentionBody(tags: Tag[]) {
    return tags.map((tag, i) => ({
      id: i,
      mentionText: tag.displayName,
      mentioned: this.getMentionedObject(tag),
    }));
  }

  /**
   * Get mentioned object based on tag type
   */
  private getMentionedObject(tag: Tag) {
    switch (tag.type) {
      case "team":
        return {
          conversation: {
            id: tag.id,
            displayName: tag.displayName,
            conversationIdentityType: "team",
          },
        };
      case "tag":
        return {
          tag: {
            id: tag.id,
            displayName: tag.displayName,
            conversationIdentityType: "channel",
          },
        };
      case "user":
        return {
          user: {
            id: tag.id,
            displayName: tag.displayName,
            userIdentityType: "aadUser",
          },
        };
      default:
        throw new Error("unrecognized tag type");
    }
  }

  /**
   * Create a fluent adaptive card builder for composing Teams messages
   *
   * @returns AdaptiveCardBuilder instance
   *
   * @example
   * await teams.compose()
   *   .team('team-id-here')
   *   .channel('channel-id-here')
   *   .card({
   *     type: 'AdaptiveCard',
   *     version: '1.4',
   *     body: [{ type: 'TextBlock', text: 'Hello Teams!' }]
   *   })
   *   .mentionTeam('team-id', 'Team Name')
   *   .send();
   *
   * @example
   * // With multiple mentions
   * await teams.compose()
   *   .team('team-id')
   *   .channel('channel-id')
   *   .card(myCard)
   *   .mentionUser('user-id', 'John Doe')
   *   .mentionTag('tag-id', 'Engineering Team')
   *   .send();
   */
  compose(): AdaptiveCardBuilder {
    return new AdaptiveCardBuilder(this);
  }
}

/**
 * Fluent adaptive card builder for composing and sending Teams messages
 * Use via teams.compose()
 */
export class AdaptiveCardBuilder {
  private teams: Teams;
  private teamId: string = "";
  private channelId: string = "";
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private adaptiveCard: any = null;
  private tags: Tag[] = [];

  constructor(teams: Teams) {
    this.teams = teams;
  }

  /**
   * Set the target team ID
   *
   * @param teamId - Microsoft Teams team ID
   */
  team(teamId: string): AdaptiveCardBuilder {
    this.teamId = teamId;
    return this;
  }

  /**
   * Set the target channel ID
   *
   * @param channelId - Microsoft Teams channel ID
   */
  channel(channelId: string): AdaptiveCardBuilder {
    this.channelId = channelId;
    return this;
  }

  /**
   * Set the adaptive card content
   *
   * @param card - Adaptive card JSON object
   *
   * @example
   * .card({
   *   type: 'AdaptiveCard',
   *   version: '1.4',
   *   body: [
   *     { type: 'TextBlock', text: 'Hello!', size: 'Large' }
   *   ],
   *   actions: [
   *     { type: 'Action.OpenUrl', title: 'Learn More', url: 'https://example.com' }
   *   ]
   * })
   */
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  card(card: any): AdaptiveCardBuilder {
    this.adaptiveCard = card;
    return this;
  }

  /**
   * Add a team mention
   *
   * @param id - Team ID
   * @param displayName - Display name for the mention
   */
  mentionTeam(id: string, displayName: string): AdaptiveCardBuilder {
    this.tags.push({ id, displayName, type: "team" });
    return this;
  }

  /**
   * Add a tag mention
   *
   * @param id - Tag ID
   * @param displayName - Display name for the tag
   */
  mentionTag(id: string, displayName: string): AdaptiveCardBuilder {
    this.tags.push({ id, displayName, type: "tag" });
    return this;
  }

  /**
   * Add a user mention
   *
   * @param id - User ID (Azure AD user ID)
   * @param displayName - Display name for the user
   */
  mentionUser(id: string, displayName: string): AdaptiveCardBuilder {
    this.tags.push({ id, displayName, type: "user" });
    return this;
  }

  /**
   * Add multiple mentions at once
   *
   * @param tags - Array of tag objects
   *
   * @example
   * .mentions([
   *   { id: 'team-id', displayName: 'Engineering', type: 'team' },
   *   { id: 'user-id', displayName: 'John Doe', type: 'user' }
   * ])
   */
  mentions(tags: Tag[]): AdaptiveCardBuilder {
    this.tags.push(...tags);
    return this;
  }

  /**
   * Clear all mentions
   */
  clearMentions(): AdaptiveCardBuilder {
    this.tags = [];
    return this;
  }

  /**
   * Get the current configuration (useful for debugging)
   */
  getConfig() {
    return {
      teamId: this.teamId,
      channelId: this.channelId,
      card: this.adaptiveCard,
      tags: this.tags,
    };
  }

  /**
   * Send the adaptive card to Teams
   *
   * @returns Response data from Microsoft Graph API
   * @throws Error if team ID, channel ID, or card is not set
   */
  async send() {
    if (!this.teamId) {
      throw new Error("Team ID is required. Use .team() to set it.");
    }
    if (!this.channelId) {
      throw new Error("Channel ID is required. Use .channel() to set it.");
    }
    if (!this.adaptiveCard) {
      throw new Error("Adaptive card is required. Use .card() to set it.");
    }

    return await this.teams.postAdaptiveCard(
      this.teamId,
      this.channelId,
      this.adaptiveCard,
      this.tags.length > 0 ? this.tags : undefined
    );
  }
}
