import Axon, { AxonError } from "axios-fluent";
import { AzureAuth } from "../core/auth";
import { AzureConfig } from "../types";

/**
 * SharePoint service for Microsoft Graph API
 * Handles SharePoint list operations and site management
 */
export class SharePoint {
  private auth: AzureAuth;
  private siteId?: string;

  /**
   * Create a new SharePoint service instance
   *
   * @param config - Optional config or AzureAuth instance
   * @param siteId - Optional site ID (can be set later or per-operation)
   *
   * @example
   * const sharepoint = new SharePoint();
   * // Set site ID later
   * await sharepoint.setSiteId('your-site-id');
   *
   * @example
   * // With site ID in constructor
   * const sharepoint = new SharePoint(undefined, 'your-site-id');
   */
  constructor(config?: AzureConfig | AzureAuth, siteId?: string) {
    if (config instanceof AzureAuth) {
      this.auth = config;
    } else {
      this.auth = new AzureAuth(config);
    }

    if (siteId) {
      this.siteId = siteId;
    }
  }

  /**
   * Set the site ID for this instance
   *
   * @param siteId - SharePoint site ID
   */
  setSiteId(siteId: string): void {
    this.siteId = siteId;
  }

  /**
   * Get the current site ID
   *
   * @returns Current site ID or undefined
   */
  getSiteId(): string | undefined {
    return this.siteId;
  }

  /**
   * Search for SharePoint sites
   *
   * @param query - Search query
   * @returns Array of sites
   *
   * @example
   * const sites = await sharepoint.searchSites('Engineering');
   * console.log(sites); // [{ id: '...', displayName: 'Engineering Site', webUrl: '...' }]
   */
  async searchSites(query: string) {
    try {
      const token = await this.auth.getAccessToken();
      const url = `https://graph.microsoft.com/v1.0/sites?search=${query}`;
      const res = await Axon.new().bearer(token).get(url);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      return res.data.value.map((site: any) => ({
        id: site.id,
        displayName: site.displayName,
        name: site.name,
        webUrl: site.webUrl,
        description: site.description,
      }));
    } catch (error) {
      this.auth.handleApiError(error as AxonError);
    }
  }

  /**
   * Get a specific site by hostname and site path
   *
   * @param hostname - SharePoint hostname (e.g., 'contoso.sharepoint.com')
   * @param sitePath - Site path (e.g., '/sites/team')
   * @returns Site information
   *
   * @example
   * const site = await sharepoint.getSiteByPath('contoso.sharepoint.com', '/sites/engineering');
   */
  async getSiteByPath(hostname: string, sitePath: string) {
    try {
      const token = await this.auth.getAccessToken();
      const url = `https://graph.microsoft.com/v1.0/sites/${hostname}:${sitePath}`;
      const res = await Axon.new().bearer(token).get(url);
      return {
        id: res.data.id,
        displayName: res.data.displayName,
        name: res.data.name,
        webUrl: res.data.webUrl,
        description: res.data.description,
      };
    } catch (error) {
      this.auth.handleApiError(error as AxonError);
    }
  }

  /**
   * Get all lists in a site
   *
   * @param siteId - Optional site ID (uses instance siteId if not provided)
   * @returns Array of lists
   *
   * @example
   * const lists = await sharepoint.getLists();
   * console.log(lists); // [{ id: '...', displayName: 'Tasks', ... }]
   */
  async getLists(siteId?: string) {
    try {
      const targetSiteId = siteId || this.siteId;
      if (!targetSiteId) {
        throw new Error("Site ID is required. Provide it in constructor, setSiteId(), or as parameter.");
      }

      const token = await this.auth.getAccessToken();
      const url = `https://graph.microsoft.com/v1.0/sites/${targetSiteId}/lists`;
      const res = await Axon.new().bearer(token).get(url);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      return res.data.value.map((list: any) => ({
        id: list.id,
        displayName: list.displayName,
        name: list.name,
        description: list.description,
        webUrl: list.webUrl,
      }));
    } catch (error) {
      this.auth.handleApiError(error as AxonError);
    }
  }

  /**
   * Get a SharePoint list by ID or display name
   *
   * @param listIdOrName - List ID or display name
   * @param siteId - Optional site ID (uses instance siteId if not provided)
   * @returns List information
   *
   * @example
   * const list = await sharepoint.getList('Tasks');
   */
  async getList(listIdOrName: string, siteId?: string) {
    try {
      const targetSiteId = siteId || this.siteId;
      if (!targetSiteId) {
        throw new Error("Site ID is required. Provide it in constructor, setSiteId(), or as parameter.");
      }

      const token = await this.auth.getAccessToken();
      const url = `https://graph.microsoft.com/v1.0/sites/${targetSiteId}/lists/${listIdOrName}`;
      const res = await Axon.new().bearer(token).get(url);
      return res.data;
    } catch (error) {
      this.auth.handleApiError(error as AxonError);
    }
  }

  /**
   * Get all items from a list
   *
   * @param listId - List ID or display name
   * @param options - Query options (filter, orderby, top, expand)
   * @param siteId - Optional site ID (uses instance siteId if not provided)
   * @returns Array of list items
   *
   * @example
   * const items = await sharepoint.getListItems('Tasks', {
   *   filter: "fields/Status eq 'Active'",
   *   orderby: 'createdDateTime desc',
   *   top: 10,
   *   expand: 'fields'
   * });
   */
  async getListItems(
    listId: string,
    options?: {
      filter?: string;
      orderby?: string;
      top?: number;
      expand?: string;
    },
    siteId?: string
  ) {
    try {
      const targetSiteId = siteId || this.siteId;
      if (!targetSiteId) {
        throw new Error("Site ID is required. Provide it in constructor, setSiteId(), or as parameter.");
      }

      const token = await this.auth.getAccessToken();
      const url = `https://graph.microsoft.com/v1.0/sites/${targetSiteId}/lists/${listId}/items`;

      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const params: any = {};
      if (options?.filter) params.$filter = options.filter;
      if (options?.orderby) params.$orderby = options.orderby;
      if (options?.top) params.$top = options.top;
      if (options?.expand) params.$expand = options.expand;

      const res = await Axon.new()
        .bearer(token)
        .params(params)
        .get(url);

      return res.data.value;
    } catch (error) {
      this.auth.handleApiError(error as AxonError);
    }
  }

  /**
   * Get a specific item from a list
   *
   * @param listId - List ID or display name
   * @param itemId - Item ID
   * @param expand - Optional fields to expand (e.g., 'fields')
   * @param siteId - Optional site ID (uses instance siteId if not provided)
   * @returns List item
   *
   * @example
   * const item = await sharepoint.getListItem('Tasks', '123', 'fields');
   */
  async getListItem(
    listId: string,
    itemId: string,
    expand?: string,
    siteId?: string
  ) {
    try {
      const targetSiteId = siteId || this.siteId;
      if (!targetSiteId) {
        throw new Error("Site ID is required. Provide it in constructor, setSiteId(), or as parameter.");
      }

      const token = await this.auth.getAccessToken();
      const url = `https://graph.microsoft.com/v1.0/sites/${targetSiteId}/lists/${listId}/items/${itemId}`;

      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const params: any = {};
      if (expand) params.$expand = expand;

      const res = await Axon.new()
        .bearer(token)
        .params(params)
        .get(url);

      return res.data;
    } catch (error) {
      this.auth.handleApiError(error as AxonError);
    }
  }

  /**
   * Create a new item in a list
   *
   * @param listId - List ID or display name
   * @param fields - Field values for the new item
   * @param siteId - Optional site ID (uses instance siteId if not provided)
   * @returns Created item
   *
   * @example
   * const item = await sharepoint.createListItem('Tasks', {
   *   Title: 'New Task',
   *   Status: 'Active',
   *   Priority: 'High'
   * });
   */
  async createListItem(
    listId: string,
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    fields: Record<string, any>,
    siteId?: string
  ) {
    try {
      const targetSiteId = siteId || this.siteId;
      if (!targetSiteId) {
        throw new Error("Site ID is required. Provide it in constructor, setSiteId(), or as parameter.");
      }

      const token = await this.auth.getAccessToken();
      const url = `https://graph.microsoft.com/v1.0/sites/${targetSiteId}/lists/${listId}/items`;

      const payload = { fields };

      const res = await Axon.new().bearer(token).post(url, payload);
      return res.data;
    } catch (error) {
      this.auth.handleApiError(error as AxonError);
    }
  }

  /**
   * Update an existing item in a list
   *
   * @param listId - List ID or display name
   * @param itemId - Item ID
   * @param fields - Field values to update
   * @param siteId - Optional site ID (uses instance siteId if not provided)
   * @returns Updated item
   *
   * @example
   * const item = await sharepoint.updateListItem('Tasks', '123', {
   *   Status: 'Completed'
   * });
   */
  async updateListItem(
    listId: string,
    itemId: string,
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    fields: Record<string, any>,
    siteId?: string
  ) {
    try {
      const targetSiteId = siteId || this.siteId;
      if (!targetSiteId) {
        throw new Error("Site ID is required. Provide it in constructor, setSiteId(), or as parameter.");
      }

      const token = await this.auth.getAccessToken();
      const url = `https://graph.microsoft.com/v1.0/sites/${targetSiteId}/lists/${listId}/items/${itemId}`;

      const payload = { fields };

      const res = await Axon.new().bearer(token).patch(url, payload);
      return res.data;
    } catch (error) {
      this.auth.handleApiError(error as AxonError);
    }
  }

  /**
   * Delete an item from a list
   *
   * @param listId - List ID or display name
   * @param itemId - Item ID to delete
   * @param siteId - Optional site ID (uses instance siteId if not provided)
   *
   * @example
   * await sharepoint.deleteListItem('Tasks', '123');
   */
  async deleteListItem(listId: string, itemId: string, siteId?: string) {
    try {
      const targetSiteId = siteId || this.siteId;
      if (!targetSiteId) {
        throw new Error("Site ID is required. Provide it in constructor, setSiteId(), or as parameter.");
      }

      const token = await this.auth.getAccessToken();
      const url = `https://graph.microsoft.com/v1.0/sites/${targetSiteId}/lists/${listId}/items/${itemId}`;
      await Axon.new().bearer(token).delete(url);
    } catch (error) {
      this.auth.handleApiError(error as AxonError);
    }
  }

  /**
   * Delete multiple items from a list
   *
   * @param listId - List ID or display name
   * @param itemIds - Array of item IDs to delete
   * @param siteId - Optional site ID (uses instance siteId if not provided)
   *
   * @example
   * await sharepoint.deleteListItems('Tasks', ['123', '456', '789']);
   */
  async deleteListItems(
    listId: string,
    itemIds: string[],
    siteId?: string
  ) {
    try {
      await Promise.all(
        itemIds.map((itemId) => this.deleteListItem(listId, itemId, siteId))
      );
    } catch (error) {
      this.auth.handleApiError(error as AxonError);
    }
  }

  /**
   * Query list items and process them (useful for task queues)
   *
   * @param listId - List ID or display name
   * @param filter - OData filter expression
   * @param processor - Function to process each item
   * @param deleteAfterProcess - Whether to delete items after processing (default: false)
   * @param siteId - Optional site ID (uses instance siteId if not provided)
   * @returns Array of processed results
   *
   * @example
   * // Process and delete items
   * const results = await sharepoint.queryAndProcess(
   *   'Tasks',
   *   "fields/Status eq 'Pending'",
   *   (item) => ({ id: item.id, title: item.fields.Title }),
   *   true
   * );
   */
  async queryAndProcess<T>(
    listId: string,
    filter: string,
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    processor: (item: any) => T,
    deleteAfterProcess: boolean = false,
    siteId?: string
  ): Promise<T[]> {
    try {
      const items = await this.getListItems(
        listId,
        {
          filter,
          expand: "fields",
          orderby: "createdDateTime asc",
        },
        siteId
      );

      if (!items || items.length === 0) {
        return [];
      }

      const results = items.map(processor);

      if (deleteAfterProcess) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const itemIds = items.map((item: any) => item.id);
        await this.deleteListItems(listId, itemIds, siteId);
      }

      return results;
    } catch (error) {
      this.auth.handleApiError(error as AxonError);
      return [];
    }
  }

  /**
   * Get the latest item from a list
   *
   * @param listId - List ID or display name
   * @param orderBy - Field to order by (default: 'createdDateTime desc')
   * @param filter - Optional OData filter expression
   * @param siteId - Optional site ID (uses instance siteId if not provided)
   * @returns Latest item or undefined
   *
   * @example
   * const latestTask = await sharepoint.getLatestItem('Tasks');
   */
  async getLatestItem(
    listId: string,
    orderBy: string = "createdDateTime desc",
    filter?: string,
    siteId?: string
  ) {
    try {
      const items = await this.getListItems(
        listId,
        {
          filter,
          orderby: orderBy,
          top: 1,
          expand: "fields",
        },
        siteId
      );

      return items && items.length > 0 ? items[0] : undefined;
    } catch (error) {
      this.auth.handleApiError(error as AxonError);
    }
  }

  /**
   * Get columns (fields) for a list
   *
   * @param listId - List ID or display name
   * @param siteId - Optional site ID (uses instance siteId if not provided)
   * @returns Array of column definitions
   *
   * @example
   * const columns = await sharepoint.getListColumns('Tasks');
   * console.log(columns); // [{ name: 'Title', displayName: 'Title', ... }]
   */
  async getListColumns(listId: string, siteId?: string) {
    try {
      const targetSiteId = siteId || this.siteId;
      if (!targetSiteId) {
        throw new Error("Site ID is required. Provide it in constructor, setSiteId(), or as parameter.");
      }

      const token = await this.auth.getAccessToken();
      const url = `https://graph.microsoft.com/v1.0/sites/${targetSiteId}/lists/${listId}/columns`;
      const res = await Axon.new().bearer(token).get(url);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      return res.data.value.map((column: any) => ({
        id: column.id,
        name: column.name,
        displayName: column.displayName,
        columnGroup: column.columnGroup,
        description: column.description,
        hidden: column.hidden,
        readOnly: column.readOnly,
      }));
    } catch (error) {
      this.auth.handleApiError(error as AxonError);
    }
  }
}
