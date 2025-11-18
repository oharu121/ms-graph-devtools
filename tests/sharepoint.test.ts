import { describe, it, expect, beforeEach, vi } from "vitest";
import { SharePoint } from "../src/services/SharePoint";
import { AzureAuth } from "../src/core/auth";
import Axon from "axios-fluent";

// Mock axios-fluent
vi.mock("axios-fluent");

describe("SharePoint Service", () => {
  let sharepoint: SharePoint;
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
    vi.spyOn(mockAuth, "handleApiError");
  });

  describe("Constructor", () => {
    it("should create instance with AzureAuth instance", () => {
      const sp = new SharePoint(mockAuth);
      expect(sp).toBeInstanceOf(SharePoint);
    });

    it("should create instance with config object", () => {
      const sp = new SharePoint({
        clientId: "test-client",
        clientSecret: "test-secret",
        tenantId: "test-tenant",
        refreshToken: "test-token",
      });
      expect(sp).toBeInstanceOf(SharePoint);
    });

    it("should create instance with siteId", () => {
      const sp = new SharePoint(undefined, "site-123");
      expect(sp.getSiteId()).toBe("site-123");
    });

    it("should create instance with config and siteId", () => {
      const sp = new SharePoint(mockAuth, "site-456");
      expect(sp.getSiteId()).toBe("site-456");
    });
  });

  describe("setSiteId and getSiteId", () => {
    beforeEach(() => {
      sharepoint = new SharePoint(mockAuth);
    });

    it("should set and get siteId", () => {
      sharepoint.setSiteId("new-site-id");
      expect(sharepoint.getSiteId()).toBe("new-site-id");
    });

    it("should return undefined when siteId not set", () => {
      expect(sharepoint.getSiteId()).toBeUndefined();
    });
  });

  describe("searchSites", () => {
    beforeEach(() => {
      sharepoint = new SharePoint(mockAuth);
    });

    it("should search for sites", async () => {
      const mockSites = [
        {
          id: "site-1",
          displayName: "Engineering Site",
          name: "engineering",
          webUrl: "https://contoso.sharepoint.com/sites/engineering",
          description: "Engineering team site",
        },
        {
          id: "site-2",
          displayName: "Marketing Site",
          name: "marketing",
          webUrl: "https://contoso.sharepoint.com/sites/marketing",
          description: "Marketing team site",
        },
      ];

      const mockAxonInstance = {
        bearer: vi.fn().mockReturnThis(),
        get: vi.fn().mockResolvedValue({
          data: { value: mockSites },
        }),
      };

      (Axon.new as any).mockReturnValue(mockAxonInstance);

      const result = await sharepoint.searchSites("Engineering");

      expect(mockAuth.getAccessToken).toHaveBeenCalled();
      expect(mockAxonInstance.bearer).toHaveBeenCalledWith("mock-access-token");
      expect(mockAxonInstance.get).toHaveBeenCalledWith(
        "https://graph.microsoft.com/v1.0/sites?search=Engineering"
      );
      expect(result).toEqual(mockSites);
    });

    it("should handle API errors", async () => {
      const mockError = new Error("Search failed");
      const mockAxonInstance = {
        bearer: vi.fn().mockReturnThis(),
        get: vi.fn().mockRejectedValue(mockError),
      };

      (Axon.new as any).mockReturnValue(mockAxonInstance);

      await expect(sharepoint.searchSites("test")).rejects.toThrow(
        "Search failed"
      );
      expect(mockAuth.handleApiError).toHaveBeenCalledWith(mockError);
    });
  });

  describe("getSiteByPath", () => {
    beforeEach(() => {
      sharepoint = new SharePoint(mockAuth);
    });

    it("should get site by hostname and path", async () => {
      const mockSite = {
        id: "site-123",
        displayName: "Engineering",
        name: "engineering",
        webUrl: "https://contoso.sharepoint.com/sites/engineering",
        description: "Engineering team",
      };

      const mockAxonInstance = {
        bearer: vi.fn().mockReturnThis(),
        get: vi.fn().mockResolvedValue({
          data: mockSite,
        }),
      };

      (Axon.new as any).mockReturnValue(mockAxonInstance);

      const result = await sharepoint.getSiteByPath(
        "contoso.sharepoint.com",
        "/sites/engineering"
      );

      expect(mockAxonInstance.get).toHaveBeenCalledWith(
        "https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/sites/engineering"
      );
      expect(result).toEqual(mockSite);
    });
  });

  describe("getLists", () => {
    beforeEach(() => {
      sharepoint = new SharePoint(mockAuth, "site-123");
    });

    it("should get lists from site", async () => {
      const mockLists = [
        {
          id: "list-1",
          displayName: "Tasks",
          name: "tasks",
          description: "Task list",
          webUrl: "https://contoso.sharepoint.com/sites/site/Lists/Tasks",
        },
        {
          id: "list-2",
          displayName: "Documents",
          name: "documents",
          description: "Document library",
          webUrl: "https://contoso.sharepoint.com/sites/site/Documents",
        },
      ];

      const mockAxonInstance = {
        bearer: vi.fn().mockReturnThis(),
        get: vi.fn().mockResolvedValue({
          data: { value: mockLists },
        }),
      };

      (Axon.new as any).mockReturnValue(mockAxonInstance);

      const result = await sharepoint.getLists();

      expect(mockAxonInstance.get).toHaveBeenCalledWith(
        "https://graph.microsoft.com/v1.0/sites/site-123/lists"
      );
      expect(result).toEqual(mockLists);
    });

    it("should get lists with explicit siteId parameter", async () => {
      const mockAxonInstance = {
        bearer: vi.fn().mockReturnThis(),
        get: vi.fn().mockResolvedValue({
          data: { value: [] },
        }),
      };

      (Axon.new as any).mockReturnValue(mockAxonInstance);

      await sharepoint.getLists("different-site");

      expect(mockAxonInstance.get).toHaveBeenCalledWith(
        "https://graph.microsoft.com/v1.0/sites/different-site/lists"
      );
    });

    it("should throw error when siteId not provided", async () => {
      const sp = new SharePoint(mockAuth);

      await expect(sp.getLists()).rejects.toThrow(
        "Site ID is required. Provide it in constructor, setSiteId(), or as parameter."
      );
    });
  });

  describe("getList", () => {
    beforeEach(() => {
      sharepoint = new SharePoint(mockAuth, "site-123");
    });

    it("should get specific list", async () => {
      const mockList = {
        id: "list-1",
        displayName: "Tasks",
        name: "tasks",
      };

      const mockAxonInstance = {
        bearer: vi.fn().mockReturnThis(),
        get: vi.fn().mockResolvedValue({
          data: mockList,
        }),
      };

      (Axon.new as any).mockReturnValue(mockAxonInstance);

      const result = await sharepoint.getList("Tasks");

      expect(mockAxonInstance.get).toHaveBeenCalledWith(
        "https://graph.microsoft.com/v1.0/sites/site-123/lists/Tasks"
      );
      expect(result).toEqual(mockList);
    });

    it("should throw error when siteId not provided", async () => {
      const sp = new SharePoint(mockAuth);

      await expect(sp.getList("Tasks")).rejects.toThrow(
        "Site ID is required"
      );
    });
  });

  describe("getListItems", () => {
    beforeEach(() => {
      sharepoint = new SharePoint(mockAuth, "site-123");
    });

    it("should get list items without options", async () => {
      const mockItems = [
        { id: "1", fields: { Title: "Item 1" } },
        { id: "2", fields: { Title: "Item 2" } },
      ];

      const mockAxonInstance = {
        bearer: vi.fn().mockReturnThis(),
        params: vi.fn().mockReturnThis(),
        get: vi.fn().mockResolvedValue({
          data: { value: mockItems },
        }),
      };

      (Axon.new as any).mockReturnValue(mockAxonInstance);

      const result = await sharepoint.getListItems("Tasks");

      expect(mockAxonInstance.get).toHaveBeenCalledWith(
        "https://graph.microsoft.com/v1.0/sites/site-123/lists/Tasks/items"
      );
      expect(result).toEqual(mockItems);
    });

    it("should get list items with filter options", async () => {
      const mockAxonInstance = {
        bearer: vi.fn().mockReturnThis(),
        params: vi.fn().mockReturnThis(),
        get: vi.fn().mockResolvedValue({
          data: { value: [] },
        }),
      };

      (Axon.new as any).mockReturnValue(mockAxonInstance);

      await sharepoint.getListItems("Tasks", {
        filter: "fields/Status eq 'Active'",
        orderby: "createdDateTime desc",
        top: 10,
        expand: "fields",
      });

      expect(mockAxonInstance.params).toHaveBeenCalledWith({
        $filter: "fields/Status eq 'Active'",
        $orderby: "createdDateTime desc",
        $top: 10,
        $expand: "fields",
      });
    });

    it("should throw error when siteId not provided", async () => {
      const sp = new SharePoint(mockAuth);

      await expect(sp.getListItems("Tasks")).rejects.toThrow(
        "Site ID is required"
      );
    });
  });

  describe("getListItem", () => {
    beforeEach(() => {
      sharepoint = new SharePoint(mockAuth, "site-123");
    });

    it("should get specific list item", async () => {
      const mockItem = { id: "123", fields: { Title: "Task 1" } };

      const mockAxonInstance = {
        bearer: vi.fn().mockReturnThis(),
        params: vi.fn().mockReturnThis(),
        get: vi.fn().mockResolvedValue({
          data: mockItem,
        }),
      };

      (Axon.new as any).mockReturnValue(mockAxonInstance);

      const result = await sharepoint.getListItem("Tasks", "123");

      expect(mockAxonInstance.get).toHaveBeenCalledWith(
        "https://graph.microsoft.com/v1.0/sites/site-123/lists/Tasks/items/123"
      );
      expect(result).toEqual(mockItem);
    });

    it("should get list item with expand option", async () => {
      const mockAxonInstance = {
        bearer: vi.fn().mockReturnThis(),
        params: vi.fn().mockReturnThis(),
        get: vi.fn().mockResolvedValue({
          data: {},
        }),
      };

      (Axon.new as any).mockReturnValue(mockAxonInstance);

      await sharepoint.getListItem("Tasks", "123", "fields");

      expect(mockAxonInstance.params).toHaveBeenCalledWith({
        $expand: "fields",
      });
    });
  });

  describe("createListItem", () => {
    beforeEach(() => {
      sharepoint = new SharePoint(mockAuth, "site-123");
    });

    it("should create list item", async () => {
      const fields = {
        Title: "New Task",
        Status: "Active",
        Priority: "High",
      };

      const mockCreated = { id: "456", fields };

      const mockAxonInstance = {
        bearer: vi.fn().mockReturnThis(),
        post: vi.fn().mockResolvedValue({
          data: mockCreated,
        }),
      };

      (Axon.new as any).mockReturnValue(mockAxonInstance);

      const result = await sharepoint.createListItem("Tasks", fields);

      expect(mockAxonInstance.post).toHaveBeenCalledWith(
        "https://graph.microsoft.com/v1.0/sites/site-123/lists/Tasks/items",
        { fields }
      );
      expect(result).toEqual(mockCreated);
    });

    it("should throw error when siteId not provided", async () => {
      const sp = new SharePoint(mockAuth);

      await expect(
        sp.createListItem("Tasks", { Title: "Test" })
      ).rejects.toThrow("Site ID is required");
    });
  });

  describe("updateListItem", () => {
    beforeEach(() => {
      sharepoint = new SharePoint(mockAuth, "site-123");
    });

    it("should update list item", async () => {
      const fields = { Status: "Completed" };
      const mockUpdated = { id: "123", fields };

      const mockAxonInstance = {
        bearer: vi.fn().mockReturnThis(),
        patch: vi.fn().mockResolvedValue({
          data: mockUpdated,
        }),
      };

      (Axon.new as any).mockReturnValue(mockAxonInstance);

      const result = await sharepoint.updateListItem("Tasks", "123", fields);

      expect(mockAxonInstance.patch).toHaveBeenCalledWith(
        "https://graph.microsoft.com/v1.0/sites/site-123/lists/Tasks/items/123",
        { fields }
      );
      expect(result).toEqual(mockUpdated);
    });
  });

  describe("deleteListItem", () => {
    beforeEach(() => {
      sharepoint = new SharePoint(mockAuth, "site-123");
    });

    it("should delete list item", async () => {
      const mockAxonInstance = {
        bearer: vi.fn().mockReturnThis(),
        delete: vi.fn().mockResolvedValue({}),
      };

      (Axon.new as any).mockReturnValue(mockAxonInstance);

      await sharepoint.deleteListItem("Tasks", "123");

      expect(mockAxonInstance.delete).toHaveBeenCalledWith(
        "https://graph.microsoft.com/v1.0/sites/site-123/lists/Tasks/items/123"
      );
    });
  });

  describe("deleteListItems", () => {
    beforeEach(() => {
      sharepoint = new SharePoint(mockAuth, "site-123");
    });

    it("should delete multiple list items", async () => {
      const mockAxonInstance = {
        bearer: vi.fn().mockReturnThis(),
        delete: vi.fn().mockResolvedValue({}),
      };

      (Axon.new as any).mockReturnValue(mockAxonInstance);

      await sharepoint.deleteListItems("Tasks", ["1", "2", "3"]);

      expect(mockAxonInstance.delete).toHaveBeenCalledTimes(3);
    });

    it("should handle deletion errors", async () => {
      const mockError = new Error("Delete failed");
      const mockAxonInstance = {
        bearer: vi.fn().mockReturnThis(),
        delete: vi.fn().mockRejectedValue(mockError),
      };

      (Axon.new as any).mockReturnValue(mockAxonInstance);

      await expect(
        sharepoint.deleteListItems("Tasks", ["1"])
      ).rejects.toThrow("Delete failed");
    });
  });

  describe("queryAndProcess", () => {
    beforeEach(() => {
      sharepoint = new SharePoint(mockAuth, "site-123");
    });

    it("should query and process items without deletion", async () => {
      const mockItems = [
        { id: "1", fields: { Title: "Task 1", Status: "Pending" } },
        { id: "2", fields: { Title: "Task 2", Status: "Pending" } },
      ];

      const mockAxonInstance = {
        bearer: vi.fn().mockReturnThis(),
        params: vi.fn().mockReturnThis(),
        get: vi.fn().mockResolvedValue({
          data: { value: mockItems },
        }),
      };

      (Axon.new as any).mockReturnValue(mockAxonInstance);

      const processor = (item: any) => ({
        id: item.id,
        title: item.fields.Title,
      });

      const result = await sharepoint.queryAndProcess(
        "Tasks",
        "fields/Status eq 'Pending'",
        processor,
        false
      );

      expect(result).toEqual([
        { id: "1", title: "Task 1" },
        { id: "2", title: "Task 2" },
      ]);
    });

    it("should query, process, and delete items", async () => {
      const mockItems = [
        { id: "1", fields: { Title: "Task 1" } },
        { id: "2", fields: { Title: "Task 2" } },
      ];

      const mockAxonInstance = {
        bearer: vi.fn().mockReturnThis(),
        params: vi.fn().mockReturnThis(),
        get: vi.fn().mockResolvedValue({
          data: { value: mockItems },
        }),
        delete: vi.fn().mockResolvedValue({}),
      };

      (Axon.new as any).mockReturnValue(mockAxonInstance);

      const processor = (item: any) => item.fields.Title;

      await sharepoint.queryAndProcess(
        "Tasks",
        "fields/Status eq 'Done'",
        processor,
        true
      );

      expect(mockAxonInstance.delete).toHaveBeenCalledTimes(2);
    });

    it("should return empty array when no items found", async () => {
      const mockAxonInstance = {
        bearer: vi.fn().mockReturnThis(),
        params: vi.fn().mockReturnThis(),
        get: vi.fn().mockResolvedValue({
          data: { value: [] },
        }),
      };

      (Axon.new as any).mockReturnValue(mockAxonInstance);

      const result = await sharepoint.queryAndProcess(
        "Tasks",
        "fields/Status eq 'Active'",
        (item) => item,
        false
      );

      expect(result).toEqual([]);
    });

    it("should return empty array on error", async () => {
      const mockError = new Error("Query failed");
      const mockAxonInstance = {
        bearer: vi.fn().mockReturnThis(),
        params: vi.fn().mockReturnThis(),
        get: vi.fn().mockRejectedValue(mockError),
      };

      (Axon.new as any).mockReturnValue(mockAxonInstance);

      // Mock handleApiError to not throw (just log)
      vi.spyOn(mockAuth, "handleApiError").mockImplementation(() => {
        // Don't throw, just return
      });

      const result = await sharepoint.queryAndProcess(
        "Tasks",
        "filter",
        (item) => item,
        false
      );

      expect(result).toEqual([]);
    });
  });

  describe("getLatestItem", () => {
    beforeEach(() => {
      sharepoint = new SharePoint(mockAuth, "site-123");
    });

    it("should get latest item with default ordering", async () => {
      const mockItems = [
        { id: "1", fields: { Title: "Latest Task" } },
      ];

      const mockAxonInstance = {
        bearer: vi.fn().mockReturnThis(),
        params: vi.fn().mockReturnThis(),
        get: vi.fn().mockResolvedValue({
          data: { value: mockItems },
        }),
      };

      (Axon.new as any).mockReturnValue(mockAxonInstance);

      const result = await sharepoint.getLatestItem("Tasks");

      expect(mockAxonInstance.params).toHaveBeenCalledWith({
        $orderby: "createdDateTime desc",
        $top: 1,
        $expand: "fields",
      });
      expect(result).toEqual(mockItems[0]);
    });

    it("should get latest item with custom ordering", async () => {
      const mockAxonInstance = {
        bearer: vi.fn().mockReturnThis(),
        params: vi.fn().mockReturnThis(),
        get: vi.fn().mockResolvedValue({
          data: { value: [] },
        }),
      };

      (Axon.new as any).mockReturnValue(mockAxonInstance);

      await sharepoint.getLatestItem("Tasks", "modified desc");

      expect(mockAxonInstance.params).toHaveBeenCalledWith({
        $orderby: "modified desc",
        $top: 1,
        $expand: "fields",
      });
    });

    it("should get latest item with filter", async () => {
      const mockAxonInstance = {
        bearer: vi.fn().mockReturnThis(),
        params: vi.fn().mockReturnThis(),
        get: vi.fn().mockResolvedValue({
          data: { value: [] },
        }),
      };

      (Axon.new as any).mockReturnValue(mockAxonInstance);

      await sharepoint.getLatestItem(
        "Tasks",
        "createdDateTime desc",
        "fields/Status eq 'Active'"
      );

      expect(mockAxonInstance.params).toHaveBeenCalledWith({
        $filter: "fields/Status eq 'Active'",
        $orderby: "createdDateTime desc",
        $top: 1,
        $expand: "fields",
      });
    });

    it("should return undefined when no items found", async () => {
      const mockAxonInstance = {
        bearer: vi.fn().mockReturnThis(),
        params: vi.fn().mockReturnThis(),
        get: vi.fn().mockResolvedValue({
          data: { value: [] },
        }),
      };

      (Axon.new as any).mockReturnValue(mockAxonInstance);

      const result = await sharepoint.getLatestItem("Tasks");

      expect(result).toBeUndefined();
    });
  });

  describe("getListColumns", () => {
    beforeEach(() => {
      sharepoint = new SharePoint(mockAuth, "site-123");
    });

    it("should get list columns", async () => {
      const mockColumns = [
        {
          id: "col-1",
          name: "Title",
          displayName: "Title",
          columnGroup: "Custom Columns",
          description: "Title field",
          hidden: false,
          readOnly: false,
        },
        {
          id: "col-2",
          name: "Status",
          displayName: "Status",
          columnGroup: "Custom Columns",
          description: "Status field",
          hidden: false,
          readOnly: false,
        },
      ];

      const mockAxonInstance = {
        bearer: vi.fn().mockReturnThis(),
        get: vi.fn().mockResolvedValue({
          data: { value: mockColumns },
        }),
      };

      (Axon.new as any).mockReturnValue(mockAxonInstance);

      const result = await sharepoint.getListColumns("Tasks");

      expect(mockAxonInstance.get).toHaveBeenCalledWith(
        "https://graph.microsoft.com/v1.0/sites/site-123/lists/Tasks/columns"
      );
      expect(result).toEqual(mockColumns);
    });

    it("should throw error when siteId not provided", async () => {
      const sp = new SharePoint(mockAuth);

      await expect(sp.getListColumns("Tasks")).rejects.toThrow(
        "Site ID is required"
      );
    });
  });
});
