import type { Client } from "@microsoft/microsoft-graph-client";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { beforeEach, describe, expect, it, vi } from "vitest";
import type { GraphService } from "../../services/graph.js";
import { registerMailTools } from "../mail.js";

// Mock the Graph service
const mockGraphService = {
  getClient: vi.fn(),
} as unknown as GraphService;

// Mock the MCP server
const mockServer = {
  registerTool: vi.fn(),
} as unknown as McpServer;

// Mock client responses
const mockClient = {
  api: vi.fn(),
} as unknown as Client;

function getHandler(name: string): (args?: any) => Promise<any> {
  const call = vi.mocked(mockServer.registerTool).mock.calls.find(([n]) => n === name);
  return call?.[2] as unknown as (args: any) => Promise<any>;
}

describe("Mail Tools", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    mockGraphService.getClient = vi.fn().mockResolvedValue(mockClient);
  });

  describe("registerMailTools", () => {
    it("should register all mail tools", () => {
      registerMailTools(mockServer, mockGraphService, false);

      expect(mockServer.registerTool).toHaveBeenCalledTimes(3);
      expect(mockServer.registerTool).toHaveBeenCalledWith(
        "list_mail_messages",
        expect.objectContaining({ title: "List Mail Messages" }),
        expect.any(Function)
      );
      expect(mockServer.registerTool).toHaveBeenCalledWith(
        "get_mail_message",
        expect.objectContaining({ title: "Get Mail Message" }),
        expect.any(Function)
      );
      expect(mockServer.registerTool).toHaveBeenCalledWith(
        "list_mail_folders",
        expect.objectContaining({ title: "List Mail Folders" }),
        expect.any(Function)
      );
    });
  });

  describe("list_mail_messages", () => {
    let handler: (args?: any) => Promise<any>;

    beforeEach(() => {
      registerMailTools(mockServer, mockGraphService, false);
      handler = getHandler("list_mail_messages");
    });

    it("should return mail messages from the default inbox folder", async () => {
      const mockMessages = [
        {
          id: "msg1",
          subject: "Engineering weekly update",
          from: { emailAddress: { name: "Jane Doe", address: "jane@example.com" } },
          receivedDateTime: "2026-08-30T10:00:00Z",
          bodyPreview: "Here's what shipped this week...",
          isRead: false,
          hasAttachments: false,
          importance: "normal",
          webLink: "https://outlook.office.com/mail/msg1",
        },
      ];

      const mockApiChain = { get: vi.fn().mockResolvedValue({ value: mockMessages }) };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await handler({ folderId: "inbox", limit: 20 });

      expect(mockClient.api).toHaveBeenCalledWith(
        expect.stringContaining("/me/mailFolders/inbox/messages?")
      );
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed).toEqual([
        {
          id: "msg1",
          subject: "Engineering weekly update",
          from: "Jane Doe",
          receivedDateTime: "2026-08-30T10:00:00Z",
          bodyPreview: "Here's what shipped this week...",
          isRead: false,
          hasAttachments: false,
          importance: "normal",
          webLink: "https://outlook.office.com/mail/msg1",
        },
      ]);
    });

    it("should use $search and skip $filter/$orderby when search is provided", async () => {
      const mockApiChain = { get: vi.fn().mockResolvedValue({ value: [] }) };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      await handler({ folderId: "inbox", search: "tracklog", limit: 20 });

      const calledUrl = vi.mocked(mockClient.api).mock.calls[0][0] as string;
      expect(calledUrl).toContain('$search="tracklog"');
      expect(calledUrl).not.toContain("$filter");
      expect(calledUrl).not.toContain("$orderby");
    });

    it("should return a friendly message when no messages are found", async () => {
      const mockApiChain = { get: vi.fn().mockResolvedValue({ value: [] }) };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await handler({ folderId: "inbox", limit: 20 });

      expect(result.content[0].text).toBe("No messages found matching the specified filters.");
    });

    it("should handle API errors gracefully", async () => {
      mockClient.api = vi.fn().mockReturnValue({
        get: vi.fn().mockRejectedValue(new Error("API Error")),
      });

      const result = await handler({ folderId: "inbox", limit: 20 });

      expect(result.content[0].text).toBe("❌ Error: API Error");
    });
  });

  describe("get_mail_message", () => {
    let handler: (args?: any) => Promise<any>;

    beforeEach(() => {
      registerMailTools(mockServer, mockGraphService, false);
      handler = getHandler("get_mail_message");
    });

    it("should return full message details", async () => {
      const mockMessage = {
        id: "msg1",
        subject: "Engineering weekly update",
        from: { emailAddress: { name: "Jane Doe", address: "jane@example.com" } },
        toRecipients: [{ emailAddress: { name: "Rafael Patro", address: "rafael@example.com" } }],
        receivedDateTime: "2026-08-30T10:00:00Z",
        importance: "normal",
        hasAttachments: false,
        body: { content: "<p>Hello team</p>", contentType: "html" },
        webLink: "https://outlook.office.com/mail/msg1",
      };

      mockClient.api = vi.fn().mockReturnValue({ get: vi.fn().mockResolvedValue(mockMessage) });

      const result = await handler({ messageId: "msg1", contentFormat: "markdown" });

      expect(mockClient.api).toHaveBeenCalledWith("/me/messages/msg1");
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.subject).toBe("Engineering weekly update");
      expect(parsed.from).toBe("Jane Doe");
      expect(parsed.toRecipients).toEqual(["Rafael Patro"]);
      expect(parsed.body).toBe("Hello team");
    });

    it("should handle API errors gracefully", async () => {
      mockClient.api = vi.fn().mockReturnValue({
        get: vi.fn().mockRejectedValue(new Error("Not found")),
      });

      const result = await handler({ messageId: "bad-id", contentFormat: "markdown" });

      expect(result.content[0].text).toBe("❌ Error: Not found");
    });
  });

  describe("list_mail_folders", () => {
    let handler: (args?: any) => Promise<any>;

    beforeEach(() => {
      registerMailTools(mockServer, mockGraphService, false);
      handler = getHandler("list_mail_folders");
    });

    it("should list top-level folders by default", async () => {
      const mockFolders = [
        {
          id: "folder1",
          displayName: "Inbox",
          unreadItemCount: 5,
          totalItemCount: 100,
          childFolderCount: 2,
        },
      ];
      mockClient.api = vi
        .fn()
        .mockReturnValue({ get: vi.fn().mockResolvedValue({ value: mockFolders }) });

      const result = await handler({});

      expect(mockClient.api).toHaveBeenCalledWith("/me/mailFolders?$top=100");
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed).toEqual([
        {
          id: "folder1",
          displayName: "Inbox",
          unreadItemCount: 5,
          totalItemCount: 100,
          childFolderCount: 2,
        },
      ]);
    });

    it("should list child folders when parentFolderId is provided", async () => {
      mockClient.api = vi.fn().mockReturnValue({ get: vi.fn().mockResolvedValue({ value: [] }) });

      await handler({ parentFolderId: "folder1" });

      expect(mockClient.api).toHaveBeenCalledWith("/me/mailFolders/folder1/childFolders?$top=100");
    });
  });
});
