import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import type { GraphService } from "../services/graph.js";
import type {
  GraphApiResponse,
  MailFolder,
  MailFolderSummary,
  MailMessageDetail,
  MailMessageSummary,
  Message,
} from "../types/graph.js";
import { formatMessageContent } from "../utils/html-to-markdown.js";

/**
 * Registers all mail-related MCP tools (read-only: requires Mail.Read).
 * Tools: list_mail_messages, get_mail_message, list_mail_folders.
 */
export function registerMailTools(
  server: McpServer,
  graphService: GraphService,
  _readOnly: boolean
) {
  server.registerTool(
    "list_mail_messages",
    {
      title: "List Mail Messages",
      description:
        "List recent messages from a mail folder (default: Inbox). Supports filtering by unread status, sender, and date range, or free-text keyword search.",
      inputSchema: {
        folderId: z
          .string()
          .optional()
          .default("inbox")
          .describe(
            "Mail folder ID or well-known folder name (e.g. 'inbox', 'sentitems'). Use list_mail_folders to find custom folder IDs."
          ),
        search: z
          .string()
          .optional()
          .describe(
            "Free-text keyword search (subject/body/sender). When provided, other filters (unreadOnly, from, since, until) are ignored — Graph API doesn't support combining $search with $filter."
          ),
        unreadOnly: z.boolean().optional().default(false).describe("Only return unread messages"),
        from: z.string().optional().describe("Filter by sender email address"),
        since: z.string().optional().describe("Get messages received since this ISO datetime"),
        until: z.string().optional().describe("Get messages received until this ISO datetime"),
        limit: z
          .number()
          .min(1)
          .max(100)
          .optional()
          .default(20)
          .describe("Number of messages to retrieve (default: 20, max: 100)"),
      },
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    async ({ folderId, search, unreadOnly, from, since, until, limit }) => {
      try {
        const client = await graphService.getClient();

        const queryParams: string[] = [`$top=${limit}`];
        queryParams.push(
          "$select=id,subject,from,receivedDateTime,bodyPreview,isRead,hasAttachments,importance,webLink"
        );

        if (search) {
          queryParams.push(`$search="${search.replace(/"/g, '\\"')}"`);
        } else {
          const filters: string[] = [];
          if (unreadOnly) filters.push("isRead eq false");
          if (from) filters.push(`from/emailAddress/address eq '${from}'`);
          if (since) filters.push(`receivedDateTime ge ${new Date(since).toISOString()}`);
          if (until) filters.push(`receivedDateTime le ${new Date(until).toISOString()}`);

          if (filters.length > 0) {
            queryParams.push(`$filter=${filters.join(" and ")}`);
          }
          queryParams.push("$orderby=receivedDateTime desc");
        }

        const response = (await client
          .api(`/me/mailFolders/${folderId}/messages?${queryParams.join("&")}`)
          .get()) as GraphApiResponse<Message>;

        if (!response?.value?.length) {
          return {
            content: [{ type: "text", text: "No messages found matching the specified filters." }],
          };
        }

        const messages: MailMessageSummary[] = response.value.map((message: Message) => ({
          id: message.id,
          subject: message.subject,
          from: message.from?.emailAddress?.name || message.from?.emailAddress?.address,
          receivedDateTime: message.receivedDateTime,
          bodyPreview: message.bodyPreview,
          isRead: message.isRead,
          hasAttachments: message.hasAttachments,
          importance: message.importance,
          webLink: message.webLink,
        }));

        return {
          content: [{ type: "text", text: JSON.stringify(messages, null, 2) }],
        };
      } catch (error: unknown) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return { content: [{ type: "text", text: `❌ Error: ${errorMessage}` }] };
      }
    }
  );

  server.registerTool(
    "get_mail_message",
    {
      title: "Get Mail Message",
      description:
        "Retrieve the full content of a specific email message, including body, sender, and recipients.",
      inputSchema: {
        messageId: z.string().describe("Message ID (from list_mail_messages)"),
        contentFormat: z
          .enum(["raw", "markdown"])
          .optional()
          .default("markdown")
          .describe(
            'Format for the message body. "markdown" (default) converts HTML to clean Markdown. "raw" returns original HTML.'
          ),
      },
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    async ({ messageId, contentFormat }) => {
      try {
        const client = await graphService.getClient();
        const message = (await client.api(`/me/messages/${messageId}`).get()) as Message;

        const detail: MailMessageDetail = {
          id: message.id,
          subject: message.subject,
          from: message.from?.emailAddress?.name || message.from?.emailAddress?.address,
          toRecipients: message.toRecipients
            ?.map((r) => r.emailAddress?.name || r.emailAddress?.address)
            .filter((r): r is string => !!r),
          receivedDateTime: message.receivedDateTime,
          importance: message.importance,
          hasAttachments: message.hasAttachments,
          body: formatMessageContent(message.body?.content, contentFormat),
          webLink: message.webLink,
        };

        return {
          content: [{ type: "text", text: JSON.stringify(detail, null, 2) }],
        };
      } catch (error: unknown) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return { content: [{ type: "text", text: `❌ Error: ${errorMessage}` }] };
      }
    }
  );

  server.registerTool(
    "list_mail_folders",
    {
      title: "List Mail Folders",
      description:
        "List mail folders (e.g. Inbox, custom rule folders). Use this to find folder IDs for list_mail_messages.",
      inputSchema: {
        parentFolderId: z
          .string()
          .optional()
          .describe("List child folders of this folder ID. Omit to list top-level folders."),
      },
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    async ({ parentFolderId }) => {
      try {
        const client = await graphService.getClient();
        const endpoint = parentFolderId
          ? `/me/mailFolders/${parentFolderId}/childFolders?$top=100`
          : "/me/mailFolders?$top=100";

        const response = (await client.api(endpoint).get()) as GraphApiResponse<MailFolder>;

        if (!response?.value?.length) {
          return { content: [{ type: "text", text: "No mail folders found." }] };
        }

        const folders: MailFolderSummary[] = response.value.map((folder: MailFolder) => ({
          id: folder.id,
          displayName: folder.displayName,
          unreadItemCount: folder.unreadItemCount,
          totalItemCount: folder.totalItemCount,
          childFolderCount: folder.childFolderCount,
        }));

        return {
          content: [{ type: "text", text: JSON.stringify(folders, null, 2) }],
        };
      } catch (error: unknown) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return { content: [{ type: "text", text: `❌ Error: ${errorMessage}` }] };
      }
    }
  );
}
