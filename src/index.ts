#!/usr/bin/env node
import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import {
  CallToolRequestSchema,
  ErrorCode,
  ListToolsRequestSchema,
  McpError,
} from "@modelcontextprotocol/sdk/types.js";
import { google } from "googleapis";
import { readFileSync, appendFileSync } from "fs";
import { fileURLToPath } from "url";
import { dirname, join } from "path";

function log(message: string): void {
  const timestamp = new Date().toISOString();
  const logMessage = `[${timestamp}] ${message}`;
  console.log(logMessage);

  try {
      appendFileSync("mcp-server.log", logMessage + "\n");
  } catch (error) {
      console.error("Failed to write to log file:", (error as Error).message);
  }
}

// Environment variables required for OAuth
// Load .env file
const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);
const envPath = join(__dirname, "..", ".env");

try {
  const envContent = readFileSync(envPath, "utf8");
  const envVars: Record<string, string> = {};

  envContent.split("\n").forEach((line: string) => {
    line = line.trim();
    if (line && !line.startsWith("#")) {
      const [key, ...valueParts] = line.split("=");
      if (key && valueParts.length > 0) {
          let value = valueParts.join("=");
          // Remove quotes and trailing comma
          value = value.replace(/^["']|["'],?$/g, "");
          envVars[key] = value;
          process.env[key] = value;
      }
    }
  });

    console.log("✅ Loaded environment variables from .env file");
} catch (error) {
    console.log("⚠️  Could not load .env file, using system environment variables");
}

const CLIENT_ID = process.env.GOOGLE_CLIENT_ID;
const CLIENT_SECRET = process.env.GOOGLE_CLIENT_SECRET;
const REFRESH_TOKEN = process.env.GOOGLE_REFRESH_TOKEN;
if (!CLIENT_ID || !CLIENT_SECRET || !REFRESH_TOKEN) {
  throw new Error(
    "Required Google OAuth credentials not found in environment variables"
  );
}

class GoogleWorkspaceServer {
  private server: Server;
  private auth: any;
  private gmail: any;
  private calendar: any;

  constructor() {
    this.server = new Server(
      {
        name: "google-workspace-server",
        version: "0.1.0",
      },
      {
        capabilities: {
          tools: {},
        },
      }
    );

    // Set up OAuth2 client
    this.auth = new google.auth.OAuth2(CLIENT_ID, CLIENT_SECRET);
    this.auth.setCredentials({ refresh_token: REFRESH_TOKEN });

    // Initialize API clients
    this.gmail = google.gmail({ version: "v1", auth: this.auth });
    this.calendar = google.calendar({ version: "v3", auth: this.auth });

    this.setupToolHandlers();

    // Error handling
    this.server.onerror = (error) => console.error("[MCP Error]", error);
    process.on("SIGINT", async () => {
      await this.server.close();
      process.exit(0);
    });
  }

  private setupToolHandlers() {
    this.server.setRequestHandler(ListToolsRequestSchema, async () => ({
      tools: [
        {
          name: "get_current_datetime",
          description: "Get current date and time in Korean format",
          inputSchema: {
              type: "object",
              properties: {},
          },
        },
        {
          name: "list_emails",
          description: "List recent emails from Gmail inbox",
          inputSchema: {
            type: "object",
            properties: {
              maxResults: {
                type: "number",
                description: "Maximum number of emails to return (default: 10)",
              },
              query: {
                type: "string",
                description: "Search query to filter emails",
              },
            },
          },
        },
        {
          name: "search_emails",
          description: "Search emails with advanced query",
          inputSchema: {
            type: "object",
            properties: {
              query: {
                type: "string",
                description:
                  'Gmail search query (e.g., "from:example@gmail.com has:attachment"). Examples:\n' +
                  '- "from:alice@example.com" (Emails from Alice)\n' +
                  '- "to:bob@example.com" (Emails sent to Bob)\n' +
                  '- "subject:Meeting Update" (Emails with "Meeting Update" in the subject)\n' +
                  '- "has:attachment filename:pdf" (Emails with PDF attachments)\n' +
                  '- "after:2024/01/01 before:2024/02/01" (Emails between specific dates)\n' +
                  '- "is:unread" (Unread emails)\n' +
                  '- "from:@company.com has:attachment" (Emails from a company domain with attachments)',
              },
              maxResults: {
                type: "number",
                description: "Maximum number of emails to return (default: 10)",
              },
            },
            required: ["query"],
          },
        },
        {
          name: "send_email",
          description: "Send a new email",
          inputSchema: {
            type: "object",
            properties: {
              to: {
                type: "string",
                description: "Recipient email address",
              },
              subject: {
                type: "string",
                description: "Email subject",
              },
              body: {
                type: "string",
                description: "Email body (can include HTML)",
              },
              cc: {
                type: "string",
                description: "CC recipients (comma-separated)",
              },
              bcc: {
                type: "string",
                description: "BCC recipients (comma-separated)",
              },
            },
            required: ["to", "subject", "body"],
          },
        },
        {
          name: "modify_email",
          description: "Modify email labels (archive, trash, mark read/unread)",
          inputSchema: {
            type: "object",
            properties: {
              id: {
                type: "string",
                description: "Email ID",
              },
              addLabels: {
                type: "array",
                items: { type: "string" },
                description: "Labels to add",
              },
              removeLabels: {
                type: "array",
                items: { type: "string" },
                description: "Labels to remove",
              },
            },
            required: ["id"],
          },
        },
        {
          name: "list_events",
          description: "List calendar events from Google Calendar. You can fetch events from the past, future, or any specific date range by using the optional parameters 'timeMin' and 'timeMax'. Both accept ISO 8601 datetime strings (e.g., '2025-08-13T14:00:00+09:00'). If omitted, the default range starts from now.",
          inputSchema: {
            type: "object",
            properties: {
              maxResults: {
                type: "number",
                description: "Maximum number of events to return (default: 30)",
              },
              timeMin: {
                type: "string",
                description: "Start time in ISO format. Can be in the past.",
              },
              timeMax: {
                type: "string",
                description: "End time in ISO format. Can be in the future or past.",
              },
            },
          },
        },
        {
          name: "create_event",
          description: "Create a new calendar event",
          inputSchema: {
            type: "object",
            properties: {
              summary: {
                type: "string",
                description: "Event title",
              },
              location: {
                type: "string",
                description: "Event location",
              },
              description: {
                type: "string",
                description: "Event description",
              },
              start: {
                type: "string",
                description: "Start time in ISO format",
              },
              end: {
                type: "string",
                description: "End time in ISO format",
              },
              attendees: {
                type: "array",
                items: { type: "string" },
                description: "List of attendee email addresses",
              },
            },
            required: ["summary", "start", "end"],
          },
        },
        {
          name: "update_event",
          description: "Update an existing calendar event",
          inputSchema: {
            type: "object",
            properties: {
              eventId: {
                type: "string",
                description: "Event ID to update",
              },
              summary: {
                type: "string",
                description: "New event title",
              },
              location: {
                type: "string",
                description: "New event location",
              },
              description: {
                type: "string",
                description: "New event description",
              },
              start: {
                type: "string",
                description: "New start time in ISO format",
              },
              end: {
                type: "string",
                description: "New end time in ISO format",
              },
              attendees: {
                type: "array",
                items: { type: "string" },
                description: "New list of attendee email addresses",
              },
            },
            required: ["eventId"],
          },
        },
        {
          name: "delete_event",
          description: "Delete a calendar event",
          inputSchema: {
            type: "object",
            properties: {
              eventId: {
                type: "string",
                description: "Event ID to delete",
              },
            },
            required: ["eventId"],
          },
        },
      ],
    }));

    this.server.setRequestHandler(CallToolRequestSchema, async (request) => {
      const args = request.params.arguments || {};

      switch (request.params.name) {
        case "get_current_datetime":
          return await this.handleGetCurrentDateTime();
        case "list_emails":
          return await this.handleListEmails(args);
        case "search_emails":
          return await this.handleSearchEmails(args);
        case "send_email":
          return await this.handleSendEmail(args);
        case "modify_email":
          return await this.handleModifyEmail(args);
        case "list_events":
          return await this.handleListEvents(args);
        case "create_event":
          return await this.handleCreateEvent(args);
        case "update_event":
          return await this.handleUpdateEvent(args);
        case "delete_event":
          return await this.handleDeleteEvent(args);
        default:
          throw new McpError(ErrorCode.MethodNotFound, `Unknown tool: ${request.params.name}`);
      }
    });
  }

  private async handleGetCurrentDateTime() {
    const now = new Date();
    const year = now.getFullYear();
    const month = now.getMonth() + 1;
    const day = now.getDate();
    const hours = now.getHours().toString().padStart(2, '0');
    const minutes = now.getMinutes().toString().padStart(2, '0');
    const seconds = now.getSeconds().toString().padStart(2, '0');

    const weekdays = ['일요일', '월요일', '화요일', '수요일', '목요일', '금요일', '토요일'];
    const weekday = weekdays[now.getDay()];

    const currentDateTime = `${year}년 ${month}월 ${day}일 ${weekday} ${hours}:${minutes}:${seconds}`;

    log(`📅 get_current_datetime called: ${currentDateTime}`);

    return {
      content: [
        {
          type: "text",
          text: currentDateTime,
        },
      ],
    };
  }

  private async handleListEmails(args: any) {
    try {
      const maxResults = args?.maxResults || 10;
      const query = args?.query || "";
      const getEmailBody = (payload: any): string => {
        if (!payload) return "";
        if (payload.body && payload.body.data) {
          return Buffer.from(payload.body.data, "base64").toString("utf-8");
        }
        if (payload.parts && payload.parts.length > 0) {
          for (const part of payload.parts) {
            if (part.mimeType === "text/plain") {
              return Buffer.from(part.body.data, "base64").toString("utf-8");
            }
          }
        }
        return "(No body content)";
      };
      const response = await this.gmail.users.messages.list({
        userId: "me",
        maxResults,
        q: query,
      });
      const messages = response.data.messages || [];
      const emailDetails = await Promise.all(
        messages.map(async (msg: any) => {
          const detail = await this.gmail.users.messages.get({
            userId: "me",
            id: msg.id,
          });
          const headers = detail.data.payload?.headers;
          const subject =
            headers?.find((h: any) => h.name === "Subject")?.value || "";
          const from =
            headers?.find((h: any) => h.name === "From")?.value || "";
          const date =
            headers?.find((h: any) => h.name === "Date")?.value || "";
          const body = getEmailBody(detail.data.payload);
          return {
            id: msg.id,
            subject,
            from,
            date,
            body,
          };
        })
      );
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(emailDetails, null, 2),
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error fetching emails: ${(error as Error).message}`,
          },
        ],
        isError: true,
      };
    }
  }

  private async handleSearchEmails(args: any) {
    try {
      const maxResults = args?.maxResults || 10;
      const query = args?.query || "";
      const getEmailBody = (payload: any): string => {
        if (!payload) return "";
        if (payload.body && payload.body.data) {
          return Buffer.from(payload.body.data, "base64").toString("utf-8");
        }
        if (payload.parts && payload.parts.length > 0) {
          for (const part of payload.parts) {
            if (part.mimeType === "text/plain") {
              return Buffer.from(part.body.data, "base64").toString("utf-8");
            }
          }
        }
        return "(No body content)";
      };
      const response = await this.gmail.users.messages.list({
        userId: "me",
        maxResults,
        q: query,
      });
      const messages = response.data.messages || [];
      const emailDetails = await Promise.all(
        messages.map(async (msg: any) => {
          const detail = await this.gmail.users.messages.get({
            userId: "me",
            id: msg.id,
          });
          const headers = detail.data.payload?.headers;
          const subject =
            headers?.find((h: any) => h.name === "Subject")?.value || "";
          const from =
            headers?.find((h: any) => h.name === "From")?.value || "";
          const date =
            headers?.find((h: any) => h.name === "Date")?.value || "";
          const body = getEmailBody(detail.data.payload);
          // Use helper function to extract the email body correctly
          return {
            id: msg.id,
            subject,
            from,
            date,
            body,
            labels: detail.data.labelIds || [],
          };
        })
      );
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(emailDetails, null, 2),
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error fetching emails: ${(error as Error).message}`,
          },
        ],
        isError: true,
      };
    }
  }

  private async handleSendEmail(args: any) {
    try {
      const { to, subject, body, cc, bcc } = args;

      const headers = [
        'Content-Type: text/plain; charset="UTF-8"',
        "MIME-Version: 1.0",
        `To: ${to}`,
        cc ? `Cc: ${cc}` : null,
        bcc ? `Bcc: ${bcc}` : null,
        `Subject: ${subject}`,
      ]
        .filter(Boolean)
        .join("\r\n");

      // Ensure proper separation between headers and body
      const email = `${headers}\r\n\r\n${body}`;

      // Encode in base64url
      const encodedMessage = Buffer.from(email)
        .toString("base64")
        .replace(/\+/g, "-")
        .replace(/\//g, "_")
        .replace(/=+$/, "");

      // Send the email
      const response = await this.gmail.users.messages.send({
        userId: "me",
        requestBody: {
          raw: encodedMessage,
        },
      });

      return {
        content: [
          {
            type: "text",
            text: `Email sent successfully. Message ID: ${response.data.id}`,
          },
        ],
      };
    } catch (error: any) {
      return {
        content: [
          {
            type: "text",
            text: `Error sending email: ${error.message}`,
          },
        ],
        isError: true,
      };
    }
  }

  private async handleModifyEmail(args: any) {
    try {
      const { id, addLabels = [], removeLabels = [] } = args;

      const response = await this.gmail.users.messages.modify({
        userId: "me",
        id,
        requestBody: {
          addLabelIds: addLabels,
          removeLabelIds: removeLabels,
        },
      });

      return {
        content: [
          {
            type: "text",
            text: `Email modified successfully. Updated labels for message ID: ${response.data.id}`,
          },
        ],
      };
    } catch (error: any) {
      return {
        content: [
          {
            type: "text",
            text: `Error modifying email: ${error.message}`,
          },
        ],
        isError: true,
      };
    }
  }

  private async handleCreateEvent(args: any) {
    try {
      const raw = args;
      const parsedArgs = typeof raw === "string" ? JSON.parse(raw) : raw;

      const {
        summary,
        location,
        description,
        start,
        end,
        attendees = [],
      } = parsedArgs;

      const event = {
        summary,
        location,
        description,
        start: {
          dateTime: start,
          timeZone: Intl.DateTimeFormat().resolvedOptions().timeZone,
        },
        end: {
          dateTime: end,
          timeZone: Intl.DateTimeFormat().resolvedOptions().timeZone,
        },
        attendees: attendees.map((email: string) => ({ email })),
      };

      const response = await this.calendar.events.insert({
        calendarId: "primary",
        requestBody: event,
      });

      return {
        content: [
          {
            type: "text",
            text: `Event created successfully. Event ID: ${response.data.id}`,
          },
        ],
      };
    } catch (error: any) {
      return {
        content: [
          {
            type: "text",
            text: `Error creating event: ${error.message}`,
          },
        ],
        isError: true,
      };
    }
  }

  private async handleUpdateEvent(args: any) {
    try {
      const raw = args;
      const parsedArgs = typeof raw === "string" ? JSON.parse(raw) : raw;

      const { eventId, summary, location, description, start, end, attendees } =
      parsedArgs;

      const event: any = {};
      if (summary) event.summary = summary;
      if (location) event.location = location;
      if (description) event.description = description;
      if (start) {
        event.start = {
          dateTime: start,
          timeZone: Intl.DateTimeFormat().resolvedOptions().timeZone,
        };
      }
      if (end) {
        event.end = {
          dateTime: end,
          timeZone: Intl.DateTimeFormat().resolvedOptions().timeZone,
        };
      }
      if (attendees) {
        event.attendees = attendees.map((email: string) => ({ email }));
      }

      const response = await this.calendar.events.patch({
        calendarId: "primary",
        eventId,
        requestBody: event,
      });

      return {
        content: [
          {
            type: "text",
            text: `Event updated successfully. Event ID: ${response.data.id}`,
          },
        ],
      };
    } catch (error: any) {
      return {
        content: [
          {
            type: "text",
            text: `Error updating event: ${error.message}`,
          },
        ],
        isError: true,
      };
    }
  }

  private async handleDeleteEvent(args: any) {
    try {
      const raw = args;
      const parsedArgs = typeof raw === "string" ? JSON.parse(raw) : raw;
      const { eventId } = parsedArgs;

      await this.calendar.events.delete({
        calendarId: "primary",
        eventId,
      });

      return {
        content: [
          {
            type: "text",
            text: `Event deleted successfully. Event ID: ${eventId}`,
          },
        ],
      };
    } catch (error: any) {
      return {
        content: [
          {
            type: "text",
            text: `Error deleting event: ${error.message}`,
          },
        ],
        isError: true,
      };
    }
  }

  private async handleListEvents(args: any) {
    log(`📅 list_events called with args: ${JSON.stringify(args)}`);
    try {
      const raw = args;
      const parsedArgs = typeof raw === "string" ? JSON.parse(raw) : raw;

      const maxResults = parsedArgs?.maxResults || 10;
      const timeMin = parsedArgs?.timeMin || new Date().toISOString();
      const timeMax = parsedArgs?.timeMax;
      log(`📅 timeMin: ${timeMin}, timeMax: ${timeMax}`);

      const response = await this.calendar.events.list({
        calendarId: "primary",
        timeMin,
        timeMax,
        maxResults,
        singleEvents: true,
        orderBy: "startTime",
      });
      const events = response.data.items?.map((event: any) => ({
        id: event.id,
        summary: event.summary,
        start: event.start,
        end: event.end,
        location: event.location,
      }));
      log(`📅 Found ${events?.length || 0} events`);
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(events, null, 2),
          },
        ],
      };
    } catch (error) {
      log(`❌ list_events error: ${(error as Error).message}`);
      return {
        content: [
          {
            type: "text",
            text: `Error fetching calendar events: ${(error as Error).message}`,
          },
        ],
        isError: true,
      };
    }
  }

  async run() {
    const transport = new StdioServerTransport();
    await this.server.connect(transport);
    console.error("Google Workspace MCP server running on stdio");
  }
}

const server = new GoogleWorkspaceServer();
server.run().catch(console.error);
