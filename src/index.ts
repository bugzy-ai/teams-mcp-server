#!/usr/bin/env node

/**
 * Teams MCP Server
 *
 * MCP server for posting messages to Microsoft Teams via Bot Connector API.
 * Used by the Bugzy agent to send notifications and responses to Teams channels.
 *
 * Environment variables (set by execution context):
 * - TEAMS_BOT_APP_ID: Bot application ID
 * - TEAMS_BOT_APP_PASSWORD: Bot application password
 * - TEAMS_SERVICE_URL: Bot Connector service URL (from conversation reference)
 * - TEAMS_CONVERSATION_ID: Conversation/channel ID (from conversation reference)
 * - TEAMS_THREAD_ID: Optional - for replying to specific thread
 */

import {
  ListToolsRequestSchema,
  CallToolRequestSchema,
} from '@modelcontextprotocol/sdk/types.js';
import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { zodToJsonSchema } from 'zod-to-json-schema';
import dotenv from 'dotenv';
import {
  PostMessageRequestSchema,
  PostRichMessageRequestSchema,
  type AdaptiveCard,
} from './schemas.js';

dotenv.config();

// Token cache for Bot Connector API
let tokenCache: { token: string; expiresAt: number } | null = null;

/**
 * Get an access token for the Bot Connector API
 */
async function getBotConnectorToken(): Promise<string> {
  const now = Date.now();

  // Check cache (with 5 minute buffer before expiry)
  if (tokenCache && tokenCache.expiresAt > now + 5 * 60 * 1000) {
    return tokenCache.token;
  }

  const appId = process.env.TEAMS_BOT_APP_ID;
  const appPassword = process.env.TEAMS_BOT_APP_PASSWORD;

  if (!appId || !appPassword) {
    throw new Error(
      'TEAMS_BOT_APP_ID and TEAMS_BOT_APP_PASSWORD must be set'
    );
  }

  // Request token from Microsoft identity platform
  const tokenUrl =
    'https://login.microsoftonline.com/botframework.com/oauth2/v2.0/token';

  const body = new URLSearchParams({
    grant_type: 'client_credentials',
    client_id: appId,
    client_secret: appPassword,
    scope: 'https://api.botframework.com/.default',
  });

  const response = await fetch(tokenUrl, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded',
    },
    body: body.toString(),
  });

  if (!response.ok) {
    const errorText = await response.text();
    console.error('[TeamsMCP] Token request failed:', errorText);
    throw new Error(`Failed to get Bot Connector token: ${response.status}`);
  }

  let data: { access_token: string; expires_in: number };
  try {
    data = (await response.json()) as { access_token: string; expires_in: number };
  } catch {
    throw new Error('Failed to parse Bot Connector token response as JSON');
  }

  if (!data.access_token || typeof data.expires_in !== 'number') {
    throw new Error('Invalid token response: missing access_token or expires_in');
  }

  // Cache the token
  tokenCache = {
    token: data.access_token,
    expiresAt: now + data.expires_in * 1000,
  };

  return data.access_token;
}

/**
 * Send an activity to Teams via Bot Connector API
 */
async function sendActivity(
  serviceUrl: string,
  conversationId: string,
  activity: Record<string, unknown>,
  replyToId?: string
): Promise<{ id: string }> {
  const token = await getBotConnectorToken();

  // Remove trailing slash from service URL if present
  const baseUrl = serviceUrl.endsWith('/') ? serviceUrl.slice(0, -1) : serviceUrl;

  let url: string;
  if (replyToId) {
    // Reply to existing activity (thread reply)
    url = `${baseUrl}/v3/conversations/${encodeURIComponent(conversationId)}/activities/${encodeURIComponent(replyToId)}`;
  } else {
    // New activity in conversation
    url = `${baseUrl}/v3/conversations/${encodeURIComponent(conversationId)}/activities`;
  }

  const response = await fetch(url, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      type: 'message',
      ...activity,
    }),
  });

  if (!response.ok) {
    const errorText = await response.text();
    console.error('[TeamsMCP] Send activity failed:', errorText);
    throw new Error(`Failed to send activity: ${response.status} - ${errorText}`);
  }

  let result: { id: string };
  try {
    result = (await response.json()) as { id: string };
  } catch {
    throw new Error('Failed to parse send activity response as JSON');
  }

  return result;
}

/**
 * Get required environment variables
 */
function getEnvConfig(): { serviceUrl: string; conversationId: string; threadId?: string } {
  const serviceUrl = process.env.TEAMS_SERVICE_URL;
  const conversationId = process.env.TEAMS_CONVERSATION_ID;
  const threadId = process.env.TEAMS_THREAD_ID;

  if (!serviceUrl) {
    throw new Error('TEAMS_SERVICE_URL is not set');
  }
  if (!conversationId) {
    throw new Error('TEAMS_CONVERSATION_ID is not set');
  }

  return { serviceUrl, conversationId, threadId };
}

/**
 * Resolve thread context from args and config
 * Returns the replyToId (if any) and appropriate action description
 */
function resolveThreadContext(
  argsThreadId: string | undefined,
  configThreadId: string | undefined,
  messageType: 'message' | 'rich'
): { replyToId: string | undefined; actionType: string } {
  const replyToId = argsThreadId || configThreadId;
  const actionType = replyToId
    ? 'Reply sent to thread'
    : messageType === 'rich'
      ? 'Rich message posted'
      : 'Message posted';
  return { replyToId, actionType };
}

function createServer(): Server {
  const server = new Server(
    {
      name: 'teams-mcp-server',
      version: '0.2.0',
    },
    {
      capabilities: {
        tools: {},
      },
    }
  );

  server.setRequestHandler(ListToolsRequestSchema, async () => {
    return {
      tools: [
        {
          name: 'teams_post_message',
          description:
            'Post a plain text message to the connected Microsoft Teams channel. Use this for simple messages and notifications.',
          inputSchema: zodToJsonSchema(PostMessageRequestSchema),
        },
        {
          name: 'teams_post_rich_message',
          description:
            'Post a rich structured message to Microsoft Teams with Adaptive Card support. Use this for formatted content like test results, status updates, or any content that benefits from structured layout.',
          inputSchema: zodToJsonSchema(PostRichMessageRequestSchema),
        },
      ],
    };
  });

  server.setRequestHandler(CallToolRequestSchema, async (request) => {
    try {
      if (!request.params) {
        throw new Error('Params are required');
      }

      // Get environment configuration
      const config = getEnvConfig();

      switch (request.params.name) {
        case 'teams_post_message': {
          const args = PostMessageRequestSchema.parse(request.params.arguments);
          const { replyToId, actionType } = resolveThreadContext(
            args.thread_id,
            config.threadId,
            'message'
          );

          const result = await sendActivity(
            config.serviceUrl,
            config.conversationId,
            {
              text: args.text,
              textFormat: 'markdown',
            },
            replyToId
          );

          return {
            content: [
              {
                type: 'text',
                text: `${actionType} successfully (id: ${result.id})`,
              },
            ],
          };
        }

        case 'teams_post_rich_message': {
          const args = PostRichMessageRequestSchema.parse(
            request.params.arguments
          );
          const { replyToId, actionType } = resolveThreadContext(
            args.thread_id,
            config.threadId,
            'rich'
          );

          let activity: Record<string, unknown>;

          if (args.card) {
            // Post with Adaptive Card attachment
            activity = {
              text: args.text || '',
              attachments: [
                {
                  contentType: 'application/vnd.microsoft.card.adaptive',
                  content: args.card,
                },
              ],
            };
          } else {
            // Plain text with markdown
            activity = {
              text: args.text || '',
              textFormat: 'markdown',
            };
          }

          const result = await sendActivity(
            config.serviceUrl,
            config.conversationId,
            activity,
            replyToId
          );

          return {
            content: [
              {
                type: 'text',
                text: `${actionType} successfully (id: ${result.id})`,
              },
            ],
          };
        }

        default:
          throw new Error(`Unknown tool: ${request.params.name}`);
      }
    } catch (error) {
      console.error('[TeamsMCP] Error handling request:', error);
      const errorMessage =
        error instanceof Error ? error.message : 'Unknown error occurred';
      throw new Error(errorMessage);
    }
  });

  return server;
}

async function runStdioServer() {
  // Validate required environment variables on startup
  const appId = process.env.TEAMS_BOT_APP_ID;
  const appPassword = process.env.TEAMS_BOT_APP_PASSWORD;

  if (!appId || !appPassword) {
    console.error(
      'TEAMS_BOT_APP_ID and TEAMS_BOT_APP_PASSWORD must be set.'
    );
    console.error(
      'These are the Azure Bot credentials, set by the execution environment.'
    );
    process.exit(1);
  }

  const server = createServer();
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error('Teams MCP Server v0.2.0 running on stdio (Bot Connector API)');
}

async function main() {
  await runStdioServer();
}

main().catch((error) => {
  console.error('Fatal error in main():', error);
  process.exit(1);
});
