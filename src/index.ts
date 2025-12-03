#!/usr/bin/env node

import {
  ListToolsRequestSchema,
  CallToolRequestSchema,
} from '@modelcontextprotocol/sdk/types.js';
import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { zodToJsonSchema } from 'zod-to-json-schema';
import { Client } from '@microsoft/microsoft-graph-client';
import dotenv from 'dotenv';
import {
  ListTeamsRequestSchema,
  ListChannelsRequestSchema,
  PostMessageRequestSchema,
  PostRichMessageRequestSchema,
  GetChannelHistoryRequestSchema,
  GetThreadRepliesRequestSchema,
  ListTeamsResponseSchema,
  ListChannelsResponseSchema,
  GetMessagesResponseSchema,
  GetRepliesResponseSchema,
} from './schemas.js';

dotenv.config();

if (!process.env.TEAMS_ACCESS_TOKEN) {
  console.error(
    'TEAMS_ACCESS_TOKEN is not set. Please set it in your environment or .env file.'
  );
  process.exit(1);
}

// Initialize Microsoft Graph client with access token
const graphClient = Client.init({
  authProvider: (done) => {
    done(null, process.env.TEAMS_ACCESS_TOKEN!);
  },
});

function createServer(): Server {
  const server = new Server(
    {
      name: 'teams-mcp-server',
      version: '0.0.1',
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
          name: 'teams_list_teams',
          description: 'List all Microsoft Teams that the user has joined',
          inputSchema: zodToJsonSchema(ListTeamsRequestSchema),
        },
        {
          name: 'teams_list_channels',
          description: 'List all channels in a Microsoft Teams team',
          inputSchema: zodToJsonSchema(ListChannelsRequestSchema),
        },
        {
          name: 'teams_post_message',
          description:
            'Post a plain text or HTML message to a Microsoft Teams channel or reply to a thread',
          inputSchema: zodToJsonSchema(PostMessageRequestSchema),
        },
        {
          name: 'teams_post_rich_message',
          description:
            'Post a rich structured message to Microsoft Teams with Adaptive Card support. Can post to channels or reply to threads. Supports text, images, fact sets, column layouts, and more.',
          inputSchema: zodToJsonSchema(PostRichMessageRequestSchema),
        },
        {
          name: 'teams_get_channel_history',
          description:
            'Get recent messages from a Microsoft Teams channel. Returns messages in reverse chronological order.',
          inputSchema: zodToJsonSchema(GetChannelHistoryRequestSchema),
        },
        {
          name: 'teams_get_thread_replies',
          description: 'Get all replies in a message thread',
          inputSchema: zodToJsonSchema(GetThreadRepliesRequestSchema),
        },
      ],
    };
  });

  server.setRequestHandler(CallToolRequestSchema, async (request) => {
    try {
      if (!request.params) {
        throw new Error('Params are required');
      }

      switch (request.params.name) {
        case 'teams_list_teams': {
          const response = await graphClient.api('/me/joinedTeams').get();
          const parsed = ListTeamsResponseSchema.parse(response);
          return {
            content: [{ type: 'text', text: JSON.stringify(parsed.value) }],
          };
        }

        case 'teams_list_channels': {
          const args = ListChannelsRequestSchema.parse(
            request.params.arguments
          );
          const response = await graphClient
            .api(`/teams/${args.team_id}/channels`)
            .get();
          const parsed = ListChannelsResponseSchema.parse(response);
          return {
            content: [{ type: 'text', text: JSON.stringify(parsed.value) }],
          };
        }

        case 'teams_post_message': {
          const args = PostMessageRequestSchema.parse(request.params.arguments);

          const messageBody = {
            body: {
              contentType: 'html',
              content: args.text,
            },
          };

          let endpoint: string;
          if (args.reply_to_id) {
            // Reply to a thread
            endpoint = `/teams/${args.team_id}/channels/${args.channel_id}/messages/${args.reply_to_id}/replies`;
          } else {
            // New message in channel
            endpoint = `/teams/${args.team_id}/channels/${args.channel_id}/messages`;
          }

          await graphClient.api(endpoint).post(messageBody);

          const actionType = args.reply_to_id
            ? 'Reply sent to thread'
            : 'Message posted';
          return {
            content: [{ type: 'text', text: `${actionType} successfully` }],
          };
        }

        case 'teams_post_rich_message': {
          const args = PostRichMessageRequestSchema.parse(
            request.params.arguments
          );

          let messageBody: Record<string, unknown>;

          if (args.card) {
            // Post with Adaptive Card attachment
            messageBody = {
              body: {
                contentType: 'html',
                content: args.text || '',
              },
              attachments: [
                {
                  id: '1',
                  contentType: 'application/vnd.microsoft.card.adaptive',
                  content: JSON.stringify(args.card),
                },
              ],
            };
          } else {
            // Plain HTML message
            messageBody = {
              body: {
                contentType: 'html',
                content: args.text,
              },
            };
          }

          let endpoint: string;
          if (args.reply_to_id) {
            endpoint = `/teams/${args.team_id}/channels/${args.channel_id}/messages/${args.reply_to_id}/replies`;
          } else {
            endpoint = `/teams/${args.team_id}/channels/${args.channel_id}/messages`;
          }

          await graphClient.api(endpoint).post(messageBody);

          const actionType = args.reply_to_id
            ? 'Reply sent to thread'
            : 'Rich message posted';
          return {
            content: [{ type: 'text', text: `${actionType} successfully` }],
          };
        }

        case 'teams_get_channel_history': {
          const args = GetChannelHistoryRequestSchema.parse(
            request.params.arguments
          );
          const response = await graphClient
            .api(
              `/teams/${args.team_id}/channels/${args.channel_id}/messages`
            )
            .top(args.top || 20)
            .get();
          const parsed = GetMessagesResponseSchema.parse(response);
          return {
            content: [{ type: 'text', text: JSON.stringify(parsed.value) }],
          };
        }

        case 'teams_get_thread_replies': {
          const args = GetThreadRepliesRequestSchema.parse(
            request.params.arguments
          );
          const response = await graphClient
            .api(
              `/teams/${args.team_id}/channels/${args.channel_id}/messages/${args.message_id}/replies`
            )
            .top(args.top || 20)
            .get();
          const parsed = GetRepliesResponseSchema.parse(response);
          return {
            content: [{ type: 'text', text: JSON.stringify(parsed.value) }],
          };
        }

        default:
          throw new Error(`Unknown tool: ${request.params.name}`);
      }
    } catch (error) {
      console.error('Error handling request:', error);
      const errorMessage =
        error instanceof Error ? error.message : 'Unknown error occurred';
      throw new Error(errorMessage);
    }
  });

  return server;
}

async function runStdioServer() {
  const server = createServer();
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error('Teams MCP Server running on stdio');
}

async function main() {
  // Run with stdio transport
  await runStdioServer();
}

main().catch((error) => {
  console.error('Fatal error in main():', error);
  process.exit(1);
});
