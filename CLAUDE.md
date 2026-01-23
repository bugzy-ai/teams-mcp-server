# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a Model Context Protocol (MCP) server that allows AI assistants to post messages to Microsoft Teams channels via the Bot Connector API. It's used by the Bugzy agent to send notifications and responses to Teams.

**Version 0.2.0** - Rewritten to use Bot Connector API (previously used Microsoft Graph API).

## Development Commands

### Building and Running
- `npm run build` - Compile TypeScript to JavaScript in `/dist`
- `npm run dev` - Start server in development mode with hot reloading
- `npm start` - Run the production build

### Code Quality
- `npm run lint` - Run both ESLint and Prettier checks
- `npm run fix` - Auto-fix all linting issues

## Architecture

### Bot Connector API (New in v0.2.0)

The server now uses the **Bot Connector REST API** instead of Microsoft Graph API. This enables:
- Posting messages as the Bugzy bot (not as a user)
- Integration with Azure Bot Service
- Proper bot identity in Teams

### Environment Variables

Required variables (set by the execution context):

```bash
# Bot credentials (from Azure Bot Service)
TEAMS_BOT_APP_ID=<Azure Bot App ID>
TEAMS_BOT_APP_PASSWORD=<Azure Bot App Password>

# Conversation context (from stored conversation reference)
TEAMS_SERVICE_URL=<Bot Connector service URL, e.g., https://smba.trafficmanager.net/...>
TEAMS_CONVERSATION_ID=<Teams conversation/channel ID>

# Optional
TEAMS_BOT_TENANT_ID=<Azure AD tenant ID for single-tenant bots>
TEAMS_THREAD_ID=<Activity ID to reply to in a thread>
```

**Note on Tenant ID:** For single-tenant bots (registered in a specific Azure AD tenant), set `TEAMS_BOT_TENANT_ID` to your tenant GUID. This changes the token endpoint from `botframework.com` to your tenant-specific endpoint. Multi-tenant bots can omit this variable.

### Available Tools

**2 tools for posting to Teams:**

| Tool | Description |
|------|-------------|
| `teams_post_message` | Post a plain text/markdown message to the connected channel |
| `teams_post_rich_message` | Post an Adaptive Card (rich structured message) to the connected channel |

**Note:** Unlike v0.1.x, these tools do NOT take `team_id` or `channel_id` parameters. The target conversation is determined by environment variables, which are set by the Bugzy execution context.

### Core Structure

1. **Schemas** (`src/schemas.ts`):
   - Adaptive Card schema for rich messages
   - Request schemas for the two tools

2. **Main Server** (`src/index.ts`):
   - Bot Connector token acquisition (client credentials flow)
   - Activity sending via Bot Connector REST API
   - Tool registration and request handling

## Key Implementation Notes

1. **Token Caching**: Bot Connector tokens are cached for ~1 hour with automatic refresh.

2. **Thread Replies**: Use `thread_id` parameter or `TEAMS_THREAD_ID` env var to reply in a thread.

3. **Adaptive Cards**: Use the `teams_post_rich_message` tool with a `card` parameter for structured content:
   ```json
   {
     "type": "AdaptiveCard",
     "version": "1.4",
     "body": [
       { "type": "TextBlock", "text": "Test Results", "weight": "Bolder" },
       { "type": "TextBlock", "text": "All 45 tests passed!", "wrap": true }
     ]
   }
   ```

4. **Markdown Support**: Plain text messages support markdown formatting via the `textFormat: 'markdown'` property.

## Differences from v0.1.x (Graph API)

| Aspect | v0.1.x (Graph API) | v0.2.0 (Bot Connector) |
|--------|-------------------|------------------------|
| API | Microsoft Graph | Bot Connector REST |
| Auth | User access token | Bot client credentials |
| Tools | 6 tools (list, post, read) | 2 tools (post only) |
| Params | Required `team_id`, `channel_id` | From environment |
| Identity | Posts as user | Posts as Bugzy bot |

## Adaptive Card Examples

### Simple Text Message
```json
{
  "type": "AdaptiveCard",
  "version": "1.4",
  "body": [
    {
      "type": "TextBlock",
      "text": "Test Results",
      "weight": "Bolder",
      "size": "Medium"
    },
    {
      "type": "TextBlock",
      "text": "All 45 tests passed!",
      "wrap": true,
      "color": "Good"
    }
  ]
}
```

### Fact Set (Key-Value Pairs)
```json
{
  "type": "AdaptiveCard",
  "version": "1.4",
  "body": [
    {
      "type": "FactSet",
      "facts": [
        { "title": "Passed", "value": "45" },
        { "title": "Failed", "value": "2" },
        { "title": "Duration", "value": "3m 42s" }
      ]
    }
  ]
}
```

### With Actions
```json
{
  "type": "AdaptiveCard",
  "version": "1.4",
  "body": [
    { "type": "TextBlock", "text": "Deployment Ready", "weight": "Bolder" }
  ],
  "actions": [
    {
      "type": "Action.Submit",
      "title": "Approve",
      "style": "positive",
      "data": { "action": "approve" }
    },
    {
      "type": "Action.Submit",
      "title": "Reject",
      "style": "destructive",
      "data": { "action": "reject" }
    }
  ]
}
```

## Common Tasks

### Adding a New Feature to Card Schema
1. Update the Zod schema in `src/schemas.ts`
2. Test with `npm run build` to ensure types are correct
3. Update this documentation with examples

## Publishing

```bash
npm run clean && npm run build
npm publish
```

The package is published to npmjs as `@bugzy-ai/teams-mcp-server`.
