# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a Model Context Protocol (MCP) server that provides AI assistants with standardized access to Microsoft Teams APIs via Microsoft Graph. It's written in TypeScript and uses stdio transport for direct integration.

## Development Commands

### Building and Running
- `npm run build` - Compile TypeScript to JavaScript in `/dist`
- `npm run dev` - Start server in development mode with hot reloading
- `npm start` - Run the production build

### Code Quality
- `npm run lint` - Run both ESLint and Prettier checks
- `npm run fix` - Auto-fix all linting issues

## Architecture

### Core Structure
The server follows a schema-driven design pattern:

1. **Request/Response Schemas** (`src/schemas.ts`):
   - All Microsoft Graph API interactions are validated with Zod schemas
   - Request schemas define input parameters
   - Response schemas filter API responses to only necessary fields
   - Adaptive Card schemas for rich message formatting

2. **Main Server** (`src/index.ts`):
   - Stdio-only transport for CLI integration
   - Tool registration and request handling
   - Microsoft Graph client initialization

### Transport Mode
- **Stdio only**: For CLI integration (Claude Desktop, etc.)

### Available Tools

**6 tools for Teams operations:**

| Tool | Description | Microsoft Graph API |
|------|-------------|---------------------|
| `teams_list_teams` | List all teams the user has joined | `GET /me/joinedTeams` |
| `teams_list_channels` | List channels in a team | `GET /teams/{id}/channels` |
| `teams_post_message` | Post text/HTML message to channel or thread | `POST /teams/{id}/channels/{id}/messages` |
| `teams_post_rich_message` | Post Adaptive Card to channel or thread | Same endpoint with attachment |
| `teams_get_channel_history` | Get recent channel messages | `GET /teams/{id}/channels/{id}/messages` |
| `teams_get_thread_replies` | Get replies to a message | `GET .../messages/{id}/replies` |

### Environment Requirements
Must set in environment or `.env` file:
- `TEAMS_ACCESS_TOKEN`: Microsoft Graph OAuth token

### Required Microsoft Graph Permissions
Your app registration needs these API permissions:
- `Team.ReadBasic.All` - To list teams
- `Channel.ReadBasic.All` - To list channels
- `ChannelMessage.Send` - To post messages
- `ChannelMessage.Read.All` - To read channel messages

## Key Implementation Notes

1. **Team ID Required**: Unlike Slack's flat channel structure, Teams channels are nested under teams. Most operations require both `team_id` and `channel_id`.

2. **Thread Replies**: Use `reply_to_id` (message ID) to reply to threads, not timestamps like Slack.

3. **Adaptive Cards**: Rich messages use Microsoft's Adaptive Card format, not Slack Block Kit:
   ```json
   {
     "type": "AdaptiveCard",
     "version": "1.4",
     "body": [
       { "type": "TextBlock", "text": "Hello!", "weight": "Bolder" }
     ]
   }
   ```

4. **HTML Content**: Plain messages support HTML formatting (`contentType: 'html'`).

5. **Type Safety**: All Microsoft Graph responses are parsed through Zod schemas.

6. **ES Modules**: Project uses `"type": "module"` - use ES import syntax.

## Common Tasks

### Adding a New Teams Tool
1. Define request/response schemas in `src/schemas.ts`
2. Add tool registration in `src/index.ts` server setup
3. Implement handler following existing pattern: validate → API call → parse → return
4. Update this documentation

### Differences from Slack MCP

| Aspect | Slack | Teams |
|--------|-------|-------|
| Channel structure | Flat list | Nested under teams |
| Thread reference | `thread_ts` timestamp | `reply_to_id` message ID |
| Rich messages | Block Kit JSON | Adaptive Cards JSON |
| API client | `@slack/web-api` | `@microsoft/microsoft-graph-client` |
| Token env var | `SLACK_BOT_TOKEN` | `TEAMS_ACCESS_TOKEN` |

### Known API Limitations
1. **Message History**: Limited to 50 messages per request
2. **Rate Limiting**: Microsoft Graph rate limits apply
3. **Permissions**: Operations require appropriate Graph API permissions
4. **Reactions**: Microsoft Graph has limited reaction support compared to Slack

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
      "wrap": true
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
        { "title": "Skipped", "value": "3" }
      ]
    }
  ]
}
```

### Two-Column Layout
```json
{
  "type": "AdaptiveCard",
  "version": "1.4",
  "body": [
    {
      "type": "ColumnSet",
      "columns": [
        {
          "type": "Column",
          "width": "auto",
          "items": [{ "type": "TextBlock", "text": "Status:", "weight": "Bolder" }]
        },
        {
          "type": "Column",
          "width": "stretch",
          "items": [{ "type": "TextBlock", "text": "Success", "color": "Good" }]
        }
      ]
    }
  ]
}
```
