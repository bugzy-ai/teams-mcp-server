import { z } from 'zod';

/**
 * Teams MCP Server Schemas
 *
 * Simplified schemas for Bot Connector API integration.
 * Only 2 tools: teams_post_message and teams_post_rich_message
 */

//
// Adaptive Card Schemas
//

const TextBlockSchema = z.object({
  type: z.literal('TextBlock'),
  text: z.string(),
  weight: z.enum(['Default', 'Lighter', 'Bolder']).optional(),
  size: z
    .enum(['Default', 'Small', 'Medium', 'Large', 'ExtraLarge'])
    .optional(),
  color: z
    .enum([
      'Default',
      'Dark',
      'Light',
      'Accent',
      'Good',
      'Warning',
      'Attention',
    ])
    .optional(),
  wrap: z.boolean().optional(),
  isSubtle: z.boolean().optional(),
  maxLines: z.number().optional(),
  horizontalAlignment: z.enum(['Left', 'Center', 'Right']).optional(),
  spacing: z
    .enum(['None', 'Small', 'Default', 'Medium', 'Large', 'ExtraLarge'])
    .optional(),
  separator: z.boolean().optional(),
});

const ImageSchema = z.object({
  type: z.literal('Image'),
  url: z.string().url(),
  altText: z.string().optional(),
  size: z.enum(['Auto', 'Stretch', 'Small', 'Medium', 'Large']).optional(),
  style: z.enum(['Default', 'Person']).optional(),
  horizontalAlignment: z.enum(['Left', 'Center', 'Right']).optional(),
  width: z.string().optional(),
  height: z.string().optional(),
});

const FactSchema = z.object({
  title: z.string(),
  value: z.string(),
});

const FactSetSchema = z.object({
  type: z.literal('FactSet'),
  facts: z.array(FactSchema).max(10),
  spacing: z
    .enum(['None', 'Small', 'Default', 'Medium', 'Large', 'ExtraLarge'])
    .optional(),
  separator: z.boolean().optional(),
});

// Forward declaration for recursive types
const CardElementSchema: z.ZodType<unknown> = z.lazy(() =>
  z.union([
    TextBlockSchema,
    ImageSchema,
    FactSetSchema,
    ColumnSetSchema,
    ContainerSchema,
    ActionSetSchema,
  ])
);

const ColumnSchema = z.object({
  type: z.literal('Column'),
  width: z.union([z.string(), z.number()]).optional(),
  items: z.array(CardElementSchema).optional(),
  verticalContentAlignment: z.enum(['Top', 'Center', 'Bottom']).optional(),
  spacing: z
    .enum(['None', 'Small', 'Default', 'Medium', 'Large', 'ExtraLarge'])
    .optional(),
});

const ColumnSetSchema = z.object({
  type: z.literal('ColumnSet'),
  columns: z.array(ColumnSchema),
  spacing: z
    .enum(['None', 'Small', 'Default', 'Medium', 'Large', 'ExtraLarge'])
    .optional(),
  separator: z.boolean().optional(),
});

const ContainerSchema = z.object({
  type: z.literal('Container'),
  items: z.array(CardElementSchema),
  style: z.enum(['Default', 'Emphasis', 'Good', 'Attention', 'Warning']).optional(),
  spacing: z
    .enum(['None', 'Small', 'Default', 'Medium', 'Large', 'ExtraLarge'])
    .optional(),
  separator: z.boolean().optional(),
});

const ActionSchema = z.object({
  type: z.enum(['Action.Submit', 'Action.OpenUrl', 'Action.ShowCard']),
  title: z.string(),
  data: z.unknown().optional(),
  url: z.string().optional(),
  style: z.enum(['default', 'positive', 'destructive']).optional(),
});

const ActionSetSchema = z.object({
  type: z.literal('ActionSet'),
  actions: z.array(ActionSchema),
});

export const AdaptiveCardSchema = z.object({
  type: z.literal('AdaptiveCard'),
  version: z.string().default('1.4'),
  body: z
    .array(
      z.union([
        TextBlockSchema,
        ImageSchema,
        FactSetSchema,
        ColumnSetSchema,
        ContainerSchema,
        ActionSetSchema,
      ])
    )
    .max(50),
  actions: z.array(ActionSchema).optional(),
  $schema: z.string().optional(),
});

export type AdaptiveCard = z.infer<typeof AdaptiveCardSchema>;

//
// Request Schemas for Bot Connector API
//

/**
 * Schema for teams_post_message tool
 *
 * Posts a plain text message to the connected Teams channel.
 * Channel info comes from environment variables (TEAMS_SERVICE_URL, TEAMS_CONVERSATION_ID).
 */
export const PostMessageRequestSchema = z.object({
  text: z
    .string()
    .describe('The message text to post. Supports markdown formatting.'),
  thread_id: z
    .string()
    .optional()
    .describe(
      'Activity ID to reply to in a thread. If not provided and TEAMS_THREAD_ID env var is set, will reply to that thread.'
    ),
});

/**
 * Schema for teams_post_rich_message tool
 *
 * Posts a rich message with Adaptive Card to the connected Teams channel.
 */
export const PostRichMessageRequestSchema = z
  .object({
    text: z
      .string()
      .optional()
      .describe(
        'Fallback text for notifications and screen readers. Required if card is not provided.'
      ),
    card: AdaptiveCardSchema.optional().describe(
      'Adaptive Card JSON for rich formatting. Use this for structured content like test results, status updates, etc.'
    ),
    thread_id: z
      .string()
      .optional()
      .describe(
        'Activity ID to reply to in a thread. If not provided and TEAMS_THREAD_ID env var is set, will reply to that thread.'
      ),
  })
  .refine((data) => data.text || data.card, {
    message: 'Either text or card must be provided',
    path: ['text', 'card'],
  });
