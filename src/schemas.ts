import { z } from 'zod';

//
// Basic schemas for Microsoft Graph API responses
//

export const TeamSchema = z
  .object({
    id: z.string(),
    displayName: z.string().nullish(),
    description: z.string().nullish(),
    isArchived: z.boolean().nullish(),
    createdDateTime: z.string().nullish(),
  })
  .strip();

export const ChannelSchema = z
  .object({
    id: z.string(),
    displayName: z.string().nullish(),
    description: z.string().nullish(),
    membershipType: z.string().nullish(),
    createdDateTime: z.string().nullish(),
    webUrl: z.string().nullish(),
  })
  .strip();

const MessageFromSchema = z
  .object({
    user: z
      .object({
        displayName: z.string().nullish(),
        id: z.string().nullish(),
      })
      .nullish(),
    application: z
      .object({
        displayName: z.string().nullish(),
        id: z.string().nullish(),
      })
      .nullish(),
  })
  .strip();

const MessageBodySchema = z
  .object({
    content: z.string().nullish(),
    contentType: z.string().nullish(),
  })
  .strip();

export const MessageSchema = z
  .object({
    id: z.string(),
    body: MessageBodySchema.nullish(),
    from: MessageFromSchema.nullish(),
    createdDateTime: z.string().nullish(),
    lastModifiedDateTime: z.string().nullish(),
    replyToId: z.string().nullish(),
    subject: z.string().nullish(),
    importance: z.string().nullish(),
    webUrl: z.string().nullish(),
  })
  .strip();

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
      ])
    )
    .max(50),
  $schema: z.string().optional(),
});

//
// Request Schemas
//

export const ListTeamsRequestSchema = z.object({}).describe('No parameters required - lists all teams the user has joined');

export const ListChannelsRequestSchema = z.object({
  team_id: z.string().describe('The ID of the team to list channels for'),
});

export const PostMessageRequestSchema = z.object({
  team_id: z.string().describe('The ID of the team'),
  channel_id: z.string().describe('The ID of the channel to post to'),
  text: z
    .string()
    .describe('The message text to post (supports HTML formatting)'),
  reply_to_id: z
    .string()
    .optional()
    .describe('Message ID to reply to in a thread'),
});

export const PostRichMessageRequestSchema = z
  .object({
    team_id: z.string().describe('The ID of the team'),
    channel_id: z.string().describe('The ID of the channel to post to'),
    text: z
      .string()
      .optional()
      .describe('Fallback text for notifications. Required if card is not provided.'),
    card: AdaptiveCardSchema.optional().describe(
      'Adaptive Card JSON for rich formatting. Required if text is not provided.'
    ),
    reply_to_id: z
      .string()
      .optional()
      .describe('Message ID to reply to in a thread'),
  })
  .refine((data) => data.text || data.card, {
    message: 'Either text or card must be provided',
    path: ['text', 'card'],
  });

export const GetChannelHistoryRequestSchema = z.object({
  team_id: z.string().describe('The ID of the team'),
  channel_id: z
    .string()
    .describe(
      'The ID of the channel. Use this to get recent messages from a channel.'
    ),
  top: z
    .number()
    .int()
    .min(1)
    .max(50)
    .optional()
    .default(20)
    .describe('Maximum number of messages to retrieve (default 20, max 50)'),
});

export const GetThreadRepliesRequestSchema = z.object({
  team_id: z.string().describe('The ID of the team'),
  channel_id: z.string().describe('The ID of the channel containing the thread'),
  message_id: z
    .string()
    .describe('The ID of the parent message to get replies for'),
  top: z
    .number()
    .int()
    .min(1)
    .max(50)
    .optional()
    .default(20)
    .describe('Maximum number of replies to retrieve (default 20, max 50)'),
});

//
// Response Schemas
//

const BaseResponseSchema = z
  .object({
    '@odata.context': z.string().optional(),
    '@odata.count': z.number().optional(),
    '@odata.nextLink': z.string().optional(),
  })
  .strip();

export const ListTeamsResponseSchema = BaseResponseSchema.extend({
  value: z.array(TeamSchema),
});

export const ListChannelsResponseSchema = BaseResponseSchema.extend({
  value: z.array(ChannelSchema),
});

export const GetMessagesResponseSchema = BaseResponseSchema.extend({
  value: z.array(MessageSchema),
});

export const GetRepliesResponseSchema = BaseResponseSchema.extend({
  value: z.array(MessageSchema),
});
