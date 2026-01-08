import { McpTool } from '../../../../types';
import * as lark from '@larksuiteoapi/node-sdk';
import { z } from 'zod';
import { CallToolResult } from '@modelcontextprotocol/sdk/types';
import {
  extractDocumentId,
  BlockType,
  markdownToBlocks,
  textToBlock,
  DocumentBlock,
  TextElement,
  TextContent,
  CODE_LANGUAGE_MAP,
} from '../../../../utils';

// Tool name type
export type docxEditToolName =
  | 'docx.builtin.updateTitle'
  | 'docx.builtin.append'
  | 'docx.builtin.replace'
  | 'docx.builtin.edit';

// ============ Local Type Definitions (API Response Types) ============

/** Block list item (API response) */
interface BlockItem {
  block_id?: string;
  parent_id?: string;
  block_type?: number;
  text?: TextContent;
  code?: {
    style?: { language?: number; wrap?: boolean };
    elements: TextElement[];
  };
}

/** Update request */
interface UpdateRequest {
  block_id: string;
  update_text_elements: {
    elements: TextElement[];
  };
}

/**
 * Helper function: Get all document blocks with pagination
 */
async function getAllDocumentBlocks(
  client: lark.Client,
  documentId: string,
  userAccessToken?: string,
  useUAT?: boolean,
): Promise<BlockItem[]> {
  const allItems: BlockItem[] = [];
  let pageToken: string | undefined;
  const PAGE_SIZE = 500;

  do {
    const response =
      userAccessToken && useUAT
        ? await client.docx.v1.documentBlock.list(
            {
              path: { document_id: documentId },
              params: {
                page_size: PAGE_SIZE,
                document_revision_id: -1,
                page_token: pageToken,
              },
            },
            lark.withUserAccessToken(userAccessToken),
          )
        : await client.docx.v1.documentBlock.list({
            path: { document_id: documentId },
            params: {
              page_size: PAGE_SIZE,
              document_revision_id: -1,
              page_token: pageToken,
            },
          });

    const items = (response.data?.items || []) as BlockItem[];
    allItems.push(...items);
    pageToken = response.data?.page_token;
  } while (pageToken);

  return allItems;
}

/**
 * Helper function: Get all block children with pagination
 */
async function getAllBlockChildren(
  client: lark.Client,
  documentId: string,
  blockId: string,
  userAccessToken?: string,
  useUAT?: boolean,
): Promise<BlockItem[]> {
  const allItems: BlockItem[] = [];
  let pageToken: string | undefined;
  const PAGE_SIZE = 500;

  do {
    const response =
      userAccessToken && useUAT
        ? await client.docx.v1.documentBlockChildren.get(
            {
              path: { document_id: documentId, block_id: blockId },
              params: {
                document_revision_id: -1,
                page_size: PAGE_SIZE,
                page_token: pageToken,
              },
            },
            lark.withUserAccessToken(userAccessToken),
          )
        : await client.docx.v1.documentBlockChildren.get({
            path: { document_id: documentId, block_id: blockId },
            params: {
              document_revision_id: -1,
              page_size: PAGE_SIZE,
              page_token: pageToken,
            },
          });

    const items = (response.data?.items || []) as BlockItem[];
    allItems.push(...items);
    pageToken = response.data?.page_token;
  } while (pageToken);

  return allItems;
}

// ============ Tool Definitions ============

/**
 * Update document title
 */
export const larkDocxUpdateTitleTool: McpTool = {
  project: 'docx',
  name: 'docx.builtin.updateTitle',
  accessTokens: ['user', 'tenant'],
  description:
    '[Feishu/Lark]-Docs-Document-Update Title-Update the title of a Feishu document. Provide document ID or URL and the new title.',
  schema: {
    data: z.object({
      document_id: z.string().describe('Document ID or URL. Supports direct document_id or full Feishu document link'),
      title: z.string().describe('New document title'),
    }),
    useUAT: z.boolean().describe('Use user identity for the request, otherwise use application identity').optional(),
  },
  customHandler: async (client, params, options): Promise<CallToolResult> => {
    try {
      const { userAccessToken } = options || {};
      const documentId = extractDocumentId(params.data.document_id);

      // Use documentBlock.patch to update document title (Page block body)
      const response =
        userAccessToken && params.useUAT
          ? await client.docx.v1.documentBlock.patch(
              {
                path: {
                  document_id: documentId,
                  block_id: documentId, // Document root block ID equals document_id
                },
                data: {
                  update_text_elements: {
                    elements: [
                      {
                        text_run: {
                          content: params.data.title,
                        },
                      },
                    ],
                  },
                },
                params: {
                  document_revision_id: -1, // Use latest version
                },
              },
              lark.withUserAccessToken(userAccessToken),
            )
          : await client.docx.v1.documentBlock.patch({
              path: {
                document_id: documentId,
                block_id: documentId,
              },
              data: {
                update_text_elements: {
                  elements: [
                    {
                      text_run: {
                        content: params.data.title,
                      },
                    },
                  ],
                },
              },
              params: {
                document_revision_id: -1,
              },
            });

      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify({
              success: true,
              message: 'Document title updated successfully',
              document_id: documentId,
              new_title: params.data.title,
              data: response.data,
            }),
          },
        ],
      };
    } catch (error) {
      return {
        isError: true,
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify({
              success: false,
              error: (error as any)?.response?.data || (error as any)?.message || error,
            }),
          },
        ],
      };
    }
  },
};

/**
 * Append content to document
 */
export const larkDocxAppendTool: McpTool = {
  project: 'docx',
  name: 'docx.builtin.append',
  accessTokens: ['user', 'tenant'],
  description:
    '[Feishu/Lark]-Docs-Document-Append Content-Append content to the end of a document. Supports appending text, headings, lists and other block types.',
  schema: {
    data: z.object({
      document_id: z.string().describe('Document ID or URL'),
      content: z.string().describe('Content to append (supports Markdown format)'),
      block_type: z
        .enum(['text', 'heading1', 'heading2', 'heading3', 'bullet', 'ordered', 'code', 'quote', 'divider', 'auto'])
        .describe(
          'Block type: text-plain text, heading1-6-headings, bullet-unordered list, ordered-ordered list, code-code block, quote-quote, divider-divider, auto-automatically parse from Markdown format',
        )
        .default('auto'),
      code_language: z
        .string()
        .describe('Code block language (only valid when block_type is code), e.g.: javascript, python, go')
        .optional(),
    }),
    useUAT: z.boolean().describe('Use user identity for the request, otherwise use application identity').optional(),
  },
  customHandler: async (client, params, options): Promise<CallToolResult> => {
    try {
      const { userAccessToken } = options || {};
      const documentId = extractDocumentId(params.data.document_id);

      let children: DocumentBlock[];

      if (params.data.block_type === 'auto') {
        // Auto parse Markdown
        children = markdownToBlocks(params.data.content);
      } else {
        // Create single block based on specified block_type
        const blockTypeMap: Record<string, number> = {
          text: BlockType.Text,
          heading1: BlockType.Heading1,
          heading2: BlockType.Heading2,
          heading3: BlockType.Heading3,
          bullet: BlockType.Bullet,
          ordered: BlockType.Ordered,
          code: BlockType.Code,
          quote: BlockType.Quote,
          divider: BlockType.Divider,
        };

        const blockType = blockTypeMap[params.data.block_type] || BlockType.Text;

        if (blockType === BlockType.Divider) {
          // Divider
          children = [{ block_type: BlockType.Divider }];
        } else if (blockType === BlockType.Code) {
          // Code block
          const language = CODE_LANGUAGE_MAP[(params.data.code_language || '').toLowerCase()] || 1;
          children = [
            {
              block_type: BlockType.Code,
              code: {
                style: { language, wrap: false },
                elements: [{ text_run: { content: params.data.content } }],
              },
            },
          ];
        } else {
          children = [textToBlock(params.data.content, blockType)];
        }
      }

      // Get all document blocks to determine insertion position (with pagination)
      const items = await getAllDocumentBlocks(client, documentId, userAccessToken, params.useUAT);
      const insertIndex = items.length > 0 ? items.length - 1 : 0;

      // Create child blocks
      const response =
        userAccessToken && params.useUAT
          ? await client.docx.v1.documentBlockChildren.create(
              {
                path: {
                  document_id: documentId,
                  block_id: documentId,
                },
                data: {
                  children,
                  index: insertIndex,
                },
                params: {
                  document_revision_id: -1,
                },
              },
              lark.withUserAccessToken(userAccessToken),
            )
          : await client.docx.v1.documentBlockChildren.create({
              path: {
                document_id: documentId,
                block_id: documentId,
              },
              data: {
                children,
                index: insertIndex,
              },
              params: {
                document_revision_id: -1,
              },
            });

      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify({
              success: true,
              message: 'Content appended successfully',
              document_id: documentId,
              blocks_created: children.length,
              data: response.data,
            }),
          },
        ],
      };
    } catch (error) {
      return {
        isError: true,
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify({
              success: false,
              error: (error as any)?.response?.data || (error as any)?.message || error,
            }),
          },
        ],
      };
    }
  },
};

/**
 * Replace document content
 */
export const larkDocxReplaceTool: McpTool = {
  project: 'docx',
  name: 'docx.builtin.replace',
  accessTokens: ['user', 'tenant'],
  description:
    '[Feishu/Lark]-Docs-Document-Replace Content-Find and replace specified text in a document. Supports replacing in the entire document or a specific block.',
  schema: {
    data: z.object({
      document_id: z.string().describe('Document ID or URL'),
      search_text: z.string().describe('Text to search for'),
      replace_text: z.string().describe('Replacement text'),
      block_id: z.string().describe('Block ID to replace in. If not specified, searches the entire document').optional(),
      replace_all: z.boolean().describe('Whether to replace all matches. Default is true').default(true),
    }),
    useUAT: z.boolean().describe('Use user identity for the request, otherwise use application identity').optional(),
  },
  customHandler: async (client, params, options): Promise<CallToolResult> => {
    try {
      const { userAccessToken } = options || {};
      const documentId = extractDocumentId(params.data.document_id);
      const { search_text, replace_text, block_id, replace_all } = params.data;

      // Validate search_text is not empty
      if (!search_text) {
        return {
          isError: true,
          content: [
            {
              type: 'text' as const,
              text: JSON.stringify({
                success: false,
                message: 'search_text cannot be empty',
              }),
            },
          ],
        };
      }

      // Get all document blocks (with pagination)
      const items = await getAllDocumentBlocks(client, documentId, userAccessToken, params.useUAT);
      const updateRequests: UpdateRequest[] = [];
      let replacedCount = 0;

      for (const block of items) {
        // If block_id is specified, only process that block
        if (block_id && block.block_id !== block_id) {
          continue;
        }

        // Get block text content
        const textContent = block.text?.elements || block.code?.elements || [];
        let hasMatch = false;
        const newElements: TextElement[] = [];

        for (const element of textContent) {
          if (element.text_run?.content) {
            const content = element.text_run.content;
            if (content.includes(search_text)) {
              hasMatch = true;
              const newContent = replace_all
                ? content.split(search_text).join(replace_text)
                : content.replace(search_text, replace_text);
              // Correctly count replacements: count all matches when replace_all, otherwise count 1
              const escapedText = search_text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
              const matches = content.match(new RegExp(escapedText, 'g')) || [];
              replacedCount += replace_all ? matches.length : (matches.length > 0 ? 1 : 0);
              newElements.push({
                text_run: {
                  content: newContent,
                  text_element_style: element.text_run.text_element_style,
                },
              });
            } else {
              newElements.push(element);
            }
          } else {
            newElements.push(element);
          }
        }

        if (hasMatch && block.block_id) {
          updateRequests.push({
            block_id: block.block_id,
            update_text_elements: {
              elements: newElements,
            },
          });
        }

        // If not replacing all and found a match, stop
        if (!replace_all && replacedCount > 0) {
          break;
        }
      }

      if (updateRequests.length === 0) {
        return {
          content: [
            {
              type: 'text' as const,
              text: JSON.stringify({
                success: true,
                message: 'No matching text found',
                document_id: documentId,
                search_text,
                replaced_count: 0,
              }),
            },
          ],
        };
      }

      // Batch update blocks
      const response =
        userAccessToken && params.useUAT
          ? await client.docx.v1.documentBlock.batchUpdate(
              {
                path: { document_id: documentId },
                data: { requests: updateRequests },
                params: { document_revision_id: -1 },
              },
              lark.withUserAccessToken(userAccessToken),
            )
          : await client.docx.v1.documentBlock.batchUpdate({
              path: { document_id: documentId },
              data: { requests: updateRequests },
              params: { document_revision_id: -1 },
            });

      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify({
              success: true,
              message: 'Replacement successful',
              document_id: documentId,
              search_text,
              replace_text,
              replaced_count: replacedCount,
              blocks_updated: updateRequests.length,
              data: response.data,
            }),
          },
        ],
      };
    } catch (error) {
      return {
        isError: true,
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify({
              success: false,
              error: (error as any)?.response?.data || (error as any)?.message || error,
            }),
          },
        ],
      };
    }
  },
};

/**
 * Simplified Markdown editing interface
 */
export const larkDocxEditTool: McpTool = {
  project: 'docx',
  name: 'docx.builtin.edit',
  accessTokens: ['user', 'tenant'],
  description:
    '[Feishu/Lark]-Docs-Document-Edit Document-Edit Feishu document using Markdown format. Supports inserting content at specified position, replacing block content, or clearing and writing new content. Note: replace_block and clear_and_write modes involve multiple API calls and are not atomic operations, which may cause race conditions in high-concurrency editing scenarios.',
  schema: {
    data: z.object({
      document_id: z.string().describe('Document ID or URL'),
      markdown: z.string().describe('Content in Markdown format'),
      mode: z
        .enum(['append', 'prepend', 'replace_block', 'clear_and_write'])
        .describe(
          'Edit mode: append-append to end, prepend-insert at beginning, replace_block-replace specific block, clear_and_write-clear document and write new content',
        )
        .default('append'),
      block_id: z.string().describe('Block ID to replace (required only when mode is replace_block)').optional(),
    }),
    useUAT: z.boolean().describe('Use user identity for the request, otherwise use application identity').optional(),
  },
  customHandler: async (client, params, options): Promise<CallToolResult> => {
    try {
      const { userAccessToken } = options || {};
      const documentId = extractDocumentId(params.data.document_id);
      const { markdown, mode, block_id } = params.data;

      // Convert Markdown to document blocks
      const blocks = markdownToBlocks(markdown);

      if (blocks.length === 0) {
        return {
          content: [
            {
              type: 'text' as const,
              text: JSON.stringify({
                success: false,
                message: 'Markdown content is empty or cannot be parsed',
              }),
            },
          ],
        };
      }

      let response: { data?: unknown } | undefined;

      switch (mode) {
        case 'append': {
          // Get all document blocks to determine insertion position (with pagination)
          const appendItems = await getAllDocumentBlocks(client, documentId, userAccessToken, params.useUAT);
          const insertIndex = appendItems.length > 0 ? appendItems.length - 1 : 0;

          response =
            userAccessToken && params.useUAT
              ? await client.docx.v1.documentBlockChildren.create(
                  {
                    path: { document_id: documentId, block_id: documentId },
                    data: { children: blocks, index: insertIndex },
                    params: { document_revision_id: -1 },
                  },
                  lark.withUserAccessToken(userAccessToken),
                )
              : await client.docx.v1.documentBlockChildren.create({
                  path: { document_id: documentId, block_id: documentId },
                  data: { children: blocks, index: insertIndex },
                  params: { document_revision_id: -1 },
                });
          break;
        }

        case 'prepend': {
          response =
            userAccessToken && params.useUAT
              ? await client.docx.v1.documentBlockChildren.create(
                  {
                    path: { document_id: documentId, block_id: documentId },
                    data: { children: blocks, index: 0 },
                    params: { document_revision_id: -1 },
                  },
                  lark.withUserAccessToken(userAccessToken),
                )
              : await client.docx.v1.documentBlockChildren.create({
                  path: { document_id: documentId, block_id: documentId },
                  data: { children: blocks, index: 0 },
                  params: { document_revision_id: -1 },
                });
          break;
        }

        case 'replace_block': {
          if (!block_id) {
            return {
              isError: true,
              content: [
                {
                  type: 'text' as const,
                  text: JSON.stringify({
                    success: false,
                    message: 'replace_block mode requires block_id',
                  }),
                },
              ],
            };
          }

          // First delete original block, then insert new blocks at same position
          // Get block info to determine parent and position
          const blockInfo =
            userAccessToken && params.useUAT
              ? await client.docx.v1.documentBlock.get(
                  {
                    path: { document_id: documentId, block_id },
                    params: { document_revision_id: -1 },
                  },
                  lark.withUserAccessToken(userAccessToken),
                )
              : await client.docx.v1.documentBlock.get({
                  path: { document_id: documentId, block_id },
                  params: { document_revision_id: -1 },
                });

          const parentId = blockInfo.data?.block?.parent_id || documentId;

          // Get all parent block's children to determine position (with pagination)
          const children = await getAllBlockChildren(client, documentId, parentId, userAccessToken, params.useUAT);
          const blockIndex = children.findIndex((c: BlockItem) => c.block_id === block_id);

          // Delete original block and get new document_revision_id
          const deleteResponse = await (userAccessToken && params.useUAT
            ? client.docx.v1.documentBlockChildren.batchDelete(
                {
                  path: { document_id: documentId, block_id: parentId },
                  data: { start_index: blockIndex, end_index: blockIndex + 1 },
                  params: { document_revision_id: -1 },
                },
                lark.withUserAccessToken(userAccessToken),
              )
            : client.docx.v1.documentBlockChildren.batchDelete({
                path: { document_id: documentId, block_id: parentId },
                data: { start_index: blockIndex, end_index: blockIndex + 1 },
                params: { document_revision_id: -1 },
              }));

          // Use revision_id from delete operation to reduce race condition risk
          const newRevisionId = deleteResponse.data?.document_revision_id || -1;

          // Insert new blocks at same position
          response =
            userAccessToken && params.useUAT
              ? await client.docx.v1.documentBlockChildren.create(
                  {
                    path: { document_id: documentId, block_id: parentId },
                    data: { children: blocks, index: blockIndex },
                    params: { document_revision_id: newRevisionId },
                  },
                  lark.withUserAccessToken(userAccessToken),
                )
              : await client.docx.v1.documentBlockChildren.create({
                  path: { document_id: documentId, block_id: parentId },
                  data: { children: blocks, index: blockIndex },
                  params: { document_revision_id: newRevisionId },
                });
          break;
        }

        case 'clear_and_write': {
          // Get all children (with pagination)
          const existingChildren = await getAllBlockChildren(client, documentId, documentId, userAccessToken, params.useUAT);

          // Delete all existing children (if any)
          if (existingChildren.length > 0) {
            await (userAccessToken && params.useUAT
              ? client.docx.v1.documentBlockChildren.batchDelete(
                  {
                    path: { document_id: documentId, block_id: documentId },
                    data: { start_index: 0, end_index: existingChildren.length },
                    params: { document_revision_id: -1 },
                  },
                  lark.withUserAccessToken(userAccessToken),
                )
              : client.docx.v1.documentBlockChildren.batchDelete({
                  path: { document_id: documentId, block_id: documentId },
                  data: { start_index: 0, end_index: existingChildren.length },
                  params: { document_revision_id: -1 },
                }));
          }

          // Write new content
          response =
            userAccessToken && params.useUAT
              ? await client.docx.v1.documentBlockChildren.create(
                  {
                    path: { document_id: documentId, block_id: documentId },
                    data: { children: blocks, index: 0 },
                    params: { document_revision_id: -1 },
                  },
                  lark.withUserAccessToken(userAccessToken),
                )
              : await client.docx.v1.documentBlockChildren.create({
                  path: { document_id: documentId, block_id: documentId },
                  data: { children: blocks, index: 0 },
                  params: { document_revision_id: -1 },
                });
          break;
        }
      }

      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify({
              success: true,
              message: 'Document edited successfully',
              document_id: documentId,
              mode,
              blocks_created: blocks.length,
              data: response?.data,
            }),
          },
        ],
      };
    } catch (error) {
      return {
        isError: true,
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify({
              success: false,
              error: (error as any)?.response?.data || (error as any)?.message || error,
            }),
          },
        ],
      };
    }
  },
};

export const docxEditTools = [larkDocxUpdateTitleTool, larkDocxAppendTool, larkDocxReplaceTool, larkDocxEditTool];
