import { McpTool } from '../../../../types';
import { z } from 'zod';
import { CallToolResult } from '@modelcontextprotocol/sdk/types';
import {
  DocxEditToolName,
  handleUpdateTitle,
  handleAppend,
  handleReplace,
  handleEdit,
} from '../../../shared/docx';

// Re-export type
export type docxEditToolName = DocxEditToolName;

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
      document_id: z
        .string()
        .describe('Document ID or URL. Supports direct document_id or full Feishu document link'),
      title: z.string().describe('New document title'),
    }),
    useUAT: z
      .boolean()
      .describe('Use user identity for the request, otherwise use application identity')
      .optional(),
  },
  customHandler: async (client, params, options): Promise<CallToolResult> => {
    const { userAccessToken } = options || {};
    return handleUpdateTitle(client, params, userAccessToken, {
      success: 'Document title updated successfully',
    });
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
    '[Feishu/Lark]-Docs-Document-Append Content-Append content to the end of a document. Supports Markdown format including headings, lists, code blocks, tables, images, etc. Images in Markdown will be automatically uploaded.',
  schema: {
    data: z.object({
      document_id: z.string().describe('Document ID or URL'),
      content: z
        .string()
        .describe(
          'Content to append (supports Markdown format including **bold**, *italic*, [link](url), ![image](url), tables, etc.)',
        ),
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
      upload_images: z
        .boolean()
        .describe('Whether to automatically upload image URLs in Markdown (default true)')
        .default(true),
    }),
    useUAT: z
      .boolean()
      .describe('Use user identity for the request, otherwise use application identity')
      .optional(),
  },
  customHandler: async (client, params, options): Promise<CallToolResult> => {
    const { userAccessToken } = options || {};
    return handleAppend(client, params, userAccessToken, {
      success: 'Content appended successfully',
    });
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
      block_id: z
        .string()
        .describe('Block ID to replace in. If not specified, searches the entire document')
        .optional(),
      replace_all: z.boolean().describe('Whether to replace all matches. Default is true').default(true),
    }),
    useUAT: z
      .boolean()
      .describe('Use user identity for the request, otherwise use application identity')
      .optional(),
  },
  customHandler: async (client, params, options): Promise<CallToolResult> => {
    const { userAccessToken } = options || {};
    return handleReplace(client, params, userAccessToken, {
      success: 'Replacement successful',
      notFound: 'No matching text found',
      emptySearchText: 'search_text cannot be empty',
    });
  },
};

/**
 * Edit document with Markdown
 */
export const larkDocxEditTool: McpTool = {
  project: 'docx',
  name: 'docx.builtin.edit',
  accessTokens: ['user', 'tenant'],
  description:
    '[Feishu/Lark]-Docs-Document-Edit Document-Edit Feishu document using Markdown format. Supports inserting content at specified position, replacing block content, or clearing and writing new content. Supports full Markdown syntax including headings, lists, code blocks, tables, images (auto-uploaded), etc. Note: replace_block and clear_and_write modes involve multiple API calls and are not atomic operations.',
  schema: {
    data: z.object({
      document_id: z.string().describe('Document ID or URL'),
      markdown: z
        .string()
        .describe(
          'Content in Markdown format (supports **bold**, *italic*, [link](url), ![image](url), tables, etc.)',
        ),
      mode: z
        .enum(['append', 'prepend', 'replace_block', 'clear_and_write'])
        .describe(
          'Edit mode: append-append to end, prepend-insert at beginning, replace_block-replace specific block, clear_and_write-clear document and write new content',
        )
        .default('append'),
      block_id: z.string().describe('Block ID to replace (required only when mode is replace_block)').optional(),
      upload_images: z
        .boolean()
        .describe('Whether to automatically upload image URLs in Markdown (default true)')
        .default(true),
    }),
    useUAT: z
      .boolean()
      .describe('Use user identity for the request, otherwise use application identity')
      .optional(),
  },
  customHandler: async (client, params, options): Promise<CallToolResult> => {
    const { userAccessToken } = options || {};
    return handleEdit(client, params, userAccessToken, {
      success: 'Document edited successfully',
      emptyContent: 'Markdown content is empty or cannot be parsed',
      blockIdRequired: 'replace_block mode requires block_id',
    });
  },
};

export const docxEditTools = [larkDocxUpdateTitleTool, larkDocxAppendTool, larkDocxReplaceTool, larkDocxEditTool];
