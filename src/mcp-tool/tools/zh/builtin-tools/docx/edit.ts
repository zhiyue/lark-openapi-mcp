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

// 重新导出类型
export type docxEditToolName = DocxEditToolName;

// ============ 工具定义 ============

/**
 * 更新文档标题
 */
export const larkDocxUpdateTitleTool: McpTool = {
  project: 'docx',
  name: 'docx.builtin.updateTitle',
  accessTokens: ['user', 'tenant'],
  description:
    '[飞书/Lark] - 云文档-文档 - 更新文档标题 - 更新飞书文档的标题。需要提供文档 ID 或文档 URL，以及新的标题。',
  schema: {
    data: z.object({
      document_id: z
        .string()
        .describe('文档 ID 或文档 URL。支持直接传入 document_id 或完整的飞书文档链接'),
      title: z.string().describe('新的文档标题'),
    }),
    useUAT: z.boolean().describe('使用用户身份请求，否则为应用身份').optional(),
  },
  customHandler: async (client, params, options): Promise<CallToolResult> => {
    const { userAccessToken } = options || {};
    return handleUpdateTitle(client, params, userAccessToken, {
      success: '文档标题更新成功',
    });
  },
};

/**
 * 追加内容到文档
 */
export const larkDocxAppendTool: McpTool = {
  project: 'docx',
  name: 'docx.builtin.append',
  accessTokens: ['user', 'tenant'],
  description:
    '[飞书/Lark] - 云文档-文档 - 追加内容 - 在文档末尾追加内容。支持 Markdown 格式，包括标题、列表、代码块、表格、图片等。如果内容包含图片 URL，会自动上传图片。',
  schema: {
    data: z.object({
      document_id: z.string().describe('文档 ID 或文档 URL'),
      content: z
        .string()
        .describe('要追加的内容（支持 Markdown 格式，包括 **加粗**、*斜体*、[链接](url)、![图片](url)、表格等）'),
      block_type: z
        .enum(['text', 'heading1', 'heading2', 'heading3', 'bullet', 'ordered', 'code', 'quote', 'divider', 'auto'])
        .describe(
          '块类型：text-普通文本, heading1-6-标题, bullet-无序列表, ordered-有序列表, code-代码块, quote-引用, divider-分割线, auto-自动根据 Markdown 格式解析',
        )
        .default('auto'),
      code_language: z
        .string()
        .describe('代码块语言（仅当 block_type 为 code 时有效），如：javascript, python, go 等')
        .optional(),
      upload_images: z.boolean().describe('是否自动上传 Markdown 中的图片 URL（默认 true）').default(true),
    }),
    useUAT: z.boolean().describe('使用用户身份请求，否则为应用身份').optional(),
  },
  customHandler: async (client, params, options): Promise<CallToolResult> => {
    const { userAccessToken } = options || {};
    return handleAppend(client, params, userAccessToken, {
      success: '内容追加成功',
    });
  },
};

/**
 * 替换文档内容
 */
export const larkDocxReplaceTool: McpTool = {
  project: 'docx',
  name: 'docx.builtin.replace',
  accessTokens: ['user', 'tenant'],
  description:
    '[飞书/Lark] - 云文档-文档 - 替换内容 - 在文档中查找并替换指定文本。支持在整个文档或指定块中进行替换操作。',
  schema: {
    data: z.object({
      document_id: z.string().describe('文档 ID 或文档 URL'),
      search_text: z.string().describe('要搜索的文本'),
      replace_text: z.string().describe('替换后的文本'),
      block_id: z.string().describe('指定要替换的块 ID。如果不指定，则在整个文档中搜索替换').optional(),
      replace_all: z.boolean().describe('是否替换所有匹配项。默认为 true').default(true),
    }),
    useUAT: z.boolean().describe('使用用户身份请求，否则为应用身份').optional(),
  },
  customHandler: async (client, params, options): Promise<CallToolResult> => {
    const { userAccessToken } = options || {};
    return handleReplace(client, params, userAccessToken, {
      success: '替换成功',
      notFound: '未找到匹配的文本',
      emptySearchText: 'search_text 不能为空',
    });
  },
};

/**
 * 简化的 Markdown 编辑接口
 */
export const larkDocxEditTool: McpTool = {
  project: 'docx',
  name: 'docx.builtin.edit',
  accessTokens: ['user', 'tenant'],
  description:
    '[飞书/Lark] - 云文档-文档 - 编辑文档 - 使用 Markdown 格式编辑飞书文档。支持在指定位置插入内容、替换指定块的内容，或清空文档后写入新内容。支持完整的 Markdown 语法，包括标题、列表、代码块、表格、图片（自动上传）等。注意：replace_block 和 clear_and_write 模式涉及多个 API 调用，非原子操作。',
  schema: {
    data: z.object({
      document_id: z.string().describe('文档 ID 或文档 URL'),
      markdown: z
        .string()
        .describe('Markdown 格式的内容（支持 **加粗**、*斜体*、[链接](url)、![图片](url)、表格等）'),
      mode: z
        .enum(['append', 'prepend', 'replace_block', 'clear_and_write'])
        .describe(
          '编辑模式：append-追加到文档末尾, prepend-插入到文档开头, replace_block-替换指定块, clear_and_write-清空文档后写入',
        )
        .default('append'),
      block_id: z.string().describe('要替换的块 ID（仅在 mode 为 replace_block 时需要）').optional(),
      upload_images: z.boolean().describe('是否自动上传 Markdown 中的图片 URL（默认 true）').default(true),
    }),
    useUAT: z.boolean().describe('使用用户身份请求，否则为应用身份').optional(),
  },
  customHandler: async (client, params, options): Promise<CallToolResult> => {
    const { userAccessToken } = options || {};
    return handleEdit(client, params, userAccessToken, {
      success: '文档编辑成功',
      emptyContent: 'Markdown 内容为空或无法解析',
      blockIdRequired: 'replace_block 模式需要指定 block_id',
    });
  },
};

export const docxEditTools = [larkDocxUpdateTitleTool, larkDocxAppendTool, larkDocxReplaceTool, larkDocxEditTool];
