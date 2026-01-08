import { McpTool } from '../../../../types';
import * as lark from '@larksuiteoapi/node-sdk';
import { z } from 'zod';
import { CallToolResult } from '@modelcontextprotocol/sdk/types';
import {
  extractDocumentId,
  BlockType,
  markdownToBlocks,
  markdownToBlocksAsync,
  textToBlock,
  parseInlineMarkdown,
  DocumentBlock,
  TextElement,
  TextContent,
  CODE_LANGUAGE_MAP,
} from '../../../../utils';

// 工具名称类型
export type docxEditToolName =
  | 'docx.builtin.updateTitle'
  | 'docx.builtin.append'
  | 'docx.builtin.replace'
  | 'docx.builtin.edit';

// ============ 本地类型定义（API 返回类型） ============

/** 块列表项（API 返回） */
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

/** 更新请求 */
interface UpdateRequest {
  block_id: string;
  update_text_elements: {
    elements: TextElement[];
  };
}

// ============ 辅助函数 ============

/**
 * 分页获取文档的所有块
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
              params: { page_size: PAGE_SIZE, document_revision_id: -1, page_token: pageToken },
            },
            lark.withUserAccessToken(userAccessToken),
          )
        : await client.docx.v1.documentBlock.list({
            path: { document_id: documentId },
            params: { page_size: PAGE_SIZE, document_revision_id: -1, page_token: pageToken },
          });

    const items = (response.data?.items || []) as BlockItem[];
    allItems.push(...items);
    pageToken = response.data?.page_token;
  } while (pageToken);

  return allItems;
}

/**
 * 分页获取块的所有子块
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
              params: { document_revision_id: -1, page_size: PAGE_SIZE, page_token: pageToken },
            },
            lark.withUserAccessToken(userAccessToken),
          )
        : await client.docx.v1.documentBlockChildren.get({
            path: { document_id: documentId, block_id: blockId },
            params: { document_revision_id: -1, page_size: PAGE_SIZE, page_token: pageToken },
          });

    const items = (response.data?.items || []) as BlockItem[];
    allItems.push(...items);
    pageToken = response.data?.page_token;
  } while (pageToken);

  return allItems;
}

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
      document_id: z.string().describe('文档 ID 或文档 URL。支持直接传入 document_id 或完整的飞书文档链接'),
      title: z.string().describe('新的文档标题'),
    }),
    useUAT: z.boolean().describe('使用用户身份请求，否则为应用身份').optional(),
  },
  customHandler: async (client, params, options): Promise<CallToolResult> => {
    try {
      const { userAccessToken } = options || {};
      const documentId = extractDocumentId(params.data.document_id);

      const response =
        userAccessToken && params.useUAT
          ? await client.docx.v1.documentBlock.patch(
              {
                path: { document_id: documentId, block_id: documentId },
                data: { update_text_elements: { elements: [{ text_run: { content: params.data.title } }] } },
                params: { document_revision_id: -1 },
              },
              lark.withUserAccessToken(userAccessToken),
            )
          : await client.docx.v1.documentBlock.patch({
              path: { document_id: documentId, block_id: documentId },
              data: { update_text_elements: { elements: [{ text_run: { content: params.data.title } }] } },
              params: { document_revision_id: -1 },
            });

      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify({
              success: true,
              message: '文档标题更新成功',
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
      content: z.string().describe('要追加的内容（支持 Markdown 格式，包括 **加粗**、*斜体*、[链接](url)、![图片](url)、表格等）'),
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
    try {
      const { userAccessToken } = options || {};
      const documentId = extractDocumentId(params.data.document_id);

      let children: DocumentBlock[];

      if (params.data.block_type === 'auto') {
        // 自动解析 Markdown，支持图片上传
        if (params.data.upload_images) {
          children = await markdownToBlocksAsync(params.data.content, {
            uploadImages: true,
            client,
            documentId,
            userAccessToken,
            useUAT: params.useUAT,
          });
        } else {
          children = markdownToBlocks(params.data.content);
        }
      } else {
        // 根据指定的 block_type 创建单个块
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
          children = [{ block_type: BlockType.Divider }];
        } else if (blockType === BlockType.Code) {
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

      // 获取文档所有块以确定插入位置
      const items = await getAllDocumentBlocks(client, documentId, userAccessToken, params.useUAT);
      const insertIndex = items.length > 0 ? items.length - 1 : 0;

      // 创建子块
      const response =
        userAccessToken && params.useUAT
          ? await client.docx.v1.documentBlockChildren.create(
              {
                path: { document_id: documentId, block_id: documentId },
                data: { children, index: insertIndex },
                params: { document_revision_id: -1 },
              },
              lark.withUserAccessToken(userAccessToken),
            )
          : await client.docx.v1.documentBlockChildren.create({
              path: { document_id: documentId, block_id: documentId },
              data: { children, index: insertIndex },
              params: { document_revision_id: -1 },
            });

      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify({
              success: true,
              message: '内容追加成功',
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
    try {
      const { userAccessToken } = options || {};
      const documentId = extractDocumentId(params.data.document_id);
      const { search_text, replace_text, block_id, replace_all } = params.data;

      // 验证 search_text 不为空
      if (!search_text) {
        return {
          isError: true,
          content: [
            {
              type: 'text' as const,
              text: JSON.stringify({
                success: false,
                message: 'search_text 不能为空',
              }),
            },
          ],
        };
      }

      // 获取文档所有块
      const items = await getAllDocumentBlocks(client, documentId, userAccessToken, params.useUAT);
      const updateRequests: UpdateRequest[] = [];
      let replacedCount = 0;

      for (const block of items) {
        if (block_id && block.block_id !== block_id) continue;

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
              const escapedText = search_text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
              const matches = content.match(new RegExp(escapedText, 'g')) || [];
              replacedCount += replace_all ? matches.length : matches.length > 0 ? 1 : 0;
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
            update_text_elements: { elements: newElements },
          });
        }

        if (!replace_all && replacedCount > 0) break;
      }

      if (updateRequests.length === 0) {
        return {
          content: [
            {
              type: 'text' as const,
              text: JSON.stringify({
                success: true,
                message: '未找到匹配的文本',
                document_id: documentId,
                search_text,
                replaced_count: 0,
              }),
            },
          ],
        };
      }

      // 批量更新块
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
              message: '替换成功',
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
      markdown: z.string().describe('Markdown 格式的内容（支持 **加粗**、*斜体*、[链接](url)、![图片](url)、表格等）'),
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
    try {
      const { userAccessToken } = options || {};
      const documentId = extractDocumentId(params.data.document_id);
      const { markdown, mode, block_id } = params.data;

      // 将 Markdown 转换为文档块（支持图片上传）
      const blocks = params.data.upload_images
        ? await markdownToBlocksAsync(markdown, {
            uploadImages: true,
            client,
            documentId,
            userAccessToken,
            useUAT: params.useUAT,
          })
        : markdownToBlocks(markdown);

      if (blocks.length === 0) {
        return {
          content: [
            {
              type: 'text' as const,
              text: JSON.stringify({ success: false, message: 'Markdown 内容为空或无法解析' }),
            },
          ],
        };
      }

      let response: { data?: unknown } | undefined;

      switch (mode) {
        case 'append': {
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
                  text: JSON.stringify({ success: false, message: 'replace_block 模式需要指定 block_id' }),
                },
              ],
            };
          }

          // 获取块信息以确定父块和位置
          const blockInfo =
            userAccessToken && params.useUAT
              ? await client.docx.v1.documentBlock.get(
                  { path: { document_id: documentId, block_id }, params: { document_revision_id: -1 } },
                  lark.withUserAccessToken(userAccessToken),
                )
              : await client.docx.v1.documentBlock.get({
                  path: { document_id: documentId, block_id },
                  params: { document_revision_id: -1 },
                });

          const parentId = blockInfo.data?.block?.parent_id || documentId;
          const children = await getAllBlockChildren(client, documentId, parentId, userAccessToken, params.useUAT);
          const blockIndex = children.findIndex((c: BlockItem) => c.block_id === block_id);

          // 删除原块并获取新的 revision_id
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

          const newRevisionId = deleteResponse.data?.document_revision_id || -1;

          // 在同一位置插入新块
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
          const existingChildren = await getAllBlockChildren(client, documentId, documentId, userAccessToken, params.useUAT);

          // 删除所有现有子块
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

          // 写入新内容
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
              message: '文档编辑成功',
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
