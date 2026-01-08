import { McpTool } from '../../../../types';
import * as lark from '@larksuiteoapi/node-sdk';
import { z } from 'zod';
import { CallToolResult } from '@modelcontextprotocol/sdk/types';

// 工具名称类型
export type docxEditToolName =
  | 'docx.builtin.updateTitle'
  | 'docx.builtin.append'
  | 'docx.builtin.replace'
  | 'docx.builtin.edit';

// ============ 类型定义 ============

/** 文本元素样式 */
interface TextElementStyle {
  bold?: boolean;
  italic?: boolean;
  strikethrough?: boolean;
  underline?: boolean;
  inline_code?: boolean;
  background_color?: number;
  text_color?: number;
  link?: { url: string };
}

/** 文本元素 */
interface TextElement {
  text_run?: {
    content: string;
    text_element_style?: TextElementStyle;
  };
  mention_user?: {
    user_id: string;
    text_element_style?: TextElementStyle;
  };
}

/** 文本块样式 */
interface TextStyle {
  align?: number;
  done?: boolean;
  folded?: boolean;
  language?: number;
  wrap?: boolean;
}

/** 文本块内容 */
interface TextContent {
  style?: TextStyle;
  elements: TextElement[];
}

/** 代码块内容 */
interface CodeContent {
  style?: {
    language?: number;
    wrap?: boolean;
  };
  elements: TextElement[];
}

/** 文档块类型 */
interface DocumentBlock {
  block_type: number;
  text?: TextContent;
  code?: CodeContent;
}

/** 块列表项（API 返回） */
interface BlockItem {
  block_id?: string;
  parent_id?: string;
  block_type?: number;
  text?: TextContent;
  code?: CodeContent;
}

/** 更新请求 */
interface UpdateRequest {
  block_id: string;
  update_text_elements: {
    elements: TextElement[];
  };
}

/**
 * 辅助函数：从 URL 或直接 ID 中提取 document_id
 */
function extractDocumentId(documentIdOrUrl: string): string {
  // 如果是 URL，提取 document_id
  const urlMatch = documentIdOrUrl.match(/\/docx\/([a-zA-Z0-9]+)/);
  if (urlMatch) {
    return urlMatch[1];
  }
  // Wiki URL 格式
  const wikiMatch = documentIdOrUrl.match(/\/wiki\/([a-zA-Z0-9]+)/);
  if (wikiMatch) {
    return wikiMatch[1];
  }
  // 否则直接返回
  return documentIdOrUrl;
}

/**
 * 辅助函数：将简单文本转换为文档块结构
 */
function textToBlock(text: string, blockType: number = 2): DocumentBlock {
  return {
    block_type: blockType,
    text: {
      elements: [
        {
          text_run: {
            content: text,
          },
        },
      ],
    },
  };
}

/**
 * 辅助函数：将 Markdown 转换为文档块数组
 */
function markdownToBlocks(markdown: string): DocumentBlock[] {
  const lines = markdown.split('\n');
  const blocks: DocumentBlock[] = [];
  let codeBlock: string[] | null = null;
  let codeLanguage = 1; // PlainText

  const languageMap: Record<string, number> = {
    plaintext: 1,
    abap: 2,
    ada: 3,
    apache: 4,
    apex: 5,
    assembly: 6,
    bash: 7,
    csharp: 8,
    'c++': 9,
    cpp: 9,
    c: 10,
    cobol: 11,
    css: 12,
    coffeescript: 13,
    d: 14,
    dart: 15,
    delphi: 16,
    django: 17,
    dockerfile: 18,
    erlang: 19,
    fortran: 20,
    go: 22,
    groovy: 23,
    html: 24,
    http: 26,
    haskell: 27,
    json: 28,
    java: 29,
    javascript: 30,
    js: 30,
    julia: 31,
    kotlin: 32,
    latex: 33,
    lisp: 34,
    lua: 36,
    matlab: 37,
    makefile: 38,
    markdown: 39,
    nginx: 40,
    'objective-c': 41,
    objc: 41,
    php: 43,
    perl: 44,
    powershell: 46,
    prolog: 47,
    protobuf: 48,
    python: 49,
    py: 49,
    r: 50,
    ruby: 52,
    rust: 53,
    sas: 54,
    scss: 55,
    sql: 56,
    scala: 57,
    scheme: 58,
    shell: 60,
    swift: 61,
    thrift: 62,
    typescript: 63,
    ts: 63,
    vbscript: 64,
    vb: 65,
    xml: 66,
    yaml: 67,
    yml: 67,
    cmake: 68,
    diff: 69,
    graphql: 71,
    toml: 75,
  };

  for (const line of lines) {
    // 处理代码块
    if (line.startsWith('```')) {
      if (codeBlock === null) {
        // 开始代码块
        codeBlock = [];
        const lang = line.slice(3).trim().toLowerCase();
        codeLanguage = languageMap[lang] || 1;
      } else {
        // 结束代码块
        blocks.push({
          block_type: 14, // Code
          code: {
            style: {
              language: codeLanguage,
              wrap: false,
            },
            elements: [
              {
                text_run: {
                  content: codeBlock.join('\n'),
                },
              },
            ],
          },
        });
        codeBlock = null;
      }
      continue;
    }

    if (codeBlock !== null) {
      codeBlock.push(line);
      continue;
    }

    // 空行跳过
    if (line.trim() === '') {
      continue;
    }

    // 处理标题
    const h1Match = line.match(/^# (.+)$/);
    if (h1Match) {
      blocks.push(textToBlock(h1Match[1], 3)); // Heading1
      continue;
    }

    const h2Match = line.match(/^## (.+)$/);
    if (h2Match) {
      blocks.push(textToBlock(h2Match[1], 4)); // Heading2
      continue;
    }

    const h3Match = line.match(/^### (.+)$/);
    if (h3Match) {
      blocks.push(textToBlock(h3Match[1], 5)); // Heading3
      continue;
    }

    const h4Match = line.match(/^#### (.+)$/);
    if (h4Match) {
      blocks.push(textToBlock(h4Match[1], 6)); // Heading4
      continue;
    }

    const h5Match = line.match(/^##### (.+)$/);
    if (h5Match) {
      blocks.push(textToBlock(h5Match[1], 7)); // Heading5
      continue;
    }

    const h6Match = line.match(/^###### (.+)$/);
    if (h6Match) {
      blocks.push(textToBlock(h6Match[1], 8)); // Heading6
      continue;
    }

    // 处理无序列表
    const bulletMatch = line.match(/^[-*+] (.+)$/);
    if (bulletMatch) {
      blocks.push(textToBlock(bulletMatch[1], 12)); // Bullet
      continue;
    }

    // 处理有序列表
    const orderedMatch = line.match(/^\d+\. (.+)$/);
    if (orderedMatch) {
      blocks.push(textToBlock(orderedMatch[1], 13)); // Ordered
      continue;
    }

    // 处理引用
    const quoteMatch = line.match(/^> (.+)$/);
    if (quoteMatch) {
      blocks.push(textToBlock(quoteMatch[1], 15)); // Quote
      continue;
    }

    // 处理待办事项
    const todoUncheckedMatch = line.match(/^- \[ \] (.+)$/);
    if (todoUncheckedMatch) {
      blocks.push({
        block_type: 17, // Todo
        text: {
          style: { done: false },
          elements: [{ text_run: { content: todoUncheckedMatch[1] } }],
        },
      });
      continue;
    }

    const todoCheckedMatch = line.match(/^- \[x\] (.+)$/i);
    if (todoCheckedMatch) {
      blocks.push({
        block_type: 17, // Todo
        text: {
          style: { done: true },
          elements: [{ text_run: { content: todoCheckedMatch[1] } }],
        },
      });
      continue;
    }

    // 处理分割线
    if (/^[-*_]{3,}$/.test(line.trim())) {
      blocks.push({ block_type: 22 }); // Divider
      continue;
    }

    // 默认处理为普通文本
    blocks.push(textToBlock(line, 2)); // Text
  }

  // 处理未闭合的代码块
  if (codeBlock !== null) {
    blocks.push({
      block_type: 14, // Code
      code: {
        style: {
          language: codeLanguage,
          wrap: false,
        },
        elements: [
          {
            text_run: {
              content: codeBlock.join('\n'),
            },
          },
        ],
      },
    });
  }

  return blocks;
}

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

      // 使用 documentBlock.patch 更新文档标题（Page 块的 body）
      const response =
        userAccessToken && params.useUAT
          ? await client.docx.v1.documentBlock.patch(
              {
                path: {
                  document_id: documentId,
                  block_id: documentId, // 文档根块 ID 等于 document_id
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
                  document_revision_id: -1, // 使用最新版本
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
    '[飞书/Lark] - 云文档-文档 - 追加内容 - 在文档末尾追加内容。支持追加文本、标题、列表等多种类型的内容块。',
  schema: {
    data: z.object({
      document_id: z.string().describe('文档 ID 或文档 URL'),
      content: z.string().describe('要追加的内容（支持 Markdown 格式）'),
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
    }),
    useUAT: z.boolean().describe('使用用户身份请求，否则为应用身份').optional(),
  },
  customHandler: async (client, params, options): Promise<CallToolResult> => {
    try {
      const { userAccessToken } = options || {};
      const documentId = extractDocumentId(params.data.document_id);

      let children: DocumentBlock[];

      if (params.data.block_type === 'auto') {
        // 自动解析 Markdown
        children = markdownToBlocks(params.data.content);
      } else {
        // 根据指定的 block_type 创建单个块
        const blockTypeMap: Record<string, number> = {
          text: 2,
          heading1: 3,
          heading2: 4,
          heading3: 5,
          bullet: 12,
          ordered: 13,
          code: 14,
          quote: 15,
          divider: 22,
        };

        const blockType = blockTypeMap[params.data.block_type] || 2;

        if (blockType === 22) {
          // 分割线
          children = [{ block_type: 22 }];
        } else if (blockType === 14) {
          // 代码块
          const languageMap: Record<string, number> = {
            javascript: 30,
            js: 30,
            typescript: 63,
            ts: 63,
            python: 49,
            py: 49,
            go: 22,
            java: 29,
            rust: 53,
            cpp: 9,
            c: 10,
            sql: 56,
            json: 28,
            yaml: 67,
            shell: 60,
            bash: 7,
          };
          const language = languageMap[(params.data.code_language || '').toLowerCase()] || 1;
          children = [
            {
              block_type: 14,
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

      // 获取文档块列表以确定插入位置
      const blocksResponse =
        userAccessToken && params.useUAT
          ? await client.docx.v1.documentBlock.list(
              {
                path: { document_id: documentId },
                params: { page_size: 500, document_revision_id: -1 },
              },
              lark.withUserAccessToken(userAccessToken),
            )
          : await client.docx.v1.documentBlock.list({
              path: { document_id: documentId },
              params: { page_size: 500, document_revision_id: -1 },
            });

      const items = blocksResponse.data?.items || [];
      const insertIndex = items.length > 0 ? items.length - 1 : 0;

      // 创建子块
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

      // 获取文档块列表
      const blocksResponse =
        userAccessToken && params.useUAT
          ? await client.docx.v1.documentBlock.list(
              {
                path: { document_id: documentId },
                params: { page_size: 500, document_revision_id: -1 },
              },
              lark.withUserAccessToken(userAccessToken),
            )
          : await client.docx.v1.documentBlock.list({
              path: { document_id: documentId },
              params: { page_size: 500, document_revision_id: -1 },
            });

      const items = blocksResponse.data?.items || [];
      const updateRequests: UpdateRequest[] = [];
      let replacedCount = 0;

      for (const block of items) {
        // 如果指定了 block_id，只处理该块
        if (block_id && block.block_id !== block_id) {
          continue;
        }

        // 获取块的文本内容
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
              // 正确计算替换数量：replace_all 时计算所有匹配，否则只计 1
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

        // 如果不是替换所有，且已经找到一个匹配，就停止
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
    '[飞书/Lark] - 云文档-文档 - 编辑文档 - 使用 Markdown 格式编辑飞书文档。支持在指定位置插入内容、替换指定块的内容，或清空文档后写入新内容。注意：replace_block 和 clear_and_write 模式涉及多个 API 调用，非原子操作，在高并发编辑场景下可能存在竞态条件。',
  schema: {
    data: z.object({
      document_id: z.string().describe('文档 ID 或文档 URL'),
      markdown: z.string().describe('Markdown 格式的内容'),
      mode: z
        .enum(['append', 'prepend', 'replace_block', 'clear_and_write'])
        .describe(
          '编辑模式：append-追加到文档末尾, prepend-插入到文档开头, replace_block-替换指定块, clear_and_write-清空文档后写入',
        )
        .default('append'),
      block_id: z.string().describe('要替换的块 ID（仅在 mode 为 replace_block 时需要）').optional(),
    }),
    useUAT: z.boolean().describe('使用用户身份请求，否则为应用身份').optional(),
  },
  customHandler: async (client, params, options): Promise<CallToolResult> => {
    try {
      const { userAccessToken } = options || {};
      const documentId = extractDocumentId(params.data.document_id);
      const { markdown, mode, block_id } = params.data;

      // 将 Markdown 转换为文档块
      const blocks = markdownToBlocks(markdown);

      if (blocks.length === 0) {
        return {
          content: [
            {
              type: 'text' as const,
              text: JSON.stringify({
                success: false,
                message: 'Markdown 内容为空或无法解析',
              }),
            },
          ],
        };
      }

      let response: { data?: unknown } | undefined;

      switch (mode) {
        case 'append': {
          // 获取文档块列表以确定插入位置
          const blocksResponse =
            userAccessToken && params.useUAT
              ? await client.docx.v1.documentBlock.list(
                  {
                    path: { document_id: documentId },
                    params: { page_size: 500, document_revision_id: -1 },
                  },
                  lark.withUserAccessToken(userAccessToken),
                )
              : await client.docx.v1.documentBlock.list({
                  path: { document_id: documentId },
                  params: { page_size: 500, document_revision_id: -1 },
                });

          const items = blocksResponse.data?.items || [];
          const insertIndex = items.length > 0 ? items.length - 1 : 0;

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
                    message: 'replace_block 模式需要指定 block_id',
                  }),
                },
              ],
            };
          }

          // 先删除原块，再在同一位置插入新块
          // 获取块信息以确定父块和位置
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

          // 获取父块的子块列表以确定位置
          const childrenResponse =
            userAccessToken && params.useUAT
              ? await client.docx.v1.documentBlockChildren.get(
                  {
                    path: { document_id: documentId, block_id: parentId },
                    params: { document_revision_id: -1, page_size: 500 },
                  },
                  lark.withUserAccessToken(userAccessToken),
                )
              : await client.docx.v1.documentBlockChildren.get({
                  path: { document_id: documentId, block_id: parentId },
                  params: { document_revision_id: -1, page_size: 500 },
                });

          const children = childrenResponse.data?.items || [];
          const blockIndex = children.findIndex((c: BlockItem) => c.block_id === block_id);

          // 删除原块，并获取新的 document_revision_id
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

          // 使用删除操作返回的 revision_id 进行插入，减少竞态条件风险
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
          // 获取所有子块
          const childrenResponse =
            userAccessToken && params.useUAT
              ? await client.docx.v1.documentBlockChildren.get(
                  {
                    path: { document_id: documentId, block_id: documentId },
                    params: { document_revision_id: -1, page_size: 500 },
                  },
                  lark.withUserAccessToken(userAccessToken),
                )
              : await client.docx.v1.documentBlockChildren.get({
                  path: { document_id: documentId, block_id: documentId },
                  params: { document_revision_id: -1, page_size: 500 },
                });

          const existingChildren = childrenResponse.data?.items || [];

          // 删除所有现有子块（如果有的话）
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
