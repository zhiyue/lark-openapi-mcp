import { McpTool } from '../../../../types';
import * as lark from '@larksuiteoapi/node-sdk';
import { z } from 'zod';
import { CallToolResult } from '@modelcontextprotocol/sdk/types';

// Tool name type
export type docxEditToolName =
  | 'docx.builtin.updateTitle'
  | 'docx.builtin.append'
  | 'docx.builtin.replace'
  | 'docx.builtin.edit';

// ============ Type Definitions ============

/** Text element style */
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

/** Text element */
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

/** Text block style */
interface TextStyle {
  align?: number;
  done?: boolean;
  folded?: boolean;
  language?: number;
  wrap?: boolean;
}

/** Text block content */
interface TextContent {
  style?: TextStyle;
  elements: TextElement[];
}

/** Code block content */
interface CodeContent {
  style?: {
    language?: number;
    wrap?: boolean;
  };
  elements: TextElement[];
}

/** Document block type */
interface DocumentBlock {
  block_type: number;
  text?: TextContent;
  code?: CodeContent;
}

/** Block list item (API response) */
interface BlockItem {
  block_id?: string;
  parent_id?: string;
  block_type?: number;
  text?: TextContent;
  code?: CodeContent;
}

/** Update request */
interface UpdateRequest {
  block_id: string;
  update_text_elements: {
    elements: TextElement[];
  };
}

/**
 * Helper function: Extract document_id from URL or direct ID
 */
function extractDocumentId(documentIdOrUrl: string): string {
  // If it's a URL, extract document_id
  const urlMatch = documentIdOrUrl.match(/\/docx\/([a-zA-Z0-9]+)/);
  if (urlMatch) {
    return urlMatch[1];
  }
  // Wiki URL format
  const wikiMatch = documentIdOrUrl.match(/\/wiki\/([a-zA-Z0-9]+)/);
  if (wikiMatch) {
    return wikiMatch[1];
  }
  // Otherwise return directly
  return documentIdOrUrl;
}

/**
 * Helper function: Convert simple text to document block structure
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
 * Helper function: Convert Markdown to document block array
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
    // Handle code blocks
    if (line.startsWith('```')) {
      if (codeBlock === null) {
        // Start code block
        codeBlock = [];
        const lang = line.slice(3).trim().toLowerCase();
        codeLanguage = languageMap[lang] || 1;
      } else {
        // End code block
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

    // Skip empty lines
    if (line.trim() === '') {
      continue;
    }

    // Handle headings
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

    // Handle unordered list
    const bulletMatch = line.match(/^[-*+] (.+)$/);
    if (bulletMatch) {
      blocks.push(textToBlock(bulletMatch[1], 12)); // Bullet
      continue;
    }

    // Handle ordered list
    const orderedMatch = line.match(/^\d+\. (.+)$/);
    if (orderedMatch) {
      blocks.push(textToBlock(orderedMatch[1], 13)); // Ordered
      continue;
    }

    // Handle quote
    const quoteMatch = line.match(/^> (.+)$/);
    if (quoteMatch) {
      blocks.push(textToBlock(quoteMatch[1], 15)); // Quote
      continue;
    }

    // Handle todo items
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

    // Handle divider
    if (/^[-*_]{3,}$/.test(line.trim())) {
      blocks.push({ block_type: 22 }); // Divider
      continue;
    }

    // Default to plain text
    blocks.push(textToBlock(line, 2)); // Text
  }

  // Handle unclosed code block
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
          // Divider
          children = [{ block_type: 22 }];
        } else if (blockType === 14) {
          // Code block
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

      // Get document block list to determine insertion position
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

      // Get document block list
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
          // Get document block list to determine insertion position
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

          // Get parent block's children list to determine position
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
          // Get all children
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
