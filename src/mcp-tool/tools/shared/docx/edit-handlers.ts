import * as lark from '@larksuiteoapi/node-sdk';
import { CallToolResult } from '@modelcontextprotocol/sdk/types';
import {
  extractDocumentId,
  BlockType,
  markdownToBlocks,
  markdownToBlocksAsync,
  textToBlock,
  DocumentBlock,
  TextElement,
  CODE_LANGUAGE_MAP,
  extractErrorMessage,
} from '../../../utils';
import { BlockItem, UpdateRequest } from './edit-types';

// ============ Helper Functions ============

const PAGE_SIZE = 500;

/**
 * Get all document blocks with pagination
 */
export async function getAllDocumentBlocks(
  client: lark.Client,
  documentId: string,
  userAccessToken?: string,
  useUAT?: boolean,
): Promise<BlockItem[]> {
  const allItems: BlockItem[] = [];
  let pageToken: string | undefined;

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
 * Get all block children with pagination
 */
export async function getAllBlockChildren(
  client: lark.Client,
  documentId: string,
  blockId: string,
  userAccessToken?: string,
  useUAT?: boolean,
): Promise<BlockItem[]> {
  const allItems: BlockItem[] = [];
  let pageToken: string | undefined;

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

// ============ Response Builders ============

function successResponse(data: Record<string, unknown>): CallToolResult {
  return {
    content: [{ type: 'text' as const, text: JSON.stringify({ success: true, ...data }) }],
  };
}

function errorResponse(error: unknown): CallToolResult {
  return {
    isError: true,
    content: [{ type: 'text' as const, text: JSON.stringify({ success: false, error: extractErrorMessage(error) }) }],
  };
}

function validationErrorResponse(message: string): CallToolResult {
  return {
    isError: true,
    content: [{ type: 'text' as const, text: JSON.stringify({ success: false, message }) }],
  };
}

// ============ Handler Implementations ============

interface UpdateTitleParams {
  data: { document_id: string; title: string };
  useUAT?: boolean;
}

interface UpdateTitleMessages {
  success: string;
}

export async function handleUpdateTitle(
  client: lark.Client,
  params: UpdateTitleParams,
  userAccessToken?: string,
  messages: UpdateTitleMessages = { success: 'Document title updated successfully' },
): Promise<CallToolResult> {
  try {
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

    return successResponse({
      message: messages.success,
      document_id: documentId,
      new_title: params.data.title,
      data: response.data,
    });
  } catch (error) {
    return errorResponse(error);
  }
}

interface AppendParams {
  data: {
    document_id: string;
    content: string;
    block_type: string;
    code_language?: string;
    upload_images?: boolean;
  };
  useUAT?: boolean;
}

interface AppendMessages {
  success: string;
}

export async function handleAppend(
  client: lark.Client,
  params: AppendParams,
  userAccessToken?: string,
  messages: AppendMessages = { success: 'Content appended successfully' },
): Promise<CallToolResult> {
  try {
    const documentId = extractDocumentId(params.data.document_id);
    let children: DocumentBlock[];

    if (params.data.block_type === 'auto') {
      if (params.data.upload_images !== false) {
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

    const items = await getAllDocumentBlocks(client, documentId, userAccessToken, params.useUAT);
    const insertIndex = items.length > 0 ? items.length - 1 : 0;

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

    return successResponse({
      message: messages.success,
      document_id: documentId,
      blocks_created: children.length,
      data: response.data,
    });
  } catch (error) {
    return errorResponse(error);
  }
}

interface ReplaceParams {
  data: {
    document_id: string;
    search_text: string;
    replace_text: string;
    block_id?: string;
    replace_all?: boolean;
  };
  useUAT?: boolean;
}

interface ReplaceMessages {
  success: string;
  notFound: string;
  emptySearchText: string;
}

export async function handleReplace(
  client: lark.Client,
  params: ReplaceParams,
  userAccessToken?: string,
  messages: ReplaceMessages = {
    success: 'Replacement successful',
    notFound: 'No matching text found',
    emptySearchText: 'search_text cannot be empty',
  },
): Promise<CallToolResult> {
  try {
    const documentId = extractDocumentId(params.data.document_id);
    const { search_text, replace_text, block_id, replace_all = true } = params.data;

    if (!search_text) {
      return validationErrorResponse(messages.emptySearchText);
    }

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
      return successResponse({
        message: messages.notFound,
        document_id: documentId,
        search_text,
        replaced_count: 0,
      });
    }

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

    return successResponse({
      message: messages.success,
      document_id: documentId,
      search_text,
      replace_text,
      replaced_count: replacedCount,
      blocks_updated: updateRequests.length,
      data: response.data,
    });
  } catch (error) {
    return errorResponse(error);
  }
}

interface EditParams {
  data: {
    document_id: string;
    markdown: string;
    mode: 'append' | 'prepend' | 'replace_block' | 'clear_and_write';
    block_id?: string;
    upload_images?: boolean;
  };
  useUAT?: boolean;
}

interface EditMessages {
  success: string;
  emptyContent: string;
  blockIdRequired: string;
}

export async function handleEdit(
  client: lark.Client,
  params: EditParams,
  userAccessToken?: string,
  messages: EditMessages = {
    success: 'Document edited successfully',
    emptyContent: 'Markdown content is empty or cannot be parsed',
    blockIdRequired: 'replace_block mode requires block_id',
  },
): Promise<CallToolResult> {
  try {
    const documentId = extractDocumentId(params.data.document_id);
    const { markdown, mode, block_id } = params.data;

    const blocks =
      params.data.upload_images !== false
        ? await markdownToBlocksAsync(markdown, {
            uploadImages: true,
            client,
            documentId,
            userAccessToken,
            useUAT: params.useUAT,
          })
        : markdownToBlocks(markdown);

    if (blocks.length === 0) {
      return validationErrorResponse(messages.emptyContent);
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
          return validationErrorResponse(messages.blockIdRequired);
        }

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
        const existingChildren = await getAllBlockChildren(
          client,
          documentId,
          documentId,
          userAccessToken,
          params.useUAT,
        );

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

    return successResponse({
      message: messages.success,
      document_id: documentId,
      mode,
      blocks_created: blocks.length,
      data: response?.data,
    });
  } catch (error) {
    return errorResponse(error);
  }
}
