import { BlockType } from './block-types';
import * as lark from '@larksuiteoapi/node-sdk';
import { Readable } from 'stream';
import { ReadStream } from 'fs';

// ============ 类型定义 ============

/** 文本元素样式 */
export interface TextElementStyle {
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
export interface TextElement {
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
export interface TextStyle {
  align?: number;
  done?: boolean;
  folded?: boolean;
  language?: number;
  wrap?: boolean;
}

/** 文本块内容 */
export interface TextContent {
  style?: TextStyle;
  elements: TextElement[];
}

/** 代码块内容 */
export interface CodeContent {
  style?: {
    language?: number;
    wrap?: boolean;
  };
  elements: TextElement[];
}

/** 图片块内容 */
export interface ImageContent {
  align?: number;
  caption?: { content?: string };
  token?: string;
  width?: number;
  height?: number;
}

/** 表格块内容 */
export interface TableContent {
  property: {
    row_size: number;
    column_size: number;
  };
  cells?: DocumentBlock[];
}

/** 文档块类型 */
export interface DocumentBlock {
  block_type: number;
  text?: TextContent;
  code?: CodeContent;
  image?: ImageContent;
  table?: TableContent;
  /** 子块（用于嵌套列表等） */
  children?: DocumentBlock[];
}

// ============ 语言映射 ============

export const CODE_LANGUAGE_MAP: Record<string, number> = {
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
  htmlbars: 25,
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
  gherkin: 70,
  graphql: 71,
  properties: 73,
  solidity: 74,
  toml: 75,
};

// ============ 内联样式解析 ============

interface InlineMatch {
  start: number;
  end: number;
  type: string;
  text: string;
  url?: string;
}

/**
 * 解析内联 Markdown 样式
 * 支持: **加粗**, *斜体*, ~~删除线~~, `代码`, [链接](url)
 */
export function parseInlineMarkdown(text: string): TextElement[] {
  const elements: TextElement[] = [];

  const patterns = [
    { regex: /\[([^\]]+)\]\(([^)]+)\)/g, type: 'link' },
    { regex: /\*\*\*([^*]+)\*\*\*|___([^_]+)___/g, type: 'bold_italic' },
    { regex: /\*\*([^*]+)\*\*|__([^_]+)__/g, type: 'bold' },
    { regex: /(?<!\*)\*([^*]+)\*(?!\*)|(?<!_)_([^_]+)_(?!_)/g, type: 'italic' },
    { regex: /~~([^~]+)~~/g, type: 'strikethrough' },
    { regex: /`([^`]+)`/g, type: 'inline_code' },
  ];

  const matches: InlineMatch[] = [];

  for (const pattern of patterns) {
    let match;
    const regex = new RegExp(pattern.regex.source, pattern.regex.flags);
    while ((match = regex.exec(text)) !== null) {
      const start = match.index;
      const end = match.index + match[0].length;
      const overlaps = matches.some((m) => (start >= m.start && start < m.end) || (end > m.start && end <= m.end));

      if (!overlaps) {
        matches.push({
          start,
          end,
          type: pattern.type,
          text: pattern.type === 'link' ? match[1] : match[1] || match[2],
          url: pattern.type === 'link' ? match[2] : undefined,
        });
      }
    }
  }

  matches.sort((a, b) => a.start - b.start);

  if (matches.length === 0) {
    return [{ text_run: { content: text } }];
  }

  let lastEnd = 0;
  for (const match of matches) {
    if (match.start > lastEnd) {
      elements.push({ text_run: { content: text.slice(lastEnd, match.start) } });
    }

    const style: TextElementStyle = {};
    switch (match.type) {
      case 'bold':
        style.bold = true;
        break;
      case 'italic':
        style.italic = true;
        break;
      case 'bold_italic':
        style.bold = true;
        style.italic = true;
        break;
      case 'strikethrough':
        style.strikethrough = true;
        break;
      case 'inline_code':
        style.inline_code = true;
        break;
      case 'link':
        style.link = { url: match.url || '' };
        break;
    }

    elements.push({
      text_run: {
        content: match.text,
        text_element_style: style,
      },
    });

    lastEnd = match.end;
  }

  if (lastEnd < text.length) {
    elements.push({ text_run: { content: text.slice(lastEnd) } });
  }

  return elements;
}

// ============ 辅助函数 ============

/**
 * 将简单文本转换为文档块结构
 */
export function textToBlock(text: string, blockType: number = BlockType.Text): DocumentBlock {
  return {
    block_type: blockType,
    text: {
      elements: parseInlineMarkdown(text),
    },
  };
}

/**
 * 解析表格行
 */
export function parseTableRow(line: string): string[] {
  const trimmed = line.trim();
  const content = trimmed.startsWith('|') ? trimmed.slice(1) : trimmed;
  const withoutEnd = content.endsWith('|') ? content.slice(0, -1) : content;
  return withoutEnd.split('|').map((cell) => cell.trim());
}

/**
 * 检查是否是表格分隔行
 */
export function isTableSeparator(line: string): boolean {
  return /^\|?[\s\-:|]+\|?$/.test(line.trim());
}

/**
 * 将表格数据转换为表格块
 */
export function tableToBlock(tableRows: string[][]): DocumentBlock {
  const rowSize = tableRows.length;
  const columnSize = Math.max(...tableRows.map((row) => row.length));

  const cells: DocumentBlock[] = [];
  for (const row of tableRows) {
    for (let col = 0; col < columnSize; col++) {
      const cellContent = row[col] || '';
      cells.push({
        block_type: BlockType.Text,
        text: {
          elements: parseInlineMarkdown(cellContent),
        },
      });
    }
  }

  return {
    block_type: BlockType.Table,
    table: {
      property: {
        row_size: rowSize,
        column_size: columnSize,
      },
      cells,
    },
  };
}

// ============ 图片上传 ============

export interface ImageUploadOptions {
  client: lark.Client;
  documentId: string;
  userAccessToken?: string;
  useUAT?: boolean;
}

/**
 * 从 URL 下载图片并上传到飞书
 */
export async function uploadImageFromUrl(
  imageUrl: string,
  options: ImageUploadOptions,
): Promise<string | null> {
  try {
    const { client, documentId, userAccessToken, useUAT } = options;

    // 下载图片
    const response = await fetch(imageUrl);
    if (!response.ok) {
      console.error(`Failed to download image: ${response.statusText}`);
      return null;
    }

    const arrayBuffer = await response.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);

    // 从 URL 获取文件名
    const urlPath = new URL(imageUrl).pathname;
    const fileName = urlPath.split('/').pop() || 'image.png';

    // 创建可读流
    const stream = Readable.from(buffer) as ReadStream;

    // 上传到飞书
    const uploadData = {
      file_name: fileName,
      parent_type: 'docx_image' as const,
      parent_node: documentId,
      size: buffer.length,
      file: stream,
    };

    const uploadResponse =
      userAccessToken && useUAT
        ? await client.drive.media.uploadAll({ data: uploadData }, lark.withUserAccessToken(userAccessToken))
        : await client.drive.media.uploadAll({ data: uploadData });

    return uploadResponse?.file_token || null;
  } catch (error) {
    console.error('Failed to upload image:', error);
    return null;
  }
}

// ============ 行令牌类型 ============

/** 行类型枚举 */
enum LineType {
  Empty = 'empty',
  Heading = 'heading',
  Bullet = 'bullet',
  Ordered = 'ordered',
  Quote = 'quote',
  TodoUnchecked = 'todo_unchecked',
  TodoChecked = 'todo_checked',
  CodeFence = 'code_fence',
  CodeContent = 'code_content',
  TableRow = 'table_row',
  TableSeparator = 'table_separator',
  Divider = 'divider',
  Image = 'image',
  Text = 'text',
}

/** 行令牌 */
interface LineToken {
  type: LineType;
  content: string;
  indent: number;
  raw: string;
  // 特定类型的额外数据
  meta?: {
    level?: number; // 标题级别
    language?: string; // 代码语言
    url?: string; // 图片 URL
    alt?: string; // 图片描述
    cells?: string[]; // 表格单元格
  };
}

/** 列表项（带缩进层级） */
interface ListItem {
  block: DocumentBlock;
  indent: number;
  type: 'bullet' | 'ordered';
}

// ============ 行解析器 ============

/**
 * 计算行的缩进级别（空格数）
 */
function getIndent(line: string): number {
  const match = line.match(/^(\s*)/);
  return match ? match[1].length : 0;
}

/**
 * 解析单行为令牌
 */
function tokenizeLine(line: string, inCodeBlock: boolean): LineToken {
  const indent = getIndent(line);
  const trimmed = line.trim();
  const contentAfterIndent = line.slice(indent);

  // 代码块内容
  if (inCodeBlock && !trimmed.startsWith('```')) {
    return { type: LineType.CodeContent, content: line, indent: 0, raw: line };
  }

  // 代码块边界
  if (trimmed.startsWith('```')) {
    const language = trimmed.slice(3).trim().toLowerCase();
    return {
      type: LineType.CodeFence,
      content: '',
      indent: 0,
      raw: line,
      meta: { language },
    };
  }

  // 空行
  if (trimmed === '') {
    return { type: LineType.Empty, content: '', indent, raw: line };
  }

  // 分割线（必须在列表检测之前）
  if (/^[-*_]{3,}$/.test(trimmed) && !/^[-*+]\s/.test(trimmed)) {
    return { type: LineType.Divider, content: '', indent, raw: line };
  }

  // 标题
  const headingMatch = trimmed.match(/^(#{1,6}) (.+)$/);
  if (headingMatch) {
    return {
      type: LineType.Heading,
      content: headingMatch[2],
      indent,
      raw: line,
      meta: { level: headingMatch[1].length },
    };
  }

  // 待办事项（未完成）- 支持缩进
  const todoUncheckedMatch = contentAfterIndent.match(/^[-*+] \[ \] (.+)$/);
  if (todoUncheckedMatch) {
    return {
      type: LineType.TodoUnchecked,
      content: todoUncheckedMatch[1],
      indent,
      raw: line,
    };
  }

  // 待办事项（已完成）- 支持缩进
  const todoCheckedMatch = contentAfterIndent.match(/^[-*+] \[x\] (.+)$/i);
  if (todoCheckedMatch) {
    return {
      type: LineType.TodoChecked,
      content: todoCheckedMatch[1],
      indent,
      raw: line,
    };
  }

  // 无序列表 - 支持缩进
  const bulletMatch = contentAfterIndent.match(/^[-*+] (.+)$/);
  if (bulletMatch) {
    return {
      type: LineType.Bullet,
      content: bulletMatch[1],
      indent,
      raw: line,
    };
  }

  // 有序列表 - 支持缩进
  const orderedMatch = contentAfterIndent.match(/^\d+\. (.+)$/);
  if (orderedMatch) {
    return {
      type: LineType.Ordered,
      content: orderedMatch[1],
      indent,
      raw: line,
    };
  }

  // 引用
  const quoteMatch = trimmed.match(/^> (.*)$/);
  if (quoteMatch) {
    return {
      type: LineType.Quote,
      content: quoteMatch[1],
      indent,
      raw: line,
    };
  }

  // 表格分隔行
  if (isTableSeparator(trimmed)) {
    return { type: LineType.TableSeparator, content: '', indent, raw: line };
  }

  // 表格行
  if (trimmed.startsWith('|') || trimmed.includes('|')) {
    return {
      type: LineType.TableRow,
      content: trimmed,
      indent,
      raw: line,
      meta: { cells: parseTableRow(trimmed) },
    };
  }

  // 图片
  const imageMatch = trimmed.match(/^!\[([^\]]*)\]\(([^)]+)\)$/);
  if (imageMatch) {
    return {
      type: LineType.Image,
      content: '',
      indent,
      raw: line,
      meta: { alt: imageMatch[1], url: imageMatch[2] },
    };
  }

  // 默认：普通文本
  return { type: LineType.Text, content: trimmed, indent, raw: line };
}

// ============ 列表处理器 ============

/**
 * 将扁平的列表项转换为嵌套结构
 */
function buildNestedList(items: ListItem[]): DocumentBlock[] {
  if (items.length === 0) return [];

  const result: DocumentBlock[] = [];
  const stack: { block: DocumentBlock; indent: number }[] = [];

  for (const item of items) {
    const block = item.block;

    // 找到正确的父级
    while (stack.length > 0 && stack[stack.length - 1].indent >= item.indent) {
      stack.pop();
    }

    if (stack.length === 0) {
      // 顶级列表项
      result.push(block);
    } else {
      // 子列表项
      const parent = stack[stack.length - 1].block;
      if (!parent.children) {
        parent.children = [];
      }
      parent.children.push(block);
    }

    // 压入栈中，以便后续项可以成为其子项
    stack.push({ block, indent: item.indent });
  }

  return result;
}

// ============ Markdown 解析器 ============

export interface MarkdownParserOptions {
  /** 是否上传图片（需要提供 client 和 documentId） */
  uploadImages?: boolean;
  /** Lark 客户端 */
  client?: lark.Client;
  /** 文档 ID */
  documentId?: string;
  /** 用户访问令牌 */
  userAccessToken?: string;
  /** 是否使用用户令牌 */
  useUAT?: boolean;
}

interface ImagePlaceholder {
  index: number;
  url: string;
  alt: string;
}

/**
 * 将 Markdown 转换为文档块数组（同步版本，图片为占位符）
 */
export function markdownToBlocks(markdown: string): DocumentBlock[] {
  return parseMarkdownInternal(markdown).blocks;
}

/**
 * 将 Markdown 转换为文档块数组（异步版本，支持图片上传）
 */
export async function markdownToBlocksAsync(
  markdown: string,
  options: MarkdownParserOptions = {},
): Promise<DocumentBlock[]> {
  const { blocks, imagePlaceholders } = parseMarkdownInternal(markdown);

  // 如果需要上传图片且有图片占位符
  if (options.uploadImages && options.client && options.documentId && imagePlaceholders.length > 0) {
    const uploadOptions: ImageUploadOptions = {
      client: options.client,
      documentId: options.documentId,
      userAccessToken: options.userAccessToken,
      useUAT: options.useUAT,
    };

    // 收集所有图片占位符（包括嵌套的）
    const allPlaceholders = collectImagePlaceholders(blocks, imagePlaceholders);

    // 并行上传所有图片
    const uploadPromises = allPlaceholders.map(async (placeholder) => {
      const token = await uploadImageFromUrl(placeholder.url, uploadOptions);
      return { path: placeholder.path, token };
    });

    const results = await Promise.all(uploadPromises);

    // 更新图片块的 token
    for (const result of results) {
      if (result.token) {
        const block = getBlockByPath(blocks, result.path);
        if (block?.image) {
          block.image.token = result.token;
        }
      }
    }
  }

  return blocks;
}

/** 图片占位符（带路径） */
interface ImagePlaceholderWithPath {
  path: number[];
  url: string;
  alt: string;
}

/**
 * 收集所有图片占位符（包括嵌套块中的）
 */
function collectImagePlaceholders(
  blocks: DocumentBlock[],
  rootPlaceholders: ImagePlaceholder[],
): ImagePlaceholderWithPath[] {
  const result: ImagePlaceholderWithPath[] = [];

  // 添加根级别的占位符
  for (const p of rootPlaceholders) {
    result.push({ path: [p.index], url: p.url, alt: p.alt });
  }

  // 递归收集嵌套块中的图片
  function collectFromBlocks(blocks: DocumentBlock[], parentPath: number[]) {
    for (let i = 0; i < blocks.length; i++) {
      const block = blocks[i];
      const currentPath = [...parentPath, i];

      if (block.block_type === BlockType.Image && block.image && !block.image.token) {
        // 如果这个图片块还没有被根占位符记录，添加它
        const alreadyRecorded = result.some((p) => p.path.length === currentPath.length && p.path.every((v, j) => v === currentPath[j]));
        if (!alreadyRecorded && block.image.caption?.content) {
          // 从 caption 中提取 URL（如果有的话）
          const captionContent = block.image.caption.content;
          // 尝试匹配 Markdown 图片格式 ![alt](url)
          const mdImageMatch = captionContent.match(/!\[([^\]]*)\]\(([^)]+)\)/);
          if (mdImageMatch) {
            result.push({ path: currentPath, url: mdImageMatch[2], alt: mdImageMatch[1] });
          } else {
            // 尝试匹配纯 URL 格式
            const urlMatch = captionContent.match(/^(https?:\/\/[^\s]+)$/);
            if (urlMatch) {
              result.push({ path: currentPath, url: urlMatch[1], alt: '' });
            }
          }
        }
      }

      if (block.children) {
        collectFromBlocks(block.children, currentPath);
      }
    }
  }

  collectFromBlocks(blocks, []);
  return result;
}

/**
 * 根据路径获取块
 */
function getBlockByPath(blocks: DocumentBlock[], path: number[]): DocumentBlock | null {
  if (path.length === 0) return null;

  let current: DocumentBlock | undefined = blocks[path[0]];
  for (let i = 1; i < path.length; i++) {
    if (!current?.children) return null;
    current = current.children[path[i]];
  }

  return current || null;
}

/**
 * 内部解析函数
 */
function parseMarkdownInternal(markdown: string): {
  blocks: DocumentBlock[];
  imagePlaceholders: ImagePlaceholder[];
} {
  const lines = markdown.split('\n');
  const blocks: DocumentBlock[] = [];
  const imagePlaceholders: ImagePlaceholder[] = [];

  // 状态变量
  let inCodeBlock = false;
  let codeLines: string[] = [];
  let codeLanguage = 1;
  let tableRows: string[][] = [];
  let listItems: ListItem[] = [];

  // 辅助函数：刷新待处理的列表项
  const flushList = () => {
    if (listItems.length > 0) {
      const nestedBlocks = buildNestedList(listItems);
      blocks.push(...nestedBlocks);
      listItems = [];
    }
  };

  // 辅助函数：刷新待处理的表格
  const flushTable = () => {
    if (tableRows.length > 0) {
      blocks.push(tableToBlock(tableRows));
      tableRows = [];
    }
  };

  for (const line of lines) {
    const token = tokenizeLine(line, inCodeBlock);

    // 处理代码块
    if (token.type === LineType.CodeFence) {
      if (!inCodeBlock) {
        // 进入代码块前，刷新其他待处理内容
        flushList();
        flushTable();
        inCodeBlock = true;
        codeLines = [];
        codeLanguage = CODE_LANGUAGE_MAP[token.meta?.language || ''] || 1;
      } else {
        // 退出代码块
        blocks.push({
          block_type: BlockType.Code,
          code: {
            style: { language: codeLanguage, wrap: false },
            elements: [{ text_run: { content: codeLines.join('\n') } }],
          },
        });
        inCodeBlock = false;
      }
      continue;
    }

    if (token.type === LineType.CodeContent) {
      codeLines.push(token.content);
      continue;
    }

    // 处理表格
    if (token.type === LineType.TableRow) {
      flushList();
      if (token.meta?.cells) {
        tableRows.push(token.meta.cells);
      }
      continue;
    }

    if (token.type === LineType.TableSeparator) {
      // 跳过表格分隔行
      continue;
    }

    // 非表格行时，刷新表格（此时已经处理过 TableRow 和 TableSeparator）
    if (tableRows.length > 0) {
      flushTable();
    }

    // 处理列表项（支持嵌套）
    if (token.type === LineType.Bullet || token.type === LineType.Ordered) {
      const blockType = token.type === LineType.Bullet ? BlockType.Bullet : BlockType.Ordered;
      listItems.push({
        block: textToBlock(token.content, blockType),
        indent: token.indent,
        type: token.type === LineType.Bullet ? 'bullet' : 'ordered',
      });
      continue;
    }

    // 非列表行时，刷新列表（此时已经处理过 Bullet 和 Ordered）
    if (listItems.length > 0) {
      flushList();
    }

    // 处理其他类型
    switch (token.type) {
      case LineType.Empty:
        // 空行可能结束列表或表格
        break;

      case LineType.Heading: {
        const headingTypes = [
          BlockType.Heading1,
          BlockType.Heading2,
          BlockType.Heading3,
          BlockType.Heading4,
          BlockType.Heading5,
          BlockType.Heading6,
        ];
        const level = token.meta?.level || 1;
        blocks.push(textToBlock(token.content, headingTypes[level - 1]));
        break;
      }

      case LineType.Quote:
        blocks.push(textToBlock(token.content, BlockType.Quote));
        break;

      case LineType.TodoUnchecked:
        blocks.push({
          block_type: BlockType.Todo,
          text: {
            style: { done: false },
            elements: parseInlineMarkdown(token.content),
          },
        });
        break;

      case LineType.TodoChecked:
        blocks.push({
          block_type: BlockType.Todo,
          text: {
            style: { done: true },
            elements: parseInlineMarkdown(token.content),
          },
        });
        break;

      case LineType.Divider:
        blocks.push({ block_type: BlockType.Divider });
        break;

      case LineType.Image:
        imagePlaceholders.push({
          index: blocks.length,
          url: token.meta?.url || '',
          alt: token.meta?.alt || '',
        });
        blocks.push({
          block_type: BlockType.Image,
          image: {
            caption: token.meta?.alt ? { content: token.meta.alt } : undefined,
          },
        });
        break;

      case LineType.Text:
        blocks.push(textToBlock(token.content, BlockType.Text));
        break;
    }
  }

  // 处理未闭合的内容
  if (inCodeBlock && codeLines.length >= 0) {
    blocks.push({
      block_type: BlockType.Code,
      code: {
        style: { language: codeLanguage, wrap: false },
        elements: [{ text_run: { content: codeLines.join('\n') } }],
      },
    });
  }

  flushList();
  flushTable();

  return { blocks, imagePlaceholders };
}
