import { describe, it, expect } from 'vitest';
import {
  markdownToBlocks,
  parseInlineMarkdown,
  textToBlock,
  parseTableRow,
  isTableSeparator,
  tableToBlock,
  DocumentBlock,
} from '../../../src/mcp-tool/utils/markdown-parser';
import { BlockType } from '../../../src/mcp-tool/utils/block-types';

describe('parseInlineMarkdown', () => {
  it('should parse plain text', () => {
    const result = parseInlineMarkdown('Hello world');
    expect(result).toHaveLength(1);
    expect(result[0].text_run?.content).toBe('Hello world');
  });

  it('should parse bold text', () => {
    const result = parseInlineMarkdown('Hello **bold** world');
    expect(result).toHaveLength(3);
    expect(result[0].text_run?.content).toBe('Hello ');
    expect(result[1].text_run?.content).toBe('bold');
    expect(result[1].text_run?.text_element_style?.bold).toBe(true);
    expect(result[2].text_run?.content).toBe(' world');
  });

  it('should parse italic text', () => {
    const result = parseInlineMarkdown('Hello *italic* world');
    expect(result).toHaveLength(3);
    expect(result[1].text_run?.content).toBe('italic');
    expect(result[1].text_run?.text_element_style?.italic).toBe(true);
  });

  it('should parse bold italic text', () => {
    const result = parseInlineMarkdown('Hello ***bold italic*** world');
    expect(result).toHaveLength(3);
    expect(result[1].text_run?.content).toBe('bold italic');
    expect(result[1].text_run?.text_element_style?.bold).toBe(true);
    expect(result[1].text_run?.text_element_style?.italic).toBe(true);
  });

  it('should parse strikethrough text', () => {
    const result = parseInlineMarkdown('Hello ~~deleted~~ world');
    expect(result).toHaveLength(3);
    expect(result[1].text_run?.content).toBe('deleted');
    expect(result[1].text_run?.text_element_style?.strikethrough).toBe(true);
  });

  it('should parse inline code', () => {
    const result = parseInlineMarkdown('Hello `code` world');
    expect(result).toHaveLength(3);
    expect(result[1].text_run?.content).toBe('code');
    expect(result[1].text_run?.text_element_style?.inline_code).toBe(true);
  });

  it('should parse links', () => {
    const result = parseInlineMarkdown('Check out [Google](https://google.com)');
    expect(result).toHaveLength(2);
    expect(result[0].text_run?.content).toBe('Check out ');
    expect(result[1].text_run?.content).toBe('Google');
    expect(result[1].text_run?.text_element_style?.link?.url).toBe('https://google.com');
  });

  it('should parse multiple styles in same text', () => {
    const result = parseInlineMarkdown('**bold** and *italic* and `code`');
    expect(result).toHaveLength(5);
    expect(result[0].text_run?.text_element_style?.bold).toBe(true);
    expect(result[2].text_run?.text_element_style?.italic).toBe(true);
    expect(result[4].text_run?.text_element_style?.inline_code).toBe(true);
  });
});

describe('markdownToBlocks', () => {
  describe('headings', () => {
    it('should parse h1 heading', () => {
      const result = markdownToBlocks('# Hello');
      expect(result).toHaveLength(1);
      expect(result[0].block_type).toBe(BlockType.Heading1);
    });

    it('should parse h2-h6 headings', () => {
      const markdown = `## H2
### H3
#### H4
##### H5
###### H6`;
      const result = markdownToBlocks(markdown);
      expect(result).toHaveLength(5);
      expect(result[0].block_type).toBe(BlockType.Heading2);
      expect(result[1].block_type).toBe(BlockType.Heading3);
      expect(result[2].block_type).toBe(BlockType.Heading4);
      expect(result[3].block_type).toBe(BlockType.Heading5);
      expect(result[4].block_type).toBe(BlockType.Heading6);
    });
  });

  describe('lists', () => {
    it('should parse unordered list', () => {
      const result = markdownToBlocks('- Item 1\n- Item 2');
      expect(result).toHaveLength(2);
      expect(result[0].block_type).toBe(BlockType.Bullet);
      expect(result[1].block_type).toBe(BlockType.Bullet);
    });

    it('should parse ordered list', () => {
      const result = markdownToBlocks('1. First\n2. Second');
      expect(result).toHaveLength(2);
      expect(result[0].block_type).toBe(BlockType.Ordered);
      expect(result[1].block_type).toBe(BlockType.Ordered);
    });

    it('should parse nested unordered list', () => {
      const markdown = `- Parent 1
  - Child 1.1
  - Child 1.2
- Parent 2
  - Child 2.1`;
      const result = markdownToBlocks(markdown);
      expect(result).toHaveLength(2);
      expect(result[0].block_type).toBe(BlockType.Bullet);
      expect(result[0].children).toHaveLength(2);
      expect(result[0].children?.[0].block_type).toBe(BlockType.Bullet);
      expect(result[1].children).toHaveLength(1);
    });

    it('should parse deeply nested list', () => {
      const markdown = `- Level 1
  - Level 2
    - Level 3
      - Level 4`;
      const result = markdownToBlocks(markdown);
      expect(result).toHaveLength(1);
      expect(result[0].children).toHaveLength(1);
      expect(result[0].children?.[0].children).toHaveLength(1);
      expect(result[0].children?.[0].children?.[0].children).toHaveLength(1);
    });

    it('should parse mixed nested list', () => {
      const markdown = `1. Ordered parent
  - Unordered child
  - Another child
2. Second ordered`;
      const result = markdownToBlocks(markdown);
      expect(result).toHaveLength(2);
      expect(result[0].block_type).toBe(BlockType.Ordered);
      expect(result[0].children).toHaveLength(2);
      expect(result[0].children?.[0].block_type).toBe(BlockType.Bullet);
    });
  });

  describe('code blocks', () => {
    it('should parse code block', () => {
      const markdown = '```javascript\nconst x = 1;\n```';
      const result = markdownToBlocks(markdown);
      expect(result).toHaveLength(1);
      expect(result[0].block_type).toBe(BlockType.Code);
      expect(result[0].code?.style?.language).toBe(30); // JavaScript
      expect(result[0].code?.elements[0].text_run?.content).toBe('const x = 1;');
    });

    it('should handle unclosed code block', () => {
      const markdown = '```python\nprint("hello")';
      const result = markdownToBlocks(markdown);
      expect(result).toHaveLength(1);
      expect(result[0].block_type).toBe(BlockType.Code);
    });
  });

  describe('quotes', () => {
    it('should parse blockquote', () => {
      const result = markdownToBlocks('> This is a quote');
      expect(result).toHaveLength(1);
      expect(result[0].block_type).toBe(BlockType.Quote);
    });
  });

  describe('todo items', () => {
    it('should parse unchecked todo', () => {
      const result = markdownToBlocks('- [ ] Todo item');
      expect(result).toHaveLength(1);
      expect(result[0].block_type).toBe(BlockType.Todo);
      expect(result[0].text?.style?.done).toBe(false);
    });

    it('should parse checked todo', () => {
      const result = markdownToBlocks('- [x] Done item');
      expect(result).toHaveLength(1);
      expect(result[0].block_type).toBe(BlockType.Todo);
      expect(result[0].text?.style?.done).toBe(true);
    });
  });

  describe('dividers', () => {
    it('should parse horizontal rule with dashes', () => {
      const result = markdownToBlocks('---');
      expect(result).toHaveLength(1);
      expect(result[0].block_type).toBe(BlockType.Divider);
    });

    it('should parse horizontal rule with asterisks', () => {
      const result = markdownToBlocks('***');
      expect(result).toHaveLength(1);
      expect(result[0].block_type).toBe(BlockType.Divider);
    });
  });

  describe('images', () => {
    it('should parse image', () => {
      const result = markdownToBlocks('![Alt text](https://example.com/image.png)');
      expect(result).toHaveLength(1);
      expect(result[0].block_type).toBe(BlockType.Image);
      expect(result[0].image?.caption?.content).toBe('Alt text');
    });
  });

  describe('tables', () => {
    it('should parse simple table', () => {
      const markdown = `| Header 1 | Header 2 |
| --- | --- |
| Cell 1 | Cell 2 |
| Cell 3 | Cell 4 |`;
      const result = markdownToBlocks(markdown);
      expect(result).toHaveLength(1);
      expect(result[0].block_type).toBe(BlockType.Table);
      expect(result[0].table?.property.row_size).toBe(3);
      expect(result[0].table?.property.column_size).toBe(2);
    });
  });

  describe('complex document', () => {
    it('should parse a complex markdown document', () => {
      const markdown = `# Title

This is a paragraph with **bold** and *italic* text.

## Section 1

- Item 1
  - Nested item
- Item 2

\`\`\`typescript
const hello = "world";
\`\`\`

> A blockquote

---

![Image](https://example.com/img.png)`;

      const result = markdownToBlocks(markdown);
      expect(result.length).toBeGreaterThan(5);
      expect(result[0].block_type).toBe(BlockType.Heading1);
    });
  });
});

describe('table utilities', () => {
  describe('parseTableRow', () => {
    it('should parse table row', () => {
      const result = parseTableRow('| a | b | c |');
      expect(result).toEqual(['a', 'b', 'c']);
    });

    it('should handle row without trailing pipe', () => {
      const result = parseTableRow('| a | b | c');
      expect(result).toEqual(['a', 'b', 'c']);
    });
  });

  describe('isTableSeparator', () => {
    it('should identify table separator', () => {
      expect(isTableSeparator('| --- | --- |')).toBe(true);
      expect(isTableSeparator('|:---:|:---:|')).toBe(true);
      expect(isTableSeparator('| --- |')).toBe(true);
    });

    it('should reject non-separator lines', () => {
      expect(isTableSeparator('| cell | cell |')).toBe(false);
    });
  });
});

describe('textToBlock', () => {
  it('should create text block with inline markdown parsed', () => {
    const result = textToBlock('Hello **world**', BlockType.Text);
    expect(result.block_type).toBe(BlockType.Text);
    expect(result.text?.elements).toHaveLength(2);
  });
});
