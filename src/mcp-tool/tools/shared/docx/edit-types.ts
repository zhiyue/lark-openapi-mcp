import { TextElement, TextContent } from '../../../utils';

/**
 * Block item from API response
 */
export interface BlockItem {
  block_id?: string;
  parent_id?: string;
  block_type?: number;
  text?: TextContent;
  code?: {
    style?: { language?: number; wrap?: boolean };
    elements: TextElement[];
  };
}

/**
 * Update request for batch update
 */
export interface UpdateRequest {
  block_id: string;
  update_text_elements: {
    elements: TextElement[];
  };
}

/**
 * Tool name type for docx edit tools
 */
export type DocxEditToolName =
  | 'docx.builtin.updateTitle'
  | 'docx.builtin.append'
  | 'docx.builtin.replace'
  | 'docx.builtin.edit';
