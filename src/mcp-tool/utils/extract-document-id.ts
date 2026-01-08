/**
 * Helper function: Extract document_id from URL or direct ID
 * Supported URL formats:
 * - /docx/xxx
 * - /wiki/xxx
 * - /base/xxx
 * - /sheets/xxx
 * - /doc/xxx
 * - /bitable/xxx
 * - /mindnote/xxx
 * - /file/xxx
 * - URLs with query parameters (e.g., ?from=...)
 * - Document IDs may contain underscores
 */
export function extractDocumentId(documentIdOrUrl: string): string {
  // If it's a URL, extract document_id (supports docx, wiki, base, sheets, etc.)
  const urlMatch = documentIdOrUrl.match(
    /\/(?:docx|wiki|base|sheets|doc|bitable|mindnote|file)\/([a-zA-Z0-9_]+)/,
  );
  if (urlMatch) {
    // Remove possible query parameters
    return urlMatch[1].split('?')[0];
  }
  // If passing ID directly, it may also have query parameters
  return documentIdOrUrl.split('?')[0];
}
