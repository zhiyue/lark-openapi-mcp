/**
 * API Error type definition
 */
export interface ApiError {
  response?: { data?: unknown };
  message?: string;
}

/**
 * Extract error message from various error types
 * Handles Lark API errors, standard errors, and unknown error types
 */
export function extractErrorMessage(error: unknown): unknown {
  const err = error as ApiError;
  return err?.response?.data || err?.message || error;
}
