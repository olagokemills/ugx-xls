/**
 * Strips ASCII control characters (U+0000–U+001F, U+007F) from a column
 * header string and trims surrounding whitespace.
 *
 * Control characters can corrupt Excel/CSV parsers, so all user-supplied
 * keys pass through here before being written to a sheet.
 *
 * @param header - Raw column header string.
 * @returns Sanitized header with control chars removed and whitespace trimmed.
 */
export declare function sanitizeHeader(header: string): string;
/**
 * Hard-truncates a cell value to `maxCellChars` characters.
 *
 * Excel silently drops content beyond ~32 767 chars per cell. This
 * configurable limit (default 10 000 via `WriteOptions.maxCellChars`)
 * prevents silent data loss in downstream tooling.
 *
 * @param value        - String value to truncate.
 * @param maxCellChars - Maximum character count allowed.
 * @returns The original string if within the limit, otherwise a truncated copy.
 */
export declare function enforceLimits(value: string, maxCellChars: number): string;
