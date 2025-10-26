export function sanitizeHeader(header: string): string {
  return header.replace(/[\u0000-\u001F\u007F]/g, "").trim();
}

export function enforceLimits(value: string, maxCellChars: number): string {
  if (!value) return value;
  return value.length > maxCellChars ? value.slice(0, maxCellChars) : value;
}
