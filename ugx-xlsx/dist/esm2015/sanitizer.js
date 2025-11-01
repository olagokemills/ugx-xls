export function sanitizeHeader(header) {
    return header.replace(/[\u0000-\u001F\u007F]/g, "").trim();
}
export function enforceLimits(value, maxCellChars) {
    if (!value)
        return value;
    return value.length > maxCellChars ? value.slice(0, maxCellChars) : value;
}
