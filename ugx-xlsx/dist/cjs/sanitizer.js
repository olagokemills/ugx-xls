"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.sanitizeHeader = sanitizeHeader;
exports.enforceLimits = enforceLimits;
function sanitizeHeader(header) {
    return header.replace(/[\u0000-\u001F\u007F]/g, "").trim();
}
function enforceLimits(value, maxCellChars) {
    if (!value)
        return value;
    return value.length > maxCellChars ? value.slice(0, maxCellChars) : value;
}
