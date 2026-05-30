# ugx-xlsx

Safe, TypeScript-first Excel & CSV export for Angular and modern web apps.  
Convert plain JSON arrays into `.xlsx` or `.csv` files **directly in the browser** — no backend required.

[![npm](https://img.shields.io/npm/v/ugx-xlsx)](https://www.npmjs.com/package/ugx-xlsx)
[![downloads](https://img.shields.io/npm/dm/ugx-xlsx)](https://www.npmjs.com/package/ugx-xlsx)
[![license](https://img.shields.io/npm/l/ugx-xlsx)](./package.json)

---

## Installation

```bash
npm install ugx-xlsx
```

---

## Usage

### Export to Excel (browser)

```typescript
import { exportJsonToExcel } from 'ugx-xlsx';

const rows = [
  { name: 'Alice', score: 42 },
  { name: 'Bob',   score: 98 },
];

await exportJsonToExcel(rows, 'results');
// → triggers download of results.xlsx
```

### Export to CSV (browser)

```typescript
import { exportJsonToCsv } from 'ugx-xlsx';

await exportJsonToCsv(rows, 'results', { includeBom: true });
// → triggers download of results.csv (BOM ensures correct encoding in Excel on Windows)
```

### Generate a buffer (Node.js / server-side)

```typescript
import { jsonToWorkbookBuffer } from 'ugx-xlsx';
import fs from 'node:fs';

const buf = await jsonToWorkbookBuffer(rows, { sheetName: 'Scores' });
fs.writeFileSync('results.xlsx', Buffer.from(buf));
```

### Generate a CSV string

```typescript
import { jsonToCsvString } from 'ugx-xlsx';

const csv = jsonToCsvString(rows, { includeBom: true });
fs.writeFileSync('results.csv', csv, 'utf8');
```

---

## API

| Function | Returns | Description |
|----------|---------|-------------|
| `exportJsonToExcel(rows, fileName?, opts?)` | `Promise<void>` | Browser XLSX download |
| `exportJsonToCsv(data, fileName?, opts?)` | `Promise<void>` | Browser CSV download |
| `jsonToWorkbookBuffer(rows, opts?)` | `Promise<ArrayBuffer \| Uint8Array>` | Raw XLSX buffer (Node + browser) |
| `jsonToCsvString(rows, opts?)` | `string` | RFC 4180 CSV string |

### `WriteOptions`

```typescript
interface WriteOptions {
  sheetName?: string;        // default: "Sheet1"
  sanitizeHeaders?: boolean; // default: true
  maxRows?: number;          // default: 100_000
  maxCols?: number;          // default: 1_000
  maxCellChars?: number;     // default: 10_000
}
```

### `CsvOptions`

```typescript
interface CsvOptions {
  sanitizeHeaders?: boolean;
  maxRows?: number;
  maxCols?: number;
  maxCellChars?: number;
  headers?: string[];   // explicit column order
  includeBom?: boolean; // prepend UTF-8 BOM (recommended for Excel on Windows)
}
```

### `SpreadsheetError`

Thrown on invalid input or limit violations. Has a `code` property:
`"INVALID_INPUT"` | `"LIMIT_ROWS"` | `"LIMIT_COLS"`.

---

## Full documentation

See the [repository README](https://github.com/olagokemills/ugx-xls#readme) for the complete API reference and contributing guide.

---

## License

MIT
