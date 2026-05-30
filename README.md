# ugx-xlsx

> Safe, TypeScript-first Excel & CSV export for Angular and modern web apps.

[![npm](https://img.shields.io/npm/v/ugx-xlsx)](https://www.npmjs.com/package/ugx-xlsx)
[![downloads](https://img.shields.io/npm/dm/ugx-xlsx)](https://www.npmjs.com/package/ugx-xlsx)
[![license](https://img.shields.io/npm/l/ugx-xlsx)](./ugx-xlsx/package.json)

---

## Features

- **JSON → XLSX and JSON → CSV** — one function call, no backend required
- **Browser-native** — runs entirely client-side using the File API
- **Node.js compatible** — `jsonToWorkbookBuffer` returns a `Uint8Array` in Node
- **Built-in safety limits** — configurable caps on rows, columns, and cell length
- **Data sanitisation** — strips control characters, hard-truncates oversized cells
- **TypeScript-first** — strict types, full TSDoc, ships its own `.d.ts` files
- **Zero Angular runtime dependency** — works in any modern web framework

---

## Installation

```bash
npm install ugx-xlsx
```

---

## Quick Start

### Angular

```typescript
import { exportJsonToExcel, exportJsonToCsv } from 'ugx-xlsx';

@Component({ ... })
export class ReportComponent {
  private rows = [
    { name: 'Alice', department: 'Engineering', salary: 95000 },
    { name: 'Bob',   department: 'Design',      salary: 82000 },
  ];

  downloadExcel() {
    exportJsonToExcel(this.rows, `staff-${new Date().toISOString().slice(0, 10)}`);
  }

  downloadCsv() {
    exportJsonToCsv(this.rows, 'staff', { includeBom: true });
  }
}
```

### Node.js

```typescript
import { jsonToWorkbookBuffer, jsonToCsvString } from 'ugx-xlsx';
import fs from 'node:fs';

const rows = [
  { product: 'Widget', qty: 120, price: 9.99 },
  { product: 'Gadget', qty: 45,  price: 24.50 },
];

// Write XLSX to disk
const buf = await jsonToWorkbookBuffer(rows, { sheetName: 'Inventory' });
fs.writeFileSync('inventory.xlsx', Buffer.from(buf));

// Write CSV to disk
const csv = jsonToCsvString(rows, { includeBom: true });
fs.writeFileSync('inventory.csv', csv, 'utf8');
```

---

## API Reference

### `exportJsonToExcel(rows, fileName?, opts?)`

Triggers a browser download of an XLSX file.

| Parameter  | Type           | Default    | Description                                        |
|------------|----------------|------------|----------------------------------------------------|
| `rows`     | `object[]`     | —          | Array of plain objects to export                   |
| `fileName` | `string`       | `"export"` | Download filename (`.xlsx` appended automatically) |
| `opts`     | `WriteOptions` | see below  | Sheet name, safety limits, sanitisation            |

---

### `exportJsonToCsv(data, fileName?, opts?)`

Triggers a browser download of a CSV file.

| Parameter  | Type                       | Default    | Description                                       |
|------------|----------------------------|------------|---------------------------------------------------|
| `data`     | `object[]`                 | —          | Array of plain objects to export                  |
| `fileName` | `string`                   | `"export"` | Download filename (`.csv` appended automatically) |
| `opts`     | `{ includeBom?: boolean }` | —          | Set `includeBom: true` for Excel on Windows       |

---

### `jsonToWorkbookBuffer(rows, opts?)`

Converts rows to a raw XLSX buffer — works in both Node.js and browser.

Returns `Promise<ArrayBuffer | Uint8Array>`.

---

### `jsonToCsvString(rows, opts?)`

Converts rows to an RFC 4180-compliant CSV string.

Returns `string`.

---

### `WriteOptions`

```typescript
interface WriteOptions {
  sheetName?: string;        // default: "Sheet1"
  sanitizeHeaders?: boolean; // default: true  — strips ASCII control chars from keys
  maxRows?: number;          // default: 100_000
  maxCols?: number;          // default: 1_000
  maxCellChars?: number;     // default: 10_000 — cells longer than this are truncated
}
```

---

### `CsvOptions`

```typescript
interface CsvOptions {
  sanitizeHeaders?: boolean; // default: true
  maxRows?: number;          // default: 100_000
  maxCols?: number;          // default: 1_000
  maxCellChars?: number;     // default: 10_000
  headers?: string[];        // explicit column order; defaults to first-row key order
  includeBom?: boolean;      // default: false — prepend UTF-8 BOM for Excel on Windows
}
```

---

### `SpreadsheetError`

Thrown when a safety limit is exceeded or invalid input is passed.

```typescript
try {
  await exportJsonToExcel(massiveDataset, 'report');
} catch (err) {
  if (err instanceof SpreadsheetError) {
    console.error(err.code);    // "INVALID_INPUT" | "LIMIT_ROWS" | "LIMIT_COLS"
    console.error(err.message); // human-readable description
  }
}
```

---

## Running the Demo App

```bash
git clone https://github.com/olagokemills/ugx-xls.git
cd ugx-xls
npm install
ng serve
```

Open `http://localhost:4200` — click **Export to Excel** or **Export to CSV** to see the library in action.

---

## Building the Library

```bash
cd ugx-xlsx
npm install
npm run build      # compiles CJS + ESM2015 + ESM2020 into dist/
npm run type-check # TypeScript strict check with no emit
npm test           # Vitest unit tests
```

---

## Contributing

See [CONTRIBUTING.md](./CONTRIBUTING.md) for setup instructions, code style, and the PR process.

---

## License

[MIT](./ugx-xlsx/package.json)
