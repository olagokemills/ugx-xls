# ugx-xlsx

A lightweight Excel export utility for Angular applications.  
Easily convert arrays or objects into `.xlsx` files **directly in the browser** — no backend required.

![npm](https://img.shields.io/npm/v/ugx-xlsx)
![downloads](https://img.shields.io/npm/dm/ugx-xlsx)
![license](https://img.shields.io/npm/l/ugx-xlsx)

---

## ✨ Features

- Simple API — one method to export data to Excel
- Supports arrays of objects and arrays of arrays
- Custom file name, sheet name, and headers
- Runs fully in the browser
- Built for Angular + TypeScript

---

## 📦 Installation

Install with npm:

```bash
npm install ugx-xlsx

yarn add ugx-xlsx

```

## 📘 How to Use ugx-xlsx in Angular

```
import { UgxXlsxService } from 'ugx-xlsx';
```

Inject the Service in Your Component

```
constructor(private ugxXlsx: UgxXlsxService) {}
```

### Prepare Your Data

You can export arrays of objects or arrays of arrays.

```
const data = [
  { name: 'John Doe', age: 30, email: 'john@mail.com' },
  { name: 'Jane Smith', age: 25, email: 'jane@mail.com' }
];
```

### Export to Excel

```
this.ugxXlsx.exportToExcel(data, {
  fileName: 'users'
});


<button (click)="downloadExcel()">Download Excel</button>

downloadExcel() {
  const rows = [
    { id: 1, item: 'Apples', qty: 10 },
    { id: 2, item: 'Bananas', qty: 6 }
  ];

  this.ugxXlsx.exportToExcel(rows, { fileName: 'inventory' });
}

export interface UgxXlsxOptions {
  fileName?: string;   // default: 'export'
  sheetName?: string;  // default: 'Sheet1'
  headers?: string[];  // optional: custom header labels
}
```

### Contributing

Contributions and suggestions are welcome.

Fork the repository

Create a new branch

Make your changes

Open a pull request

https://github.com/olagokemills/ugx-xls
