---
name: Bug report
about: Something isn't working as expected
labels: bug
---

## Description

<!-- A clear, concise description of the bug. -->

## Minimal reproduction

```typescript
// Paste the smallest snippet that triggers the bug
import { exportJsonToExcel } from 'ugx-xlsx';

const rows = [ /* ... */ ];
exportJsonToExcel(rows, 'export');
// Expected: ...
// Actual: ...
```

## Environment

| Field           | Value |
|-----------------|-------|
| ugx-xlsx version | `npm list ugx-xlsx` |
| Node.js version  | `node -v` |
| Browser / OS     | e.g. Chrome 124 / macOS 14 |
| Angular version  | `ng version` (if applicable) |

## Expected behaviour

<!-- What should happen? -->

## Actual behaviour

<!-- What actually happens? Include the full error message and stack trace if available. -->
