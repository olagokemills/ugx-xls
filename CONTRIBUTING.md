# Contributing to ugx-xls

Thank you for taking the time to contribute! This guide covers everything you need to open a great pull request.

---

## Prerequisites

| Tool        | Minimum version |
|-------------|----------------|
| Node.js     | 18             |
| npm         | 9              |
| Angular CLI | 19 (for the demo app only) |

---

## Local Setup

```bash
# 1. Fork and clone
git clone https://github.com/<your-username>/ugx-xls.git
cd ugx-xls

# 2. Install demo-app dependencies
npm install

# 3. Install library dependencies
cd ugx-xlsx && npm install && cd ..
```

---

## Development Workflow

### Run the demo app

```bash
# From the repo root
npm start          # or: ng serve
# Open http://localhost:4200
```

The demo app imports the library from `ugx-xlsx/dist/`. Rebuild the library after source changes:

```bash
cd ugx-xlsx && npm run build
```

### Run library tests

```bash
cd ugx-xlsx
npm test           # Vitest in watch mode
npm run type-check # TypeScript strict check (no emit)
```

---

## Code Style

- **TypeScript strict mode** is enforced — all code must pass `npm run type-check` without errors.
- **TSDoc on all public exports** — every exported function, type, and interface must have a JSDoc block with at minimum a one-line summary. Public functions should include `@param`, `@returns`, and `@throws` where applicable.
- **No redundant comments** — don't describe *what* the code does if the identifier names already make it clear. Only comment on *why* something is done a non-obvious way.
- **No new dependencies** without a discussion in an issue first. The library aims to stay lean.

---

## Commit Messages

Use [Conventional Commits](https://www.conventionalcommits.org/):

```
feat: add streaming export support
fix: prevent BOM from being written twice
docs: improve WriteOptions TSDoc examples
chore: upgrade vitest to 4.1.0
test: add edge case for empty rows
refactor: extract header normalisation to shared util
```

---

## Submitting a Pull Request

1. **Open an issue first** for non-trivial changes — this avoids duplicate work and clarifies scope.
2. **Branch from `main`** using a descriptive name:
   ```
   feat/streaming-export
   fix/double-bom
   docs/contributing-guide
   ```
3. **Keep PRs focused** — one feature or fix per PR.
4. **Before pushing**, make sure all of the following pass locally:
   ```bash
   cd ugx-xlsx
   npm run type-check
   npm test
   ```
5. **Fill in the PR template** — describe what changed and why, and link any related issues.

---

## Reporting Bugs

Open a [GitHub Issue](https://github.com/olagokemills/ugx-xls/issues) and include:

- `ugx-xlsx` version (`npm list ugx-xlsx`)
- Node.js / browser version
- Minimal reproducible code snippet
- Expected vs. actual behaviour

---

## Feature Requests

Open an issue with the **enhancement** label. Describe the use case (not just the solution) so we can discuss the best approach before implementation begins.
