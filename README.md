# Excel Filter App — Web

A browser-based JavaScript port of `filter_ids.py`.  
Upload an `.xlsx` file and instantly see all **IDs** that match the filter criteria — no server, no install, everything runs in the browser.

---

## Filter criteria

| Column | Position | Required value |
|--------|----------|----------------|
| ID     | 1        | *(returned)*   |
| Status | 3        | `DONE` or `CLOSED` |
| Group  | 5        | `HIS_GROUP`    |
| Type   | 6        | `Defect (Bug)` |

---

## Project structure

```
web/
├── index.html          # Single-page UI
├── src/
│   ├── filter.js       # Core filter logic (port of filter_ids.py)
│   ├── app.js          # UI wiring (file picker, drag-and-drop, render)
│   └── filter.test.js  # Jest unit tests
└── package.json
```

---

## Running locally

No build step is required. Because `app.js` uses ES modules (`type="module"`), the
page **must** be served over HTTP — opening `index.html` directly as a `file://` URL
will block module loading in most browsers.

The simplest options:

```bash
# Option A — Node http-server (one-time global install)
npx http-server web -p 8080
# then open http://localhost:8080

# Option B — Python (if available)
cd web
python -m http.server 8080
# then open http://localhost:8080

# Option C — VS Code Live Server extension
# Right-click index.html → "Open with Live Server"
```

---

## Running tests

```bash
cd web
npm install
npm test
```

Tests cover `matches()` and `extractIds()` in [`src/filter.js`](src/filter.js) and mirror
the Python `test_filter_ids.py` suite.

---

## Dependencies

| Library | Version | Purpose |
|---------|---------|---------|
| [SheetJS (xlsx)](https://sheetjs.com/) | 0.20.3 | Parse `.xlsx` files in the browser |
| [Jest](https://jestjs.io/) | ^29.7.0 | Unit test runner (dev only) |

SheetJS is loaded at runtime from `cdn.sheetjs.com` (pinned to `0.20.3`).  
No other runtime dependencies.

---

## How it works

1. User drops or picks an `.xlsx` file.
2. The browser reads the file as an `ArrayBuffer` via the `FileReader` API.
3. **SheetJS** parses the workbook and converts the first sheet to an array-of-arrays.
4. [`src/filter.js`](src/filter.js) iterates every data row and applies the same three
   conditions as the Python original (whitespace trimmed, null-safe).
5. Matching IDs are rendered in the results list — all client-side, nothing is uploaded.
