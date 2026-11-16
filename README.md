# Excel Viewer

Converts spreadsheets with multiple tabs into plain text. Drop a file on the left, pick a format on the right, copy or download the result.

Everything runs in the browser, the file is never uploaded.

## What it does

- Reads `.xlsx`, `.xls`, `.xlsm`, `.ods` and `.csv`
- Previews each sheet, with a checkbox to include or skip it
- Outputs as:
  - **XML + CSV** each sheet as CSV inside a `<sheet>` tag
  - **Markdown** a heading and a table per sheet
  - **JSON** one array of row objects per sheet, using the first row as keys
  - **Schema** column names, guessed types, fill counts and sample values
- Options to trim empty rows and columns, limit rows per sheet, fill merged cells, use ISO dates, show formulas instead of values, and write compact JSON

## Running it locally

```bash
npm install
npm run dev
```

## Stack

- Vue 3 with `<script setup>`
- Vite
- Tailwind CSS v4
- SheetJS, vendored in `vendor/` because npm blocks remote tarballs
