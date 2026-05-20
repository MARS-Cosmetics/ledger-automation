# Mars vs Brand — Ledger Reconciliation

Browser-only React + TypeScript app for reconciling Mars and Brand ledgers.

**All processing runs in the browser.** Files never leave the user's laptop. No backend, no uploads, no data sent anywhere.

## What it does

1. Upload Mars ledger (.xlsx) and Brand ledger (.xlsx)
2. Pick the correct sheet from each workbook
3. Confirm column mapping (auto-detected, override if wrong)
4. Run reconciliation — produces:
   - **Recon Summary** — category-wise totals + grand total
   - **Mars Cosmetics** sheet — Mars rows annotated with `Amount_Brand`, `Difference`, `Remarks`
   - **Brand** sheet — Brand rows annotated with `Amount_Mars`, `Diff`, `Remarks`
5. Download the styled output workbook

### Match logic

- Filter both sides to `Period == "Recon"`
- Match key: Brand `Reference` (or `Correct Ref No` if filled) ↔ Mars `Vch/Bill No`
- Remarks: `Matched` · `Amount Mismatch` · `Not Booked by Brand` · `Not Booked by Mars`
- Many-to-one splits are summed; remarks note the split count

## Run locally

```bash
npm install
npm run dev
```

Opens at http://localhost:5173.

## Build

```bash
npm run build
```

Outputs static assets to `dist/` — deploy to any static host (Cloudflare Pages, Netlify, etc.).

## Deploy on Cloudflare Pages

1. In Cloudflare Pages, connect this GitHub repo (`MARS-Cosmetics/ledger-automation`)
2. Framework preset: **Vite**
3. Build command: `npm run build`
4. Build output directory: `dist`
5. Push to `main` → auto-deploy

## Project layout

```
src/
  App.tsx                    main UI flow
  main.tsx                   entry point
  index.css                  Tailwind base
  components/
    FileUploader.tsx         drag-and-drop upload
    SheetPicker.tsx          choose sheet from a multi-sheet workbook
    ColumnMapper.tsx         confirm logical-field-to-column mapping
    ResultView.tsx           summary + previews + download button
  lib/
    types.ts                 shared types
    columnDetect.ts          fuzzy header matching, synonym tables
    reconcile.ts             pure reconciliation logic
    excelRead.ts             ExcelJS-based input parser
    excelWrite.ts            styled output workbook
```
