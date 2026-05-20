import ExcelJS from 'exceljs';
import { canonicalizeCategory } from './reconcile';
import type { ReconResult, Row } from './types';

// ── fills & format ────────────────────────────────────────────────────────────
const HEADER_FILL      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDCE6F1' } } as const;
const TOTAL_FILL       = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } } as const;
const NOT_BOOKED_FILL  = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFCE4D6' } } as const;
const MISMATCH_FILL    = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF4CE' } } as const;
const MATCH_FILL       = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2EFDA' } } as const;
const REVERSAL_FILL    = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDDEBF7' } } as const;
const TITLE_FILL       = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F3864' } } as const;

const MONEY_FMT = '#,##,##0.00;(#,##,##0.00);"-"';

// Columns for the combined "Annex" sheet (Sheet 4) — exact names as specified
const ANNEX_HEADERS = [
  'Particular',
  'Category',
  'Sub-Category',
  'Date',
  'Invoice No',
  'Amount_MARS',
  'Amount_Brand',
  'Difference',
] as const;

// Canonical display order for category sections
const CAT_ORDER = ['Opening Balance', 'Invoice', 'Contra_Inv', 'CN', 'Contra_CN', 'DN'];

// Categories that should be merged into one section in the Annex sheet.
// label  → displayed as the section title
// cats   → ordered list: first category's rows appear before second category's rows
const ANNEX_MERGED_GROUPS: { label: string; cats: string[] }[] = [
  { label: 'Invoice & Contra Invoice', cats: ['Invoice', 'Contra_Inv'] },
];


// ── public API ────────────────────────────────────────────────────────────────
export async function buildReconWorkbook(
  res: ReconResult,
  generatedAt: Date = new Date(),
): Promise<ArrayBuffer> {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Mars Recon Tool';
  wb.created = generatedAt;

  writeSummary(wb, res, generatedAt);                                      // Sheet 1
  writeDataSheet(wb, 'Mars Cosmetics', res.mars.headers, res.mars.rows); // Sheet 2
  writeDataSheet(wb, 'Brand', res.brand.headers, res.brand.rows);        // Sheet 3
  writeAnnexSheet(wb, res);                                               // Sheet 4

  const buf = await wb.xlsx.writeBuffer();
  return buf as ArrayBuffer;
}

// ── Sheet 4: Annex — all transactions, grouped category-wise ─────────────────
function writeAnnexSheet(wb: ExcelJS.Workbook, res: ReconResult): void {
  const mc = res.marsCols;
  const bc = res.brandCols;

  const marsVchCol  = mc.vch_no;
  const marsAmtCol  = mc.net_amount;
  const marsCatCol  = mc.category;
  const marsDateCol = mc.date;

  const brandRefCol  = bc.reference;
  const brandAmtCol  = bc.net_amount;
  const brandCatCol  = bc.category;
  const brandDateCol = bc.date;

  if (!marsVchCol || !marsAmtCol || !brandRefCol || !brandAmtCol) return;

  // Collect all distinct categories from both sides, excluding Opening Balance
  const catSet = new Set<string>();
  for (const r of res.mars.rows) {
    const cat = normCat(marsCatCol ? String(r[marsCatCol] ?? '') : '');
    if (cat !== 'Opening Balance') catSet.add(cat);
  }
  for (const r of res.brand.rows) {
    if (toNum(r['Amount_Mars']) !== 0) continue;
    const cat = normCat(brandCatCol ? String(r[brandCatCol] ?? '') : '');
    if (cat !== 'Opening Balance') catSet.add(cat);
  }

  const cats = [...catSet].sort((a, b) => {
    const ia = CAT_ORDER.indexOf(a);
    const ib = CAT_ORDER.indexOf(b);
    if (ia !== -1 && ib !== -1) return ia - ib;
    if (ia !== -1) return -1;
    if (ib !== -1) return 1;
    return a.localeCompare(b);
  });

  const ws = wb.addWorksheet('Annex');
  const ncols = ANNEX_HEADERS.length;

  // Helper: collect Annex rows for an ordered list of category names.
  // Rows are emitted category by category in the given order —
  // all rows for cats[0] first, then all rows for cats[1], etc.
  // Particular column shows the raw ledger value, not any merged label.
  const collectRows = (orderedCats: string[]): Row[] => {
    const rows: Row[] = [];
    for (const cat of orderedCats) {
      for (const r of res.mars.rows) {
        const rawCat = marsCatCol ? String(r[marsCatCol] ?? '') : '';
        if (normCat(rawCat) !== cat) continue;
        rows.push({
          'Particular':    rawCat,
          'Category':      String(r['Remarks'] ?? ''),
          'Sub-Category':  '',
          'Date':          marsDateCol ? (r[marsDateCol] ?? null) : null,
          'Invoice No':    String(r[marsVchCol] ?? ''),
          'Amount_MARS':   toNum(r[marsAmtCol]),
          'Amount_Brand':  toNum(r['Amount_Brand']),
          'Difference':    toNum(r['Difference']),
        });
      }
      for (const r of res.brand.rows) {
        if (toNum(r['Amount_Mars']) !== 0) continue;
        const rawCat = brandCatCol ? String(r[brandCatCol] ?? '') : '';
        if (normCat(rawCat) !== cat) continue;
        rows.push({
          'Particular':    rawCat,
          'Category':      String(r['Remarks'] ?? ''),
          'Sub-Category':  '',
          'Date':          brandDateCol ? (r[brandDateCol] ?? null) : null,
          'Invoice No':    String(r[brandRefCol] ?? ''),
          'Amount_MARS':   0,
          'Amount_Brand':  toNum(r[brandAmtCol]),
          'Difference':    toNum(r['Diff']),
        });
      }
    }
    return rows;
  };

  // Helper: write a titled section into the sheet
  const writeSection = (label: string, rows: Row[]) => {
    if (rows.length === 0) return;
    const titleRow = ws.addRow([label]);
    titleRow.height = 22;
    ws.mergeCells(titleRow.number, 1, titleRow.number, ncols);
    for (let c = 1; c <= ncols; c++) {
      const cell = titleRow.getCell(c);
      cell.fill = TITLE_FILL;
      cell.font = { bold: true, size: 11, color: { argb: 'FFFFFFFF' } };
      cell.alignment = { vertical: 'middle' };
    }
    const hRow = ws.addRow([...ANNEX_HEADERS]);
    styleHeader(hRow);
    for (const r of rows) {
      const values = ANNEX_HEADERS.map((h) => normalizeForCell(r[h]));
      const row = ws.addRow(values);
      for (const c of [6, 7, 8]) row.getCell(c).numFmt = MONEY_FMT;
      const fill = remarkFill(String(r['Category'] ?? ''));
      if (fill) row.eachCell((cell) => { cell.fill = fill; });
    }
    ws.addRow([]);
    ws.addRow([]);
  };

  // Track which individual cats have already been handled via a merged group
  const handledCats = new Set<string>();

  for (const cat of cats) {
    if (handledCats.has(cat)) continue;

    const mergedGroup = ANNEX_MERGED_GROUPS.find((g) => g.cats.includes(cat));

    if (mergedGroup) {
      // Mark all cats in this group so they are not written individually later
      mergedGroup.cats.forEach((c) => handledCats.add(c));
      // collectRows respects the group order: Invoice rows first, Contra_Inv rows after
      writeSection(mergedGroup.label, collectRows(mergedGroup.cats));
    } else {
      writeSection(cat, collectRows([cat]));
    }
  }

  // Column widths
  [20, 32, 16, 14, 32, 18, 18, 16].forEach((w, i) => { ws.getColumn(i + 1).width = w; });
}


// ── helpers ───────────────────────────────────────────────────────────────────
function normCat(raw: string): string {
  const c = canonicalizeCategory(raw.trim());
  return c || raw.trim() || 'Uncategorized';
}


function remarkFill(remark: string): ExcelJS.Fill | null {
  if (remark === 'Match')                          return MATCH_FILL;
  if (remark === 'Amount Mismatch')                return MISMATCH_FILL;
  if (remark.startsWith('Not Booked'))             return NOT_BOOKED_FILL;
  if (remark.toLowerCase().includes('reversal'))   return REVERSAL_FILL;
  return null;
}

function toNum(v: unknown): number {
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  if (v === null || v === undefined) return 0;
  const s = String(v).replace(/,/g, '').trim();
  if (!s || s.toLowerCase() === 'nan') return 0;
  const n = Number(s);
  return isFinite(n) ? n : 0;
}

function styleHeader(row: ExcelJS.Row): void {
  row.eachCell((cell) => {
    cell.font = { bold: true };
    cell.fill = HEADER_FILL;
  });
}

function normalizeForCell(v: unknown): string | number | Date | null {
  if (v === null || v === undefined) return null;
  if (typeof v === 'number' || typeof v === 'string' || v instanceof Date) return v;
  if (typeof v === 'boolean') return v ? 'Yes' : 'No';
  return String(v);
}

function formatTimestamp(d: Date): string {
  const pad = (n: number) => n.toString().padStart(2, '0');
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
}

export function buildOutputFilename(d: Date = new Date()): string {
  const pad = (n: number) => n.toString().padStart(2, '0');
  return `recon_${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}_${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}.xlsx`;
}

// ── Full ledger sheets (Mars Cosmetics + Brand) ───────────────────────────────
const MONEY_HEADERS = new Set([
  'Amount_Mars', 'Amount_Brand', 'Difference', 'Diff',
  'Net Amount', 'Net amount', 'Debit(Rs.)', 'Credit(Rs.)',
  'Debit', 'Credit', 'Amount in local currency',
]);

function writeDataSheet(wb: ExcelJS.Workbook, name: string, headers: string[], rows: Row[]): void {
  const ws = wb.addWorksheet(name);
  ws.views = [{ state: 'frozen', ySplit: 1 }];

  const headerRow = ws.addRow(headers);
  styleHeader(headerRow);

  const moneyColIdx: number[] = [];
  headers.forEach((h, i) => {
    if (MONEY_HEADERS.has(h) || /amount|debit|credit|diff/i.test(h)) moneyColIdx.push(i + 1);
  });
  const remarksIdx = headers.indexOf('Remarks');

  for (const r of rows) {
    const values = headers.map((h) => normalizeForCell(r[h]));
    const row = ws.addRow(values);
    for (const c of moneyColIdx) row.getCell(c).numFmt = MONEY_FMT;
    if (remarksIdx >= 0) {
      const rem = String(values[remarksIdx] ?? '');
      const fill = remarkFill(rem);
      if (fill) row.eachCell((cell) => { cell.fill = fill; });
    }
  }

  for (let i = 0; i < headers.length; i++) {
    const col = ws.getColumn(i + 1);
    let max = headers[i].length;
    col.eachCell({ includeEmpty: false }, (cell) => {
      const v = cell.value;
      if (v === null || v === undefined) return;
      const s = v instanceof Date ? v.toISOString().slice(0, 10) : String(v);
      if (s.length > max) max = s.length;
    });
    col.width = Math.min(Math.max(12, max + 2), 40);
  }
}

// ── Recon Summary sheet ───────────────────────────────────────────────────────
function writeSummary(wb: ExcelJS.Workbook, res: ReconResult, generatedAt: Date): void {
  const ws = wb.addWorksheet('Recon Summary');
  ws.columns = [
    { key: 'particulars', width: 30 },
    { key: 'mars', width: 18 },
    { key: 'brand', width: 18 },
    { key: 'diff', width: 18 },
  ];

  const titleRow = ws.addRow(['Recon Summary – Category Wise']);
  titleRow.getCell(1).font = { bold: true, size: 13 };
  ws.mergeCells(titleRow.number, 1, titleRow.number, 4);
  ws.addRow([]);

  const tableHeader = ws.addRow(['Particulars', 'Amount_Mars', 'Amount_Brand', 'Difference']);
  styleHeader(tableHeader);

  for (const r of res.summary) {
    const row = ws.addRow([r.Particulars, r.Amount_Mars, r.Amount_Brand, r.Difference]);
    for (let c = 2; c <= 4; c++) row.getCell(c).numFmt = MONEY_FMT;
    if (r.Particulars === 'Grand Total') {
      row.eachCell((cell) => {
        cell.font = { bold: true };
        cell.fill = TOTAL_FILL;
      });
    }
  }

  ws.addRow([]);
  ws.addRow([]);

  const statsTitle = ws.addRow(['Match Statistics']);
  statsTitle.getCell(1).font = { bold: true, size: 12 };

  const statsHeader = ws.addRow(['Metric', 'Mars side', 'Brand side']);
  styleHeader(statsHeader);

  const s = res.stats;
  ws.addRow(['Recon period rows', s.mars_recon_rows, s.brand_recon_rows]);
  ws.addRow(['Match', s.mars_match, s.brand_match]);
  ws.addRow(['Amount Mismatch', s.mars_mismatch, s.brand_mismatch]);
  ws.addRow(['Reversal', '', s.brand_reversal]);
  ws.addRow(['Not Booked in Brands Ledger', s.mars_not_booked_by_brand, '']);
  ws.addRow(['Not Booked by Mars', '', s.brand_not_booked_by_mars]);

  ws.addRow([]);
  ws.addRow(['Generated at', formatTimestamp(generatedAt)]);
}

