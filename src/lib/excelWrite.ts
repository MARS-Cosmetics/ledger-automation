import ExcelJS from 'exceljs';
import { canonicalizeCategory } from './reconcile';
import type { ReconResult, Row } from './types';

const MONEY_FMT = '#,##,##0.00;(#,##,##0.00);"-"';
const BASE_FONT = { name: 'Calibri', size: 9 } as const;
const THIN_BORDER: ExcelJS.Borders = {
  top:      { style: 'thin' },
  left:     { style: 'thin' },
  bottom:   { style: 'thin' },
  right:    { style: 'thin' },
  diagonal: {},
};

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

// Columns for the "Open Points Summary" sheet (Sheet 5)
const OPS_HEADERS = [
  'Particular',
  'Category',
  'Sub-Category',
  'Count',
  'Amount_MARS',
  'Amount_Brand',
  'Difference',
  'Annexure',
  'Action On',
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
  writeOpenPointsSummary(wb, res);                                        // Sheet 5

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
  // Within each category, rows are grouped by remark (no intermixing).
  // Particular column shows the raw ledger value, not any merged label.
  const collectRows = (orderedCats: string[]): Row[] => {
    const rows: Row[] = [];
    for (const cat of orderedCats) {
      const catRows: Row[] = [];
      for (const r of res.mars.rows) {
        const rawCat = marsCatCol ? String(r[marsCatCol] ?? '') : '';
        if (normCat(rawCat) !== cat) continue;
        catRows.push({
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
        catRows.push({
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
      catRows.sort((a, b) =>
        remarkSortOrder(String(a['Category'] ?? '')) - remarkSortOrder(String(b['Category'] ?? ''))
      );
      rows.push(...catRows);
    }
    return rows;
  };

  // Helper: write a titled section into the sheet
  const writeSection = (label: string, rows: Row[]) => {
    if (rows.length === 0) return;
    const titleRow = ws.addRow([label]);
    titleRow.height = 18;
    ws.mergeCells(titleRow.number, 1, titleRow.number, ncols);
    titleRow.font = { ...BASE_FONT, bold: true };
    titleRow.alignment = { vertical: 'middle' };
    borderRow(titleRow, ncols);
    const hRow = ws.addRow([...ANNEX_HEADERS]);
    styleHeader(hRow, ncols);
    for (const r of rows) {
      const values = ANNEX_HEADERS.map((h) => normalizeForCell(r[h]));
      const row = ws.addRow(values);
      row.font = BASE_FONT;
      borderRow(row, ncols);
      for (const c of [6, 7, 8]) row.getCell(c).numFmt = MONEY_FMT;
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


// ── Sheet 5: Open Points Summary — category × remark summary with formulas ────
function writeOpenPointsSummary(wb: ExcelJS.Workbook, res: ReconResult): void {
  const mc = res.marsCols;
  const bc = res.brandCols;
  const marsCatCol = mc.category;
  const brandCatCol = bc.category;

  if (!mc.vch_no || !mc.net_amount || !bc.reference || !bc.net_amount) return;

  // ── collect same category set as Annex (excluding Opening Balance) ──────────
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

  // ── collect unique (canonCat, remark) pairs in Annex section order ──────────
  type SummaryPair = { canonCat: string; remark: string; annexureLabel: string };
  const pairs: SummaryPair[] = [];
  const handledCats = new Set<string>();
  const seenKeys = new Set<string>(); // "canonCat||remark"

  const collectPairs = (orderedCats: string[], annexureLabel: string): void => {
    const temp: SummaryPair[] = [];
    for (const cat of orderedCats) {
      for (const r of res.mars.rows) {
        const rawCat = marsCatCol ? String(r[marsCatCol] ?? '') : '';
        if (normCat(rawCat) !== cat) continue;
        const remark = String(r['Remarks'] ?? '');
        const key = `${cat}||${remark}`;
        if (!seenKeys.has(key)) { seenKeys.add(key); temp.push({ canonCat: cat, remark, annexureLabel }); }
      }
      for (const r of res.brand.rows) {
        if (toNum(r['Amount_Mars']) !== 0) continue;
        const rawCat = brandCatCol ? String(r[brandCatCol] ?? '') : '';
        if (normCat(rawCat) !== cat) continue;
        const remark = String(r['Remarks'] ?? '');
        const key = `${cat}||${remark}`;
        if (!seenKeys.has(key)) { seenKeys.add(key); temp.push({ canonCat: cat, remark, annexureLabel }); }
      }
    }
    temp.sort((a, b) => remarkSortOrder(a.remark) - remarkSortOrder(b.remark));
    pairs.push(...temp);
  };

  for (const cat of cats) {
    if (handledCats.has(cat)) continue;
    const mergedGroup = ANNEX_MERGED_GROUPS.find((g) => g.cats.includes(cat));
    if (mergedGroup) {
      mergedGroup.cats.forEach((c) => handledCats.add(c));
      collectPairs(mergedGroup.cats, mergedGroup.label);
    } else {
      handledCats.add(cat);
      collectPairs([cat], cat);
    }
  }

  if (pairs.length === 0) return;

  // ── write sheet ─────────────────────────────────────────────────────────────
  const ws = wb.addWorksheet('Open Points Summary');
  ws.views = [{ state: 'frozen', ySplit: 2 }];
  const ncols = OPS_HEADERS.length;

  // Title row
  const titleRow = ws.addRow(['Open Points Summary']);
  titleRow.height = 18;
  ws.mergeCells(titleRow.number, 1, titleRow.number, ncols);
  titleRow.font = { ...BASE_FONT, bold: true };
  titleRow.alignment = { vertical: 'middle' };
  borderRow(titleRow, ncols);

  // Header row
  styleHeader(ws.addRow([...OPS_HEADERS]), ncols);

  // Data rows — Count/Amount columns use COUNTIFS/SUMIFS referencing the Annex sheet
  // Annex col F = Amount_MARS, col G = Amount_Brand  (positions defined in ANNEX_HEADERS)
  for (const { canonCat, remark, annexureLabel } of pairs) {
    const row = ws.addRow([]);
    const rn = row.number;
    row.font = BASE_FONT;
    borderRow(row, ncols);

    row.getCell(1).value = canonCat;
    row.getCell(2).value = remark;
    row.getCell(3).value = '';
    row.getCell(4).value = { formula: `=COUNTIFS('Annex'!$A:$A,$A${rn},'Annex'!$B:$B,$B${rn})` };
    row.getCell(5).value = { formula: `=SUMIFS('Annex'!$F:$F,'Annex'!$A:$A,$A${rn},'Annex'!$B:$B,$B${rn})` };
    row.getCell(6).value = { formula: `=SUMIFS('Annex'!$G:$G,'Annex'!$A:$A,$A${rn},'Annex'!$B:$B,$B${rn})` };
    row.getCell(7).value = { formula: `=E${rn}+F${rn}` };
    row.getCell(8).value = annexureLabel;
    row.getCell(9).value = { formula: `=IF($B${rn}="Match","Closed","Open Point")` };

    for (const c of [5, 6, 7]) row.getCell(c).numFmt = MONEY_FMT;
  }

  // Column widths
  [22, 30, 14, 10, 18, 18, 16, 28, 14].forEach((w, i) => { ws.getColumn(i + 1).width = w; });
}


// ── helpers ───────────────────────────────────────────────────────────────────
function normCat(raw: string): string {
  const c = canonicalizeCategory(raw.trim());
  return c || raw.trim() || 'Uncategorized';
}


function remarkSortOrder(remark: string): number {
  if (remark === 'Match') return 0;
  if (remark === 'Amount Mismatch') return 1;
  if (remark.startsWith('Not Booked in Brands')) return 2;
  if (remark.startsWith('Not Booked by Mars')) return 3;
  if (remark.startsWith('Not Booked')) return 4;
  if (remark.toLowerCase().includes('reversal')) return 5;
  return 6;
}


function toNum(v: unknown): number {
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  if (v === null || v === undefined) return 0;
  const s = String(v).replace(/,/g, '').trim();
  if (!s || s.toLowerCase() === 'nan') return 0;
  const n = Number(s);
  return isFinite(n) ? n : 0;
}

function borderRow(row: ExcelJS.Row, numCols: number): void {
  for (let c = 1; c <= numCols; c++) row.getCell(c).border = THIN_BORDER;
}

function styleHeader(row: ExcelJS.Row, numCols: number): void {
  row.font = { ...BASE_FONT, bold: true };
  borderRow(row, numCols);
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
  styleHeader(headerRow, headers.length);

  const moneyColIdx: number[] = [];
  headers.forEach((h, i) => {
    if (MONEY_HEADERS.has(h) || /amount|debit|credit|diff/i.test(h)) moneyColIdx.push(i + 1);
  });
  for (const r of rows) {
    const values = headers.map((h) => normalizeForCell(r[h]));
    const row = ws.addRow(values);
    row.font = BASE_FONT;
    borderRow(row, headers.length);
    for (const c of moneyColIdx) row.getCell(c).numFmt = MONEY_FMT;
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
  titleRow.font = { ...BASE_FONT, bold: true };
  ws.mergeCells(titleRow.number, 1, titleRow.number, 4);
  ws.addRow([]);

  const tableHeader = ws.addRow(['Particulars', 'Amount_Mars', 'Amount_Brand', 'Difference']);
  styleHeader(tableHeader, 4);

  for (const r of res.summary) {
    const row = ws.addRow([r.Particulars, r.Amount_Mars, r.Amount_Brand, r.Difference]);
    for (let c = 2; c <= 4; c++) row.getCell(c).numFmt = MONEY_FMT;
    row.font = r.Particulars === 'Grand Total' ? { ...BASE_FONT, bold: true } : BASE_FONT;
    borderRow(row, 4);
  }

  ws.addRow([]);
  ws.addRow([]);

  const statsTitle = ws.addRow(['Match Statistics']);
  statsTitle.font = { ...BASE_FONT, bold: true };

  const statsHeader = ws.addRow(['Metric', 'Mars side', 'Brand side']);
  styleHeader(statsHeader, 3);

  const s = res.stats;
  const statsData: [string, number | string, number | string][] = [
    ['Recon period rows',           s.mars_recon_rows,       s.brand_recon_rows],
    ['Match',                       s.mars_match,            s.brand_match],
    ['Amount Mismatch',             s.mars_mismatch,         s.brand_mismatch],
    ['Reversal',                    '',                      s.brand_reversal],
    ['Not Booked in Brands Ledger', s.mars_not_booked_by_brand, ''],
    ['Not Booked by Mars',          '',                      s.brand_not_booked_by_mars],
  ];
  for (const data of statsData) {
    const row = ws.addRow(data);
    row.font = BASE_FONT;
    borderRow(row, 3);
  }

  ws.addRow([]);
  const genRow = ws.addRow(['Generated at', formatTimestamp(generatedAt)]);
  genRow.font = BASE_FONT;
  borderRow(genRow, 2);
}

