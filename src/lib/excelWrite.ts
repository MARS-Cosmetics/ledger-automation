import ExcelJS from 'exceljs';
import type { ReconResult, Row } from './types';

const HEADER_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDCE6F1' } } as const;
const TOTAL_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } } as const;
const NOT_BOOKED_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFCE4D6' } } as const;
const MISMATCH_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF4CE' } } as const;
const MONEY_FMT = '#,##,##0.00;(#,##,##0.00);"-"';

const MONEY_HEADERS = new Set([
  'Amount_Mars',
  'Amount_Brand',
  'Difference',
  'Diff',
  'Net Amount',
  'Net amount',
  'Debit(Rs.)',
  'Credit(Rs.)',
  'Debit',
  'Credit',
  'Amount in local currency',
]);

export async function buildReconWorkbook(res: ReconResult, generatedAt: Date = new Date()): Promise<ArrayBuffer> {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Mars Recon Tool';
  wb.created = generatedAt;

  writeSummary(wb, res, generatedAt);
  writeDataSheet(wb, 'Mars Cosmetics', res.mars.headers, res.mars.rows);
  writeDataSheet(wb, 'Brand', res.brand.headers, res.brand.rows);

  const buf = await wb.xlsx.writeBuffer();
  return buf as ArrayBuffer;
}

function writeSummary(wb: ExcelJS.Workbook, res: ReconResult, generatedAt: Date): void {
  const ws = wb.addWorksheet('Recon Summary');
  ws.columns = [
    { key: 'particulars', width: 30 },
    { key: 'mars', width: 18 },
    { key: 'brand', width: 18 },
    { key: 'diff', width: 18 },
  ];

  const stats: [string, string | number][] = [
    ['Mars rows in Recon period', res.stats.mars_recon_rows],
    ['Brand rows in Recon period', res.stats.brand_recon_rows],
    ['Mars rows matched', res.stats.mars_matched_rows],
    ['Mars rows not booked by Brand', res.stats.mars_unmatched_rows],
    ['Brand rows not booked by Mars', res.stats.brand_unmatched_rows],
    ['Generated at', formatTimestamp(generatedAt)],
  ];

  const statsHeader = ws.addRow(['Metric', 'Value']);
  styleHeader(statsHeader);
  for (const [label, val] of stats) ws.addRow([label, val]);

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
}

function writeDataSheet(wb: ExcelJS.Workbook, name: string, headers: string[], rows: Row[]): void {
  const ws = wb.addWorksheet(name);
  ws.views = [{ state: 'frozen', ySplit: 1 }];

  const headerRow = ws.addRow(headers);
  styleHeader(headerRow);

  const moneyColIdx: number[] = [];
  headers.forEach((h, i) => {
    if (MONEY_HEADERS.has(h)) moneyColIdx.push(i + 1);
    else if (/amount|debit|credit|diff/i.test(h)) moneyColIdx.push(i + 1);
  });
  const remarksIdx = headers.indexOf('Remarks');

  for (const r of rows) {
    const values = headers.map((h) => normalizeForCell(r[h]));
    const row = ws.addRow(values);
    for (const c of moneyColIdx) row.getCell(c).numFmt = MONEY_FMT;
    if (remarksIdx >= 0) {
      const rem = String(values[remarksIdx] ?? '');
      if (rem.startsWith('Not Booked')) {
        row.eachCell((cell) => {
          cell.fill = NOT_BOOKED_FILL;
        });
      } else if (rem.startsWith('Amount Mismatch')) {
        row.eachCell((cell) => {
          cell.fill = MISMATCH_FILL;
        });
      }
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
