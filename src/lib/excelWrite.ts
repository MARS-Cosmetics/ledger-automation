import ExcelJS from 'exceljs';
import type { ReconResult, Row } from './types';

const HEADER_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDCE6F1' } } as const;
const TOTAL_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } } as const;
const NOT_BOOKED_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFCE4D6' } } as const;
const MISMATCH_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF4CE' } } as const;
const MATCH_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2EFDA' } } as const;
const REVERSAL_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDDEBF7' } } as const;
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

  // 1. Write Reconciliation Summary sheet
  writeReconciliationSummary(wb, res, generatedAt);

  // 2. Write Annexure sheets
  writeAnnexureSheets(wb, res);

  // 3. Write Open Points Summary sheet
  writeOpenPointsSummary(wb, res);

  // 4. Write original data sheets
  writeDataSheet(wb, 'Mars Cosmetics', res.mars.headers, res.mars.rows);
  writeDataSheet(wb, 'Brand', res.brand.headers, res.brand.rows);

  const buf = await wb.xlsx.writeBuffer();
  return buf as ArrayBuffer;
}

function writeReconciliationSummary(wb: ExcelJS.Workbook, res: ReconResult, generatedAt: Date): void {
  const ws = wb.addWorksheet('Reconciliation Summary');
  ws.columns = [
    { key: 'particulars', width: 30 },
    { key: 'mars', width: 18 },
    { key: 'brand', width: 18 },
    { key: 'diff', width: 18 },
  ];

  const titleRow = ws.addRow(['Reconciliation Summary – Category Wise']);
  titleRow.getCell(1).font = { bold: true, size: 13 };
  ws.mergeCells(titleRow.number, 1, titleRow.number, 4);
  ws.addRow([]);

  const tableHeader = ws.addRow(['Category', 'Amount MARS', 'Amount Brand', 'Difference']);
  styleHeader(tableHeader);

  for (const r of res.summary) {
    const row = ws.addRow([r.Particulars, r.Amount_Mars, r.Amount_Brand, r.Difference]);
    for (let c = 2; c <= 4; c++) row.getCell(c).numFmt = MONEY_FMT;
    if (r.Particulars === 'Closing Balance' || r.Particulars === 'Opening Balance') {
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
  ws.addRow(['Not Booked by Brand', s.mars_not_booked_by_brand, '']);
  ws.addRow(['Not Booked by Mars', '', s.brand_not_booked_by_mars]);

  ws.addRow([]);
  ws.addRow(['Generated at', formatTimestamp(generatedAt)]);
}

function writeAnnexureSheets(wb: ExcelJS.Workbook, res: ReconResult): void {
  for (const annexure of res.annexures) {
    const sheetName = `Annexure_${annexure.category.replace(/\s+/g, '_')}`;
    const ws = wb.addWorksheet(sheetName);

    const headers = [
      'Particular Categorized in Both Ledgers',
      'Category (Remark Given by System)',
      'Sub Category',
      'Date',
      'Invoice/Voucher Number',
      'Amount MARS',
      'Amount Brand',
      'Difference',
    ];

    const headerRow = ws.addRow(headers);
    styleHeader(headerRow);
    ws.views = [{ state: 'frozen', ySplit: 1 }];

    for (const annexureRow of annexure.rows) {
      const row = ws.addRow([
        annexureRow.Particular,
        annexureRow.Category,
        annexureRow.SubCategory,
        annexureRow.Date,
        annexureRow.ReferenceNumber,
        annexureRow.Amount_Mars,
        annexureRow.Amount_Brand,
        annexureRow.Difference,
      ]);

      // Format money columns
      for (let c = 6; c <= 8; c++) {
        row.getCell(c).numFmt = MONEY_FMT;
      }
    }

    // Add totals row
    ws.addRow([]);
    const totalRow = ws.addRow([
      'TOTAL',
      '',
      '',
      '',
      '',
      `=SUM(F2:F${annexure.rows.length + 1})`,
      `=SUM(G2:G${annexure.rows.length + 1})`,
      `=SUM(H2:H${annexure.rows.length + 1})`,
    ]);

    totalRow.eachCell((cell, colNum: number) => {
      cell.font = { bold: true };
      cell.fill = TOTAL_FILL;
      if (colNum >= 6 && colNum <= 8) {
        cell.numFmt = MONEY_FMT;
      }
    });

    // Set column widths
    ws.columns.forEach((col, idx) => {
      let width = headers[idx]?.length ?? 15;
      width = Math.min(Math.max(12, width + 2), 50);
      col.width = width;
    });
  }
}

function writeOpenPointsSummary(wb: ExcelJS.Workbook, res: ReconResult): void {
  const ws = wb.addWorksheet('Open Points Summary');

  const headers = [
    'Particular',
    'Category',
    'Sub Category',
    'Count',
    'Amount MARS',
    'Amount Brand',
    'Difference',
    'Annexure Link',
    'Action On',
  ];

  const headerRow = ws.addRow(headers);
  styleHeader(headerRow);
  ws.views = [{ state: 'frozen', ySplit: 1 }];

  for (let i = 0; i < res.openPoints.length; i++) {
    const openPoint = res.openPoints[i];

    const row = ws.addRow([
      openPoint.Particular,
      openPoint.Category,
      openPoint.SubCategory,
      openPoint.Count,
      openPoint.Amount_Mars,
      openPoint.Amount_Brand,
      openPoint.Difference,
      openPoint.AnnexureLink,
      openPoint.ActionOn,
    ]);

    // Format money columns
    for (let c = 5; c <= 7; c++) {
      row.getCell(c).numFmt = MONEY_FMT;
    }

    // Add hyperlink to annexure
    const linkCell = row.getCell(8);
    linkCell.value = {
      text: openPoint.AnnexureLink,
      hyperlink: `#'${openPoint.AnnexureLink}'!A1`,
    };
    linkCell.font = { color: { argb: 'FF0563C1' }, underline: true };

    // Color code the action
    if (openPoint.ActionOn === 'Closed') {
      row.getCell(9).fill = MATCH_FILL;
    } else {
      row.getCell(9).fill = NOT_BOOKED_FILL;
    }
  }

  // Add a summary row
  ws.addRow([]);
  const summaryRow = ws.addRow([
    'TOTAL',
    '',
    '',
    `=SUM(D2:D${res.openPoints.length + 1})`,
    `=SUM(E2:E${res.openPoints.length + 1})`,
    `=SUM(F2:F${res.openPoints.length + 1})`,
    `=SUM(G2:G${res.openPoints.length + 1})`,
    '',
    '',
  ]);

  summaryRow.eachCell((cell, colNum: number) => {
    cell.font = { bold: true };
    cell.fill = TOTAL_FILL;
    if (colNum === 4) {
      cell.numFmt = '0';
    } else if (colNum >= 5 && colNum <= 7) {
      cell.numFmt = MONEY_FMT;
    }
  });

  // Set column widths
  ws.columns.forEach((col, idx) => {
    let width = headers[idx]?.length ?? 15;
    width = Math.min(Math.max(12, width + 2), 50);
    col.width = width;
  });
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
      } else if (rem === 'Amount Mismatch') {
        row.eachCell((cell) => {
          cell.fill = MISMATCH_FILL;
        });
      } else if (rem === 'Match') {
        row.eachCell((cell) => {
          cell.fill = MATCH_FILL;
        });
      } else if (rem.startsWith('Reversal')) {
        row.eachCell((cell) => {
          cell.fill = REVERSAL_FILL;
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
