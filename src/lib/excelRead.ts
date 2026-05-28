import ExcelJS from 'exceljs';
import type { Cell, Row, Sheet, Workbook } from './types';

export async function readWorkbook(file: File): Promise<Workbook> {
  const buffer = await file.arrayBuffer();
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);

  const sheets: Sheet[] = [];
  wb.eachSheet((ws) => {
    sheets.push(parseSheet(ws));
  });
  return { sheets };
}

function parseSheet(ws: ExcelJS.Worksheet): Sheet {
  const allRows: Cell[][] = [];
  ws.eachRow({ includeEmpty: false }, (row) => {
    const values: Cell[] = [];
    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      values[colNumber - 1] = extractCellValue(cell);
    });
    allRows.push(values);
  });

  const headerIndex = pickHeaderRow(allRows);
  const headerRowRaw = allRows[headerIndex] ?? [];
  const headers = uniqueHeaders(headerRowRaw);

  const dataRows = allRows.slice(headerIndex + 1);
  const rows: Row[] = [];
  for (const r of dataRows) {
    if (isBlankRow(r)) continue;
    const obj: Row = {};
    for (let i = 0; i < headers.length; i++) {
      obj[headers[i]] = r[i] ?? null;
    }
    rows.push(obj);
  }
  return { name: ws.name, headers, rows };
}

function extractCellValue(cell: ExcelJS.Cell): Cell {
  const v = cell.value;
  if (v === null || v === undefined) return null;
  if (typeof v === 'string') return v.trim() === '' ? null : v;
  if (typeof v === 'number' || typeof v === 'boolean') return v;
  if (v instanceof Date) return v;
  if (typeof v === 'object') {
    // Formula error cell: { error: '#VALUE!' } etc.
    if ('error' in v) return null;
    // Rich text
    if ('richText' in v && Array.isArray((v as { richText: unknown }).richText)) {
      const text = (v as { richText: { text: string }[] }).richText.map((p) => p.text).join('').trim();
      return text === '' ? null : text;
    }
    // Hyperlink / formula with text
    if ('text' in v && typeof (v as { text: unknown }).text === 'string') {
      const text = (v as { text: string }).text.trim();
      return text === '' ? null : text;
    }
    // Formula result
    if ('result' in v) {
      const r = (v as { result: unknown }).result;
      if (r === null || r === undefined) return null;
      if (typeof r === 'string') return r.trim() === '' ? null : r;
      if (typeof r === 'number' || typeof r === 'boolean') return r;
      if (r instanceof Date) return r;
      // Formula error in result
      if (typeof r === 'object' && r !== null && 'error' in r) return null;
    }
  }
  return String(v);
}

function isBlankRow(r: Cell[]): boolean {
  return r.every((c) => c === null || c === undefined || (typeof c === 'string' && c.trim() === ''));
}

function pickHeaderRow(rows: Cell[][]): number {
  const limit = Math.min(rows.length, 20);

  // Pass 1: row where every non-blank cell is a string — strongest signal for a header row
  // (real data rows always contain at least one number or date)
  for (let i = 0; i < limit; i++) {
    const r = rows[i];
    if (!r) continue;
    const nonBlank = r.filter((c) => c !== null && c !== undefined && String(c).trim() !== '');
    if (nonBlank.length >= 3 && nonBlank.every((c) => typeof c === 'string')) return i;
  }

  // Pass 2: first row with ≥2 non-blank scalar cells (catches edge cases like numeric column headers)
  for (let i = 0; i < limit; i++) {
    const r = rows[i];
    if (!r) continue;
    const nonBlank = r.filter((c) => c !== null && c !== undefined && String(c).trim() !== '');
    if (nonBlank.length >= 2 && nonBlank.every((c) => typeof c === 'string' || typeof c === 'number')) return i;
  }

  return 0;
}

function uniqueHeaders(raw: Cell[]): string[] {
  const seen = new Map<string, number>();
  const out: string[] = [];
  for (let i = 0; i < raw.length; i++) {
    let h = raw[i];
    let label = h === null || h === undefined ? '' : String(h).trim();
    if (!label) label = `Column ${i + 1}`;
    const count = seen.get(label) ?? 0;
    seen.set(label, count + 1);
    out.push(count === 0 ? label : `${label} (${count + 1})`);
  }
  return out;
}
