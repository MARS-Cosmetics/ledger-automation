import type { ColumnMapping, ReconResult, Row, SummaryRow } from './types';

export const CANONICAL_CATEGORIES = [
  'Invoice',
  'Contra_Inv',
  'CN',
  'Contra_CN',
  'DN',
  'Others',
] as const;

const CATEGORY_MAP: Record<string, string> = {
  invoice: 'Invoice',
  inv: 'Invoice',
  contra_inv: 'Contra_Inv',
  contra_invoice: 'Contra_Inv',
  contrainv: 'Contra_Inv',
  cn: 'CN',
  credit_note: 'CN',
  creditnote: 'CN',
  contra_cn: 'Contra_CN',
  contracn: 'Contra_CN',
  dn: 'DN',
  debit_note: 'DN',
  debitnote: 'DN',
};

const AMOUNT_TOLERANCE = 0.005;

function normStr(x: unknown): string {
  if (x === null || x === undefined) return '';
  const s = String(x).trim();
  return s === 'NaN' || s === 'nan' ? '' : s;
}

function normKey(x: unknown): string {
  return normStr(x).toUpperCase();
}

function normCatForMatch(x: unknown): string {
  let s = normStr(x).toLowerCase().replace(/-/g, '_').replace(/\s+/g, '_');
  while (s.includes('__')) s = s.replace(/__/g, '_');
  return s;
}

export function toCanonicalCategory(x: unknown): string {
  const s = normCatForMatch(x);
  if (!s) return 'Invoice';
  return CATEGORY_MAP[s] ?? 'Others';
}

function toNumber(x: unknown): number {
  if (typeof x === 'number') return Number.isFinite(x) ? x : 0;
  if (typeof x === 'boolean') return x ? 1 : 0;
  if (x === null || x === undefined) return 0;
  const s = String(x).replace(/,/g, '').trim();
  if (s === '' || s.toLowerCase() === 'nan') return 0;
  const n = Number(s);
  return Number.isFinite(n) ? n : 0;
}

function approxEqual(a: number, b: number): boolean {
  return Math.abs(a - b) < AMOUNT_TOLERANCE;
}

function filterRecon(rows: Row[], periodCol: string | undefined): Row[] {
  if (!periodCol) return rows;
  return rows.filter((r) => {
    const v = normStr(r[periodCol]).toLowerCase();
    return v === '' || v === 'recon';
  });
}

interface AggValue {
  amount: number;
  count: number;
}

function groupSum(rows: Row[], keyFn: (r: Row) => string, amtFn: (r: Row) => number): Map<string, AggValue> {
  const out = new Map<string, AggValue>();
  for (const r of rows) {
    const k = keyFn(r);
    if (!k) continue;
    const cur = out.get(k);
    const amt = amtFn(r);
    if (cur) {
      cur.amount += amt;
      cur.count += 1;
    } else {
      out.set(k, { amount: amt, count: 1 });
    }
  }
  return out;
}

export function reconcile(
  marsRows: Row[],
  brandRows: Row[],
  marsCols: ColumnMapping,
  brandCols: ColumnMapping,
): ReconResult {
  for (const k of ['vch_no', 'net_amount'] as const) {
    if (!marsCols[k]) throw new Error(`Mars column mapping missing required key: ${k}`);
  }
  for (const k of ['reference', 'net_amount'] as const) {
    if (!brandCols[k]) throw new Error(`Brand column mapping missing required key: ${k}`);
  }

  const marsVchCol = marsCols.vch_no!;
  const marsAmtCol = marsCols.net_amount!;
  const marsCatCol = marsCols.category;
  const marsPeriodCol = marsCols.period;

  const brandRefCol = brandCols.reference!;
  const brandCorrectRefCol = brandCols.correct_ref;
  const brandAmtCol = brandCols.net_amount!;
  const brandCatCol = brandCols.category;
  const brandPeriodCol = brandCols.period;

  const mars = filterRecon(marsRows, marsPeriodCol);
  const brand = filterRecon(brandRows, brandPeriodCol);

  const marsKey = (r: Row) => normKey(r[marsVchCol]);
  const marsAmt = (r: Row) => toNumber(r[marsAmtCol]);

  const brandKey = (r: Row) => {
    if (brandCorrectRefCol) {
      const corrected = normKey(r[brandCorrectRefCol]);
      if (corrected) return corrected;
    }
    return normKey(r[brandRefCol]);
  };
  const brandAmt = (r: Row) => toNumber(r[brandAmtCol]);

  const brandAgg = groupSum(brand, brandKey, brandAmt);
  const marsAgg = groupSum(mars, marsKey, marsAmt);

  const marsHeaders = inferHeaders(marsRows);
  const brandHeaders = inferHeaders(brandRows);

  const marsOutHeaders = [...marsHeaders, 'Amount_Brand', 'Difference', 'Remarks'];
  const brandOutHeaders = [...brandHeaders, 'Amount_Mars', 'Diff', 'Remarks'];

  const marsOut: Row[] = mars.map((r) => {
    const key = marsKey(r);
    const own = marsAmt(r);
    const brandRow = brandAgg.get(key);
    const amtBrand = brandRow?.amount ?? 0;
    const diff = own - amtBrand;
    let remarks: string;
    if (!brandRow) {
      remarks = 'Not Booked by Brand';
    } else if (!approxEqual(own, amtBrand)) {
      remarks = brandRow.count > 1 ? `Amount Mismatch (Brand split: ${brandRow.count} rows)` : 'Amount Mismatch';
    } else {
      remarks = brandRow.count > 1 ? `Matched (Brand split: ${brandRow.count} rows)` : 'Matched';
    }
    return { ...r, Amount_Brand: amtBrand, Difference: diff, Remarks: remarks };
  });

  const brandOut: Row[] = brand.map((r) => {
    const key = brandKey(r);
    const own = brandAmt(r);
    const marsRow = marsAgg.get(key);
    const amtMars = marsRow?.amount ?? 0;
    const diff = own - amtMars;
    let remarks: string;
    if (!marsRow) {
      remarks = 'Not Booked by Mars';
    } else if (!approxEqual(own, amtMars)) {
      remarks = marsRow.count > 1 ? `Amount Mismatch (Mars split: ${marsRow.count} rows)` : 'Amount Mismatch';
    } else {
      remarks = marsRow.count > 1 ? `Matched (Mars split: ${marsRow.count} rows)` : 'Matched';
    }
    return { ...r, Amount_Mars: amtMars, Diff: diff, Remarks: remarks };
  });

  const summary = buildSummary(mars, brand, marsAmtCol, brandAmtCol, marsCatCol, brandCatCol);

  const matched = marsOut.filter((r) => r.Remarks !== 'Not Booked by Brand').length;
  const unmatchedMars = marsOut.filter((r) => r.Remarks === 'Not Booked by Brand').length;
  const unmatchedBrand = brandOut.filter((r) => r.Remarks === 'Not Booked by Mars').length;

  return {
    summary,
    mars: { headers: marsOutHeaders, rows: marsOut },
    brand: { headers: brandOutHeaders, rows: brandOut },
    stats: {
      mars_recon_rows: marsOut.length,
      brand_recon_rows: brandOut.length,
      mars_matched_rows: matched,
      mars_unmatched_rows: unmatchedMars,
      brand_unmatched_rows: unmatchedBrand,
    },
  };
}

function inferHeaders(rows: Row[]): string[] {
  if (rows.length === 0) return [];
  const seen = new Set<string>();
  const out: string[] = [];
  for (const r of rows) {
    for (const k of Object.keys(r)) {
      if (!seen.has(k)) {
        seen.add(k);
        out.push(k);
      }
    }
  }
  return out;
}

function buildSummary(
  mars: Row[],
  brand: Row[],
  marsAmtCol: string,
  brandAmtCol: string,
  marsCatCol: string | undefined,
  brandCatCol: string | undefined,
): SummaryRow[] {
  const catOf = (catCol: string | undefined, r: Row): string =>
    catCol ? toCanonicalCategory(r[catCol]) : 'Invoice';

  const marsByCat = new Map<string, number>();
  for (const r of mars) {
    const c = catOf(marsCatCol, r);
    marsByCat.set(c, (marsByCat.get(c) ?? 0) + toNumber(r[marsAmtCol]));
  }
  const brandByCat = new Map<string, number>();
  for (const r of brand) {
    const c = catOf(brandCatCol, r);
    brandByCat.set(c, (brandByCat.get(c) ?? 0) + toNumber(r[brandAmtCol]));
  }

  const present = new Set<string>([...marsByCat.keys(), ...brandByCat.keys()]);
  const ordered = (CANONICAL_CATEGORIES as readonly string[]).filter((c) => present.has(c));
  const extras = [...present].filter((c) => !(CANONICAL_CATEGORIES as readonly string[]).includes(c));
  const final = [...ordered, ...extras.sort()];

  const rows: SummaryRow[] = final.map((cat) => {
    const m = marsByCat.get(cat) ?? 0;
    const b = brandByCat.get(cat) ?? 0;
    return { Particulars: cat, Amount_Mars: m, Amount_Brand: b, Difference: m - b };
  });

  const totalMars = rows.reduce((s, r) => s + r.Amount_Mars, 0);
  const totalBrand = rows.reduce((s, r) => s + r.Amount_Brand, 0);
  rows.push({
    Particulars: 'Grand Total',
    Amount_Mars: totalMars,
    Amount_Brand: totalBrand,
    Difference: totalMars - totalBrand,
  });

  return rows;
}
