import type {
  ColumnMapping,
  DiagnosticInfo,
  ReconOptions,
  ReconResult,
  Row,
  SummaryRow,
} from './types';

export const CANONICAL_CATEGORY_ORDER = [
  'Invoice',
  'Contra_Inv',
  'CN',
  'Contra_CN',
  'DN',
] as const;

const CANONICAL_NORMALIZE: Record<string, string> = {
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

// Categories that name the same thing on Mars vs Brand sides.
// Used only for the match key, not for display / summary breakdown.
const CATEGORY_MATCH_EQUIVALENCE: Record<string, string> = {
  receipt: 'payment_receipt',
  payment: 'payment_receipt',
};

export const DEFAULT_OPTIONS: ReconOptions = {
  matchMode: 'vch_and_category',
  matchToleranceRupees: 5,
  acceptBlankPeriod: false,
};

function normStr(x: unknown): string {
  if (x === null || x === undefined) return '';
  const s = String(x).trim();
  return s === 'NaN' || s === 'nan' ? '' : s;
}

function normKey(x: unknown): string {
  return normStr(x).toUpperCase();
}

function normCatKey(x: unknown): string {
  let s = normStr(x).toLowerCase().replace(/-/g, '_').replace(/\s+/g, '_');
  while (s.includes('__')) s = s.replace(/__/g, '_');
  return s;
}

function matchCatKey(x: unknown): string {
  const k = normCatKey(x);
  return CATEGORY_MATCH_EQUIVALENCE[k] ?? k;
}

export function canonicalizeCategory(x: unknown): string {
  const raw = normStr(x);
  if (!raw) return '';
  const key = normCatKey(raw);
  return CANONICAL_NORMALIZE[key] ?? raw;
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

function isReconPeriod(v: unknown, acceptBlank: boolean): boolean {
  const s = normStr(v).toLowerCase();
  if (s === 'recon') return true;
  if (acceptBlank && s === '') return true;
  return false;
}

function filterRecon(rows: Row[], periodCol: string | undefined, acceptBlank: boolean): Row[] {
  if (!periodCol) return rows;
  return rows.filter((r) => isReconPeriod(r[periodCol], acceptBlank));
}

interface AggValue {
  amount: number;
  count: number;
  rawCategories: Set<string>;
}

function makeKey(vch: string, cat: string, mode: 'vch' | 'vch_and_category'): string {
  if (mode === 'vch_and_category') return `${vch}||${cat}`;
  return vch;
}

function groupSum(
  rows: Row[],
  keyFn: (r: Row) => string,
  amtFn: (r: Row) => number,
  catFn: (r: Row) => string,
): Map<string, AggValue> {
  const out = new Map<string, AggValue>();
  for (const r of rows) {
    const k = keyFn(r);
    if (!k) continue;
    const cur = out.get(k);
    const amt = amtFn(r);
    const cat = catFn(r);
    if (cur) {
      cur.amount += amt;
      cur.count += 1;
      if (cat) cur.rawCategories.add(cat);
    } else {
      const set = new Set<string>();
      if (cat) set.add(cat);
      out.set(k, { amount: amt, count: 1, rawCategories: set });
    }
  }
  return out;
}

export function reconcile(
  marsRows: Row[],
  brandRows: Row[],
  marsCols: ColumnMapping,
  brandCols: ColumnMapping,
  options: ReconOptions = DEFAULT_OPTIONS,
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
  const brandAltRefCol = brandCols.alt_reference;
  const brandInvoiceRefCol = brandCols.invoice_ref;
  const brandAmtCol = brandCols.net_amount!;
  const brandCatCol = brandCols.category;
  const brandPeriodCol = brandCols.period;

  const mars = filterRecon(marsRows, marsPeriodCol, options.acceptBlankPeriod);
  const brand = filterRecon(brandRows, brandPeriodCol, options.acceptBlankPeriod);

  const marsCatRaw = (r: Row): string =>
    marsCatCol ? canonicalizeCategory(r[marsCatCol]) : '';
  const brandCatRaw = (r: Row): string =>
    brandCatCol ? canonicalizeCategory(r[brandCatCol]) : '';

  const marsVch = (r: Row) => normKey(r[marsVchCol]);
  const marsAmt = (r: Row) => toNumber(r[marsAmtCol]);
  const brandVch = (r: Row) => {
    const primary = normKey(r[brandRefCol]);
    if (primary) return primary;
    if (brandCorrectRefCol) {
      const corrected = normKey(r[brandCorrectRefCol]);
      if (corrected) return corrected;
    }
    if (brandAltRefCol) {
      const alt = normKey(r[brandAltRefCol]);
      if (alt) return alt;
    }
    return '';
  };
  const brandAmt = (r: Row) => toNumber(r[brandAmtCol]);

  const marsMatchKey = (r: Row) => {
    const v = marsVch(r);
    if (!v) return '';
    return makeKey(v, matchCatKey(marsCatRaw(r)), options.matchMode);
  };
  const brandMatchKey = (r: Row) => {
    const v = brandVch(r);
    if (!v) return '';
    return makeKey(v, matchCatKey(brandCatRaw(r)), options.matchMode);
  };

  const brandAgg = groupSum(brand, brandMatchKey, brandAmt, brandCatRaw);
  const marsAgg = groupSum(mars, marsMatchKey, marsAmt, marsCatRaw);

  const marsHeaders = inferHeaders(marsRows);
  const brandHeaders = inferHeaders(brandRows);

  const marsOutHeaders = [...marsHeaders, 'Amount_Brand', 'Difference', 'Remarks'];
  const brandOutHeaders = [...brandHeaders, 'Amount_Mars', 'Diff', 'Remarks'];

  const tolerance = Math.max(0, options.matchToleranceRupees);

  const marsOut: Row[] = mars.map((r) => {
    const key = marsMatchKey(r);
    const own = marsAmt(r);
    const brandHit = key ? brandAgg.get(key) : undefined;
    const amtBrand = brandHit?.amount ?? 0;
    const magnitudeDiff = Math.abs(Math.abs(own) - Math.abs(amtBrand));
    const displayDiff = own + amtBrand;
    let remarks: string;
    if (!brandHit) {
      remarks = 'Not Booked by Brand';
    } else if (magnitudeDiff <= tolerance) {
      remarks = 'Match';
    } else {
      remarks = 'Amount Mismatch';
    }
    return { ...r, Amount_Brand: amtBrand, Difference: displayDiff, Remarks: remarks };
  });

  const brandOut: Row[] = brand.map((r) => {
    const key = brandMatchKey(r);
    const own = brandAmt(r);
    const marsHit = key ? marsAgg.get(key) : undefined;
    const amtMars = marsHit?.amount ?? 0;
    const magnitudeDiff = Math.abs(Math.abs(own) - Math.abs(amtMars));
    const displayDiff = own + amtMars;
    let remarks: string;
    if (!marsHit) {
      remarks = 'Not Booked by Mars';
    } else if (magnitudeDiff <= tolerance) {
      remarks = 'Match';
    } else {
      remarks = 'Amount Mismatch';
    }
    return { ...r, Amount_Mars: amtMars, Diff: displayDiff, Remarks: remarks };
  });

  if (brandInvoiceRefCol) {
    applyReversalRemarks(brandOut, brandInvoiceRefCol, brandRefCol, brandAmtCol);
  }

  if (marsCatCol && brandCatCol) {
    matchReceiptToPaymentByAmount(
      marsOut,
      brandOut,
      marsCatCol,
      brandCatCol,
      marsAmtCol,
      brandAmtCol,
      tolerance,
    );
  }

  const summary = buildSummary(
    mars,
    brand,
    marsAmtCol,
    brandAmtCol,
    marsCatRaw,
    brandCatRaw,
  );

  const marsMatch = marsOut.filter((r) => r.Remarks === 'Match').length;
  const marsMismatch = marsOut.filter((r) => r.Remarks === 'Amount Mismatch').length;
  const marsNotBookedByBrand = marsOut.filter((r) => r.Remarks === 'Not Booked by Brand').length;
  const brandMatch = brandOut.filter((r) => r.Remarks === 'Match').length;
  const brandMismatch = brandOut.filter((r) => r.Remarks === 'Amount Mismatch').length;
  const brandNotBookedByMars = brandOut.filter((r) => r.Remarks === 'Not Booked by Mars').length;
  const brandReversal = brandOut.filter((r) => String(r.Remarks).startsWith('Reversal')).length;

  const diagnostics = buildDiagnostics(
    marsRows,
    brandRows,
    mars,
    brand,
    marsPeriodCol,
    brandPeriodCol,
    marsVch,
    brandVch,
    marsOut,
    brandOut,
    marsAmtCol,
    brandAmtCol,
  );

  return {
    summary,
    mars: { headers: marsOutHeaders, rows: marsOut },
    brand: { headers: brandOutHeaders, rows: brandOut },
    stats: {
      mars_recon_rows: marsOut.length,
      brand_recon_rows: brandOut.length,
      mars_match: marsMatch,
      mars_mismatch: marsMismatch,
      mars_not_booked_by_brand: marsNotBookedByBrand,
      brand_match: brandMatch,
      brand_mismatch: brandMismatch,
      brand_not_booked_by_mars: brandNotBookedByMars,
      brand_reversal: brandReversal,
    },
    diagnostics,
    options,
  };
}

function matchReceiptToPaymentByAmount(
  marsOut: Row[],
  brandOut: Row[],
  marsCatCol: string,
  brandCatCol: string,
  marsAmtCol: string,
  brandAmtCol: string,
  tolerance: number,
): void {
  const marsReceipts = marsOut
    .map((r, i) => ({ r, i }))
    .filter(
      ({ r }) =>
        normCatKey(canonicalizeCategory(r[marsCatCol])) === 'receipt' &&
        r.Remarks === 'Not Booked by Brand',
    );
  const brandPayments = brandOut
    .map((r, i) => ({ r, i }))
    .filter(
      ({ r }) =>
        normCatKey(canonicalizeCategory(r[brandCatCol])) === 'payment' &&
        r.Remarks === 'Not Booked by Mars',
    );

  if (marsReceipts.length === 0 || brandPayments.length === 0) return;

  const claimed = new Set<number>();
  for (const { r: marsRow } of marsReceipts) {
    const ownAbs = Math.abs(toNumber(marsRow[marsAmtCol]));
    let bestIdx = -1;
    let bestDiff = Infinity;
    for (let j = 0; j < brandPayments.length; j++) {
      if (claimed.has(j)) continue;
      const brandRow = brandPayments[j].r;
      const diff = Math.abs(Math.abs(toNumber(brandRow[brandAmtCol])) - ownAbs);
      if (diff <= tolerance && diff < bestDiff) {
        bestDiff = diff;
        bestIdx = j;
      }
    }
    if (bestIdx === -1) continue;
    claimed.add(bestIdx);
    const brandRow = brandPayments[bestIdx].r;
    const ownAmt = toNumber(marsRow[marsAmtCol]);
    const brandAmt = toNumber(brandRow[brandAmtCol]);
    marsRow.Amount_Brand = brandAmt;
    marsRow.Difference = ownAmt + brandAmt;
    marsRow.Remarks = 'Match';
    brandRow.Amount_Mars = ownAmt;
    brandRow.Diff = brandAmt + ownAmt;
    brandRow.Remarks = 'Match';
  }
}

function applyReversalRemarks(
  rows: Row[],
  invoiceRefCol: string,
  refCol: string,
  amtCol: string,
): void {
  const groups = new Map<string, number[]>();
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    const inv = normStr(r[invoiceRefCol]);
    const ref = normStr(r[refCol]);
    if (!inv || !ref) continue;
    const k = `${inv.toUpperCase()}||${ref.toUpperCase()}`;
    const list = groups.get(k);
    if (list) list.push(i);
    else groups.set(k, [i]);
  }

  for (const [, indices] of groups) {
    if (indices.length !== 2) continue;
    const a = rows[indices[0]];
    const b = rows[indices[1]];
    const amtA = toNumber(a[amtCol]);
    const amtB = toNumber(b[amtCol]);
    if (amtA === 0 || amtB === 0) continue;
    if (Math.sign(amtA) === Math.sign(amtB)) continue;
    const invRef = normStr(a[invoiceRefCol]);
    const tag = invRef ? `Reversal ${invRef}` : 'Reversal';
    a.Remarks = tag;
    b.Remarks = tag;
  }
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
  marsCatOf: (r: Row) => string,
  brandCatOf: (r: Row) => string,
): SummaryRow[] {
  const marsByCat = new Map<string, number>();
  for (const r of mars) {
    const c = marsCatOf(r) || 'Uncategorized';
    marsByCat.set(c, (marsByCat.get(c) ?? 0) + toNumber(r[marsAmtCol]));
  }
  const brandByCat = new Map<string, number>();
  for (const r of brand) {
    const c = brandCatOf(r) || 'Uncategorized';
    brandByCat.set(c, (brandByCat.get(c) ?? 0) + toNumber(r[brandAmtCol]));
  }

  const present = new Set<string>([...marsByCat.keys(), ...brandByCat.keys()]);
  const canonical = (CANONICAL_CATEGORY_ORDER as readonly string[]).filter((c) => present.has(c));
  const extras = [...present]
    .filter((c) => !(CANONICAL_CATEGORY_ORDER as readonly string[]).includes(c))
    .sort((a, b) => a.localeCompare(b));
  const final = [...canonical, ...extras];

  const rows: SummaryRow[] = final.map((cat) => {
    const m = marsByCat.get(cat) ?? 0;
    const b = brandByCat.get(cat) ?? 0;
    return { Particulars: cat, Amount_Mars: m, Amount_Brand: b, Difference: m + b };
  });

  const totalMars = rows.reduce((s, r) => s + r.Amount_Mars, 0);
  const totalBrand = rows.reduce((s, r) => s + r.Amount_Brand, 0);
  rows.push({
    Particulars: 'Grand Total',
    Amount_Mars: totalMars,
    Amount_Brand: totalBrand,
    Difference: totalMars + totalBrand,
  });

  return rows;
}

function buildDiagnostics(
  marsRowsAll: Row[],
  brandRowsAll: Row[],
  marsFiltered: Row[],
  brandFiltered: Row[],
  marsPeriodCol: string | undefined,
  brandPeriodCol: string | undefined,
  marsVchFn: (r: Row) => string,
  brandVchFn: (r: Row) => string,
  marsOut: Row[],
  brandOut: Row[],
  marsAmtCol: string,
  brandAmtCol: string,
): DiagnosticInfo {
  const periodCounts = (rows: Row[], col: string | undefined) => {
    if (!col) return [{ value: '(no period column mapped)', count: rows.length }];
    const counts = new Map<string, number>();
    for (const r of rows) {
      const v = normStr(r[col]) || '(blank)';
      counts.set(v, (counts.get(v) ?? 0) + 1);
    }
    return [...counts.entries()]
      .sort((a, b) => b[1] - a[1])
      .map(([value, count]) => ({ value, count }));
  };

  const marsVchSet = new Set<string>();
  for (const r of marsFiltered) {
    const v = marsVchFn(r);
    if (v) marsVchSet.add(v);
  }
  const brandVchSet = new Set<string>();
  for (const r of brandFiltered) {
    const v = brandVchFn(r);
    if (v) brandVchSet.add(v);
  }
  let intersection = 0;
  for (const v of marsVchSet) if (brandVchSet.has(v)) intersection++;

  const sampleUnmatchedMars = marsOut
    .filter((r) => r.Remarks === 'Not Booked by Brand')
    .slice(0, 8)
    .map((r) => marsVchFn(r))
    .filter(Boolean);

  const sampleUnmatchedBrand = brandOut
    .filter((r) => r.Remarks === 'Not Booked by Mars')
    .slice(0, 8)
    .map((r) => brandVchFn(r))
    .filter(Boolean);

  const brandColumnOverlap = computeBrandColumnOverlap(brandFiltered, marsVchSet);

  const amountSignSample: DiagnosticInfo['amount_sign_sample'] = [];
  for (const r of marsFiltered) {
    if (amountSignSample.length >= 3) break;
    const vch = marsVchFn(r);
    if (!vch) continue;
    const matchingBrand = brandFiltered.find((br) => brandVchFn(br) === vch);
    if (!matchingBrand) continue;
    const rawBrand = toNumber(matchingBrand[brandAmtCol]);
    amountSignSample.push({
      vch,
      mars_net: toNumber(r[marsAmtCol]),
      brand_amount_raw: rawBrand,
      brand_amount_used: rawBrand,
    });
  }

  return {
    mars_total_rows: marsRowsAll.length,
    brand_total_rows: brandRowsAll.length,
    mars_after_period_filter: marsFiltered.length,
    brand_after_period_filter: brandFiltered.length,
    mars_unique_vch: marsVchSet.size,
    brand_unique_vch: brandVchSet.size,
    vch_in_both: intersection,
    vch_only_in_mars: marsVchSet.size - intersection,
    vch_only_in_brand: brandVchSet.size - intersection,
    mars_period_values: periodCounts(marsRowsAll, marsPeriodCol),
    brand_period_values: periodCounts(brandRowsAll, brandPeriodCol),
    sample_unmatched_mars_vch: sampleUnmatchedMars,
    sample_unmatched_brand_vch: sampleUnmatchedBrand,
    brand_column_overlap: brandColumnOverlap,
    amount_sign_sample: amountSignSample,
  };
}

function computeBrandColumnOverlap(
  brandFiltered: Row[],
  marsVchSet: Set<string>,
): DiagnosticInfo['brand_column_overlap'] {
  if (brandFiltered.length === 0) return [];
  const headers = Object.keys(brandFiltered[0]);
  const out: { column: string; overlap: number; nonBlank: number }[] = [];
  for (const h of headers) {
    let overlap = 0;
    let nonBlank = 0;
    const seen = new Set<string>();
    for (const r of brandFiltered) {
      const v = normKey(r[h]);
      if (!v) continue;
      nonBlank++;
      if (seen.has(v)) continue;
      seen.add(v);
      if (marsVchSet.has(v)) overlap++;
    }
    if (overlap > 0) out.push({ column: h, overlap, nonBlank });
  }
  return out.sort((a, b) => b.overlap - a.overlap).slice(0, 6);
}
