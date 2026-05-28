import type { ColumnMapping } from './types';

export function normalizeHeader(h: unknown): string {
  return String(h ?? '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

export const MARS_FIELD_SYNONYMS: Record<string, string[]> = {
  vch_no: [
    'vch bill no',
    'vch no',
    'vch no.',
    'vch/bill no',
    'vch/bill no.',
    'voucher no',
    'voucher no.',
    'vch number',
    'voucher number',
    'bill no',
    'bill no.',
    'bill number',
    'invoice no',
    'invoice no.',
    'invoice number',
    'inv no',
    'inv no.',
    'document no',
    'document no.',
    'document number',
    'doc no',
    'doc no.',
    'trans no',
    'transaction no',
    'entry no',
    'vchno',
    'vch',
    'voucher',
  ],
  date: [
    'date',
    'vch date',
    'voucher date',
    'trans date',
    'transaction date',
    'entry date',
    'posting date',
    'narration date',
    'doc date',
  ],
  type: ['type', 'vch type', 'voucher type', 'txn type', 'transaction type'],
  account: ['account', 'particulars', 'ledger', 'ledger account', 'party', 'party name', 'name'],
  debit: ['debit rs', 'debit', 'dr', 'debit amount', 'debit amt', 'debit rs.'],
  credit: ['credit rs', 'credit', 'cr', 'credit amount', 'credit amt', 'credit rs.'],
  correct_ref: ['correct ref no', 'correct reference', 'corrected ref', 'manual ref', 'correct ref'],
  net_amount: [
    'net amount',
    'net amt',
    'net',
    'amount',
    'amt',
    'total amount',
    'txn amount',
    'transaction amount',
    'amount rs',
    'amount inr',
    'amount in inr',
    'amount rs.',
  ],
  period: ['period', 'recon period'],
  category: [
    'category',
    'cat',
    'type of voucher',
    'voucher category',
    'voucher type',
    'txn type',
    'transaction type',
    'doc type',
    'document type',
    'type of vch',
  ],
};

export const BRAND_FIELD_SYNONYMS: Record<string, string[]> = {
  reference: [
    'reference',
    'ref',
    'ref no',
    'ref no.',
    'reference no',
    'reference no.',
    'reference number',
    'ref number',
    'vch no',
    'voucher no',
    'bill ref',
    'invoice ref no',
    'transaction ref',
    'trans ref',
  ],
  correct_ref: ['correct ref no', 'correct reference', 'corrected ref', 'manual ref', 'correct ref'],
  alt_reference: [
    'document number',
    'doc number',
    'doc no',
    'doc no.',
    'document no',
    'document no.',
    'assignment',
    'alt ref',
    'reference 2',
  ],
  invoice_ref: ['invoice reference', 'invoice ref', 'inv ref', 'inv reference'],
  net_amount: [
    'net amount',
    'net amt',
    'amount in local currency',
    'amount local',
    'amount',
    'amt',
    'amount in lc',
    'lc amount',
    'amount lc',
    'local amount',
    'amount in doc cur',
    'doc currency amount',
    'amount in doc',
  ],
  date: [
    'posting date',
    'document date',
    'date',
    'pstng date',
    'doc date',
    'key date',
    'value date',
    'entry date',
    'invoice date',
    'trans date',
    'transaction date',
  ],
  period: ['period', 'recon period'],
  category: [
    'category',
    'cat',
    'type of voucher',
    'voucher category',
    'doc type',
    'document type',
    'txn type',
    'transaction type',
    'voucher type',
  ],
};

export const REQUIRED_MARS = ['vch_no', 'net_amount'];
export const REQUIRED_BRAND = ['reference', 'net_amount'];

export const HIDDEN_FIELDS = new Set<string>(['date', 'type', 'account', 'debit', 'credit']);

export const FIELD_LABELS: Record<string, string> = {
  vch_no: 'Vch / Bill No',
  reference: 'Reference (= Mars Vch No)',
  correct_ref: 'Correct Ref No (override)',
  alt_reference: 'Alt Reference (Document No / Assignment)',
  invoice_ref: 'Invoice Reference (for Reversal detection)',
  net_amount: 'Net Amount',
  period: 'Period',
  category: 'Category',
  date: 'Date',
  type: 'Type',
  account: 'Account',
  debit: 'Debit',
  credit: 'Credit',
};

export const FIELD_HINTS: Record<string, string> = {
  vch_no: 'The voucher number that uniquely identifies each Mars entry.',
  reference: 'The Brand column that holds the Mars Vch No (usually called "Reference").',
  correct_ref: 'Optional — used when Reference is wrong/blank and a manual override exists.',
  alt_reference: 'Optional — fallback used only when Reference and Correct Ref No are blank.',
  invoice_ref: 'Optional — when two Brand rows share the same Invoice Reference + same Reference + opposite signs, both are tagged as Reversal instead of Match.',
  net_amount: 'The signed rupee amount used for amount comparison. On Brand SAP exports this is often "Amount in local currency".',
  period: 'Used to filter rows. Only rows with Period = "Recon" enter the reconciliation.',
  category: 'Used in the match key when "Vch + Category" mode is selected, and in the Recon Summary breakdown.',
};

function score(headerNorm: string, candidate: string): number {
  if (headerNorm === candidate) return 1000;
  const hTokens = new Set(headerNorm.split(' ').filter(Boolean));
  const cTokens = new Set(candidate.split(' ').filter(Boolean));
  if (cTokens.size === 0 || hTokens.size === 0) return 0;
  let isSubset = true;
  for (const t of cTokens) {
    if (!hTokens.has(t)) {
      isSubset = false;
      break;
    }
  }
  if (isSubset) return 500 + candidate.length;
  let overlap = 0;
  for (const t of cTokens) if (hTokens.has(t)) overlap++;
  if (overlap === 0) return 0;
  return overlap * 10 + candidate.length * 0.1;
}

export function detectColumns(
  headers: string[],
  synonyms: Record<string, string[]>,
): ColumnMapping {
  const norm: Record<string, string> = {};
  for (const h of headers) norm[h] = normalizeHeader(h);
  const used = new Set<string>();
  const result: ColumnMapping = {};

  for (const [field, candidates] of Object.entries(synonyms)) {
    let bestScore = 0;
    let bestHeader: string | undefined;
    for (const header of headers) {
      if (used.has(header)) continue;
      const n = norm[header];
      if (
        field === 'net_amount' &&
        n.split(' ').some((t) => ['debit', 'credit', 'dr', 'cr'].includes(t))
      ) {
        continue;
      }
      for (const cand of candidates) {
        const s = score(n, cand);
        if (s > bestScore) {
          bestScore = s;
          bestHeader = header;
        }
      }
    }
    result[field] = bestHeader;
    if (bestHeader) used.add(bestHeader);
  }

  return result;
}

export function missingRequired(detected: ColumnMapping, required: string[]): string[] {
  return required.filter((f) => !detected[f]);
}
