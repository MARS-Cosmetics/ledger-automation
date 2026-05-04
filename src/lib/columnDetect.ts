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
    'voucher no',
    'vch number',
    'voucher number',
    'bill no',
    'vchno',
    'vch',
    'voucher',
  ],
  date: ['date', 'vch date', 'voucher date'],
  type: ['type', 'vch type', 'voucher type'],
  account: ['account', 'particulars', 'ledger', 'ledger account', 'party'],
  debit: ['debit rs', 'debit', 'dr', 'debit amount', 'debit amt'],
  credit: ['credit rs', 'credit', 'cr', 'credit amount', 'credit amt'],
  net_amount: ['net amount', 'net amt', 'net', 'amount', 'amt'],
  period: ['period'],
  category: ['category', 'cat', 'type of voucher', 'voucher category'],
};

export const BRAND_FIELD_SYNONYMS: Record<string, string[]> = {
  reference: ['reference', 'ref', 'ref no', 'reference no', 'vch no', 'voucher no', 'ref number'],
  correct_ref: ['correct ref no', 'correct reference', 'corrected ref', 'manual ref'],
  net_amount: ['net amount', 'net amt', 'net', 'amount', 'amt'],
  period: ['period'],
  category: ['category', 'cat', 'type of voucher', 'voucher category'],
};

export const REQUIRED_MARS = ['vch_no', 'net_amount'];
export const REQUIRED_BRAND = ['reference', 'net_amount'];

export const HIDDEN_FIELDS = new Set(['period', 'category']);

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
