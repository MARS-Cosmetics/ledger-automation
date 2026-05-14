export type Cell = string | number | boolean | Date | null | undefined;
export type Row = Record<string, Cell>;

export interface Sheet {
  name: string;
  headers: string[];
  rows: Row[];
}

export interface Workbook {
  sheets: Sheet[];
}

export interface ColumnMapping {
  [logicalField: string]: string | undefined;
}

export interface SummaryRow {
  Particulars: string;
  Amount_Mars: number;
  Amount_Brand: number;
  Difference: number;
}

export interface ReconStats {
  mars_recon_rows: number;
  brand_recon_rows: number;
  mars_match: number;
  mars_mismatch: number;
  mars_not_booked_by_brand: number;
  brand_match: number;
  brand_mismatch: number;
  brand_not_booked_by_mars: number;
  brand_reversal: number;
}

export interface DiagnosticInfo {
  mars_total_rows: number;
  brand_total_rows: number;
  mars_after_period_filter: number;
  brand_after_period_filter: number;
  mars_unique_vch: number;
  brand_unique_vch: number;
  vch_in_both: number;
  vch_only_in_mars: number;
  vch_only_in_brand: number;
  mars_period_values: { value: string; count: number }[];
  brand_period_values: { value: string; count: number }[];
  sample_unmatched_mars_vch: string[];
  sample_unmatched_brand_vch: string[];
  brand_column_overlap: { column: string; overlap: number; nonBlank: number }[];
  amount_sign_sample: {
    vch: string;
    mars_net: number;
    brand_amount_raw: number;
    brand_amount_used: number;
  }[];
}

export interface ReconOptions {
  matchMode: 'vch' | 'vch_and_category';
  matchToleranceRupees: number;
  acceptBlankPeriod: boolean;
}

export interface AnnexureRow {
  Particular: string;
  Category: string;
  SubCategory: string;
  Date: string;
  ReferenceNumber: string;
  Amount_Mars: number;
  Amount_Brand: number;
  Difference: number;
}

export interface OpenPointsRow {
  Particular: string;
  Category: string;
  SubCategory: string;
  Count: number;
  Amount_Mars: number;
  Amount_Brand: number;
  Difference: number;
  AnnexureLink: string;
  ActionOn: string;
}

export interface CategoryAnnexure {
  category: string;
  rows: AnnexureRow[];
}

export interface ReconResult {
  summary: SummaryRow[];
  mars: { headers: string[]; rows: Row[] };
  brand: { headers: string[]; rows: Row[] };
  annexures: CategoryAnnexure[];
  openPoints: OpenPointsRow[];
  stats: ReconStats;
  diagnostics: DiagnosticInfo;
  options: ReconOptions;
}
