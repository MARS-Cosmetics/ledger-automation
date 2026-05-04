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
  mars_matched_rows: number;
  mars_unmatched_rows: number;
  brand_unmatched_rows: number;
}

export interface ReconResult {
  summary: SummaryRow[];
  mars: { headers: string[]; rows: Row[] };
  brand: { headers: string[]; rows: Row[] };
  stats: ReconStats;
}
