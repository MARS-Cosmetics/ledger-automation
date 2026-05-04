import type { ReconResult, Row } from '../lib/types';

interface Props {
  result: ReconResult;
  onDownload: () => void;
  downloading: boolean;
}

export function ResultView({ result, onDownload, downloading }: Props) {
  const { stats, summary, mars, brand } = result;
  return (
    <div className="space-y-6">
      <div className="grid grid-cols-2 gap-4 sm:grid-cols-4">
        <Metric label="Mars (Recon) rows" value={stats.mars_recon_rows} />
        <Metric label="Brand (Recon) rows" value={stats.brand_recon_rows} />
        <Metric label="Mars not booked by Brand" value={stats.mars_unmatched_rows} highlight />
        <Metric label="Brand not booked by Mars" value={stats.brand_unmatched_rows} highlight />
      </div>

      <div className="card">
        <div className="flex items-center justify-between">
          <h3 className="text-base font-semibold">Recon Summary</h3>
          <button type="button" className="btn-primary" disabled={downloading} onClick={onDownload}>
            {downloading ? 'Building…' : 'Download workbook'}
          </button>
        </div>
        <div className="mt-3 overflow-x-auto">
          <table className="min-w-full text-sm">
            <thead className="bg-brand-100 text-left text-xs uppercase tracking-wide text-slate-600">
              <tr>
                <th className="px-3 py-2">Particulars</th>
                <th className="px-3 py-2 text-right">Amount Mars</th>
                <th className="px-3 py-2 text-right">Amount Brand</th>
                <th className="px-3 py-2 text-right">Difference</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {summary.map((r) => {
                const isTotal = r.Particulars === 'Grand Total';
                return (
                  <tr key={r.Particulars} className={isTotal ? 'bg-amber-50 font-semibold' : ''}>
                    <td className="px-3 py-2">{r.Particulars}</td>
                    <td className="px-3 py-2 text-right tabular-nums">{fmt(r.Amount_Mars)}</td>
                    <td className="px-3 py-2 text-right tabular-nums">{fmt(r.Amount_Brand)}</td>
                    <td className="px-3 py-2 text-right tabular-nums">{fmt(r.Difference)}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>

      <details className="card">
        <summary className="cursor-pointer text-sm font-semibold">
          Mars sheet preview ({mars.rows.length} rows)
        </summary>
        <div className="mt-3 overflow-x-auto">
          <PreviewTable headers={mars.headers} rows={mars.rows.slice(0, 50)} />
        </div>
      </details>

      <details className="card">
        <summary className="cursor-pointer text-sm font-semibold">
          Brand sheet preview ({brand.rows.length} rows)
        </summary>
        <div className="mt-3 overflow-x-auto">
          <PreviewTable headers={brand.headers} rows={brand.rows.slice(0, 50)} />
        </div>
      </details>
    </div>
  );
}

function Metric({ label, value, highlight }: { label: string; value: number; highlight?: boolean }) {
  return (
    <div className={`card ${highlight && value > 0 ? 'ring-1 ring-amber-300' : ''}`}>
      <div className="text-xs uppercase tracking-wide text-slate-500">{label}</div>
      <div className="mt-1 text-2xl font-semibold tabular-nums">{value.toLocaleString('en-IN')}</div>
    </div>
  );
}

function PreviewTable({ headers, rows }: { headers: string[]; rows: Row[] }) {
  return (
    <table className="min-w-full text-xs">
      <thead className="bg-slate-100 text-left uppercase tracking-wide text-slate-600">
        <tr>
          {headers.map((h) => (
            <th key={h} className="whitespace-nowrap px-2 py-1.5">
              {h}
            </th>
          ))}
        </tr>
      </thead>
      <tbody className="divide-y divide-slate-100">
        {rows.map((r, i) => (
          <tr key={i} className={rowClass(r)}>
            {headers.map((h) => (
              <td key={h} className="whitespace-nowrap px-2 py-1.5 tabular-nums">
                {renderCell(r[h])}
              </td>
            ))}
          </tr>
        ))}
      </tbody>
    </table>
  );
}

function rowClass(r: Row): string {
  const rem = String(r.Remarks ?? '');
  if (rem.startsWith('Not Booked')) return 'bg-orange-50';
  if (rem.startsWith('Amount Mismatch')) return 'bg-yellow-50';
  return '';
}

function renderCell(v: unknown): string {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  if (typeof v === 'number') return fmt(v);
  return String(v);
}

function fmt(n: number): string {
  if (!Number.isFinite(n)) return '';
  return n.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
