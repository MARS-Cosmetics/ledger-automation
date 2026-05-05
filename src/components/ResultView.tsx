import { useMemo, useState } from 'react';
import type { ReconResult, Row } from '../lib/types';

interface Props {
  result: ReconResult;
  onDownload: () => void;
  downloading: boolean;
}

type RemarkFilter = 'all' | 'Match' | 'Amount Mismatch' | 'Not Booked by Brand' | 'Not Booked by Mars';

export function ResultView({ result, onDownload, downloading }: Props) {
  const { stats, summary, mars, brand } = result;
  const [marsFilter, setMarsFilter] = useState<RemarkFilter>('all');
  const [brandFilter, setBrandFilter] = useState<RemarkFilter>('all');

  const filteredMars = useMemo(
    () => filterRows(mars.rows, marsFilter),
    [mars.rows, marsFilter],
  );
  const filteredBrand = useMemo(
    () => filterRows(brand.rows, brandFilter),
    [brand.rows, brandFilter],
  );

  return (
    <div className="space-y-6">
      <div className="grid grid-cols-2 gap-3 sm:grid-cols-3 lg:grid-cols-6">
        <Metric label="Mars Recon rows" value={stats.mars_recon_rows} />
        <Metric label="Brand Recon rows" value={stats.brand_recon_rows} />
        <Metric label="Match (Mars)" value={stats.mars_match} tone="good" />
        <Metric label="Mismatch (Mars)" value={stats.mars_mismatch} tone="warn" />
        <Metric label="Not Booked by Brand" value={stats.mars_not_booked_by_brand} tone="warn" />
        <Metric label="Not Booked by Mars" value={stats.brand_not_booked_by_mars} tone="warn" />
      </div>

      <Diagnostics result={result} />

      <div className="card">
        <div className="flex items-center justify-between">
          <h3 className="text-base font-semibold">Recon Summary – Category Wise</h3>
          <button type="button" className="btn-primary" disabled={downloading} onClick={onDownload}>
            {downloading ? 'Building…' : 'Download workbook'}
          </button>
        </div>
        <div className="mt-3 overflow-x-auto">
          <table className="min-w-full text-sm">
            <thead className="bg-brand-100 text-left text-xs uppercase tracking-wide text-slate-600">
              <tr>
                <th className="px-3 py-2">Particulars</th>
                <th className="px-3 py-2 text-right">Amount_Mars</th>
                <th className="px-3 py-2 text-right">Amount_Brand</th>
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

      <details className="card" open>
        <summary className="cursor-pointer text-sm font-semibold">
          Mars Cosmetics — Recon entries ({filteredMars.length} of {mars.rows.length})
        </summary>
        <RemarksFilter value={marsFilter} onChange={setMarsFilter} side="mars" />
        <div className="mt-3 overflow-x-auto">
          <PreviewTable headers={mars.headers} rows={filteredMars.slice(0, 200)} />
        </div>
        {filteredMars.length > 200 && (
          <p className="mt-2 text-xs text-slate-500">Showing first 200 rows. Full data is in the downloadable workbook.</p>
        )}
      </details>

      <details className="card">
        <summary className="cursor-pointer text-sm font-semibold">
          Brand — Recon entries ({filteredBrand.length} of {brand.rows.length})
        </summary>
        <RemarksFilter value={brandFilter} onChange={setBrandFilter} side="brand" />
        <div className="mt-3 overflow-x-auto">
          <PreviewTable headers={brand.headers} rows={filteredBrand.slice(0, 200)} />
        </div>
        {filteredBrand.length > 200 && (
          <p className="mt-2 text-xs text-slate-500">Showing first 200 rows. Full data is in the downloadable workbook.</p>
        )}
      </details>
    </div>
  );
}

function filterRows(rows: Row[], f: RemarkFilter): Row[] {
  if (f === 'all') return rows;
  return rows.filter((r) => String(r.Remarks ?? '') === f);
}

function Diagnostics({ result }: { result: ReconResult }) {
  const d = result.diagnostics;
  const o = result.options;
  const warn = d.vch_in_both === 0 && d.mars_after_period_filter > 0 && d.brand_after_period_filter > 0;
  return (
    <details className={`card ${warn ? 'ring-2 ring-amber-400' : ''}`} open={warn}>
      <summary className="cursor-pointer text-sm font-semibold">
        Diagnostics {warn && <span className="ml-2 text-amber-700">⚠ no Vch overlaps — open me</span>}
      </summary>
      <div className="mt-3 space-y-3 text-sm">
        <div className="rounded-md bg-slate-50 p-3 text-xs text-slate-700">
          Settings used: match key = <b>{o.matchMode === 'vch' ? 'Vch No only' : 'Vch No + Category'}</b>,
          tolerance = <b>₹{o.matchToleranceRupees}</b>, period filter ={' '}
          <b>{o.acceptBlankPeriod ? 'lenient (Recon + blank)' : 'strict Recon only'}</b>
        </div>

        <div className="grid grid-cols-2 gap-3 sm:grid-cols-4">
          <DiagBox label="Mars rows in file" value={d.mars_total_rows} />
          <DiagBox label="Mars after Period filter" value={d.mars_after_period_filter} />
          <DiagBox label="Brand rows in file" value={d.brand_total_rows} />
          <DiagBox label="Brand after Period filter" value={d.brand_after_period_filter} />
          <DiagBox label="Mars unique Vch No" value={d.mars_unique_vch} />
          <DiagBox label="Brand unique Ref" value={d.brand_unique_vch} />
          <DiagBox
            label="Vch in BOTH ledgers"
            value={d.vch_in_both}
            tone={d.vch_in_both === 0 ? 'bad' : 'good'}
          />
          <DiagBox label="Vch only in Mars" value={d.vch_only_in_mars} />
        </div>

        <div className="grid grid-cols-1 gap-3 sm:grid-cols-2">
          <PeriodList title="Mars Period column values" entries={d.mars_period_values} />
          <PeriodList title="Brand Period column values" entries={d.brand_period_values} />
        </div>

        {d.brand_column_overlap.length > 0 && (
          <div className="rounded-md border border-slate-200 p-2">
            <div className="text-xs font-semibold text-slate-600">
              Brand columns ranked by overlap with Mars Vch No
            </div>
            <p className="mt-1 text-xs text-slate-500">
              If your selected Reference column is not the top one here, change the Reference mapping above.
            </p>
            <ul className="mt-1 space-y-0.5 text-xs">
              {d.brand_column_overlap.map((c) => (
                <li key={c.column} className="flex justify-between">
                  <code className="text-slate-700">{c.column}</code>
                  <span className="tabular-nums text-slate-500">
                    {c.overlap.toLocaleString('en-IN')} match / {c.nonBlank.toLocaleString('en-IN')} non-blank
                  </span>
                </li>
              ))}
            </ul>
          </div>
        )}

        {d.amount_sign_sample.length > 0 && (
          <div className="rounded-md border border-slate-200 p-2">
            <div className="text-xs font-semibold text-slate-600">
              Amount sign check on sample matched Vch
            </div>
            <p className="mt-1 text-xs text-slate-500">
              If <b>Mars Net</b> and <b>Brand used</b> have opposite signs, toggle "Negate Brand amount sign" above.
            </p>
            <table className="mt-2 w-full text-xs">
              <thead>
                <tr className="text-left text-slate-500">
                  <th className="px-1 py-0.5">Vch</th>
                  <th className="px-1 py-0.5 text-right">Mars Net</th>
                  <th className="px-1 py-0.5 text-right">Brand raw</th>
                  <th className="px-1 py-0.5 text-right">Brand used</th>
                </tr>
              </thead>
              <tbody>
                {d.amount_sign_sample.map((s) => (
                  <tr key={s.vch}>
                    <td className="px-1 py-0.5"><code>{s.vch}</code></td>
                    <td className="px-1 py-0.5 text-right tabular-nums">{s.mars_net.toLocaleString('en-IN')}</td>
                    <td className="px-1 py-0.5 text-right tabular-nums">{s.brand_amount_raw.toLocaleString('en-IN')}</td>
                    <td className="px-1 py-0.5 text-right tabular-nums">{s.brand_amount_used.toLocaleString('en-IN')}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}

        {d.sample_unmatched_mars_vch.length > 0 && (
          <div>
            <div className="text-xs font-semibold text-slate-600">Sample Mars Vch with no Brand match</div>
            <div className="mt-1 flex flex-wrap gap-1">
              {d.sample_unmatched_mars_vch.map((v) => (
                <code key={v} className="rounded bg-slate-100 px-2 py-0.5 text-xs">
                  {v}
                </code>
              ))}
            </div>
            <p className="mt-1 text-xs text-slate-500">
              Open your Brand ledger and search the Reference / Correct Ref No column for any of these Vch
              numbers. If they exist there but didn&apos;t match, paste one back to me with both ledger rows.
            </p>
          </div>
        )}
        {d.sample_unmatched_brand_vch.length > 0 && (
          <div>
            <div className="text-xs font-semibold text-slate-600">Sample Brand Ref with no Mars match</div>
            <div className="mt-1 flex flex-wrap gap-1">
              {d.sample_unmatched_brand_vch.map((v) => (
                <code key={v} className="rounded bg-slate-100 px-2 py-0.5 text-xs">
                  {v}
                </code>
              ))}
            </div>
          </div>
        )}
      </div>
    </details>
  );
}

function DiagBox({
  label,
  value,
  tone,
}: {
  label: string;
  value: number;
  tone?: 'good' | 'bad';
}) {
  const ring =
    tone === 'good' && value > 0
      ? 'ring-1 ring-emerald-300'
      : tone === 'bad' && value === 0
        ? 'ring-1 ring-rose-400'
        : '';
  return (
    <div className={`rounded-md border border-slate-200 bg-white p-2 ${ring}`}>
      <div className="text-[11px] uppercase tracking-wide text-slate-500">{label}</div>
      <div className="text-lg font-semibold tabular-nums">{value.toLocaleString('en-IN')}</div>
    </div>
  );
}

function PeriodList({
  title,
  entries,
}: {
  title: string;
  entries: { value: string; count: number }[];
}) {
  return (
    <div className="rounded-md border border-slate-200 p-2">
      <div className="text-xs font-semibold text-slate-600">{title}</div>
      <ul className="mt-1 space-y-0.5 text-xs">
        {entries.slice(0, 8).map((e) => (
          <li key={e.value} className="flex justify-between">
            <code className="text-slate-700">{e.value}</code>
            <span className="tabular-nums text-slate-500">{e.count.toLocaleString('en-IN')}</span>
          </li>
        ))}
        {entries.length > 8 && (
          <li className="text-slate-400">…{entries.length - 8} more values</li>
        )}
      </ul>
    </div>
  );
}

function RemarksFilter({
  value,
  onChange,
  side,
}: {
  value: RemarkFilter;
  onChange: (v: RemarkFilter) => void;
  side: 'mars' | 'brand';
}) {
  const opts: RemarkFilter[] =
    side === 'mars'
      ? ['all', 'Match', 'Amount Mismatch', 'Not Booked by Brand']
      : ['all', 'Match', 'Amount Mismatch', 'Not Booked by Mars'];
  return (
    <div className="mt-3 flex flex-wrap items-center gap-2">
      <span className="text-xs uppercase tracking-wide text-slate-500">Filter:</span>
      {opts.map((o) => (
        <button
          key={o}
          type="button"
          onClick={() => onChange(o)}
          className={`rounded-full border px-3 py-1 text-xs ${
            value === o
              ? 'border-brand-600 bg-brand-600 text-white'
              : 'border-slate-200 bg-white text-slate-600 hover:bg-slate-50'
          }`}
        >
          {o === 'all' ? 'All' : o}
        </button>
      ))}
    </div>
  );
}

function Metric({
  label,
  value,
  tone,
}: {
  label: string;
  value: number;
  tone?: 'good' | 'warn';
}) {
  const ring =
    tone === 'warn' && value > 0
      ? 'ring-1 ring-amber-300'
      : tone === 'good' && value > 0
        ? 'ring-1 ring-emerald-300'
        : '';
  return (
    <div className={`card ${ring}`}>
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
  if (rem === 'Amount Mismatch') return 'bg-yellow-50';
  if (rem === 'Match') return 'bg-emerald-50';
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
  if (n === 0) return '-';
  const abs = Math.abs(n).toLocaleString('en-IN', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
  return n < 0 ? `(${abs})` : abs;
}
