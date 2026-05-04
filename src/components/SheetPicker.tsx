import type { Sheet } from '../lib/types';

interface Props {
  label: string;
  sheets: Sheet[];
  selected: string;
  onChange: (name: string) => void;
}

export function SheetPicker({ label, sheets, selected, onChange }: Props) {
  if (sheets.length <= 1) {
    const s = sheets[0];
    return (
      <div>
        <span className="label">{label} — sheet</span>
        <p className="mt-1 text-sm text-slate-700">
          {s ? `${s.name} (${s.rows.length} rows · ${s.headers.length} columns)` : 'No sheets'}
        </p>
      </div>
    );
  }
  const current = sheets.find((s) => s.name === selected) ?? sheets[0];
  return (
    <div>
      <label className="label">{label} — sheet</label>
      <select className="field" value={selected} onChange={(e) => onChange(e.target.value)}>
        {sheets.map((s) => (
          <option key={s.name} value={s.name}>
            {s.name} ({s.rows.length} rows)
          </option>
        ))}
      </select>
      <p className="mt-1 text-xs text-slate-500">
        Columns: {current.headers.join(', ')}
      </p>
    </div>
  );
}
