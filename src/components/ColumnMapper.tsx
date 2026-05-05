import { FIELD_HINTS, FIELD_LABELS } from '../lib/columnDetect';
import type { ColumnMapping } from '../lib/types';

interface Props {
  title: string;
  headers: string[];
  fields: string[];
  required: string[];
  mapping: ColumnMapping;
  onChange: (m: ColumnMapping) => void;
}

const NONE = '__none__';

export function ColumnMapper({ title, headers, fields, required, mapping, onChange }: Props) {
  const set = (field: string, value: string) => {
    const next = { ...mapping };
    if (value === NONE) delete next[field];
    else next[field] = value;
    onChange(next);
  };

  return (
    <div>
      <h3 className="text-sm font-semibold text-slate-700">{title}</h3>
      <div className="mt-3 grid grid-cols-1 gap-3 sm:grid-cols-2 lg:grid-cols-3">
        {fields.map((field) => {
          const isRequired = required.includes(field);
          const value = mapping[field] ?? NONE;
          const label = FIELD_LABELS[field] ?? field.replace(/_/g, ' ');
          const hint = FIELD_HINTS[field];
          const unmappedRequired = isRequired && value === NONE;
          return (
            <div key={field} className={unmappedRequired ? 'rounded-md ring-1 ring-rose-300 p-2' : ''}>
              <label className="label">
                {label}
                {isRequired && <span className="ml-1 text-red-500">*</span>}
              </label>
              <select className="field" value={value} onChange={(e) => set(field, e.target.value)}>
                <option value={NONE}>(none)</option>
                {headers.map((h) => (
                  <option key={h} value={h}>
                    {h}
                  </option>
                ))}
              </select>
              {hint && <p className="mt-1 text-[11px] leading-snug text-slate-500">{hint}</p>}
            </div>
          );
        })}
      </div>
    </div>
  );
}
