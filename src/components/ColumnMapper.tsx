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
      <div className="mt-3 grid grid-cols-1 gap-3 sm:grid-cols-2 lg:grid-cols-4">
        {fields.map((field) => {
          const isRequired = required.includes(field);
          const value = mapping[field] ?? NONE;
          return (
            <div key={field}>
              <label className="label">
                {field.replace(/_/g, ' ')}
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
            </div>
          );
        })}
      </div>
    </div>
  );
}
