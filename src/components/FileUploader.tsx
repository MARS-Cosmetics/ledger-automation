import { useRef } from 'react';

interface Props {
  label: string;
  file: File | null;
  onChange: (f: File | null) => void;
}

export function FileUploader({ label, file, onChange }: Props) {
  const inputRef = useRef<HTMLInputElement>(null);

  return (
    <div className="card">
      <span className="label">{label}</span>
      <div className="mt-2 flex items-center gap-3">
        <button type="button" className="btn-secondary" onClick={() => inputRef.current?.click()}>
          {file ? 'Replace file' : 'Choose Excel file'}
        </button>
        <span className="truncate text-sm text-slate-700">
          {file ? file.name : 'No file selected'}
        </span>
        {file && (
          <button
            type="button"
            className="ml-auto text-xs text-slate-500 underline-offset-2 hover:text-slate-700 hover:underline"
            onClick={() => {
              onChange(null);
              if (inputRef.current) inputRef.current.value = '';
            }}
          >
            Remove
          </button>
        )}
      </div>
      <input
        ref={inputRef}
        type="file"
        accept=".xlsx,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel"
        className="hidden"
        onChange={(e) => onChange(e.target.files?.[0] ?? null)}
      />
    </div>
  );
}
