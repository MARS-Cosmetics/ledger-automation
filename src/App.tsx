import { useEffect, useMemo, useState } from 'react';
import { ColumnMapper } from './components/ColumnMapper';
import { FileUploader } from './components/FileUploader';
import { ResultView } from './components/ResultView';
import { SheetPicker } from './components/SheetPicker';
import {
  BRAND_FIELD_SYNONYMS,
  HIDDEN_FIELDS,
  MARS_FIELD_SYNONYMS,
  REQUIRED_BRAND,
  REQUIRED_MARS,
  detectColumns,
  missingRequired,
} from './lib/columnDetect';
import { readWorkbook } from './lib/excelRead';
import { buildOutputFilename, buildReconWorkbook } from './lib/excelWrite';
import { reconcile } from './lib/reconcile';
import type { ColumnMapping, ReconResult, Workbook } from './lib/types';

interface SideState {
  workbook: Workbook | null;
  sheetName: string;
  mapping: ColumnMapping;
  error: string | null;
}

const blankSide: SideState = { workbook: null, sheetName: '', mapping: {}, error: null };

export default function App() {
  const [marsFile, setMarsFile] = useState<File | null>(null);
  const [brandFile, setBrandFile] = useState<File | null>(null);
  const [mars, setMars] = useState<SideState>(blankSide);
  const [brand, setBrand] = useState<SideState>(blankSide);
  const [reconResult, setReconResult] = useState<ReconResult | null>(null);
  const [reconError, setReconError] = useState<string | null>(null);
  const [running, setRunning] = useState(false);
  const [downloading, setDownloading] = useState(false);

  useEffect(() => {
    setReconResult(null);
    setReconError(null);
    if (!marsFile) {
      setMars(blankSide);
      return;
    }
    let cancelled = false;
    readWorkbook(marsFile)
      .then((wb) => {
        if (cancelled) return;
        const first = wb.sheets[0]?.name ?? '';
        const headers = wb.sheets[0]?.headers ?? [];
        setMars({
          workbook: wb,
          sheetName: first,
          mapping: detectColumns(headers, MARS_FIELD_SYNONYMS),
          error: null,
        });
      })
      .catch((e: Error) => !cancelled && setMars({ ...blankSide, error: e.message }));
    return () => {
      cancelled = true;
    };
  }, [marsFile]);

  useEffect(() => {
    setReconResult(null);
    setReconError(null);
    if (!brandFile) {
      setBrand(blankSide);
      return;
    }
    let cancelled = false;
    readWorkbook(brandFile)
      .then((wb) => {
        if (cancelled) return;
        const first = wb.sheets[0]?.name ?? '';
        const headers = wb.sheets[0]?.headers ?? [];
        setBrand({
          workbook: wb,
          sheetName: first,
          mapping: detectColumns(headers, BRAND_FIELD_SYNONYMS),
          error: null,
        });
      })
      .catch((e: Error) => !cancelled && setBrand({ ...blankSide, error: e.message }));
    return () => {
      cancelled = true;
    };
  }, [brandFile]);

  const marsSheet = mars.workbook?.sheets.find((s) => s.name === mars.sheetName) ?? null;
  const brandSheet = brand.workbook?.sheets.find((s) => s.name === brand.sheetName) ?? null;

  const onMarsSheetChange = (name: string) => {
    setReconResult(null);
    const sheet = mars.workbook?.sheets.find((s) => s.name === name);
    setMars((s) => ({
      ...s,
      sheetName: name,
      mapping: sheet ? detectColumns(sheet.headers, MARS_FIELD_SYNONYMS) : {},
    }));
  };
  const onBrandSheetChange = (name: string) => {
    setReconResult(null);
    const sheet = brand.workbook?.sheets.find((s) => s.name === name);
    setBrand((s) => ({
      ...s,
      sheetName: name,
      mapping: sheet ? detectColumns(sheet.headers, BRAND_FIELD_SYNONYMS) : {},
    }));
  };

  const missMars = missingRequired(mars.mapping, REQUIRED_MARS);
  const missBrand = missingRequired(brand.mapping, REQUIRED_BRAND);

  const canRun = useMemo(
    () => !!marsSheet && !!brandSheet && missMars.length === 0 && missBrand.length === 0,
    [marsSheet, brandSheet, missMars, missBrand],
  );

  const runRecon = async () => {
    if (!marsSheet || !brandSheet) return;
    setRunning(true);
    setReconError(null);
    setReconResult(null);
    try {
      const res = reconcile(marsSheet.rows, brandSheet.rows, mars.mapping, brand.mapping);
      setReconResult(res);
    } catch (e) {
      setReconError(e instanceof Error ? e.message : String(e));
    } finally {
      setRunning(false);
    }
  };

  const downloadXlsx = async () => {
    if (!reconResult) return;
    setDownloading(true);
    try {
      const buf = await buildReconWorkbook(reconResult);
      const blob = new Blob([buf], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = buildOutputFilename();
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    } finally {
      setDownloading(false);
    }
  };

  return (
    <div className="mx-auto max-w-7xl px-4 py-8 sm:px-6 lg:px-8">
      <header className="mb-6">
        <h1 className="text-2xl font-bold text-slate-900">Mars vs Brand — Ledger Reconciliation</h1>
        <p className="mt-1 text-sm text-slate-600">
          Upload both ledgers below. All processing runs in your browser — files never leave your laptop.
        </p>
      </header>

      <section className="grid grid-cols-1 gap-4 sm:grid-cols-2">
        <FileUploader label="Mars ledger" file={marsFile} onChange={setMarsFile} />
        <FileUploader label="Brand ledger" file={brandFile} onChange={setBrandFile} />
      </section>

      {(mars.error || brand.error) && (
        <div className="mt-4 rounded-md border border-red-200 bg-red-50 p-3 text-sm text-red-800">
          {mars.error && <div>Mars file: {mars.error}</div>}
          {brand.error && <div>Brand file: {brand.error}</div>}
        </div>
      )}

      {mars.workbook && brand.workbook && (
        <>
          <section className="mt-8 card space-y-6">
            <h2 className="text-lg font-semibold">1. Pick the right sheet</h2>
            <div className="grid grid-cols-1 gap-4 sm:grid-cols-2">
              <SheetPicker
                label="Mars"
                sheets={mars.workbook.sheets}
                selected={mars.sheetName}
                onChange={onMarsSheetChange}
              />
              <SheetPicker
                label="Brand"
                sheets={brand.workbook.sheets}
                selected={brand.sheetName}
                onChange={onBrandSheetChange}
              />
            </div>
          </section>

          <section className="mt-6 card space-y-6">
            <div>
              <h2 className="text-lg font-semibold">2. Confirm column mapping</h2>
              <p className="text-xs text-slate-500">
                Required fields are marked with *. Override any wrong guess before reconciling.
              </p>
            </div>
            {marsSheet && (
              <ColumnMapper
                title="Mars ledger"
                headers={marsSheet.headers}
                fields={Object.keys(MARS_FIELD_SYNONYMS).filter((f) => !HIDDEN_FIELDS.has(f))}
                required={REQUIRED_MARS}
                mapping={mars.mapping}
                onChange={(m) => setMars((s) => ({ ...s, mapping: m }))}
              />
            )}
            {brandSheet && (
              <ColumnMapper
                title="Brand ledger"
                headers={brandSheet.headers}
                fields={Object.keys(BRAND_FIELD_SYNONYMS).filter((f) => !HIDDEN_FIELDS.has(f))}
                required={REQUIRED_BRAND}
                mapping={brand.mapping}
                onChange={(m) => setBrand((s) => ({ ...s, mapping: m }))}
              />
            )}
          </section>

          <section className="mt-6 card space-y-3">
            <h2 className="text-lg font-semibold">3. Run reconciliation</h2>
            {!canRun && (
              <p className="rounded-md bg-amber-50 px-3 py-2 text-sm text-amber-800">
                Map all required (*) fields before running.
                {missMars.length > 0 && <> Mars missing: {missMars.join(', ')}.</>}
                {missBrand.length > 0 && <> Brand missing: {missBrand.join(', ')}.</>}
              </p>
            )}
            <button
              type="button"
              className="btn-primary"
              disabled={!canRun || running}
              onClick={runRecon}
            >
              {running ? 'Reconciling…' : 'Reconcile'}
            </button>
            {reconError && (
              <div className="rounded-md border border-red-200 bg-red-50 p-3 text-sm text-red-800">
                {reconError}
              </div>
            )}
          </section>

          {reconResult && (
            <section className="mt-8">
              <ResultView result={reconResult} onDownload={downloadXlsx} downloading={downloading} />
            </section>
          )}
        </>
      )}

      {!marsFile && !brandFile && (
        <p className="mt-8 text-sm text-slate-500">Upload both files to begin.</p>
      )}

      <footer className="mt-12 border-t border-slate-200 pt-4 text-xs text-slate-500">
        Runs entirely in your browser. No backend, no uploads, no data sent anywhere.
      </footer>
    </div>
  );
}
