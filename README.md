# Mars Recon Tool

Streamlit app for reconciling Mars vs Brand ledgers.

## Run locally

```bash
./run.sh
```

First run installs dependencies. Subsequent runs launch the app in your browser at http://localhost:8501.

## Update to the latest version

```bash
git pull
./run.sh
```

## Files

- `app.py` — Streamlit UI
- `column_detect.py` — auto-detects column mappings from headers
- `reconcile.py` — core reconciliation logic
- `test_reconcile.py` — tests
- `_make_sample_output.py` — utility to regenerate `sample_output.xlsx`
