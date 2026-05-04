"""Core reconciliation logic — pure pandas, no UI.

Inputs: two DataFrames (Mars and Brand ledgers) plus a column mapping for each
that names which uploaded column plays each logical role (vch_no, net_amount,
period, category, ...). Output: a dict with three DataFrames (summary, mars,
brand) and a stats block.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, Optional

import pandas as pd


CANONICAL_CATEGORIES = ["Invoice", "Contra_Inv", "CN", "Contra_CN", "DN", "Others"]

# Map normalized (lowercase, underscores collapsed) category strings to canonical.
# Anything not in this map falls into "Others".
_CATEGORY_MAP = {
    "invoice": "Invoice",
    "inv": "Invoice",
    "contra_inv": "Contra_Inv",
    "contra_invoice": "Contra_Inv",
    "contrainv": "Contra_Inv",
    "cn": "CN",
    "credit_note": "CN",
    "creditnote": "CN",
    "contra_cn": "Contra_CN",
    "contracn": "Contra_CN",
    "dn": "DN",
    "debit_note": "DN",
    "debitnote": "DN",
}


def _norm_str(x) -> str:
    if pd.isna(x):
        return ""
    return str(x).strip()


def _norm_key(x) -> str:
    return _norm_str(x).upper()


def _norm_cat_for_match(x) -> str:
    """Normalized category used as the JOIN key between Mars and Brand."""
    s = _norm_str(x).lower()
    s = s.replace("-", "_").replace(" ", "_")
    while "__" in s:
        s = s.replace("__", "_")
    return s


def to_canonical_category(x) -> str:
    """Bucket a raw category value into one of CANONICAL_CATEGORIES."""
    s = _norm_cat_for_match(x)
    return _CATEGORY_MAP.get(s, "Others")


def _filter_recon(df: pd.DataFrame, period_col: str) -> pd.DataFrame:
    if period_col not in df.columns:
        raise KeyError(f"Period column '{period_col}' not found")
    mask = df[period_col].apply(_norm_str).str.lower() == "recon"
    return df[mask].copy()


@dataclass
class ReconResult:
    summary: pd.DataFrame
    mars: pd.DataFrame
    brand: pd.DataFrame
    stats: Dict[str, int]


def reconcile(
    mars_df: pd.DataFrame,
    brand_df: pd.DataFrame,
    mars_cols: Dict[str, str],
    brand_cols: Dict[str, str],
) -> ReconResult:
    """Reconcile Mars vs Brand ledgers.

    mars_cols required keys: vch_no, net_amount, period, category
    brand_cols required keys: reference, net_amount, period, category
    Other Mars keys (date, type, account, debit, credit) are passed through if mapped.
    """
    for k in ("vch_no", "net_amount", "period", "category"):
        if k not in mars_cols:
            raise ValueError(f"Mars column mapping missing required key: {k}")
    for k in ("reference", "net_amount", "period", "category"):
        if k not in brand_cols:
            raise ValueError(f"Brand column mapping missing required key: {k}")

    mars = _filter_recon(mars_df, mars_cols["period"])
    brand = _filter_recon(brand_df, brand_cols["period"])

    mars["_vch_key"] = mars[mars_cols["vch_no"]].apply(_norm_key)
    mars["_cat_key"] = mars[mars_cols["category"]].apply(_norm_cat_for_match)
    mars["_canonical_cat"] = mars[mars_cols["category"]].apply(to_canonical_category)
    mars["_amt"] = pd.to_numeric(mars[mars_cols["net_amount"]], errors="coerce").fillna(0)

    brand["_ref_key"] = brand[brand_cols["reference"]].apply(_norm_key)
    brand["_cat_key"] = brand[brand_cols["category"]].apply(_norm_cat_for_match)
    brand["_canonical_cat"] = brand[brand_cols["category"]].apply(to_canonical_category)
    brand["_amt"] = pd.to_numeric(brand[brand_cols["net_amount"]], errors="coerce").fillna(0)

    # Aggregate the OTHER side by (key, category) so many-to-one splits sum correctly.
    brand_agg = (
        brand.groupby(["_ref_key", "_cat_key"], as_index=False)
        .agg(_brand_amt=("_amt", "sum"), _brand_rows=("_amt", "size"))
    )
    mars_agg = (
        mars.groupby(["_vch_key", "_cat_key"], as_index=False)
        .agg(_mars_amt=("_amt", "sum"), _mars_rows=("_amt", "size"))
    )

    # ---- Mars output sheet ----
    mars_joined = mars.merge(
        brand_agg.rename(columns={"_ref_key": "_vch_key"}),
        on=["_vch_key", "_cat_key"],
        how="left",
    )
    mars_joined["Amount_Brand"] = mars_joined["_brand_amt"].fillna(0)
    mars_joined["Difference"] = mars_joined["_amt"] - mars_joined["Amount_Brand"]

    def _mars_remark(row):
        if pd.isna(row["_brand_amt"]):
            return "Not Booked by Brand"
        n = int(row["_brand_rows"])
        if n > 1:
            return f"Matched (Brand split: {n} rows)"
        return "Matched"

    mars_joined["Remarks"] = mars_joined.apply(_mars_remark, axis=1)

    mars_out = mars_joined.drop(
        columns=["_vch_key", "_cat_key", "_canonical_cat", "_amt", "_brand_amt", "_brand_rows"]
    )

    # ---- Brand output sheet ----
    brand_joined = brand.merge(
        mars_agg.rename(columns={"_vch_key": "_ref_key"}),
        on=["_ref_key", "_cat_key"],
        how="left",
    )
    brand_joined["Amount_Mars"] = brand_joined["_mars_amt"].fillna(0)
    brand_joined["Diff"] = brand_joined["Amount_Mars"] - brand_joined["_amt"]

    def _brand_remark(row):
        if pd.isna(row["_mars_amt"]):
            return "Not Booked by Mars"
        n = int(row["_mars_rows"])
        if n > 1:
            return f"Matched (Mars split: {n} rows)"
        return "Matched"

    brand_joined["Remarks"] = brand_joined.apply(_brand_remark, axis=1)

    brand_out = brand_joined.drop(
        columns=["_ref_key", "_cat_key", "_canonical_cat", "_amt", "_mars_amt", "_mars_rows"]
    )

    # ---- Recon Summary ----
    mars_by_cat = mars.groupby("_canonical_cat", as_index=False)["_amt"].sum().rename(
        columns={"_amt": "Amount_Mars"}
    )
    brand_by_cat = brand.groupby("_canonical_cat", as_index=False)["_amt"].sum().rename(
        columns={"_amt": "Amount_Brand"}
    )
    summary = pd.merge(mars_by_cat, brand_by_cat, on="_canonical_cat", how="outer")
    summary["Amount_Mars"] = summary["Amount_Mars"].fillna(0)
    summary["Amount_Brand"] = summary["Amount_Brand"].fillna(0)
    summary["Difference"] = summary["Amount_Mars"] - summary["Amount_Brand"]
    summary = summary.rename(columns={"_canonical_cat": "Particulars"})

    # Ensure every canonical category is represented even if absent in data
    present = set(summary["Particulars"])
    missing_rows = [
        {"Particulars": c, "Amount_Mars": 0.0, "Amount_Brand": 0.0, "Difference": 0.0}
        for c in CANONICAL_CATEGORIES
        if c not in present
    ]
    if missing_rows:
        summary = pd.concat([summary, pd.DataFrame(missing_rows)], ignore_index=True)

    summary["_order"] = summary["Particulars"].apply(
        lambda c: CANONICAL_CATEGORIES.index(c) if c in CANONICAL_CATEGORIES else len(CANONICAL_CATEGORIES)
    )
    summary = summary.sort_values("_order").drop(columns=["_order"]).reset_index(drop=True)

    total_row = pd.DataFrame([{
        "Particulars": "Total",
        "Amount_Mars": float(summary["Amount_Mars"].sum()),
        "Amount_Brand": float(summary["Amount_Brand"].sum()),
        "Difference": float(summary["Difference"].sum()),
    }])
    summary = pd.concat([summary, total_row], ignore_index=True)
    summary = summary[["Particulars", "Amount_Mars", "Amount_Brand", "Difference"]]

    matched_mars_rows = int((mars_out["Remarks"] != "Not Booked by Brand").sum())
    unmatched_mars_rows = int((mars_out["Remarks"] == "Not Booked by Brand").sum())
    unmatched_brand_rows = int((brand_out["Remarks"] == "Not Booked by Mars").sum())

    stats = {
        "mars_recon_rows": int(len(mars_out)),
        "brand_recon_rows": int(len(brand_out)),
        "mars_matched_rows": matched_mars_rows,
        "mars_unmatched_rows": unmatched_mars_rows,
        "brand_unmatched_rows": unmatched_brand_rows,
    }

    return ReconResult(summary=summary, mars=mars_out, brand=brand_out, stats=stats)
