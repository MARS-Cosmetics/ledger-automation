"""Synthetic-data sanity check for reconcile.py.

Run: python3 test_reconcile.py
Exits 0 on success, prints what each assertion is verifying so a failure tells
you which scenario broke.
"""

from __future__ import annotations

import sys
import pandas as pd

from reconcile import reconcile, to_canonical_category, CANONICAL_CATEGORIES
from column_detect import detect_columns, MARS_FIELD_SYNONYMS, BRAND_FIELD_SYNONYMS


def _check(label, cond):
    status = "PASS" if cond else "FAIL"
    print(f"  [{status}] {label}")
    if not cond:
        raise AssertionError(label)


def build_fixtures():
    mars = pd.DataFrame([
        # matched, simple
        {"Vch No": "V001", "Date": "2026-04-01", "Type": "Sales", "Account": "Brand A",
         "Debit": 0, "Credit": 1000, "Net Amount": 1000, "Period": "Recon", "Category": "Invoice"},
        # matched, amount mismatch (will produce a Difference)
        {"Vch No": "V002", "Net Amount": 500, "Period": "Recon", "Category": "Invoice",
         "Date": "", "Type": "", "Account": "", "Debit": 0, "Credit": 500},
        # not booked by Brand
        {"Vch No": "V003", "Net Amount": 700, "Period": "Recon", "Category": "CN",
         "Date": "", "Type": "", "Account": "", "Debit": 0, "Credit": 700},
        # outside Recon period — should be excluded
        {"Vch No": "V004", "Net Amount": 9999, "Period": "Apr", "Category": "Invoice",
         "Date": "", "Type": "", "Account": "", "Debit": 0, "Credit": 0},
        # category bucketed to Others
        {"Vch No": "V005", "Net Amount": 250, "Period": "Recon", "Category": "Journal",
         "Date": "", "Type": "", "Account": "", "Debit": 0, "Credit": 0},
        # Mars side of a many-to-one (Brand has 2 rows for V006)
        {"Vch No": "V006", "Net Amount": 1000, "Period": "Recon", "Category": "DN",
         "Date": "", "Type": "", "Account": "", "Debit": 0, "Credit": 0},
        # category written with space, should normalize and match Brand's "Contra_Inv"
        {"Vch No": "V007", "Net Amount": 300, "Period": "Recon", "Category": "Contra Inv",
         "Date": "", "Type": "", "Account": "", "Debit": 0, "Credit": 0},
    ])

    brand = pd.DataFrame([
        # matches V001 exactly
        {"Reference": "V001", "Net Amount": 1000, "Period": "Recon", "Category": "Invoice"},
        # matches V002 but different amount → Difference = 500 - 480 = 20
        {"Reference": "V002", "Net Amount": 480, "Period": "Recon", "Category": "Invoice"},
        # not booked by Mars
        {"Reference": "B999", "Net Amount": 150, "Period": "Recon", "Category": "DN"},
        # outside Recon
        {"Reference": "V001", "Net Amount": 9999, "Period": "Apr", "Category": "Invoice"},
        # split: two Brand rows for V006, summing to 1000 (matches Mars)
        {"Reference": "V006", "Net Amount": 600, "Period": "Recon", "Category": "DN"},
        {"Reference": "V006", "Net Amount": 400, "Period": "Recon", "Category": "DN"},
        # matches V007, written as Contra_Inv with underscore — different surface, same canonical
        {"Reference": "V007", "Net Amount": 300, "Period": "Recon", "Category": "Contra_Inv"},
    ])
    return mars, brand


def test_canonical_category():
    print("test_canonical_category")
    _check("Invoice → Invoice", to_canonical_category("Invoice") == "Invoice")
    _check("Contra Inv → Contra_Inv", to_canonical_category("Contra Inv") == "Contra_Inv")
    _check("contra-cn → Contra_CN", to_canonical_category("contra-cn") == "Contra_CN")
    _check("Credit Note → CN", to_canonical_category("Credit Note") == "CN")
    _check("DN → DN", to_canonical_category("DN") == "DN")
    _check("Journal → Others", to_canonical_category("Journal") == "Others")
    _check("'' → Others", to_canonical_category("") == "Others")


def test_column_detection():
    print("test_column_detection")
    mars, brand = build_fixtures()
    md = detect_columns(mars, MARS_FIELD_SYNONYMS)
    bd = detect_columns(brand, BRAND_FIELD_SYNONYMS)
    _check("Mars vch_no detected", md["vch_no"] == "Vch No")
    _check("Mars net_amount detected (not Debit/Credit)", md["net_amount"] == "Net Amount")
    _check("Mars debit detected", md["debit"] == "Debit")
    _check("Mars credit detected", md["credit"] == "Credit")
    _check("Mars period detected", md["period"] == "Period")
    _check("Mars category detected", md["category"] == "Category")
    _check("Brand reference detected", bd["reference"] == "Reference")
    _check("Brand net_amount detected", bd["net_amount"] == "Net Amount")


def test_reconcile_logic():
    print("test_reconcile_logic")
    mars, brand = build_fixtures()
    mars_cols = {"vch_no": "Vch No", "net_amount": "Net Amount", "period": "Period", "category": "Category"}
    brand_cols = {"reference": "Reference", "net_amount": "Net Amount", "period": "Period", "category": "Category"}

    res = reconcile(mars, brand, mars_cols, brand_cols)

    # --- Mars sheet ---
    m = res.mars
    _check("Mars sheet excludes non-Recon rows", "V004" not in m["Vch No"].values)
    _check("Mars sheet has 6 rows (7 minus the Apr row)", len(m) == 6)

    v001 = m[m["Vch No"] == "V001"].iloc[0]
    _check("V001 matched", v001["Remarks"] == "Matched")
    _check("V001 Amount_Brand=1000", v001["Amount_Brand"] == 1000)
    _check("V001 Difference=0", v001["Difference"] == 0)

    v002 = m[m["Vch No"] == "V002"].iloc[0]
    _check("V002 matched with amount mismatch", v002["Remarks"] == "Matched")
    _check("V002 Difference = 500-480 = 20", v002["Difference"] == 20)

    v003 = m[m["Vch No"] == "V003"].iloc[0]
    _check("V003 not booked by Brand", v003["Remarks"] == "Not Booked by Brand")
    _check("V003 Amount_Brand=0 (not blank)", v003["Amount_Brand"] == 0)
    _check("V003 Difference = own amount", v003["Difference"] == 700)

    v006 = m[m["Vch No"] == "V006"].iloc[0]
    _check("V006 remarks notes Brand split", "Brand split: 2 rows" in v006["Remarks"])
    _check("V006 Amount_Brand sums Brand rows = 1000", v006["Amount_Brand"] == 1000)
    _check("V006 Difference = 0 (split sums to Mars amt)", v006["Difference"] == 0)

    v007 = m[m["Vch No"] == "V007"].iloc[0]
    _check("V007 matched across Contra Inv vs Contra_Inv", v007["Remarks"] == "Matched")
    _check("V007 Difference=0", v007["Difference"] == 0)

    # --- Brand sheet ---
    b = res.brand
    _check("Brand sheet excludes non-Recon", len(b) == 6)
    b999 = b[b["Reference"] == "B999"].iloc[0]
    _check("B999 not booked by Mars", b999["Remarks"] == "Not Booked by Mars")
    _check("B999 Amount_Mars=0", b999["Amount_Mars"] == 0)

    # --- Summary ---
    s = res.summary
    _check("Summary first row is Invoice", s.iloc[0]["Particulars"] == "Invoice")
    expected_order = CANONICAL_CATEGORIES + ["Total"]
    _check(f"Summary order = {expected_order}", list(s["Particulars"]) == expected_order)

    inv_row = s[s["Particulars"] == "Invoice"].iloc[0]
    _check("Invoice Amount_Mars = 1000+500 = 1500", inv_row["Amount_Mars"] == 1500)
    _check("Invoice Amount_Brand = 1000+480 = 1480", inv_row["Amount_Brand"] == 1480)
    _check("Invoice Difference = 20", inv_row["Difference"] == 20)

    others_row = s[s["Particulars"] == "Others"].iloc[0]
    _check("Others Amount_Mars=250 (Journal bucketed)", others_row["Amount_Mars"] == 250)

    cn_row = s[s["Particulars"] == "CN"].iloc[0]
    _check("CN Mars=700, Brand=0", cn_row["Amount_Mars"] == 700 and cn_row["Amount_Brand"] == 0)

    dn_row = s[s["Particulars"] == "DN"].iloc[0]
    _check("DN Mars=1000, Brand=600+400+150=1150", dn_row["Amount_Mars"] == 1000 and dn_row["Amount_Brand"] == 1150)

    contra_inv_row = s[s["Particulars"] == "Contra_Inv"].iloc[0]
    _check("Contra_Inv Mars=300, Brand=300", contra_inv_row["Amount_Mars"] == 300 and contra_inv_row["Amount_Brand"] == 300)

    total_row = s[s["Particulars"] == "Total"].iloc[0]
    _check("Total Mars = 1500+700+250+1000+300 = 3750", total_row["Amount_Mars"] == 3750)
    _check("Total Brand = 1480+0+0+1150+300 = 2930", total_row["Amount_Brand"] == 2930)
    _check("Total Diff = 820", total_row["Difference"] == 820)

    # --- Stats ---
    st = res.stats
    _check("stats.mars_recon_rows=6", st["mars_recon_rows"] == 6)
    _check("stats.mars_unmatched_rows=2 (V003, V005)", st["mars_unmatched_rows"] == 2)
    _check("stats.brand_unmatched_rows=1 (B999)", st["brand_unmatched_rows"] == 1)


def main():
    try:
        test_canonical_category()
        test_column_detection()
        test_reconcile_logic()
    except AssertionError as e:
        print(f"\nFAILED: {e}")
        sys.exit(1)
    print("\nAll tests passed.")


if __name__ == "__main__":
    main()
