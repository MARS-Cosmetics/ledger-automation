"""Fuzzy header matching for ledger columns.

Maps a logical field (e.g. "net_amount") to whichever column header in the
uploaded spreadsheet best represents it, using a synonym dictionary plus
token-overlap scoring. Always exposed in the UI for confirmation/override —
this is a guess, not a guarantee.
"""

from __future__ import annotations

import re
from typing import Dict, List, Optional, Tuple

import pandas as pd


def normalize_header(h) -> str:
    s = re.sub(r"[^a-zA-Z0-9]+", " ", str(h).lower())
    return re.sub(r"\s+", " ", s).strip()


MARS_FIELD_SYNONYMS: Dict[str, List[str]] = {
    "vch_no": ["vch no", "voucher no", "vch number", "voucher number", "vchno", "vch", "voucher"],
    "date": ["date", "vch date", "voucher date"],
    "type": ["type", "vch type", "voucher type"],
    "account": ["account", "particulars", "ledger", "ledger account", "party"],
    "debit": ["debit", "dr", "debit amount", "debit amt"],
    "credit": ["credit", "cr", "credit amount", "credit amt"],
    "net_amount": ["net amount", "net amt", "net", "amount", "amt"],
    "period": ["period"],
    "category": ["category", "cat", "type of voucher", "voucher category"],
}

BRAND_FIELD_SYNONYMS: Dict[str, List[str]] = {
    "reference": ["reference", "ref", "ref no", "reference no", "vch no", "voucher no", "ref number"],
    "net_amount": ["net amount", "net amt", "net", "amount", "amt"],
    "period": ["period"],
    "category": ["category", "cat", "type of voucher", "voucher category"],
}

REQUIRED_MARS = ["vch_no", "net_amount", "period", "category"]
REQUIRED_BRAND = ["reference", "net_amount", "period", "category"]


def _score(header_norm: str, candidate: str) -> float:
    if header_norm == candidate:
        return 1000.0
    h_tokens = set(header_norm.split())
    c_tokens = set(candidate.split())
    if not c_tokens or not h_tokens:
        return 0.0
    if c_tokens.issubset(h_tokens):
        # candidate fully contained — longer candidate wins (more specific)
        return 500.0 + len(candidate)
    overlap = len(h_tokens & c_tokens)
    if overlap == 0:
        return 0.0
    # partial overlap, weight by candidate length so "net amount" beats "amount"
    return overlap * 10.0 + len(candidate) * 0.1


def detect_columns(df: pd.DataFrame, synonyms: Dict[str, List[str]]) -> Dict[str, Optional[str]]:
    """For each logical field, pick the best-scoring column header.

    Greedy by field order (in the synonyms dict). A header already claimed by an
    earlier field is unavailable to later ones — so put more-specific fields
    first in the dict. We also block "amount" from claiming columns whose
    headers contain "debit"/"credit" so debit/credit always win their match.
    """
    headers = list(df.columns)
    norm: Dict[str, str] = {h: normalize_header(h) for h in headers}
    used = set()
    result: Dict[str, Optional[str]] = {}

    for field, candidates in synonyms.items():
        best: Tuple[float, Optional[str]] = (0.0, None)
        for header in headers:
            if header in used:
                continue
            n = norm[header]
            # guardrail: don't let a generic "amount" synonym swallow a debit/credit column
            if field == "net_amount" and any(t in n.split() for t in ("debit", "credit", "dr", "cr")):
                continue
            for cand in candidates:
                s = _score(n, cand)
                if s > best[0]:
                    best = (s, header)
        result[field] = best[1]
        if best[1] is not None:
            used.add(best[1])

    return result


def missing_required(detected: Dict[str, Optional[str]], required: List[str]) -> List[str]:
    return [f for f in required if not detected.get(f)]
