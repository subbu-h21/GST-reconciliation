"""
GST Reconciliation Engine
=========================
Compares Books (Inward Register) vs GSTR-2B to find matching invoices.

Pipeline:
  Stage 0  — load, clean, recompute money, group books to invoice grain
  Stage 1  — exact matching (1A / 1B with uniqueness gate / 1C collision)
  Stage 2  — standardized-invoice matching (same logic, invoice_std)
  Stage 3  — split into blue / nonblue queues (no auto-match here)
  Stage 4A — AI pass on blue queue
  Stage 4B — AI pass on nonblue queue (pile builder filters no-money-twin rows)
  Stage 5  — greedy one-to-one merge + 7-sheet output workbook

Fixes applied vs submitted version:
  - B2C rows excluded from unmatched_books (were double-listed)
  - 1B uniqueness gate added (spec requirement; masks collision clusters)
  - Stage 2 now gets work when 1B is correctly restricted
  - Type-column guard with loud failure if label set changes
  - money_match checks taxable + each tax head individually (per spec)
  - find_identity_money_mismatch uses invoice_std + uniqueness guard
  - build_piles size-gate pre-pass now records and returns its matches
  - Stage 3 q_remaining_g excludes blue 2B rows (no double-proposal risk)
  - per-pile AI logging added (spec requirement)
  - source column added to canonical schema for both books and 2B
"""

from __future__ import annotations

import asyncio
import json
import os
import re
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from pathlib import Path

import networkx as nx
import pandas as pd
from dotenv import load_dotenv
from openai import AsyncOpenAI

load_dotenv()

# ── config ────────────────────────────────────────────────────────────────────
BOOKS_PATH  = r"C:\Users\subra\OneDrive\Inward Register Feb-25.xlsx"
B2B_PATH    = r"C:\Users\subra\OneDrive\022025_29ADAFS3950J1ZS_GSTR2B_18032025_SSA Firm.xlsx"
OUTPUT_PATH = Path("gst_reconciliation_output.xlsx")

TOL                 = 2.0
PILE_SIZE_LIMIT     = 15      # rows per side before exact-invoice pre-pass inside pile
BLUE_AUTO_ACCEPT    = 0.90
BLUE_MIN_REVIEW     = 0.60
NONBLUE_AUTO_ACCEPT = 0.93
NONBLUE_MIN_REVIEW  = 0.75
AI_MODEL            = "openai/gpt-4o-mini"   # OpenRouter model ID
# AI_MODEL            = "gemini-2.5-flash-lite"
# AI_MODEL            = "openai/gpt-5-mini"     # stronger, but more expensive and slower — use only if needed

RATE_COLS = {
    "Taxable @ 0%":  0.00,
    "Taxable @ 5%":  0.05,
    "Taxable @ 12%": 0.12,
    "Taxable @ 18%": 0.18,
    "Taxable @ 28%": 0.28,
}

# ── helpers ───────────────────────────────────────────────────────────────────

def clean_str(val) -> str | None:
    """Strip, uppercase, return None for blanks / NA sentinels."""
    if pd.isna(val):
        return None
    s = str(val).strip().upper()
    return s if s not in ("", "NAN", "<NA>", "NONE", "NAT") else None


def standardize_invoice(val) -> str | None:
    """Remove all non-alphanumeric, uppercase, drop leading zeros."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    s = re.sub(r"[^A-Z0-9]", "", str(val).upper()).lstrip("0")
    return s or None


def money_match(b: pd.Series, g: pd.Series, tol: float = TOL) -> bool:
    """
    True when taxable AND total_tax are both within tol.

    We compare total_tax (cgst+sgst+igst) as a single figure rather than
    each head independently. This correctly handles place-of-supply
    differences where books record CGST+SGST but 2B reports IGST (or
    vice versa) — same taxable, same total tax, legitimately the same
    invoice.
    """
    return (
        abs(b["taxable"]   - g["taxable"])   <= tol
        and abs(b["total_tax"] - g["total_tax"]) <= tol
    )


# ── match record ──────────────────────────────────────────────────────────────

@dataclass
class Match:
    books_id:      str
    b2b_id:        str
    stage:         str    # "1" | "2"
    sub:           str    # "1A" | "1B" | "1C" | "2A" | "2B"
    match_reason:  str
    taxable_diff:  float
    total_tax_diff: float


# ═══════════════════════════════════════════════════════════════════════════════
# STAGE 0 — LOAD, CLEAN, RECOMPUTE, GROUP
# ═══════════════════════════════════════════════════════════════════════════════

_INTRA_LABELS = {"INTRA-STATE", "INTRA STATE", "INTRASTATE", "LOCAL"}
_INTER_LABELS = {"INTER-STATE", "INTER STATE", "INTERSTATE", "IMPORT", "SEZ"}


def load_books(path: str) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Load and clean the books (Inward Register).
    Returns (clean_df, dirty_df).

    Money strategy:
      1. Recompute taxable from rate-slab columns (rate cols are always populated).
      2. Use the Type column (INTRA-STATE / INTER-STATE) to split into
         CGST+SGST vs IGST.
      3. Validate: |recomputed_gross - Gross Total| <= 2.
         Rows that fail → dirty_df (skip matching, write to dirty_input sheet).

    Guard: if Type column has unexpected labels, raise immediately so the
    caller knows not to trust the tax-head split.
    """
    df = pd.read_excel(path, header=1, engine="openpyxl")
    df = df[~(df["Supplier"].isna() & df["Gross Total"].isna())].copy()

    # coerce all numeric columns
    numeric_cols = [
        "Gross Total", "Taxable Value", "CGST", "SGST", "IGST",
        "Cess", "Exempt", "Others",
        *[c for c in RATE_COLS if c in df.columns],
    ]
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    # ── Type-column guard ────────────────────────────────────────────────────
    if "Type" not in df.columns:
        raise ValueError(
            "Books file is missing a 'Type' column (expected INTRA-STATE / "
            "INTER-STATE). Cannot split CGST/SGST vs IGST without it."
        )
    raw_types = set(df["Type"].dropna().astype(str).str.strip().str.upper().unique())
    unknown   = raw_types - _INTRA_LABELS - _INTER_LABELS - {"NAN", ""}
    if unknown:
        raise ValueError(
            f"Books 'Type' column has unexpected values: {unknown}. "
            "Update _INTRA_LABELS / _INTER_LABELS in config and verify."
        )

    is_intra = df["Type"].astype(str).str.strip().str.upper().isin(_INTRA_LABELS)

    # ── recompute money from rate-slab columns ───────────────────────────────
    avail = [c for c in RATE_COLS if c in df.columns]
    if not avail:
        raise ValueError("No rate-slab columns found (Taxable @ X%). Cannot recompute money.")

    df["taxable"] = sum(df[c] for c in avail).round(2)

    # CGST = half the rate on intra-state rows, 0 on inter-state
    df["cgst"] = sum(
        df[c] * (RATE_COLS[c] / 2) for c in avail
    ).where(is_intra, 0).round(2)
    df["sgst"] = df["cgst"].copy()

    # IGST = full rate on inter-state rows, 0 on intra-state
    df["igst"] = sum(
        df[c] * RATE_COLS[c] for c in avail
    ).where(~is_intra, 0).round(2)

    df["total_tax"] = (df["cgst"] + df["sgst"] + df["igst"]).round(2)

    # fallback: where recomputed value is 0 but cached column is non-zero, use cached
    cached_taxable = pd.to_numeric(df.get("Taxable Value", 0), errors="coerce").fillna(0)
    cached_cgst    = pd.to_numeric(df.get("CGST", 0),          errors="coerce").fillna(0)
    cached_sgst    = pd.to_numeric(df.get("SGST", 0),          errors="coerce").fillna(0)
    cached_igst    = pd.to_numeric(df.get("IGST", 0),          errors="coerce").fillna(0)

    df["taxable"]   = df["taxable"].where(df["taxable"] != 0, cached_taxable)
    df["cgst"]      = df["cgst"].where(df["cgst"]       != 0, cached_cgst)
    df["sgst"]      = df["sgst"].where(df["sgst"]       != 0, cached_sgst)
    df["igst"]      = df["igst"].where(df["igst"]       != 0, cached_igst)
    df["total_tax"] = (df["cgst"] + df["sgst"] + df["igst"]).round(2)

    cess   = df["Cess"]   if "Cess"   in df.columns else 0
    exempt = df["Exempt"] if "Exempt" in df.columns else 0
    others = df["Others"] if "Others" in df.columns else 0
    df["gross_recomputed"] = (
        df["taxable"] + df["total_tax"] + cess + exempt + others
    ).round(2)
    df["gross_diff"] = (df["Gross Total"] - df["gross_recomputed"]).round(2)

    # ── clean identity fields ────────────────────────────────────────────────
    for col in ["GSTIN/UIN", "Voucher Number", "Supplier"]:
        df[col] = df[col].apply(clean_str)
    df["invoice_date"] = pd.to_datetime(df.get("Accounting Date"), errors="coerce")

    dirty = df[df["gross_diff"].abs() > TOL].copy()
    clean = df[df["gross_diff"].abs() <= TOL].copy()

    print(
        f"[Stage 0] Books: {len(df)} rows | "
        f"dirty={len(dirty)} | clean={len(clean)} | "
        f"intra={is_intra.sum()} inter={(~is_intra).sum()}"
    )
    return clean, dirty


def load_2b(path: str) -> pd.DataFrame:
    """Load GSTR-2B B2B sheet; flatten MultiIndex header; return canonical frame."""
    df = pd.read_excel(path, sheet_name="B2B", header=[4, 5], engine="openpyxl")

    sup_col = ("Trade/Legal name", "Unnamed: 1_level_1")
    val_col = ("Invoice Details", "Invoice Value(₹)")
    df = df[~(df[sup_col].isna() & df[val_col].isna())].copy()

    # flatten MultiIndex: keep level0 only when level1 is "Unnamed…"
    df.columns = [
        a if "Unnamed" in b else f"{a}::{b}"
        for a, b in df.columns
    ]

    df = df.rename(columns={
        "GSTIN of supplier":                 "gstin",
        "Trade/Legal name":                  "supplier",
        "Invoice Details::Invoice number":   "invoice_raw",
        "Invoice Details::Invoice Date":     "invoice_date",
        "Invoice Details::Invoice Value(₹)": "invoice_value",
        "Taxable Value (₹)":                 "taxable",
        "Tax Amount::Central Tax(₹)":        "cgst",
        "Tax Amount::State/UT Tax(₹)":       "sgst",
        "Tax Amount::Integrated Tax(₹)":     "igst",
    })

    for col in ["taxable", "cgst", "sgst", "igst", "invoice_value"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    df["total_tax"]    = (df["cgst"] + df["sgst"] + df["igst"]).round(2)
    df["invoice_date"] = pd.to_datetime(df.get("invoice_date"), errors="coerce")

    for col in ["gstin", "invoice_raw", "supplier"]:
        df[col] = df[col].apply(clean_str)

    df["invoice_std"] = df["invoice_raw"].apply(standardize_invoice)
    df["source"] = "2b"
    df = df.reset_index(drop=True)
    df.insert(0, "row_id", ["G" + str(i) for i in df.index])

    print(f"[Stage 0] 2B: {len(df)} rows")
    return df


def group_books(df: pd.DataFrame) -> pd.DataFrame:
    """
    Group books to invoice grain.

    Key assignment (OR, not AND — each row gets exactly one key):
      GSTIN + invoice present → key = (gstin, invoice_raw)
      GSTIN missing, invoice present → key = (supplier, invoice_raw)
      No invoice (regardless of GSTIN) → standalone key (never merged)

    The 2B is NOT grouped — it is already one row per (GSTIN + invoice).
    """
    def make_key(row):
        g = row["GSTIN/UIN"]
        v = row["Voucher Number"]
        s = row["Supplier"]
        if v:
            if g:
                return ("gstin_invoice", g, v)
            else:
                return ("supplier_invoice", s or f"_null_{row.name}", v)
        # no invoice — standalone; row.name is the original df index (unique)
        return ("standalone", f"_solo_{row.name}", f"_solo_{row.name}")

    df = df.copy()
    df[["_ktype", "_k1", "_k2"]] = df.apply(
        lambda r: pd.Series(make_key(r)), axis=1
    )

    grouped = (
        df.groupby(["_ktype", "_k1", "_k2"], dropna=False, sort=False)
        .agg(
            gstin          = ("GSTIN/UIN",      "first"),
            invoice_raw    = ("Voucher Number",  "first"),
            supplier       = ("Supplier",        "first"),
            invoice_date   = ("invoice_date",    "first"),
            taxable        = ("taxable",         "sum"),
            cgst           = ("cgst",            "sum"),
            sgst           = ("sgst",            "sum"),
            igst           = ("igst",            "sum"),
            member_row_ids = ("Voucher Number",  lambda x: list(x.index)),
        )
        .reset_index()
    )

    grouped["total_tax"]    = (grouped["cgst"] + grouped["sgst"] + grouped["igst"]).round(2)
    grouped["invoice_std"]  = grouped["invoice_raw"].apply(standardize_invoice)
    grouped["source"] = "books"
    grouped["invoice_date"] = pd.to_datetime(grouped["invoice_date"], errors="coerce")
    grouped = grouped.reset_index(drop=True)
    grouped.insert(0, "row_id", ["B" + str(i) for i in grouped.index])

    n_multi      = (grouped["member_row_ids"].apply(len) > 1).sum()
    n_no_gstin   = grouped["gstin"].isna().sum()
    n_standalone = (grouped["_ktype"] == "standalone").sum()

    print(
        f"[Stage 0] Grouped books: {len(grouped)} rows | "
        f"multi-line={n_multi} | no-GSTIN={n_no_gstin} | standalone={n_standalone}"
    )
    return grouped


# ═══════════════════════════════════════════════════════════════════════════════
# STAGE 1 & 2 — EXACT / STANDARDIZED MATCHING
# ═══════════════════════════════════════════════════════════════════════════════

def _identity(b: pd.Series, g: pd.Series, use_std: bool) -> tuple[bool, bool, bool]:
    """Return (gstin_match, supplier_match, invoice_match). Nulls never equal."""
    inv_b = b["invoice_std"] if use_std else b["invoice_raw"]
    inv_g = g["invoice_std"] if use_std else g["invoice_raw"]
    gm = bool(b["gstin"]    and g["gstin"]    and b["gstin"]    == g["gstin"])
    sm = bool(b["supplier"] and g["supplier"] and b["supplier"] == g["supplier"])
    im = bool(inv_b         and inv_g         and inv_b         == inv_g)
    return gm, sm, im


def _date_diff(b: pd.Series, g: pd.Series) -> int:
    bd, gd = b.get("invoice_date"), g.get("invoice_date")
    if pd.isnull(bd) or pd.isnull(gd):
        return 9999
    try:
        return abs((pd.Timestamp(bd) - pd.Timestamp(gd)).days)
    except Exception:
        return 9999


def _completeness(row: pd.Series) -> int:
    return sum(1 for f in ("gstin", "invoice_raw", "supplier") if row.get(f))


def _rank_key(c: dict) -> tuple:
    """Lower = better candidate. Tie-break order per spec."""
    return (
        0 if c["im"] else 1,
        0 if c["gm"] else 1,
        0 if c["sm"] else 1,
        c["total_tax_diff"],
        c["date_diff_days"],
        -(c["completeness_b"] + c["completeness_g"]),
    )


def _apply_tiebreak(
    candidates: list[dict],
    used_b: set[str],
    used_g: set[str],
) -> tuple[list[dict], set[str]]:
    """
    Pick the single best 2B candidate for each books row.
    - Genuine ties (top two share the same rank key) → books row is NOT matched;
      it stays in the pool (flows to Stage 3 blue or later).
    - Sort globally so a higher-priority books row wins a contested 2B slot.
    Returns (accepted, tied_books_ids).
    """
    by_books: dict[str, list[dict]] = defaultdict(list)
    for c in candidates:
        if c["books_id"] not in used_b and c["b2b_id"] not in used_g:
            by_books[c["books_id"]].append(c)

    tied_books: set[str] = set()
    proposals: list[dict] = []

    for bid, cands in by_books.items():
        cands.sort(key=_rank_key)
        if len(cands) >= 2 and _rank_key(cands[0]) == _rank_key(cands[1]):
            tied_books.add(bid)
        else:
            proposals.append(cands[0])

    # global sort: highest-priority books rows claim their 2B slot first
    proposals.sort(key=_rank_key)

    accepted: list[dict] = []
    for c in proposals:
        if c["books_id"] not in used_b and c["b2b_id"] not in used_g:
            accepted.append(c)
            used_b.add(c["books_id"])
            used_g.add(c["b2b_id"])

    return accepted, tied_books


def _is_unique_pair(
    b: pd.Series,
    g: pd.Series,
    books: pd.DataFrame,
    gstr2b: pd.DataFrame,
) -> bool:
    """
    Uniqueness gate for 1B / 2B matches (spec §Stage 1 — 1B rule).

    A loose match (identity + money, invoice differs) is only safe to
    auto-accept when exactly ONE books row and ONE 2B row share that
    (identity, amount±TOL) combination. If multiple rows exist on either
    side the amount is not a unique fingerprint — route to blue instead.

    Identity anchor: prefer GSTIN if both sides have it; fall back to supplier.
    """
    # pick the strongest shared identity
    use_gstin = bool(b["gstin"] and g["gstin"] and b["gstin"] == g["gstin"])

    if use_gstin:
        same_b = books[
            (books["gstin"] == b["gstin"]) &
            ((books["taxable"] - b["taxable"]).abs() <= TOL)
        ]
        same_g = gstr2b[
            (gstr2b["gstin"] == g["gstin"]) &
            ((gstr2b["taxable"] - g["taxable"]).abs() <= TOL)
        ]
    else:
        # supplier match
        same_b = books[
            (books["supplier"] == b["supplier"]) &
            ((books["taxable"] - b["taxable"]).abs() <= TOL)
        ]
        same_g = gstr2b[
            (gstr2b["supplier"] == g["supplier"]) &
            ((gstr2b["taxable"] - g["taxable"]).abs() <= TOL)
        ]

    return len(same_b) == 1 and len(same_g) == 1


def run_exact_stage(
    books: pd.DataFrame,
    gstr2b: pd.DataFrame,
    use_std: bool = False,
) -> tuple[list[Match], pd.DataFrame, pd.DataFrame]:
    """
    Stage 1 (use_std=False) or Stage 2 (use_std=True).

    Tight pass (xA / xC): money_match AND invoice matches
      → auto-match (any collision handled by tie-break → 1A or 1C)

    Loose pass (xB): money_match AND (gstin OR supplier) matches, invoice differs
      → auto-match ONLY when (identity, amount) is UNIQUE on both sides
      → collision cluster (non-unique) → stays in pool → becomes blue in Stage 3

    Tie-break: invoice > gstin > supplier > smallest money diff > closest date
    """
    lbl = "2" if use_std else "1"

    # build all candidates that have at least one identity match
    candidates: list[dict] = []
    for _, b in books.iterrows():
        for _, g in gstr2b.iterrows():
            if not money_match(b, g):
                continue
            gm, sm, im = _identity(b, g, use_std)
            if not (gm or sm):
                continue
            candidates.append({
                "books_id":       b["row_id"],
                "b2b_id":         g["row_id"],
                "gm": gm, "sm": sm, "im": im,
                "taxable_diff":   round(abs(b["taxable"]   - g["taxable"]),   2),
                "total_tax_diff": round(abs(b["total_tax"] - g["total_tax"]), 2),
                "date_diff_days": _date_diff(b, g),
                "completeness_b": _completeness(b),
                "completeness_g": _completeness(g),
            })

    used_b: set[str] = set()
    used_g: set[str] = set()
    matches: list[Match] = []

    # ── tight pass: invoice must match ───────────────────────────────────────
    tight       = [c for c in candidates if c["im"]]
    tight_cnt_b = Counter(c["books_id"] for c in tight)

    tight_accepted, tight_tied = _apply_tiebreak(tight, used_b, used_g)
    for c in tight_accepted:
        collision = tight_cnt_b[c["books_id"]] > 1
        sub    = f"{lbl}C" if collision else f"{lbl}A"
        reason = "exact identity + invoice + money"
        if collision:
            reason += " (collision resolved via tie-break)"
        matches.append(Match(
            c["books_id"], c["b2b_id"], lbl, sub, reason,
            c["taxable_diff"], c["total_tax_diff"],
        ))

    # ── loose pass: invoice differs, identity + money match ──────────────────
    # UNIQUENESS GATE: only auto-accept when (identity, amount) is unique on
    # both sides. Non-unique pairs stay free → become blue in Stage 3.
    books_idx  = books.set_index("row_id")
    gstr2b_idx = gstr2b.set_index("row_id")

    loose = [
        c for c in candidates
        if not c["im"]
        and c["books_id"] not in used_b
        and c["b2b_id"]   not in used_g
    ]
    loose_unique = [
        c for c in loose
        if _is_unique_pair(
            books_idx.loc[c["books_id"]],
            gstr2b_idx.loc[c["b2b_id"]],
            books,
            gstr2b,
        )
    ]
    loose_accepted, loose_tied = _apply_tiebreak(loose_unique, used_b, used_g)
    for c in loose_accepted:
        matches.append(Match(
            c["books_id"], c["b2b_id"], lbl, f"{lbl}B",
            "identity + money; invoice differs (unique pair — format/voucher mismatch)",
            c["taxable_diff"], c["total_tax_diff"],
        ))

    rem_books = books[~books["row_id"].isin(used_b)].copy()
    rem_2b    = gstr2b[~gstr2b["row_id"].isin(used_g)].copy()

    n_tied = len(tight_tied) + len(loose_tied)
    print(
        f"[Stage {lbl}] matches={len(matches)} | tied→blue={n_tied} | "
        f"remaining books={len(rem_books)}, 2B={len(rem_2b)}"
    )
    return matches, rem_books, rem_2b


# ═══════════════════════════════════════════════════════════════════════════════
# STAGE 3 — IDENTIFY BLUE ROWS ONLY (no auto-match)
# ═══════════════════════════════════════════════════════════════════════════════

def run_stage3(
    books: pd.DataFrame,
    gstr2b: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Classify leftover books rows into:
      BLUE:      money + (gstin OR supplier) matches, invoice differs.
                 → strong anchor, easy AI question.
      REMAINING: no identity anchor at all.
                 → pile builder filters further (only piles with a money-twin
                   reach the AI; isolated nodes go straight to unmatched).

    No rows are removed from the pool here — classification only.
    The same 2B candidate pool is passed to both 4A and 4B; the pile
    builder handles isolation of no-money-twin 2B rows.
    """
    blue_b: set[str] = set()
    blue_g: set[str] = set()

    for _, b in books.iterrows():
        for _, g in gstr2b.iterrows():
            if not money_match(b, g):
                continue
            id_match = (
                (b["gstin"]    and g["gstin"]    and b["gstin"]    == g["gstin"])
                or (b["supplier"] and g["supplier"] and b["supplier"] == g["supplier"])
            )
            if id_match:
                blue_b.add(b["row_id"])
                blue_g.add(g["row_id"])

    q_blue_b      = books[books["row_id"].isin(blue_b)].copy()
    q_blue_g      = gstr2b[gstr2b["row_id"].isin(blue_g)].copy()
    q_remaining_b = books[~books["row_id"].isin(blue_b)].copy()
    q_remaining_g = gstr2b[~gstr2b["row_id"].isin(blue_g)].copy()  # exclude blue 2B rows

    print(
        f"[Stage 3] blue_books={len(q_blue_b)} | "
        f"remaining→4B: books={len(q_remaining_b)}, 2B={len(q_remaining_g)}"
    )
    return q_blue_b, q_blue_g, q_remaining_b, q_remaining_g


# ═══════════════════════════════════════════════════════════════════════════════
# PILE BUILDER — connected components on money-match edges
# ═══════════════════════════════════════════════════════════════════════════════

def build_piles(books: pd.DataFrame, gstr2b: pd.DataFrame) -> tuple[list[dict], list[dict]]:
    """
    Partition rows into disjoint piles using connected components.

    Edge rule: (books_i, 2b_j) iff money_match(books_i, 2b_j).

    Guarantees:
      - Every row in exactly one pile (partition — no row in two piles).
      - No possible match ever crosses a pile boundary.
      - Rows with no money-twin on the other side → single-side pile → skipped
        by AI → go straight to unmatched.

    Size gate: piles with > PILE_SIZE_LIMIT rows per side get an exact-invoice
    pre-pass first; only the ambiguous remainder reaches the AI.

    Returns (piles, pre_matches) where pre_matches are the pairs resolved
    deterministically by the size-gate pre-pass.
    """
    G = nx.Graph()
    for _, b in books.iterrows():
        G.add_node(("B", b["row_id"]), data=b.to_dict())
    for _, g in gstr2b.iterrows():
        G.add_node(("G", g["row_id"]), data=g.to_dict())
    for _, b in books.iterrows():
        for _, g in gstr2b.iterrows():
            if money_match(b, g):
                G.add_edge(("B", b["row_id"]), ("G", g["row_id"]))

    piles: list[dict] = []
    pre_matches: list[dict] = []

    for comp in nx.connected_components(G):
        b_rows = [G.nodes[n]["data"] for n in comp if n[0] == "B"]
        g_rows = [G.nodes[n]["data"] for n in comp if n[0] == "G"]

        # single-side pile — no AI needed
        if not b_rows or not g_rows:
            piles.append({"books": b_rows, "2b": g_rows})
            continue

        # size gate: pre-pass exact-invoice matches within oversized piles
        if len(b_rows) > PILE_SIZE_LIMIT or len(g_rows) > PILE_SIZE_LIMIT:
            paired_b: set[str] = set()
            paired_g: set[str] = set()
            for b in b_rows:
                for g in g_rows:
                    if (
                        b["invoice_std"] and g["invoice_std"]
                        and b["invoice_std"] == g["invoice_std"]
                        and money_match(pd.Series(b), pd.Series(g))
                        and b["row_id"] not in paired_b
                        and g["row_id"] not in paired_g
                    ):
                        paired_b.add(b["row_id"])
                        paired_g.add(g["row_id"])
                        pre_matches.append({
                            "books_id":   b["row_id"],
                            "b2b_id":     g["row_id"],
                            "confidence": 1.0,
                            "reason":     "size-gate exact-invoice pre-pass",
                            "_pass":      "size_gate",
                        })

            b_rem = [r for r in b_rows if r["row_id"] not in paired_b]
            g_rem = [r for r in g_rows if r["row_id"] not in paired_g]
            if b_rem or g_rem:
                piles.append({"books": b_rem, "2b": g_rem})
        else:
            piles.append({"books": b_rows, "2b": g_rows})

    return piles, pre_matches


# ═══════════════════════════════════════════════════════════════════════════════
# STAGE 4A / 4B — AI MATCHING (match-the-following via tool calling)
# ═══════════════════════════════════════════════════════════════════════════════

_TOOL: dict = {
    "type": "function",
    "function": {
        "name": "submit_pile_matches",
        "description": (
            "Submit one-to-one matching decisions for this pile. "
            "Every row ID must appear exactly once across matches, "
            "unmatched_books, and unmatched_b2b."
        ),
        "parameters": {
            "type": "object",
            "required": ["reasoning", "matches", "unmatched_books", "unmatched_b2b"],
            "properties": {
                "reasoning": {
                    "type": "string",
                    "description": (
                        "Think step by step BEFORE committing to decisions. "
                        "This field is evaluated first — use it to reason about "
                        "each candidate pair before writing your final answer."
                    ),
                },
                "matches": {
                    "type": "array",
                    "items": {
                        "type": "object",
                        "required": ["books_id", "b2b_id", "reason", "confidence"],
                        "properties": {
                            "books_id":   {"type": "string"},
                            "b2b_id":     {"type": "string"},
                            "reason":     {"type": "string"},
                            "confidence": {
                                "type": "number",
                                "minimum": 0.0,
                                "maximum": 1.0,
                            },
                        },
                    },
                },
                "unmatched_books": {
                    "type": "array",
                    "items": {
                        "type": "object",
                        "required": ["books_id", "reason"],
                        "properties": {
                            "books_id": {"type": "string"},
                            "reason":   {"type": "string"},
                        },
                    },
                },
                "unmatched_b2b": {
                    "type": "array",
                    "items": {
                        "type": "object",
                        "required": ["b2b_id", "reason"],
                        "properties": {
                            "b2b_id": {"type": "string"},
                            "reason": {"type": "string"},
                        },
                    },
                },
            },
        },
    },
}

_SYSTEM_PROMPT = (
    "You are a GST reconciliation assistant. "
    "Your only job is to decide identity: whether a books invoice and a 2B invoice "
    "refer to the same real-world transaction. "
    "Financial amounts already match within ₹2 for every row in the pile — "
    "do NOT re-evaluate the money. Focus entirely on the identity fields: "
    "GSTIN (supplier tax ID), invoice number, and supplier name. "
    "Apply strict one-to-one matching: each books row pairs with at most one "
    "2B row and vice versa. A false match is worse than leaving a row unmatched. "
    "When rows have identical amounts and supplier names, "
    "rely on invoice dates to find the best one-to-one pairing. "
    "Sort both sides by date and match the closest pairs. "
    "A confident date-based pairing is better than returning no matches."
)


def _pile_prompt(pile: dict, pass_type: str) -> str:
    if pass_type == "blue":
        question = (
            "QUESTION: Do any of these books rows refer to the same invoice as a 2B row, "
            "even though the invoice numbers are written differently? "
            "(The GSTIN or supplier already matches — the invoice number is the only disagreement.)"
        )
        instruction = (
            "Accept a match if the invoice numbers are plausibly the same document — "
            "e.g. a supplier number vs an internal voucher, or a formatting difference. "
            "Auto-accept threshold is 0.90."
        )
    else:
        question = (
            "QUESTION: Do any of these books rows refer to the same invoice as a 2B row? "
            "No identity field (GSTIN, invoice number, supplier) has been confirmed to match "
            "by automated rules. Be conservative — only match if you are highly confident "
            "these describe the same company and the same transaction."
        )
        instruction = (
            "A wrong match here is harder to detect than a miss. "
            "When in doubt, leave a row unmatched. Auto-accept threshold is 0.93."
        )

    lines = [
        f"GST reconciliation — {pass_type} AI pass.",
        "",
        "CONTEXT:",
        "  All rows in this pile already agree on taxable value and total tax within ₹2.",
        "  Decide IDENTITY only.",
        "  One-to-one constraint: each row may be used in at most one match.",
        "  Every ID (every Bi and every Gj) must appear exactly once across",
        "  matches / unmatched_books / unmatched_b2b.",
        "",
        "COLUMN MEANINGS:",
        "  gstin    = 15-character supplier tax registration number",
        "  invoice  = invoice / voucher number as recorded on each side",
        "  supplier = trade or legal name of the supplier",
        "  taxable  = taxable value (₹)",
        "  cgst/sgst/igst = GST components (₹)",
        "",
        question,
        "",
        instruction,
        "",
        "BOOKS ROWS (internal purchase register):",
    ]

    for i, r in enumerate(pile["books"], 1):
        lines.append(
            f"  B{i}: gstin={r['gstin'] or 'N/A'}  "
            f"invoice={r['invoice_raw'] or 'N/A'}  "
            f"date={str(r['invoice_date'])[:10] if r.get('invoice_date') and str(r['invoice_date']) != 'NaT' else 'N/A'}  "
            f"supplier={r['supplier'] or 'N/A'}  "
            f"taxable={r['taxable']}  total_tax={r['total_tax']}"
        )

    lines += ["", "2B ROWS (GSTR-2B supplier-reported):"]

    for i, r in enumerate(pile["2b"], 1):
        lines.append(
            f"  G{i}: gstin={r['gstin'] or 'N/A'}  "
            f"invoice={r['invoice_raw'] or 'N/A'}  "
            f"date={str(r['invoice_date'])[:10] if r.get('invoice_date') and str(r['invoice_date']) != 'NaT' else 'N/A'}  "
            f"supplier={r['supplier'] or 'N/A'}  "
            f"taxable={r['taxable']}  total_tax={r['total_tax']}"
        )

    amounts = [r["taxable"] for r in pile["books"]]
    is_collision = len(set(amounts)) == 1 and len(pile["books"]) > 1

    if is_collision:
        lines += [
            "",
            "NOTE: All rows in this pile have identical taxable amounts. "
            "The invoice numbers are in different formats on each side (books vs 2B). "
            "Use invoice DATE as the primary matching signal — pair the books row "
            "whose date is closest to each 2B row's date. "
            "Match sequentially by date if dates are equal or very close. "
            "Only leave a row unmatched if you have a specific reason to believe "
            "it has no counterpart.",
        ]

    lines += [
        "",
        "Call submit_pile_matches.",
        "Every Bi and every Gj must appear exactly once. Use B1..Bn and G1..Gm as the IDs.",
    ]
    return "\n".join(lines)


async def _ai_pile(
    client: AsyncOpenAI,
    pile: dict,
    pass_type: str,
) -> dict | None:
    if not pile["books"] or not pile["2b"]:
        return None

    # build ID → row lookup (B1, B2, …  G1, G2, …)
    b_id_map = {f"B{i}": r for i, r in enumerate(pile["books"], 1)}
    g_id_map = {f"G{i}": r for i, r in enumerate(pile["2b"],    1)}

    try:
        response = await client.chat.completions.create(
            model=AI_MODEL,
            temperature=0,
            tools=[_TOOL],
            tool_choice={"type": "function", "function": {"name": "submit_pile_matches"}},
            messages=[
                {"role": "system", "content": _SYSTEM_PROMPT},
                {"role": "user",   "content": _pile_prompt(pile, pass_type)},
            ],
        )
    except Exception as exc:
        print(f"  [AI error — {pass_type}] {exc}")
        return None

    raw = json.loads(response.choices[0].message.tool_calls[0].function.arguments)

    # per-pile log (spec requirement)
    print(
        f"  [pile log — {pass_type}] "
        f"books={list(b_id_map)} 2b={list(g_id_map)} | "
        f"matches={[{'B': m.get('books_id'), 'G': m.get('b2b_id'), 'conf': m.get('confidence')} for m in raw.get('matches', [])]}"
    )

    # ── validate: correct IDs, no duplicates, money check ───────────────────
    seen_b: set[str] = set()
    seen_g: set[str] = set()
    clean_matches: list[dict] = []

    for m in raw.get("matches", []):
        bid, gid = m.get("books_id", ""), m.get("b2b_id", "")

        # ID must exist in this pile
        if bid not in b_id_map or gid not in g_id_map:
            print(f"    [AI warn] unknown IDs {bid}/{gid} — skipped")
            continue

        # one-to-one: no ID used twice
        if bid in seen_b or gid in seen_g:
            print(f"    [AI warn] duplicate ID {bid} or {gid} — skipped")
            continue

        # money must still hold (guard against hallucinated pairings)
        b_row = pd.Series(b_id_map[bid])
        g_row = pd.Series(g_id_map[gid])
        if not money_match(b_row, g_row):
            print(f"    [AI warn] {bid}/{gid} fails money check — skipped")
            continue

        seen_b.add(bid)
        seen_g.add(gid)
        m["_books_row_id"] = b_id_map[bid]["row_id"]
        m["_b2b_row_id"]   = g_id_map[gid]["row_id"]
        m["_pass"]         = pass_type
        clean_matches.append(m)

    raw["matches"] = clean_matches
    raw["_pass"]   = pass_type
    return raw


async def run_ai_stages(
    q_blue_b:    pd.DataFrame,
    q_blue_g:    pd.DataFrame,
    remaining_b: pd.DataFrame,
    remaining_g: pd.DataFrame,
) -> list[dict]:
    """
    Stage 4A: blue piles — easy question (invoice differs despite identity match).
    Stage 4B: remaining rows — pile builder filters; single-side piles skip AI.

    Both passes run concurrently with asyncio.gather.
    Returns a flat list of validated AI proposals (one entry per proposed pair).
    """
    client = AsyncOpenAI(
        api_key=os.getenv("OPENROUTER_API_KEY"),
        base_url="https://openrouter.ai/api/v1",
    )

    piles_blue,    pre_blue    = build_piles(q_blue_b,    q_blue_g)
    piles_nonblue, pre_nonblue = build_piles(remaining_b, remaining_g)

    n_4a_sent = sum(1 for p in piles_blue    if p["books"] and p["2b"])
    n_4b_sent = sum(1 for p in piles_nonblue if p["books"] and p["2b"])
    print(
        f"  4A piles={len(piles_blue)} (sent={n_4a_sent}) | "
        f"4B piles={len(piles_nonblue)} (sent to AI={n_4b_sent}) | "
        f"size-gate pre-matches={len(pre_blue) + len(pre_nonblue)}"
    )

    tasks = (
        [_ai_pile(client, p, "blue")    for p in piles_blue    if p["books"] and p["2b"]]
        + [_ai_pile(client, p, "nonblue") for p in piles_nonblue if p["books"] and p["2b"]]
    )

    results = await asyncio.gather(*tasks, return_exceptions=True)

    proposals: list[dict] = pre_blue + pre_nonblue   # seed with size-gate pre-pass matches
    for r in results:
        if isinstance(r, Exception):
            print(f"  [AI error] {r}")
        elif r:
            for m in r.get("matches", []):
                proposals.append({
                    "books_id":   m["_books_row_id"],
                    "b2b_id":     m["_b2b_row_id"],
                    "confidence": m.get("confidence", 0.0),
                    "reason":     m.get("reason", ""),
                    "_pass":      m.get("_pass", "unknown"),
                })

    print(f"[Stage 4A/4B] AI proposals: {len(proposals)}")
    return proposals


# ═══════════════════════════════════════════════════════════════════════════════
# STAGE 5 — GREEDY MERGE + 7-SHEET OUTPUT WORKBOOK
# ═══════════════════════════════════════════════════════════════════════════════

def _find_identity_money_mismatch(
    unmatched_b: pd.DataFrame,
    unmatched_g: pd.DataFrame,
) -> pd.DataFrame:
    """
    Among final unmatched rows, find pairs where identity matches on
    invoice_std (GSTIN+invoice OR supplier+invoice) but money does NOT match.

    Uses invoice_std (not raw) so formatting differences don't create phantom
    mismatches. Results are advisory — for the accountant to review.
    Same-invoice-different-amount is the most actionable exception class.
    """
    rows = []
    seen: set[tuple] = set()   # avoid duplicate pairs

    for _, b in unmatched_b.iterrows():
        for _, g in unmatched_g.iterrows():
            if money_match(b, g):
                continue   # money matches → not a mismatch

            std_b = b.get("invoice_std")
            std_g = g.get("invoice_std")
            if not std_b or not std_g:
                continue

            gstin_inv = (
                b["gstin"] and g["gstin"]
                and b["gstin"] == g["gstin"]
                and std_b == std_g
            )
            sup_inv = (
                b["supplier"] and g["supplier"]
                and b["supplier"] == g["supplier"]
                and std_b == std_g
            )
            if not (gstin_inv or sup_inv):
                continue

            key = (b["row_id"], g["row_id"])
            if key in seen:
                continue
            seen.add(key)

            rows.append({
                "books_row_id":   b["row_id"],
                "b2b_row_id":     g["row_id"],
                "gstin":          b.get("gstin"),
                "invoice_books":  b.get("invoice_raw"),
                "invoice_2b":     g.get("invoice_raw"),
                "supplier_books": b.get("supplier"),
                "supplier_2b":    g.get("supplier"),
                "taxable_books":  b.get("taxable"),
                "taxable_2b":     g.get("taxable"),
                "cgst_books":     b.get("cgst"),     "cgst_2b":  g.get("cgst"),
                "sgst_books":     b.get("sgst"),     "sgst_2b":  g.get("sgst"),
                "igst_books":     b.get("igst"),     "igst_2b":  g.get("igst"),
                "taxable_diff":   round(abs(b["taxable"]   - g["taxable"]),   2),
                "total_tax_diff": round(abs(b["total_tax"] - g["total_tax"]), 2),
                "anchor":         "GSTIN+invoice" if gstin_inv else "supplier+invoice",
                "flag":           "IDENTITY_MATCH_MONEY_MISMATCH",
            })

    return pd.DataFrame(rows)


def run_stage5(
    det_matches:   list[Match],
    ai_proposals:  list[dict],
    books_grouped: pd.DataFrame,
    gstr2b:        pd.DataFrame,
    dirty_books:   pd.DataFrame,
    output_path:   Path,
) -> None:
    """
    Greedy one-to-one merge of AI proposals, then write 7-sheet workbook.

    B2C fix: no-GSTIN books rows appear in no_gstin_b2c AND may also be in
    unmatched_books if they weren't matched. To avoid double-listing, we
    exclude them from unmatched_books (they're fully represented in no_gstin_b2c).
    """
    books_idx = books_grouped.set_index("row_id")
    b2b_idx   = gstr2b.set_index("row_id")

    # rows already consumed by deterministic stages
    used_b: set[str] = {m.books_id for m in det_matches}
    used_g: set[str] = {m.b2b_id   for m in det_matches}

    # greedy accept AI proposals, highest confidence first
    ai_proposals_sorted = sorted(
        ai_proposals, key=lambda x: x.get("confidence", 0), reverse=True
    )
    ai_accepted: list[dict] = []
    ai_review:   list[dict] = []

    for p in ai_proposals_sorted:
        bid, gid = p["books_id"], p["b2b_id"]
        if bid in used_b or gid in used_g:
            continue
        conf       = p.get("confidence", 0)
        is_blue    = p.get("_pass") == "blue"
        threshold  = BLUE_AUTO_ACCEPT   if is_blue else NONBLUE_AUTO_ACCEPT
        min_review = BLUE_MIN_REVIEW    if is_blue else NONBLUE_MIN_REVIEW
        if conf >= threshold:
            ai_accepted.append(p)
            used_b.add(bid)
            used_g.add(gid)
        elif conf >= min_review:
            ai_review.append(p)
        # else: confidence too low → leave both rows free → unmatched

    # ── build matched sheet ───────────────────────────────────────────────────
    matched_rows: list[dict] = []

    for m in det_matches:
        b = books_idx.loc[m.books_id]
        g = b2b_idx.loc[m.b2b_id]
        matched_rows.append({
            "match_id":       f"{m.books_id}_{m.b2b_id}",
            "stage":          m.stage,
            "sub":            m.sub,
            "books_row_ids":  str(b.get("member_row_ids", [])),
            "b2b_row_id":     m.b2b_id,
            "gstin":          b.get("gstin"),
            "invoice_books":  b.get("invoice_raw"),
            "invoice_2b":     g.get("invoice_raw"),
            "supplier_books": b.get("supplier"),
            "supplier_2b":    g.get("supplier"),
            "taxable_diff":   m.taxable_diff,
            "total_tax_diff": m.total_tax_diff,
            "match_reason":   m.match_reason,
            "ai_confidence":  "",
            "ai_reason":      "",
            "flag":           "none",
        })

    for p in ai_accepted:
        b = books_idx.loc[p["books_id"]]
        g = b2b_idx.loc[p["b2b_id"]]
        matched_rows.append({
            "match_id":       f"{p['books_id']}_{p['b2b_id']}",
            "stage":          "4",
            "sub":            "4A" if p.get("_pass") == "blue" else "4P" if p.get("_pass") == "size_gate" else "4B",
            "books_row_ids":  str(b.get("member_row_ids", [])),
            "b2b_row_id":     p["b2b_id"],
            "gstin":          b.get("gstin"),
            "invoice_books":  b.get("invoice_raw"),
            "invoice_2b":     g.get("invoice_raw"),
            "supplier_books": b.get("supplier"),
            "supplier_2b":    g.get("supplier"),
            "taxable_diff":   round(abs(b["taxable"] - g["taxable"]), 2),
            "total_tax_diff": round(abs(b["total_tax"] - g["total_tax"]), 2),
            "match_reason":   p.get("reason", ""),
            "ai_confidence":  p.get("confidence", ""),
            "ai_reason":      p.get("reason", ""),
            "flag":           "blue" if p.get("_pass") == "blue" else "none",
        })

    df_matched = pd.DataFrame(matched_rows)

    # ── unmatched rows ────────────────────────────────────────────────────────
    b2c_ids = set(books_grouped.loc[books_grouped["gstin"].isna(), "row_id"])

    unmatched_b_ids = set(books_grouped["row_id"]) - used_b - b2c_ids  # B2C excluded
    df_unmatched_b  = books_grouped[books_grouped["row_id"].isin(unmatched_b_ids)].copy()

    unmatched_g_ids = set(gstr2b["row_id"]) - used_g
    df_unmatched_g  = gstr2b[gstr2b["row_id"].isin(unmatched_g_ids)].copy()

    # ── identity-match-money-mismatch flag ────────────────────────────────────
    df_id_mismatch = _find_identity_money_mismatch(df_unmatched_b, df_unmatched_g)
    if len(df_id_mismatch):
        print(f"  [Flag] identity+invoice match, money differs: {len(df_id_mismatch)} pairs")

    df_blue_review = pd.DataFrame(ai_review) if ai_review else pd.DataFrame()
    df_b2c         = books_grouped[books_grouped["row_id"].isin(b2c_ids - used_b)].copy()

    # ── write workbook ────────────────────────────────────────────────────────
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df_matched.to_excel(    writer, sheet_name="matched",                 index=False)
        df_blue_review.to_excel(writer, sheet_name="blue_review",             index=False)
        df_unmatched_b.to_excel(writer, sheet_name="unmatched_books",         index=False)
        df_unmatched_g.to_excel(writer, sheet_name="unmatched_2b",            index=False)
        df_id_mismatch.to_excel(writer, sheet_name="identity_money_mismatch", index=False)
        df_b2c.to_excel(        writer, sheet_name="no_gstin_b2c",            index=False)
        dirty_books.to_excel(   writer, sheet_name="dirty_input",             index=False)

    print(f"\n[Stage 5] Output → {output_path}")
    print(f"  matched:                  {len(df_matched)}")
    print(f"  blue_review:              {len(df_blue_review)}")
    print(f"  unmatched_books:          {len(df_unmatched_b)}  (B2C excluded — see no_gstin_b2c)")
    print(f"  unmatched_2b:             {len(df_unmatched_g)}")
    print(f"  identity_money_mismatch:  {len(df_id_mismatch)}")
    print(f"  no_gstin_b2c:             {len(df_b2c)}  (matched B2C in matched sheet)")
    print(f"  dirty_input:              {len(dirty_books)}")

    # ── conservation assertion ────────────────────────────────────────────────
    total_b = len(df_matched) + len(df_unmatched_b) + len(df_b2c)
    total_g = len(df_matched) + len(df_unmatched_g)
    assert total_b == len(books_grouped), (
        f"Books conservation failed: {total_b} != {len(books_grouped)}"
    )
    assert total_g == len(gstr2b), (
        f"2B conservation failed: {total_g} != {len(gstr2b)}"
    )
    print("  [OK] conservation check passed")


# ═══════════════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════════════

async def main(
    books_path:  str  = BOOKS_PATH,
    b2b_path:    str  = B2B_PATH,
    output_path: Path = OUTPUT_PATH,
) -> None:
    print("\n── Stage 0: Load & Prepare ──")
    books_clean, books_dirty = load_books(books_path)
    gstr2b                   = load_2b(b2b_path)
    books_grouped            = group_books(books_clean)

    print("\n── Stage 1: Exact Matching ──")
    m1, books_r1, gstr2b_r1 = run_exact_stage(books_grouped, gstr2b, use_std=False)

    print("\n── Stage 2: Standardized Invoice ──")
    m2, books_r2, gstr2b_r2 = run_exact_stage(books_r1, gstr2b_r1, use_std=True)

    print("\n── Stage 3: Identify Blue Rows ──")
    q_blue_b, q_blue_g, q_remaining_b, q_remaining_g = run_stage3(books_r2, gstr2b_r2)

    print("\n── Stage 4A/4B: AI Matching ──")
    ai_proposals = await run_ai_stages(q_blue_b, q_blue_g, q_remaining_b, q_remaining_g)

    print("\n── Stage 5: Merge & Output ──")
    run_stage5(
        det_matches   = m1 + m2,
        ai_proposals  = ai_proposals,
        books_grouped = books_grouped,
        gstr2b        = gstr2b,
        dirty_books   = books_dirty,
        output_path   = Path(output_path),
    )


if __name__ == "__main__":
    asyncio.run(main())