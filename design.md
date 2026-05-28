# GST Reconciliation Engine — Build Spec for Claude Code

Build a transaction-level reconciliation engine comparing two Excel files:
- **Books** (internal purchase/inward register) — line-level, may have several rows per invoice.
- **2B** (GSTR-2B supplier-reported, sheet "B2B") — exactly one row per (GSTIN + invoice).

Goal: for each books invoice find its twin in the 2B and prove they are the same invoice. Do **not** merge into one combined frame. Every row ends as: matched (record step + reason), blue (review), or unmatched (books-only / 2B-only). Output must be auditable.

This file is the source of truth. Where it and the code disagree, this file wins.

---

## THE ONE MATCHING RULE (everything obeys this)

A books invoice and a 2B invoice match only when BOTH hold:
1. **Money matches** — taxable within ₹2, AND total tax (CGST+SGST+IGST) within ₹2.
2. **At least one identity matches** — GSTIN, invoice number, or supplier name.

Money is the hard gate (never matched without it). The stages differ only in how hard we work to line up one identity: exact → standardized → (blue → AI).

**Place-of-supply note:** Books records CGST+SGST for intra-state and IGST for inter-state; the 2B does the same. Where the two sides disagree on the tax head split (e.g. books shows CGST+SGST but 2B shows IGST), `money_match` still passes because we compare `taxable` and `total_tax` only — not each head individually. These rows reach Stage 3 blue review where the head-split discrepancy is visible to the accountant.

Tolerance helper:
```python
def money_match(a, b, tol=2.0):
    return (abs(a.taxable   - b.taxable)   <= tol and
            abs(a.total_tax - b.total_tax) <= tol)   # total_tax = cgst+sgst+igst
```

---

## FLOW DIAGRAM

```
┌─────────────────────────────────────────────────────────────────┐
│ STAGE 0 — PREPARE                                                 │
│  • load books (header=1) and 2B (sheet "B2B", header=[4,5])       │
│  • drop fully-empty trailing rows                                 │
│  • RECOMPUTE money on books from rate columns (don't trust cached)│
│  • clean: trim + UPPER on gstin, invoice/voucher, supplier        │
│  • blank invoice strings "" -> true null                          │
│  • GROUP BOOKS ONLY to invoice grain (see GROUPING). 2B unchanged.│
│  • no-GSTIN rows grouped by supplier+invoice; still enter pool    │
│  • standalone row (no invoice) kept ungrouped                     │
└───────────────────────────────┬───────────────────────────────────┘
                                 │  pool = (grouped books) vs (2B)
                                 ▼
┌─────────────────────────────────────────────────────────────────┐
│ STAGE 1 — EXACT                                                   │
│  candidates: money_match AND (GSTIN match OR supplier match)      │
│  1A: invoice also matches + unique pair → auto-match              │
│  1B: invoice differs + unique loose pair → auto-match (note)      │
│  1C: collision cluster → only invoice-exact unique pairs matched  │
│  Remove matched pairs before Stage 2.                             │
└───────────────────────────────┬───────────────────────────────────┘
                                 │ leftovers
                                 ▼
┌─────────────────────────────────────────────────────────────────┐
│ STAGE 2 — STANDARDIZED INVOICE                                    │
│  identical to Stage 1 but compare invoice_std not invoice_raw.    │
│  invoice_std = strip non-alphanumeric, uppercase, drop leading 0s │
│  Remove matched pairs before Stage 3.                             │
└───────────────────────────────┬───────────────────────────────────┘
                                 │ leftovers
                                 ▼
┌─────────────────────────────────────────────────────────────────┐
│ STAGE 3 — IDENTIFY BLUE ROWS ONLY (no auto-match)                 │
│  BLUE = money_match AND (GSTIN OR supplier) match, invoice DIFFERS│
│  Pull blue books rows + their candidate 2B rows → queue_blue      │
│  Everything else (all remaining books + all remaining 2B) →       │
│  queue_remaining. No filtering here — pile builder handles it.    │
└───────────────┬───────────────────────────────┬───────────────────┘
                │ queue_blue                      │ queue_remaining
                ▼                                 ▼
┌───────────────────────────────┐  ┌──────────────────────────────────┐
│ STAGE 4A — AI (blue pass)      │  │ STAGE 4B — AI (remaining pass)    │
│  Q: "same invoice despite      │  │  ALL remaining books + 2B rows    │
│  different number?"            │  │  go into pile builder.            │
│  auto-accept ≥ 0.90            │  │  Pile builder draws edges only    │
│                                │  │  where money_match. Rows with no  │
│                                │  │  money-match partner → isolated   │
│                                │  │  node → single-side pile → NOT    │
│                                │  │  sent to AI → unmatched sheet.    │
│                                │  │  Only piles with both sides → AI. │
│                                │  │  auto-accept ≥ 0.95               │
│  BOTH passes use PILE method (see PILES) + submit_pile_matches    │
│  tool schema. Output = proposals (pair, confidence, reason).       │
└───────────────┬───────────────────────────────┬───────────────────┘
                └───────────────┬───────────────┘
                                ▼
┌─────────────────────────────────────────────────────────────────┐
│ STAGE 5 — MERGE (one-to-one) + RESIDUAL + FLAG                    │
│  collect AI proposals, sort by confidence desc, accept a pair only │
│  if both rows still free; remove accepted pairs.                  │
│  Everything left → unmatched: books-only / 2B-only.               │
│  Blue-but-unconfirmed → blue_review sheet.                        │
│  SCAN final unmatched: if identity matches (GSTIN+invoice OR      │
│  supplier+invoice) but money differs → flag as                    │
│  IDENTITY_MATCH_MONEY_MISMATCH → separate sheet for accountant.  │
└─────────────────────────────────────────────────────────────────┘
```

---

## STAGE 0 — PREPARE (details)

**Load.** Books: `pd.read_excel(path, header=1)`. 2B: `pd.read_excel(path, sheet_name="B2B", header=[4,5])` (two-row header → MultiIndex; flatten to single strings).

**Canonical schema.** Map both sources into one internal frame with neutral names; keep raw columns alongside for audit. Fields:
`row_id, source ("books"/"2b"), gstin, invoice_raw, invoice_std, supplier, taxable, cgst, sgst, igst, total_tax, invoice_date, member_row_ids[]`

Column mapping (verified against the real files):
- Books: gstin=`GSTIN/UIN`, invoice=`Voucher Number`, supplier=`Supplier`, date=`Accounting Date`. Money: RECOMPUTE from the rate columns (`Taxable @ X%`, etc.), do NOT trust the cached `Taxable Value/CGST/SGST` cells (they are Excel formulas and can come through blank). Validate recomputed total ≈ `Gross Total`; rows that fail → "dirty input" list (skip matching).
- 2B (flattened): gstin=`GSTIN of supplier`, invoice=`Invoice Details::Invoice number`, supplier=`Trade/Legal name`, date=`Invoice Details::Invoice Date`, taxable=`Taxable Value (₹)`, igst=`Tax Amount::Integrated Tax(₹)`, cgst=`Tax Amount::Central Tax(₹)`, sgst=`Tax Amount::State/UT Tax(₹)`.

**Clean.** `str.strip().str.upper()` on gstin, invoice_raw, supplier. Convert `""` invoice to null. Numeric coercion + `.fillna(0)` on money.

**invoice_std.** `re.sub(r'[^A-Z0-9]','', invoice_raw.upper()).lstrip('0')`.

**GROUP BOOKS ONLY** (the answer to the grouping question):
- One key per row, chosen by what the row has:
  - GSTIN present → key = (`gstin`, `invoice_raw`)
  - GSTIN null    → key = (`supplier`, `invoice_raw`)
- It is OR (per-row choice), NOT AND. Each row gets exactly one key → each row in exactly one group → no chaining.
- Aggregate group: sum money fields; keep `member_row_ids` = the original line numbers; first() the identity fields.
- The 2B is already invoice-grain (one row per GSTIN+invoice) → DO NOT group the 2B.
- **Special case:** a row with no GSTIN AND no invoice cannot be safely grouped (supplier-only would merge unrelated invoices). Leave it ungrouped as a standalone row; let it fall through to blue/AI/unmatched. (On the sample file this is exactly 1 row.)

**Side pools:** no-GSTIN rows still participate via supplier identity; dirty-input rows are parked.

Expected after Stage 0 on the sample files: ~520 grouped books rows vs 465 2B rows; 59 multi-line groups (all GSTIN-keyed); 40 no-GSTIN rows; 1 no-GSTIN-no-invoice standalone.

---

## STAGE 1 — EXACT

Candidate condition: `money_match(b, g)` AND [ (b.gstin==g.gstin AND b.invoice_raw==g.invoice_raw) OR (b.supplier==g.supplier AND b.invoice_raw==g.invoice_raw) ]. Null identity fields never count as equal (treat null != null).

**Uniqueness resolution (1A/1B/1C):**
- 1A — full agreement (identity + invoice both exact) → auto-match.
- 1B — identity + money agree, invoice differs, AND this (identity, amount) is UNIQUE on both sides (exactly one books and one 2B at that amount) → auto-match, note "invoice differs (format/voucher)".
- 1C — identity + money agree but (identity, amount) is NOT unique (collision cluster, e.g. a supplier billing the same amount many times) → only invoice-exact pairs are taken; remaining collision rows are NOT sealed on amount → send to Stage 3 blue.

**Multiple candidates / tie-break:** if a books row has several qualifying 2B candidates, rank by: invoice-also-matches > gstin-also-matches > supplier-also-matches > smallest total money diff > closest date > most complete. Take the single best, remove that pair, leave the rest in the pool. Genuine tie (ranking can't separate) → blue, never guess.

Remove every matched pair from both pools before Stage 2.

Expected on sample: Stage 1 matches ~330–365 of 465.

---

## STAGE 2 — STANDARDIZED INVOICE

Identical to Stage 1, but compare `invoice_std` instead of `invoice_raw` (same two conditions, same 1A/1B/1C uniqueness, same tie-break). Catches `INV 001` vs `INV001`, `SPL-892` vs `MSPL-892`-type only if std makes them equal (note: std removes separators but not added letters; those go blue). Remove matched pairs.

---

## STAGE 3 — IDENTIFY BLUE ROWS ONLY (no auto-match)

Only job: find BLUE rows and pull them out.

**BLUE** = money_match AND (gstin equal OR supplier equal), but no invoice match survived Stages 1–2. Strong identity anchor — invoice number is the only disagreement.

Everything else flows into Stage 4B:
- `queue_remaining_books` = all remaining books rows NOT in blue_b.
- `queue_remaining_2b` = all remaining 2B rows NOT in blue_g (blue 2B rows are reserved for Stage 4A only — excluding them prevents the same 2B row from being proposed by both passes).

Do NOT pre-filter or classify non-blue rows further here — the pile builder does that automatically.

Expected on sample: ~12 blue books rows; ~160 remaining books rows → Stage 4B.

---

## PILES (used by both AI passes — the candidate generator)

The AI must never see every-row × every-row. Build disjoint candidate blocks:

1. Take the queue's books rows + the still-unmatched 2B rows.
2. Build a graph: node per row; edge (books_i, 2b_j) iff `money_match`.
3. **Connected components** of this graph = piles. Every row is in exactly one pile (a partition — no row in two piles).
4. For each pile:
   - 0 rows on a side → the present side is unmatched (books-only / 2b-only). No AI.
   - 1×1 → one candidate pair → one small AI confirm (or rule-check).
   - many×many → ONE pile sent as a unit (see AI INPUT).
5. **Size gate:** if a pile exceeds the model's reliable size (~10–15 rows per side), first run a deterministic EXACT-invoice pre-pass inside the pile (pair rows whose invoice_std is equal). **Record each pre-pass pair as a proposal** with confidence=1.0, reason="size-gate exact-invoice pre-pass", sub="4P". Remove those rows from the pile, then send only the ambiguous remainder to the AI. Split further if still too large. `build_piles` returns `(piles, pre_matches)` — callers must seed the proposal list with `pre_matches` before appending AI results.
6. (Optional safety) if a pile's taxable span > ₹2 it formed by chaining — flag/split at the widest internal gap. On the sample, all multi-row piles have span ₹0, so this never fires; keep it as a guard for future files.

Expected on sample: leftovers form ~249 piles — 129 books-only, 74 2B-only, 42 clean 1×1, only 4 many×many (largest 9×9). So ≤46 piles ever touch the AI.

---

## STAGES 4A / 4B — AI MATCHING (match-the-following via tool calling)

Both passes use the SAME machinery; they differ only in entry queue, prompt emphasis, and auto-accept threshold.
- 4A (blue): question is "same invoice despite a different invoice number?" — strong anchor, allow auto-accept at confidence ≥ 0.90.
- 4B (remaining): ALL remaining books + 2B rows fed into the pile builder. Pile builder draws edges only where money_match — books rows with no money-match partner in 2B become isolated nodes (single-side piles) and are skipped by AI, going directly to unmatched. Only piles with rows on both sides are sent to AI. Question is "same invoice?" — auto-accept at confidence ≥ 0.95; most go to blue_review.

**AI input per pile.** Refer to rows by short IDs (B1.. / G1..), not by re-typing values. Tell the model the schema in words (column meanings) — this measurably improves tabular accuracy. State that all rows in the pile already match on money within ₹2, so it must decide identity only. State the one-to-one constraint and demand full accounting.

**Tool schema (reason BEFORE decisions so the model thinks first):**
```json
{
  "name": "submit_pile_matches",
  "parameters": {
    "reasoning":       "string",
    "matches":         [{"books_id":"B1","b2b_id":"G3","reason":"string","confidence":0.0}],
    "unmatched_books": [{"books_id":"B2","reason":"string"}],
    "unmatched_b2b":   [{"b2b_id":"G4","reason":"string"}]
  }
}
```
Prompt requirement: every Bi and every Gj must appear exactly once across `matches`/`unmatched_*`. Each id used at most once.

**Model:** default GPT-5-mini (reasoning model, reliable tool calling); Gemini 3 Flash with thinking budget ON is the fast all-rounder alternative; Claude Haiku 4.5 if rule-adherence is the issue. Verify the chosen model supports thinking + tool calling together. Use temperature 0; log input, full output JSON, confidence, reason for every pile.

**Code still enforces one-to-one** after the tool returns — the schema can request uniqueness but cannot guarantee it. Validate every returned id exists, is used once, and re-check `money_match` on every proposed pair (reject any pair that fails). Optional accuracy booster: run hard piles through two models and auto-accept only on agreement.

---

## STAGE 5 — MERGE + RESIDUAL

1. Collect all proposals: seed with size-gate pre-pass matches (`_pass="size_gate"`, confidence=1.0, sub="4P"), then append 4A + 4B AI proposals (each: books_id, b2b_id, confidence, reason, pass).
2. Boost confidence where both models/passes agree on a pair.
3. Sort proposals by confidence desc.
4. Greedily accept a pair only if neither its books row nor its 2B row is already taken; else skip.
5. Apply three-band confidence logic (thresholds differ by pass):
   - **Blue pass:**    conf ≥ 0.90 → matched | 0.60 ≤ conf < 0.90 → blue_review | conf < 0.60 → unmatched
   - **Nonblue pass:** conf ≥ 0.93 → matched | 0.75 ≤ conf < 0.93 → blue_review | conf < 0.75 → unmatched
   - Nonblue has a higher floor (0.75) because there is no pre-confirmed identity anchor — a human reviewer has nothing solid to work with below that.
6. Remove accepted pairs. Everything left → unmatched, labelled books-only or 2B-only with a reason.

---

## OUTPUT WORKBOOK (sheets)
- `matched` — all auto-matched pairs (stages 1, 2) + AI-accepted, with audit columns.
- `blue_review` — AI proposals below auto-accept threshold; human must decide.
- `unmatched_books` — books invoices with no accepted 2B match.
- `unmatched_2b` — 2B invoices with no accepted books match.
- `identity_money_mismatch` — pairs where GSTIN+invoice OR supplier+invoice match across the final unmatched rows but money does NOT match. Flag for accountant — same invoice reference, different amounts on each side.
- `no_gstin_b2c` — books rows where GSTIN is null (B2C / unregistered purchases).
- `dirty_input` — books rows where recomputed gross total differs from Excel value by more than ₹2.

**Audit columns per matched row:** `match_id, stage, sub (1A/1B/1C/2A/2B/2C/4A/4B/4P), books_row_ids[] (list — a group has several), b2b_row_ids[], gstin, invoice, taxable_diff, total_tax_diff, supplier_score, invoice_score, ai_confidence, ai_reason, flag(none/blue/tie), match_reason`.

Sub-stage values: `1A/2A` = exact identity+invoice; `1B/2B` = loose unique pair; `1C/2C` = collision tie-break; `4A` = AI blue pass; `4B` = AI remaining pass; `4P` = size-gate deterministic pre-pass.

---

## PRIORITY SUMMARY (one line)
Money is always the gate. Then identity: exact → standardized → (blue queue → 4A AI) ∥ (all remaining → 4B pile builder → AI on piles with both sides), then one-to-one greedy merge, then residual. Auto-match only when (identity+amount) is unique; collisions need the invoice number; identity+money agree but invoice differs → blue. After all matching, scan final unmatched for identity-match-money-mismatch pairs and flag separately. Every match removes its pair before the next stage; group books only (GSTIN+invoice else supplier+invoice); never group the 2B.

## BUILD ORDER (suggested)
1. Stage 0 loader + recompute + clean + group; assert expected counts on the sample files.
2. Stage 1 exact with 1A/1B/1C + tie-break; print match count.
3. Stage 2 standardized.
4. Stage 3 split into queues.
5. Pile builder (connected components + size gate + exact-invoice pre-pass).
6. Stages 4A/4B AI (start with a mock matcher that does exact-invoice, swap in the real model).
7. Stage 5 merge + workbook writer.
Validate after each stage against the expected counts noted above before proceeding.
```