# Session Log

Complete session-by-session history of the Broadridge Envelope Reconciliation project.

## Session 1 - 2026-02-20

**Accomplished:**
- Built full reconciliation from individual Broadridge source reports (Purchase Reports + Billing Workbooks)
- Processed 37 purchase report files and 45 billing workbook files across 4 different format eras (2022-2025)
- Created 7-tab Excel reconciliation workbook from source data
- Performed contract compliance audit against both the original 2019 contract and 2024 amendment
- Identified product-level anomalies in envelope usage (Apr 2023 spike, Monthly Statements spoilage, ghost products)
- Generated executive-ready HTML report with Apex brand design system

**Key Findings:** Purchased 17,150,500 / Used 21,307,968 / Deficit (4,157,468). Data quality issues identified (2022 purchases incomplete, markup inconsistency).

## Session 2 - 2026-02-23

**Accomplished:**
- Searched Outlook for Koebel/Denci emails (130 found)
- Bug Fixes 1-3: Billing Master filter, client name filter (`is_apex()`), Jun 2024 billing workbook
- Narrowed reconciliation period to March 2022 - December 2025
- Added Running Balance column to monthly detail
- Resolved all remaining data gaps (Apr/Jun/Jul 2025 files)
- Bug Fixes 4-9: April 2025 billing sheet names, false DQ flags, hidden first row, year format, pre-2022 column names, missing Client column
- Expanded reconciliation period back to January 2020 (pre-2022 purchase reports + Billing Master 2020-2021)
- Added inventory reconciliation bridge to Executive Summary

**Deficit progression:** (4,157,468) -> (5,615,384) -> (240,384) -> (1,702,384) -> (2,031,693) -> (901,687) -> (539,156)

## Session 3 - 2026-02-23

**Accomplished:**
- Cross-validated against `2019-2022 Purchases.xlsx` (2022: exact match at 6,707,000)
- Bug Fixes 10-13: .xls format support, Jun-20 wrong file, cross-source dedup, UOM "TH"
- Added consolidated file as supplementary purchase source
- Built two-pass purchase deduplication
- Removed commodities/inventory from report; added "Usage by Product" section

**Variance progression:** (539,156) -> +202,412 (+0.7%)

## Session 4 - 2026-02-23

**Accomplished:**
- Full review of 58 Brandon Koebel emails (Aug 2023 - Feb 2024)
- Explained negative running balance periods: pre-existing buffer stock, "usage" != mailed, 10-15% wastage inflation
- Compiled envelope type mapping from Sep 12, 2023 email
- Compiled operational facts (supplier, delivery, reorder points, hand insertion cutoff, etc.)

## Session 5 - 2026-02-23

**Accomplished:**
- Reviewed 7 envelope sample PDFs, mapped to physical specs
- Added "Envelope Specifications" section to HTML report with WMS code, mail type, size, style, postage type

## Session 6 - 2026-02-23

**Accomplished:**
- Consolidated "By Envelope Type" from 35 raw variants to 9 canonical types via `normalize_envelope_type()`
- Bug Fix 14: Non-Apex envelopes (Fidelity 168K, HSBC 20K, Morgan Stanley 880K) excluded
- Full inventory efficiency audit
- Total purchased corrected from 31,133,500 to 30,065,500; variance from +202,412 to (865,588)

## Session 7 - 2026-02-23

**Accomplished:**
- Bug Fix 15: YTD file double-counting 2022 (3,716,075 duplicated records)
- Removed Data Quality section, Contract Audit section, all cost/invoiced columns from HTML report
- 2022 variance corrected from -42.0% to +14.9%. Overall variance: +2,850,487 (+9.5%)

## Session 8 - 2026-02-23

**Accomplished:**
- Consolidated Annual Summary and Monthly Detail into single table with subtotal rows
- Merged "Purchases by Envelope Type" and "Usage by Product" into side-by-side section

## Session 9 - 2026-02-23

**Accomplished:**
- Added usage by envelope type mapping (`map_usage_to_envelope_type()` based on Koebel's Sep 2023 email)
- Located pass-through paper dispute settlement ($643,457.92) in GTO Proxy & BPS Term Sheet
- Added post-settlement analysis (Mar 2022 - Dec 2025): 21,039,500 purchased / 18,469,949 used / +2,569,551 (+12.2%)
- Implied inventory = 8.7 months buffer stock (Broadridge policy: 2-3 months)
- Usage declined 39% from 462,992/mo (2022) to 284,200/mo (2025)

## Session 10 - 2026-02-23

**Accomplished:**
- CEO-lens review and full report restructure for executive readability
- Bottom Line callout, post-settlement KPIs promoted, full-period demoted
- Added inventory gauge, SKU recon table, usage trend sparkline, usage-by-product table

## Session 11 - 2026-02-23

**Accomplished:**
- Completed HTML template wiring from session 10
- Bug Fixes 16-17: `fmt_pct` percentage formatting, date-aware SKU usage mapping
- Report scoped to post-settlement (Mar 2022 - Dec 2025)
- Removed Mailed and Spoils columns from monthly detail
- Corrected SKU-level recon (NI/PFC split fixed)

## Session 12 - 2026-02-23

**Accomplished:**
- Comprehensive Outlook search (7 keyword/sender searches)
- 9 new email threads analyzed (beyond the 130 Koebel/Denci emails)
- Key findings: wastage confirmed at 5% (Denci), materials separate from annual fee (Jobanputra), interim extension Dec 2023, PostEdge limitations, ICS credits $63K

## Session 13 - 2026-02-23

**Accomplished:**
- Bug Fixes 18-19: Old format vendor cost lost, new format double-counting vendor cost. Root cause of "31.8%" markup: entirely a data parsing artifact. Actual markup = 10% (favors Apex).
- Investigated ENVCONRIDGE9X12DW deficit: 2022 production surge, not actionable
- Bug Fixes 20-21: Triple-counted May 2025 billing, Nov 2023 data recovered from misfiled workbook
- Corrected totals: 30,065,500 purchased / 26,373,417 used / +3,692,083 (+12.3%)

## Session 14 - 2026-02-23

**Accomplished:**
- Bug Fix 22: Stale May-25 footnote
- Created `generate_broadridge_report.py` for external-facing report
- Fixed variance mismatch between sections (By Type was full-period, Summary was post-settlement)
- Added monthly granularity to By Envelope Type and Usage by Envelope Type Excel tabs

## Session 15 - 2026-02-24

**Accomplished:**
- Converted both reports to dark theme with print-safe overrides
- Added cost analysis (Total Invoiced KPI, unit cost columns, 2026 projection)
- Added structured recommendations (3 action items with dollar impact)
- Added wastage discrepancy callout and email quotes to Reference section
- Removed ~200 lines of dead code

## Session 16 - 2026-02-24

**Accomplished:**
- Billing basis discrepancy analysis: contract says usage-based, actual is receipt-based ($225,870 excess — updated from $192K with per-month contract formula)
- Aligned Broadridge report with internal (wastage-adjusted variance)
- Security review of Broadridge report (5 items identified)
- Removed 4 of 5 security-reviewed risk items per user decision
- Added Denci/Koebel email quotes to wastage section
- Removed Envelope Specifications from Broadridge report
- Added Wastage and Adj. Variance columns to By Type and By SKU tables
- Removed Monthly Trend section and client filter from Broadridge report

## Session 17 - 2026-02-24

**Accomplished:**
- Added Wastage % column to all tables in Broadridge report
- Spot-checked 3 months end-to-end: all matched exactly (Oct 2022, Jun 2023, Mar 2025)
- Generic stock classification added to both reports
- Restructured Broadridge report (Summary / Items for review / By Type / Reference)
- Removed Monthly Detail and Tax Form row from Broadridge report
- Report size: 54.1 KB -> 33.3 KB

**Confidence assessment:** Overall totals ~95%, per-SKU ~75%, pre-2022 medium risk.

## Session 18 - 2026-02-24

**Accomplished:**
- Converted Broadridge report to print-ready light theme for PDF
- Replaced Var % with Buffer (Mo.) using trailing 12-month average
- Added confirmation items for Broadridge (wastage rate, inventory position)
- Added excess inventory observation with buffer months
- Reviewed contract audit provisions (MSA Section 23.S, Amendment Section 4)
- Added 39% usage decline quantification and "Used" definition

## Session 19 - 2026-02-24

**Accomplished:**
- Full numerical consistency audit: resolved wastage rounding drift (Bug Fix 23), By Type Used off by 1 (Bug Fix 24)
- Removed Buffer (Mo.) from year-by-year table
- Reduced confirmation items from 4 to 2
- Fixed em dashes, HTML entities, text wrapping
- Cross-report audit: aligned internal with Broadridge (trailing 12-month, wastage rounding, Tax Form row)
- Labeled wastage box rows (a), (b), (c) with explicit formula
- Added print page-break CSS rules

**Final audit (all pass):** Purchased 21,039,500 / Used 18,469,949 / Wastage 690,713 / Variance 1,878,838 - consistent across all 5 sections (KPI, Year-by-Year, Wastage Box, By Type, By SKU).

## Variance Progression (Full Project)

| Stage | Variance | Session |
|-------|----------|---------|
| Initial build | (4,157,468) | 1 |
| After client filter + billing master fixes | (901,687) | 2 |
| After expanding to Jan 2020 | (539,156) | 2 |
| After cross-validation + dedup | +202,412 (+0.7%) | 3 |
| After non-Apex exclusion | (865,588) (-2.9%) | 6 |
| After YTD double-counting fix | +2,850,487 (+9.5%) | 7 |
| After triple-count + Nov-23 recovery | +3,692,083 (+12.3%) | 13 |
| Post-settlement scope (final) | +1,878,838 (+8.9%) wastage-adjusted | 19 |

## Session 26 - 2026-04-28

**Accomplished — Consolidated to 2 files, fixed CPI finding with vendor price increase data:**

1. **Identified missed files from Session 25** — `Print & Mail Billing Rectification Summary - Internal.html` and `Print & Mail Billing Review - For Broadridge Review.html` were NOT updated in Session 25. Still showed F2=$168,816, total=$193,949, no 2026 row.
2. **Fixed all stale values** — Updated both missed files: F2 $168,816→$193,960, total $193,949→$219,093 (later $219,566), added 2026 row ($108,264/$83,120/$25,144), excess 2,569,551→2,673,125.
3. **Standardized titles** — All documents now use "Print & Mail Billing Rectification Summary" and "Print & Mail Billing Review" (proper capitalization).
4. **Deleted 3 redundant files** — Kept only 2: `Print & Mail Billing Rectification Summary - Internal.html` and `Print & Mail Billing Review - For Broadridge Review.html`. Removed: `Rectification Summary - Internal.html`, `Rectification Summary - For Broadridge Review.html`, `Print & Mail Billing Review - For Broadridge.html`.
5. **CPI Finding 3 recalculated from D17 invoices** — Prior $4,256 was only through Feb 2026 (14 months). Verified 15-month total from D17 invoices: $4,729. Pages corrected: 12,524,445→12,475,407.
6. **Vendor price increase incorporated** — Feb 11, 2026 Broadridge notification: Pre-print Forms +5% (Mar 1), Envelopes +4% (Mar 1), Mill Direct Rolls +3% (Feb 25). Added callout in Billing Review showing expected rate with increase ($0.0114) vs charged ($0.0115). Proves $0.0115 was CPI-driven (established Jan 2026 before notification).
7. **Final combined total: $219,566** — F1 $20,877 + F2 $193,960 + F3 $4,729.

**Files modified:**
- `Print & Mail Billing Rectification Summary - Internal.html` — all values updated
- `Print & Mail Billing Review - For Broadridge Review.html` — all values updated, CPI table restructured with vendor callout

**Files deleted:**
- `Print & Mail Billing Review - For Broadridge.html`
- `Rectification Summary - For Broadridge Review.html`
- `Rectification Summary - Internal.html`

## Session 27 - 2026-04-28

**Accomplished — COO approved usage-based switch, recovery basis email sent:**

1. **Fritz (COO) agreed to flip to usage-based billing** — Asked whether Apex has standing to recover the historical overage.
2. **Assessed recovery basis** — Confirmed no assumptions: all claims traced to documented source material (Denci emails, contract PDFs, month-by-month invoice data).
3. **Drafted internal email to Fritz and Terry Ray** — Summarized recovery basis using two pillars: (a) Denci's five shifting answers and inability to defend classification, (b) contract language clearly defining generic vs client-specific by product characteristics, not operational handling.
4. **Corrected figures** — Initially used older $225,870 figure; updated to current $219,566 total ($220K) from Session 26 rectification documents (F1+F2+F3).
5. **Email sent** — To Fritz Louis-Charles and Terry Ray, Re: Envelope Recon. Also noted broader pattern: ICS overcharging, PostEdge annual credit chasing, Prospectus erroneous charges. Fritz now has the basis to authorize formal recovery.

**No code/file changes this session — strategy and communication only.**

## Session 28 - 2026-07-29

**Accomplished — Analyzed Broadridge's proposed paper rate increase; drafted reply requesting vendor invoice:**

1. **Contract question answered** — Broadridge can pass through actual vendor material cost without advance notice (materials are pass-through, no notice clause in Amendment No. 1; only the Postage Advance Account requires pre-wire notice). Verification is via monthly invoices (Amendment §4) and audit rights (MSA §23.S).
2. **Category confirmed** — Apex paper stock = **Mill Direct Rolls – Blank Stock**, the **+3%** line on the Feb 11, 2026 Broadridge vendor notice (effective Feb 25, 2026). Notice also listed Envelopes +4% (Mar 1) and Pre-print Forms +5% (Mar 1); the +5% Pre-print line does NOT apply to our stock.
3. **Decomposed Broadridge's rates** — Proposed increase from $0.011215 → $0.011495. Established $0.011215 is NOT vendor cost: it is $0.010900 (vendor cost) × 1.0289 (2025 CPI) — CPI improperly baked into a material. The $0.011495 = $0.011215 × ~1.025, stacking a second bump on the CPI-inflated base, ≈ **+5.5% over the $0.010900 cost** vs the 3% on the notice ($0.010900 ÷ into $0.011495 = 1.0546).
4. **Correct rate** — 3% Blank Stock increase applied straight to vendor cost = $0.010900 × 1.03 = **$0.011227**. Overcharge ≈ $0.000268/page (~$2,700/yr, unverified estimate).
5. **Drafted reply email** — Collaborative tone, requests supporting vendor invoice to confirm the $0.010900 cost base; shows the 3-row rate trace (vendor cost / current / proposed) without exposing the CPI-on-materials argument. Held CPI-on-materials card back per strategy.

**Outstanding before send:** Verify the $0.010900 base against a recent D17 invoice so the cost anchor is solid.

**No code/file changes this session — analysis and communication only.**

## Session 29 - 2026-07-29

**Accomplished — Extracted full D17 paper rate history, confirmed 21# grade, corrected Session 28 cost base, redrafted reply:**

1. **Built paper rate history from source invoices** — New script `extract_paper_rates.py` reads every `Mon_C10_D17.pdf` (Jan 2023 - Mar 2026), parses the "ICS FORMS/PAPER CHARGES" section, and captures the per-page rate. All three paper lines (CONFIRM / LETTER / STATEMENT) bill at the same rate throughout. Verified two ways: displayed rate and amount ÷ quantity (invoice prints only 4 decimals, so the exact rate carries more precision).

   | Effective | Rate ($/page) | Change | Notes |
   |-----------|---------------|--------|-------|
   | Jan 2023 | $0.009400 | baseline | |
   | May 2023 | $0.009900 | +5.3% | = 21# vendor cost (see #3) |
   | Jan 2024 | $0.010900 | +10.1% | = 2023 cost + 10% Amendment margin |
   | Jan 2025 | $0.011215 | +2.89% | tracks CPI, no vendor notice |
   | Jan 2026 | $0.011519 | +2.71% | tracks CPI, predates Feb notice |

2. **CORRECTION to Session 28** — The $0.010900 that Session 28 called "vendor cost" is actually the **2024 billed rate**, not a vendor cost. The Session 28 "correct rate" of $0.011227 (= $0.010900 × 1.03) was therefore built on a wrong base and is void. The real 2023 vendor cost is **$0.0099** (21#, per #3). The CPI decomposition ($0.011215 = $0.0109 × 1.0289) was circular (treated a billed rate as a cost) and is dropped as a lead argument.

3. **Grade confirmed: 21#** — Broadridge's own 2023 procurement email (Miguel Mendoza → Joann Rogers) prices base stock at **21# Blank $0.0099** and 24# Blank $0.0112. Our May-Dec 2023 billed rate was **exactly $0.0099** = the 21# vendor cost that year. This proves (a) our statements run on basic 21# stock, and (b) paper was billed at straight vendor cost in 2023 (no margin, pre-Amendment). Screenshot confirmed to be from **2023**, so it proves grade, not current cost.

4. **Increase walk** — 2024 step to $0.0109 = $0.0099 × 1.10 (Amendment margin), contractual. The 2025 and 2026 steps (+2.9%, +2.7%) have **no vendor cost notice behind them** — the first vendor increase notice was Feb 11, 2026, after both steps were already in place — and track CPI, which the Amendment excludes for materials. The Jan 2026 rate ($0.011519) predates the Feb notice AND exceeds the $0.011495 Broadridge proposed in it.

5. **Redrafted reply email** — Asserts the 21# grade using Broadridge's own 2023 cost match (no longer a guess), isolates the 2024 margin step as legitimate, questions only the 2025/2026 escalation, and keeps a single ask: send the current vendor invoice so the rate can be priced off actual cost. CPI-on-materials point now safe to include because it is anchored to notice timing + contract language, not the void $0.0109-as-cost math. User will attach the Feb notice + the 2023 procurement screenshot as support.

**Files created:**
- `extract_paper_rates.py` — reads all D17 invoices, extracts/parses paper per-page rates, flags change points

**Files removed:**
- `dump_one.py`, `temp_page1.png` — throwaway debug artifacts
