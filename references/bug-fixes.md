# Bug Fix Log

All bug fixes applied to `build_recon_from_source.py` and `generate_html_report.py` / `generate_broadridge_report.py` during the reconciliation project.

## Bug Fix 1 — Billing Master filter (Session 2, Feb 23)
- `is_billing_duplicate()` was unconditionally skipping all "Billing Master" files, but May/Jun/Jul 2023 only exist as Billing Master files (no Billing Workbook counterpart)
- Updated `process_all_billing_workbooks()` to only skip master files when a non-master file already covers that month
- Result: +3 billing files processed, +1.5M usage records recovered

## Bug Fix 2 — Client name filter `is_apex()` (Session 2, Feb 23)
- Filter only matched "APEX" in client name, missing Ridge Clearing and Penson Financial envelope purchases
- Apr and Jun 2022 showed zero purchases because those months' POs were under "Pension Financial" and "RIDGE" client names
- Updated `is_apex()` to match APEX, RIDGE, PENSON, PENSION (with BROADRIDGE exclusion to avoid false positives)
- Result: +36 purchase records, +5.4M purchased envelopes recovered, deficit dropped from 5.6M to 901K

## Bug Fix 3 — Jun 2024 billing workbook (Session 2, Feb 23)
- The unlabeled `Apex Billing Sheet.xlsx` in the 2024 folder turned out to be December 2024 data, not June
- User provided the actual file: `Jun-24 Apex Billing Sheet.xlsx` — 212 APEX volume records for Jun 2024
- Result: Jun 2024 gap resolved

## Bug Fix 4 — April 2025 billing sheet names (Session 2 continued, Feb 23)
- `read_billing_workbook()` looked for exact sheet name `"Volume Data"` but `April-25 Billing Sheet.xlsx` had `"Volume "` (trailing space) and `"Postage"` (no "Data" suffix)
- Added flexible sheet name matching: `{sn.strip().lower(): sn for sn in sheet_names}` with fallback candidates `["volume data", "volume"]` and `["postage data", "postage"]`
- Result: +241 volume records, +400 postage records for Apr-25

## Bug Fix 5 — False "missing data" flags in DQ section (Session 2 continued, Feb 23)
- DQ section flagged months with 0 purchases or 0 usage as "missing", but May/Jun 2025 legitimately had no purchases
- Removed `zero_purchase` and `zero_usage` DQ checks entirely — 0 values mean no activity, not missing data
- DQ section now shows "No outstanding gaps" when all months have source data

## Bug Fix 6 — First data row hidden in HTML tables (Session 2 continued, Feb 23)
- `table th { position: sticky; top: 47px; }` combined with `.section-body { overflow: hidden; }` caused the first data row to render behind the sticky header
- Removed `position: sticky; top: 47px;` from table th CSS

## Bug Fix 7 — Annual summary year format (Session 2 continued, Feb 23)
- Year column displayed as "2022.0" instead of "2022" due to Pandas float coercion
- Fixed with `int()` cast: `f'<td>{int(r["Year"])}</td>'`

## Bug Fix 8 — Pre-2022 purchase report column names (Session 2 continued, Feb 23)
- 2020-era files use different column headers than 2022+
- Added fallback chains in `read_standard_purchase()`:
  - `PO Date` -> `Receipt Date` -> `Order Date`
  - `Quantity` -> `Qty`
  - `Quantity Received` -> `Qty Recd.`
  - `Receipt Amount` -> `Total Price`
  - `Mark up %` -> `Markup` -> `Mark % 1`
  - `PO Number` -> `Order ID`

## Bug Fix 9 — Missing Client column in 2020-2021 purchase reports (Session 2 continued, Feb 23)
- Some pre-2022 files have no `Client Name`/`Client` column, and contain purchases for ALL Broadridge clients (Fidelity, Morgan Stanley, UBS, etc.)
- Initial fix (blanket "assume Apex") inflated purchases to 75M — included all clients
- Corrected approach: when no Client column, check row text for "APEX"/"RIDGE" keywords; exclude "BROADRIDGE" false positives
- Jan 2021 legitimately has no Apex items — not a bug

## Bug Fix 10 — Legacy `.xls` format not supported (Session 3, Feb 23)
- `Purchase Report Aug'20(1).xls` was silently skipped because `open_workbook()` only handled `.xlsx`/`.xlsm`
- Added `xlrd`-based reader in `open_workbook()` that converts `.xls` to openpyxl-compatible workbook, including date cell conversion
- Updated file extension filters on both pre-2022 and main purchase report loops to include `.xls`
- Result: +204,000 envelopes recovered for Aug-20

## Bug Fix 11 — Jun-20 folder contains wrong file (Session 3, Feb 23)
- `06-20 Purchase report/` folder contained a duplicate of the May purchase report (`Purchase Report May (1).xls.xlsx`), not the actual June file
- Caused 4 missing June POs (-392K) and 2 duplicate May POs (+70K)
- Added `2019-2022 Purchases.xlsx` as supplementary source to fill the gap (4 Jun-20 POs with PO numbers 691862, 692617, 692749, 692612)

## Bug Fix 12 — Cross-source deduplication failure (Session 3, Feb 23)
- Consolidated file uses bare WMS codes (`ENVCONPFSN10NI`)
- Monthly files use verbose descriptions (`UNITED ENVELOPE -- ENVCONPFSN10NI`)
- Single-pass dedup by `(date, description, qty)` couldn't match these
- Implemented two-pass dedup:
  - Pass 1: `(date, description, qty)` — catches same-source duplicates
  - Pass 2: PO number match — catches cross-source duplicates with different descriptions; prefers non-consolidated record (richer data)
- Result: 109 duplicates correctly identified and removed from 349 raw records

## Bug Fix 13 — UOM "TH" not handled (Session 3, Feb 23)
- Sep-21 purchase report used `UOM = "TH"` (thousands) instead of `"M"`
- Quantities stored as 84 and 180 instead of 84,000 and 180,000
- Updated all UOM checks: `uom in ("M", "TH")` instead of `uom == "M"` (4 locations)

## Bug Fix 14 — Non-Apex envelopes in purchase records (Session 6, Feb 23)
- `2019-2022 Purchases.xlsx` (consolidated file) contains purchases for ALL Broadridge clients, not just Apex
- `read_consolidated_purchases()` checked `is_envelope()` but not `is_apex()`, allowing Fidelity/HSBC/Morgan Stanley envelopes through
- Three non-Apex envelope types identified and excluded via `normalize_envelope_type()`:
  - `ENVFIDN101.9863867.101` — Fidelity, 168,000 (Sep 2021)
  - `ENVHSBCN10EN796A05/14` — HSBC, 20,000 (Oct 2021)
  - `ENVMSTN10M017ECCCDi(5/19)` — Morgan Stanley, 880,000 (Oct 2021)
- Impact: Total purchased dropped from 31,133,500 to 30,065,500

## Bug Fix 15 — YTD file double-counting 2022 volume and postage (Session 7, Feb 23)
- `read_ytd_usage()` read Volume Data and Postage Data tabs from the YTD file, which contained identical records to the Jan-Aug 2022 individual billing workbooks
- Every month from Jan-Aug 2022 was exactly doubled (e.g., Jan-22: 511,349 billing + 511,349 YTD = 1,022,698)
- Total duplication: 3,716,075 volume records + 3,716,075 postage records
- Fix: Removed `read_ytd_usage()` call from `main()` — function is entirely redundant with individual billing workbooks
- Impact: 2022 variance corrected from -42.0% (2.7M deficit) to +14.9% (972K surplus)

## Bug Fix 16 — `fmt_pct` percentage formatting (Session 11, Feb 23)
- Any variance ratio exceeding +/-1.0 (i.e., >100%) displayed as raw ratio instead of percentage
- Example: -148.4% showed as "-1.5%" because `abs(val) < 1` branch multiplied by 100, else branch did not
- Fix: always multiply by 100 (`return f"{val * 100:.1f}%"`)

## Bug Fix 17 — SKU usage mapping not date-aware (Session 11, Feb 23)
- `map_usage_to_envelope_type()` in `build_recon_from_source.py` was static — always mapped domestic #10 confirms/letters to ENVAPXN10 (PFC)
- Before Oct 2022, those envelopes were physically ENVCONPFSN10NI (NI version)
- All pre-Oct 2022 domestic confirm/letter usage (~7.1M) was incorrectly attributed to the PFC SKU
- Fix: added `month_key` parameter; domestic #10 usage before `2022-10` maps to ENVCONPFSN10NI, after maps to ENVAPXN10 (PFC)
- Result: ENVAPXN10 (PFC) corrected from (7.3M) deficit to (182K) at -2.0%; ENVCONPFSN10NI corrected from 9.4M surplus to 2.2M at +23.0%

## Bug Fix 18 — Old format (.xlsm) vendor cost lost (Session 13, Feb 23)
- `read_standard_purchase()` line 325: `total_cost = invoiced if invoiced > 0 else receipt_amt`
- The "Mark up %" column contains the invoiced total in dollars (vendor cost x 1.10), not a percentage
- Code set both `total_cost` and `invoiced_amount` to the invoiced value, losing the vendor cost
- Fix: `total_cost = receipt_amt` (vendor cost from "Receipt Amount" column)
- Effect: Jan 2024 - Sep 2025 POs now correctly show 10% markup instead of 0%

## Bug Fix 19 — New format (.xlsx) double-counting vendor cost (Session 13, Feb 23)
- `read_new_format_purchase()` line 396: `invoiced = total_cost + markup_total`
- The "Markup Total" column contains the invoiced total (vendor cost x 1.10), not just the markup delta
- Code added vendor cost + invoiced total = 2.1x vendor cost -> appeared as 110% markup
- Fix: `invoiced = markup_total` (already the full invoiced amount)
- Effect: Oct-Dec 2025 POs now correctly show 10% markup instead of 110%

## Bug Fix 20 — Triple-counted May 2025 billing (Session 13 part 3, Feb 23)
- Three versions of the May 2025 billing file existed with different filenames
- No dedup logic existed for multiple non-master files covering the same billing month
- Added data-aware dedup: `peek_billing_month()` reads the first APEX record's Billing_Month/Billing_Year, groups files by actual billing month, keeps only the last file alphabetically per month
- Result: May-25 usage corrected from 752,406 to 250,802

## Bug Fix 21 — Nov 2023 data incorrectly skipped (Session 13 part 3, Feb 23)
- Initial filename-based dedup grouped `Apex Billing November.xlsx` (Nov 2023) with `Apex November Billing.xlsx` (Nov 2024) because both filenames contain "november"
- Data-aware dedup correctly identifies them as different months and processes both
- Result: Nov-23 usage recovered (326,113 envelopes, 396K purchased)

## Bug Fix 22 — Stale May-25 footnote (Session 14, Feb 23)
- Internal report still referenced "May-25 usage of 752K is 2.3x the trailing average" — the old triple-counted number from before Bug Fix 20
- Removed the stale investigation language; footnote now reads: "May & Jun 2025: Zero purchases confirmed (not missing data)."

## Bug Fix 23 — Wastage rounding drift, 110 envelopes (Session 19, Feb 24)
- Per-SKU-per-month `int()` truncation across ~400 rows lost 110 envelopes vs per-month aggregate
- Fix: accumulate as floats, `int(round())` at final aggregate, adjust largest SKU to foot exactly

## Bug Fix 24 — By Type Used off by 1 (Session 19, Feb 24)
- Tax Form row (0 purchased, 1 used) was hidden by `u <= 1` filter
- Fix: changed to `u == 0` — row now visible in grouped table
