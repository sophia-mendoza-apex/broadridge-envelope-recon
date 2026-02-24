# Broadridge Envelope Reconciliation Project

## Project Overview

Reconciliation of Apex Clearing's envelope purchases vs. usage from Broadridge Print & Mail Services, covering January 2020 through December 2025.

## Key Files

### Outputs (deliverables)
| File | Description |
|------|-------------|
| `Envelope Reconciliation Report.html` | **Primary deliverable** - Executive-ready HTML dashboard with Apex branding, sortable tables, SVG charts, contract audit. Self-contained, no external dependencies. |
| `Envelope Reconciliation - Source Data.xlsx` | 7-tab Excel workbook built from individual Broadridge source reports (Monthly Summary, Annual Summary, By Envelope Type, Purchase Detail, Contract Audit, Current Inventory, Data Quality) |
| `Envelope Reconciliation Mar2022-Current.xlsx` | Earlier version built from the summary workbook `P&M Postage and Material Recon.xlsx` (superseded by Source Data version) |

### Scripts (rerunnable)
| File | Description |
|------|-------------|
| `build_recon_from_source.py` | Reads ~60 monthly Purchase Reports (2020-2025) + ~53 Billing Workbooks/Masters from Broadridge source directory, consolidates into `Envelope Reconciliation - Source Data.xlsx` |
| `generate_html_report.py` | Reads the Source Data xlsx and generates the HTML report |
| `build_envelope_recon.py` | Earlier script that built recon from `P&M Postage and Material Recon.xlsx` (superseded) |
| `audit_script.py` | Contract compliance audit script |

### Source Data (read-only, do not modify)
| File | Location |
|------|----------|
| `P&M Postage and Material Recon.xlsx` | Local project directory - summary workbook with 7 sheets (Envelopes Purchased, Purchase Orders, Volume Data, Postage Data, Envelope Inventory, Sheet1, Sheet2) |
| Purchase Reports (60+ files) | `...\Print & Mail Reports\Envelope Purchase Orders\` (2022-2025) and `...\Purchase Reports (from email)\Purchase Reports\FY'20\` / `FY'21\` (2020-2021) |
| Billing Workbooks (50+ files) | `...\Print & Mail Reports\Postage and Volume Report\{2022,2023,2024,2025}\` |
| Billing Master 2020-2021 | `Billing Master - 2020 - 2021.xlsx` in project directory — 27 sheets covering all 2020-2021 volume/postage data |
| Commodities Item List | `...\Print & Mail Reports\Commodities_Item_List Jan 2026.xlsx` |

### Contracts
| File | Description |
|------|-------------|
| `GTO Print and Mail Services Schedule_Effective Jan-2019.pdf` | Original contract - 5% envelope wastage, $475K annual fee |
| `GTO Print and Mail Services Schedule Amendment No.1_Effective Jan-2024.pdf` | Amendment - 2% envelope wastage + 10% margin, term extended to Dec 2028 |

### Temp/Working Files
All `_*.py` files (e.g., `_build_final.py`, `_part2.py`, etc.) are intermediate build artifacts from the HTML generation process. Safe to delete.

## Session Log

### 2026-02-20

**Accomplished:**
- Built full reconciliation from individual Broadridge source reports (Purchase Reports + Billing Workbooks)
- Processed 37 purchase report files and 45 billing workbook files across 4 different format eras (2022-2025)
- Created 7-tab Excel reconciliation workbook from source data
- Performed contract compliance audit against both the original 2019 contract and 2024 amendment
- Identified product-level anomalies in envelope usage (Apr 2023 spike, Monthly Statements spoilage, ghost products)
- Generated executive-ready HTML report with Apex brand design system

**Key Findings:**
| Metric | Value |
|--------|-------|
| Total Purchased | 17,150,500 |
| Total Used (Volume) | 21,307,968 |
| Net Variance | (4,157,468) deficit |
| Total Invoiced | $1,394,618 |
| Contract Audit Overcharge | $72,490 (61 POs over, 51 under) |
| Current Inventory | 409,437 |

**Data Quality Issues (resolved Feb 23):**
- ~~2022 purchase data appears incomplete~~ — Resolved: client filter was too narrow, missing Ridge/Penson purchases
- Markup inconsistency: 2022-2024 invoiced amounts = purchase cost (no visible markup), 2025 shows markup

**Contract Audit Findings:**
- Pre-2024: Original contract specifies 5% wastage for envelopes, but Broadridge may have billed 10%
- Post-2024: Amendment specifies vendor price + 2% wastage + 10% margin
- Net overcharge vs. contract terms: ~$72,490
- Need vendor invoices to validate whether Receipt Amount includes or excludes wastage

**Next Steps:**
- [ ] Request missing Apr-Jul 2025 purchase reports from Broadridge (not in Outlook, not on disk)
- [ ] Request missing Apr/Jun/Jul 2025 billing workbooks from Broadridge
- [ ] Obtain 3-5 vendor invoices to validate Receipt Amount composition (wastage embedded or separate)
- [ ] Clarify markup structure discrepancy between 2022-2024 (no visible markup) and 2025 (10% markup)
- [ ] Once gaps are filled, rerun `build_recon_from_source.py` then `generate_html_report.py` to refresh outputs
- [ ] Consider drafting formal letter to Broadridge re: contract variance
- [ ] Clean up temp `_*.py` files from project directory

### 2026-02-23

**Accomplished:**
- Searched Outlook for emails from Brandon.Koebel@broadridge.com and Christopher.Denci@broadridge.com (130 emails found)
- Extracted `Purchase Reports.zip` and `Billing Master - 2020 - 2021.xlsx` attachments from Oct 2023 email — zip contains FY2019-2021 purchase reports only (no 2022 data)
- Investigated discrepancy between "missing files" narrative and actual files on disk

**Bug Fix 1 — Billing Master filter:**
- `is_billing_duplicate()` was unconditionally skipping all "Billing Master" files, but May/Jun/Jul 2023 only exist as Billing Master files (no Billing Workbook counterpart)
- Updated `process_all_billing_workbooks()` to only skip master files when a non-master file already covers that month
- Result: +3 billing files processed, +1.5M usage records recovered

**Bug Fix 2 — Client name filter (`is_apex()`):**
- Filter only matched "APEX" in client name, missing Ridge Clearing and Penson Financial envelope purchases
- Apr and Jun 2022 showed zero purchases because those months' POs were under "Pension Financial" and "RIDGE" client names
- Updated `is_apex()` to match APEX, RIDGE, PENSON, PENSION (with BROADRIDGE exclusion to avoid false positives)
- Result: +36 purchase records, +5.4M purchased envelopes recovered, deficit dropped from 5.6M to 901K

**Bug Fix 3 — Jun 2024 billing workbook:**
- The unlabeled `Apex Billing Sheet.xlsx` in the 2024 folder turned out to be December 2024 data, not June
- User provided the actual file: `Jun-24 Apex Billing Sheet.xlsx` — 212 APEX volume records for Jun 2024
- Result: Jun 2024 gap resolved

**Reconciliation period narrowed to March 2022 – December 2025:**
- Added `START_MONTH_KEY = "2022-03"` filter to `build_recon_from_source.py`
- Filters applied to monthly summary, purchase detail, contract audit, and console summary

**HTML report improvements:**
- Added **Running Balance** column to monthly detail table (cumulative purchased - used)
- Cleaned up Data Quality section: distinguishes missing purchase reports, missing billing workbooks, and months missing both; each row has "Request from Broadridge" action tag
- Months entirely absent from data (Apr/Jun/Jul 2025) now explicitly flagged
- Updated header to "March 2022 – December 2025" with dynamic generated date

**Current Key Findings (as of end of session):**
| Metric | Value |
|--------|-------|
| Reconciliation Period | Mar 2022 – Dec 2025 |
| Purchase files processed | 37 |
| Billing files processed | 49 |
| Total Purchased | 20,343,500 |
| Total Used (Volume) | 21,245,187 |
| Net Variance | (901,687) = -4.4% |
| Total Invoiced | $1,613,406 |
| Current Inventory | 409,437 |

**Progression of deficit through session:**
| Stage | Deficit | Cause |
|-------|---------|-------|
| Start of session (Feb 20 baseline) | (4,157,468) | APEX-only filter, missing billing masters |
| After Billing Master fix | (5,615,384) | Added May/Jun/Jul 2023 usage (+1.5M) |
| After client filter fix | (240,384) | Added Ridge/Penson purchases (+5.4M) |
| After Broadridge exclusion | (1,702,384) | Removed 8 false-positive Broadridge rows |
| After Jun 2024 billing fix | (2,031,693) | Added Jun 2024 usage (+329K) |
| After Mar 2022 start date | (901,687) | Dropped Jan/Feb 2022 from scope |

**Remaining data gaps (resolved later in session — see continuation below):**
| Gap | Type | Resolution |
|-----|------|------------|
| ~~May-25~~ | ~~Missing purchase report~~ | No purchases made in May 2025 (confirmed by user) |
| ~~Apr-25~~ | ~~Missing purchase report + billing workbook~~ | Files provided by user; billing sheet name fix required |
| ~~Jun-25~~ | ~~Missing purchase report + billing workbook~~ | Files provided by user; no purchases in Jun 2025 (confirmed) |
| ~~Jul-25~~ | ~~Missing purchase report + billing workbook~~ | Files provided by user |

**Client names matched in reconciliation:**
- Purchase side: APEX, Apex, APEX/RIDGE, RIDGE/APEX, RIDGE, Ridge Clearing, Pension Financial, Pension financial, APEX BCC006392
- Usage side: APEX CLEARING (all 13,762 records)

**Outlook search summary:**
- 130 emails found from Brandon Koebel and Christopher Denci
- No purchase reports or billing workbooks for 2025 found in Outlook from any sender
- `Purchase Reports.zip` from Oct 2023 contains FY2019-2021 only
- Most recent billing attachments from these senders: Jan 2024 (December 2023 billing)

**How to Refresh Outputs:**
```bash
# Step 1: Rebuild Excel from source reports
py -3 "C:\Users\smendoza\Projects\Broadridge Envelopes\build_recon_from_source.py"

# Step 2: Regenerate HTML from Excel
py -3 "C:\Users\smendoza\Projects\Broadridge Envelopes\generate_html_report.py"
```

**Purchase Report Format Eras (for future reference):**
| Era | Years | Format | Key Sheet | Notes |
|-----|-------|--------|-----------|-------|
| 0 | 2020-2021 | .xlsx | Month-named tab | Nested FY dirs, no Client column in some files, `Order Date`/`Qty`/`Total Price` headers |
| 1 | 2022 | .xlsx | Month-named tab | Simple flat table, `Markup%` is dollar amount |
| 2 | 2023 | .xlsm | `Final Data` | Added `Owner`, `Mark % 1` columns, 3-4 sheets |
| 3 | 2024 | .xlsm | `Final Data` | Same as Era 2, PO prefix changed to SPSPO |
| 4 | 2025 | .xlsx/.xlsm | `Purchase Report` or `Final Data` | Restructured columns, text quantities with commas, `Unit Cost`/`Total Cost` replace `Unit Price`/`Receipt Amount` |

### 2026-02-23 (continued)

**Accomplished:**
- Resolved all remaining data gaps — user provided missing Apr/Jun/Jul 2025 files
- Fixed April 2025 billing extraction failure (sheet name mismatch)
- Confirmed May/Jun 2025 had no purchases (not a data gap)
- Fixed HTML report CSS and formatting issues
- Added inventory reconciliation bridge to Executive Summary
- Expanded reconciliation period from March 2022 back to January 2020
- Incorporated pre-2022 purchase reports (FY'20/FY'21) and Billing Master 2020-2021

**Bug Fix 4 — April 2025 billing sheet names:**
- `read_billing_workbook()` looked for exact sheet name `"Volume Data"` but `April-25 Billing Sheet.xlsx` had `"Volume "` (trailing space) and `"Postage"` (no "Data" suffix)
- Added flexible sheet name matching: `{sn.strip().lower(): sn for sn in sheet_names}` with fallback candidates `["volume data", "volume"]` and `["postage data", "postage"]`
- Result: +241 volume records, +400 postage records for Apr-25

**Bug Fix 5 — False "missing data" flags in DQ section:**
- DQ section flagged months with 0 purchases or 0 usage as "missing", but May/Jun 2025 legitimately had no purchases
- Removed `zero_purchase` and `zero_usage` DQ checks entirely — 0 values mean no activity, not missing data
- DQ section now shows "No outstanding gaps" when all months have source data

**Bug Fix 6 — First data row hidden in HTML tables:**
- `table th { position: sticky; top: 47px; }` combined with `.section-body { overflow: hidden; }` caused the first data row to render behind the sticky header
- Removed `position: sticky; top: 47px;` from table th CSS

**Bug Fix 7 — Annual summary year format:**
- Year column displayed as "2022.0" instead of "2022" due to Pandas float coercion
- Fixed with `int()` cast: `f'<td>{int(r["Year"])}</td>'`

**Bug Fix 8 — Pre-2022 purchase report column names:**
- 2020-era files use different column headers than 2022+
- Added fallback chains in `read_standard_purchase()`:
  - `PO Date` → `Receipt Date` → `Order Date`
  - `Quantity` → `Qty`
  - `Quantity Received` → `Qty Recd.`
  - `Receipt Amount` → `Total Price`
  - `Mark up %` → `Markup` → `Mark % 1`
  - `PO Number` → `Order ID`

**Bug Fix 9 — Missing Client column in 2020-2021 purchase reports:**
- Some pre-2022 files have no `Client Name`/`Client` column, and contain purchases for ALL Broadridge clients (Fidelity, Morgan Stanley, UBS, etc.)
- Initial fix (blanket "assume Apex") inflated purchases to 75M — included all clients
- Corrected approach: when no Client column, check row text for "APEX"/"RIDGE" keywords; exclude "BROADRIDGE" false positives
- Jan 2021 legitimately has no Apex items — not a bug

**Reconciliation expanded to January 2020:**
- Changed `START_MONTH_KEY = "2020-01"`
- Added `BASE_PRE2022_PURCHASE` path constant for `Purchase Reports (from email)/Purchase Reports/FY'20` and `FY'21` directories
- Added `BILLING_MASTER_2020` path constant for `Billing Master - 2020 - 2021.xlsx`
- New function `process_pre2022_purchase_reports()` walks nested FY directories
- Updated billing pipeline to process Billing Master 2020-2021 first
- Updated all HTML report date references from "March 2022" to "January 2020"

**Inventory reconciliation bridge added:**
- Executive Summary now includes an inventory bridge table:
  - Opening inventory (pre-Jan 2020): implied = current inventory - net variance
  - \+ Envelopes purchased
  - − Envelopes used (volume)
  - = Current inventory (from Commodities Item List)

**Current Key Findings (superseded — see session 3 below):**
| Metric | Value |
|--------|-------|
| Reconciliation Period | Jan 2020 – Dec 2025 |
| Purchase files processed | 61 |
| Billing files processed | 53 |
| Total Purchased | 30,391,932 |
| Total Used (Volume) | 30,931,088 |
| Net Variance | (539,156) = -1.8% |
| Total Cost | $2,138,309 |
| Total Invoiced | $2,233,701 |

**Progression of deficit (continued):**
| Stage | Deficit | Cause |
|-------|---------|-------|
| Prior session end | (901,687) | Mar 2022 – Dec 2025 scope |
| After Apr/Jun/Jul 2025 files added | (220,388) | +3 purchase reports, +3 billing workbooks |
| After Apr-25 billing sheet fix | (547,117) | +327K Apr-25 usage recovered |
| After expanding to Jan 2020 | (539,156) | +24 months of data, net -1.8% |

**Data gaps: NONE**
- All months from January 2020 through December 2025 have source data
- May/Jun 2025 confirmed as zero-purchase months (not missing data)

### 2026-02-23 (session 3)

**Accomplished:**
- Cleaned up all temp files (4 encoded chunks, 7 Outlook search scripts)
- Cross-validated purchase data against Broadridge consolidated file `2019-2022 Purchases.xlsx`
- Added consolidated file as supplementary purchase source to fill data gaps
- Added `.xls` legacy file format support (recovered Aug-20 purchase report)
- Built two-pass purchase deduplication (description-based + PO-number-based)
- Fixed UOM `"TH"` (thousands) handling for Sep-21 purchase reports
- Removed commodities/inventory from report — now shows purchase vs usage only
- Added "Usage by Product" breakdown section to HTML report

**Cross-Validation Results:**
- Validated against `2019-2022 Purchases.xlsx` (121 POs, Feb 2019 – Dec 2022)
- 2022 purchases: **exact match** at 6,707,000
- 2021: our recon has +414K more (Pension Financial specialty items Broadridge excluded)
- 2020: -308K gap fully explained and resolved (see bug fixes below)

**Bug Fix 10 — Legacy `.xls` format not supported:**
- `Purchase Report Aug'20(1).xls` was silently skipped because `open_workbook()` only handled `.xlsx`/`.xlsm`
- Added `xlrd`-based reader in `open_workbook()` that converts `.xls` to openpyxl-compatible workbook, including date cell conversion
- Updated file extension filters on both pre-2022 and main purchase report loops to include `.xls`
- Result: +204,000 envelopes recovered for Aug-20

**Bug Fix 11 — Jun-20 folder contains wrong file:**
- `06-20 Purchase report/` folder contained a duplicate of the May purchase report (`Purchase Report May (1).xls.xlsx`), not the actual June file
- Caused 4 missing June POs (-392K) and 2 duplicate May POs (+70K)
- Added `2019-2022 Purchases.xlsx` as supplementary source to fill the gap (4 Jun-20 POs with PO numbers 691862, 692617, 692749, 692612)

**Bug Fix 12 — Cross-source deduplication failure (description mismatch):**
- Consolidated file uses bare WMS codes (`ENVCONPFSN10NI`)
- Monthly files use verbose descriptions (`UNITED ENVELOPE -- ENVCONPFSN10NI`)
- Single-pass dedup by `(date, description, qty)` couldn't match these
- Implemented two-pass dedup:
  - Pass 1: `(date, description, qty)` — catches same-source duplicates
  - Pass 2: PO number match — catches cross-source duplicates with different descriptions; prefers non-consolidated record (richer data)
- Result: 109 duplicates correctly identified and removed from 349 raw records

**Bug Fix 13 — UOM "TH" not handled:**
- Sep-21 purchase report used `UOM = "TH"` (thousands) instead of `"M"`
- Quantities stored as 84 and 180 instead of 84,000 and 180,000
- Updated all UOM checks: `uom in ("M", "TH")` instead of `uom == "M"` (4 locations)

**Report restructured — purchase vs usage only:**
- Removed: Current Inventory KPI, inventory reconciliation bridge, "Inventory" column from envelope types, "Current inventory by location" section, Commodities file processing
- Added: "Usage by Product" section — usage broken down by product name (Monthly Statements, Confirmations, etc.) with % of total bars
- Added: "Variance %" KPI card (replaces inventory card)
- Renamed: "Envelope Types" → "Purchases by envelope type"

**Progression of variance (continued):**
| Stage | Variance | Cause |
|-------|----------|-------|
| Prior (expanded to Jan 2020) | (539,156) -1.8% | Before validation fixes |
| After .xls support | +Aug-20 204K | Recovered Aug-20 purchase report |
| After dedup (too broad) | (741,156) -2.5% | Removed May-20 dupes but lost real records |
| After consolidated source + PO dedup | +202,412 +0.7% | Filled Jun-20 gap, proper cross-source dedup |

**Current Key Findings (final):**
| Metric | Value |
|--------|-------|
| Reconciliation Period | Jan 2020 – Dec 2025 |
| Purchase files processed | 63 |
| Billing files processed | 53 |
| Records extracted / unique | 349 / 240 |
| Total Purchased | 31,133,500 |
| Total Used (Volume) | 30,931,088 |
| Net Variance | +202,412 = +0.7% |
| Total Cost | $2,116,463 |
| Total Invoiced | $2,211,855 |

**Source files added this session:**
| File | Purpose |
|------|---------|
| `2019-2022 Purchases.xlsx` | Consolidated PO history from Broadridge (121 records, Feb 2019 – Dec 2022); supplementary source for gap-filling |
| `APEX Usage - Orders 2020-2022.xlsx` | Same data as above, stored in usage directory; not processed (duplicate) |

**Broadridge envelope logistics (from email context):**
- Supplier: United Envelope LLC, Mt. Pocono, PA
- Delivery: 1-2 trucks daily to Broadridge warehouse at 300 Executive Drive
- Mailing: 51 Mercedes Way (same business district, 1/2 mile from warehouse)
- Inventory: online management system with reorder points (1.5-month trend small volume, 3-month trend high volume)
- Reorder alerts: Yellow = close to reorder, Red = below reorder (expedite)
- 20 envelope types tracked in WMS

**Next Steps:**
- [ ] Obtain 3-5 vendor invoices to validate Receipt Amount composition (wastage embedded or separate)
- [ ] Clarify markup structure discrepancy between 2020-2024 (no visible markup) and 2025 (10% markup)
- [ ] Consider drafting formal letter to Broadridge re: contract variance
- [ ] Request actual Jun-20 purchase report from Broadridge (currently using consolidated file as fill-in)

### 2026-02-23 (session 4)

**Accomplished:**
- Completed full review of all 58 emails from Brandon Koebel (Aug 2023 – Feb 2024 portion)
- Compiled comprehensive analysis explaining sustained negative running balance periods
- Cleaned up temp `_search_koebel.ps1` file

**Email Analysis — Key Operational Facts from Brandon Koebel:**

The user questioned why the running balance shows negative inventory for 10 months straight and again 7 months recently. The emails reveal three structural reasons this is expected:

1. **Pre-existing buffer stock invisible to our window:**
   - Brandon (Nov 1, 2022): *"We typically try to keep a 2-3 month supply of envelopes in stock."*
   - At ~700K envelopes/month, that's 1.4M–2.1M envelopes on hand before our Jan 2020 start date
   - This inventory was purchased pre-2020 and never appears in our "purchased" totals but is available for use

2. **"Usage" ≠ envelopes actually mailed:**
   - Brandon (Oct 31, 2022): *"The usage report shows what was ordered from our inventory system and brought to the machines processing Apex jobs. For example, the first line shows 336,000 used but only 282,497 confirms mailed, which means aside from ~10% waste, the remainder of the envelopes are on the floor queued up to be used, also called 'surplus'."*
   - Usage = envelopes pulled from warehouse to production floor, including waste and floor surplus

3. **Wastage inflates usage by 10–15%:**
   - Brandon (Sep 29, 2022): *"did not account for any waste or spoilage (typically 10-15%)"*
   - Brandon (Nov 7, 2022): *"Wastage is roughly 10%... This includes envelopes that are damaged, need to be reprinted and reinserted, etc."*

**Conclusion:** The +0.7% net surplus across the full period confirms data integrity. Negative running balance stretches reflect drawdown of pre-existing buffer stock and usage inflation from wastage, not actual stockouts.

**Complete Envelope Type Mapping (from Sep 12, 2023 email):**
| WMS Code | Mail Type | Category |
|----------|-----------|----------|
| ENVMEAPEXN14PFC | Domestic Fold Statement | Statements |
| ENVMEAPEX9X12PFC | Domestic Flat Statement | Statements |
| ENVMERIDGEN14NI11/08 | Foreign Fold Statement | Statements |
| ENVMERIDGE9X12NI11/08 | Foreign Flat Statement | Statements |
| ENVCONRIDGE9X12DW | Flat Confirms (domestic + foreign) | Confirms |
| ENVAPXN10PFSCONN10IND(10/22) | Domestic Fold Confirms + Domestic Letters | Confirms/Letters |
| ENVCONPFSN10NI | Foreign Fold Confirms + Foreign Letters | Confirms/Letters |

**Additional Email Facts:**
| Date | Finding |
|------|---------|
| Sep 28, 2022 | Supplier: United Envelope LLC, Mt. Pocono PA. 1-2 trucks daily to 300 Executive Drive. Mailing at 51 Mercedes Way. Reorder on 1.5-month (small vol) or 3-month (high vol) trend. Yellow/Red inventory alerts. |
| Mar 31, 2023 | ENVAPXN10PFSCONN10IND(10/22) replaced ENVCONPFSN10NI in Oct 2022 — postal permit update for better postage rate |
| Apr 2, 2023 | Jan 2023 purchase report was missing one Ridge line item (legacy Apex name) |
| Apr 14, 2023 | PO#688124 cancelled. WMS-104050 was billed for 180,000 (not estimated qty) |
| May 5, 2023 | "Copy of Envelope Mapping.xlsx" attachment — maps envelope types to product names |
| May 12, 2023 | NI envelopes not retired, still used for foreign mail |
| Aug 28, 2023 | "Postage_No Volume Support" items = spoils (damaged pieces reprinted in separate jobs) |
| Sep 12, 2023 | Insert waste estimate: 10-20% added to projected monthly statement count |
| Sep 27, 2023 | Hand insertion cutoff: <200 envelopes = hand, ≥200 = machine. Manifested mail used for statements (USPS discount, quicker). |
| Oct 26, 2023 | "Any months missing from the purchase reports means there were no applicable purchases in that month" |
| Nov 8, 2023 | 2018 data only available as QBR totals (no detailed backup) |
| Nov 30, 2023 | Account-level detail only stored for 60 days |
| Dec 11, 2023 | All envelopes are double-window (return address visible through window) |
| Dec 20, 2023 | ADS letters SOW signed late 2022, development mid-2023, live Sep 2023 |
| Jan 22, 2024 | Account-level report shows print accounts, pages, images, job names. Each line = one envelope. No reports tie individual account to postage. |
| Feb 13, 2024 | Volumes tracked in "mailing database" — envelopes, images, sheets, postage all sourced from it |

### 2026-02-23 (session 5)

**Accomplished:**
- Reviewed envelope sample PDFs from `Products and Envelope Samples` directory (7 specification sheets + 4 sample documents)
- Mapped all 7 active envelope types to physical specs, order numbers, sizes, and mail types
- Added "Envelope Specifications" section to HTML report between "Purchases by Type" and "Usage by Product"
- Section includes WMS code, mail type, size, style, postage type, and notes for each envelope
- Added legend for code abbreviations (PFC, NI, DW, IND)

**Envelope Spec Source Files (read-only reference):**
| File | Order # | WMS Code |
|------|---------|----------|
| `926131_APEX14PFC 712_rp.pdf` | 926131 | ENVMEAPEXN14PFC |
| `851251_APEX9X12PFC 712.pdf` | 830851 | ENVMEAPEX9X12PFC |
| `942095_RIDGEPLN14 (11-08)_rp.pdf` | 942095 | ENVMERIDGEN14NI11/08 |
| `823804_RIDGEPLN9X12_11_08..pdf` | 823804 | ENVMERIDGE9X12NI11/08 |
| `992124_PFS CON N10 IND (1022)_sp.pdf` | 992124 | ENVAPXN10PFSCONN10IND(10/22) |
| `856743_PFS CON N10 (0210).pdf` | 856743 | ENVCONPFSN10NI |
| `893283_RIDHE 9x12.pdf` | 818105 | ENVCONRIDGE9X12DW |

### 2026-02-23 (session 6)

**Accomplished:**
- Conducted full inventory efficiency audit analyzing purchase patterns, wastage, cost trends, and ordering behavior
- Consolidated "By Envelope Type" tab from 35 raw description variants down to 9 canonical types
- Discovered and excluded 1,068,000 non-Apex envelopes (Fidelity 168K, HSBC 20K, Morgan Stanley 880K) that had leaked into Oct 2021 purchase data from the consolidated file
- Added `normalize_envelope_type()` function with regex-based mapping rules
- Applied non-Apex filter to monthly summary, purchase detail, contract audit, and console summary
- Added "Envelope Specifications" section to HTML report with physical specs and mail type mapping
- Regenerated Excel and HTML outputs

**Bug Fix 14 — Non-Apex envelopes in purchase records:**
- `2019-2022 Purchases.xlsx` (consolidated file) contains purchases for ALL Broadridge clients, not just Apex
- `read_consolidated_purchases()` checked `is_envelope()` but not `is_apex()`, allowing Fidelity/HSBC/Morgan Stanley envelopes through
- Three non-Apex envelope types identified and excluded via `normalize_envelope_type()`:
  - `ENVFIDN101.9863867.101` — Fidelity, 168,000 (Sep 2021)
  - `ENVHSBCN10EN796A05/14` — HSBC, 20,000 (Oct 2021)
  - `ENVMSTN10M017ECCCDi(5/19)` — Morgan Stanley, 880,000 (Oct 2021)
- Impact: Total purchased dropped from 31,133,500 to 30,065,500. Oct 2021 "3.2x over-purchase" was actually 88% non-Apex envelopes.

**Envelope type consolidation mapping:**
| Canonical Type | Variants Merged | Total Purchased |
|----------------|----------------|-----------------|
| ENVMEAPEXN14PFC | 3 variants (incl. SUPPLIER ID, STMTPFC9/24) | 9,370,000 |
| ENVAPXN10 Confirms+Letters (PFC) | 9 variants (IND 10/22, 4/25, CNFPFC, LTRPFC) | 9,342,000 |
| ENVCONPFSN10NI | 4 variants (incl. SUPPLIER ID, UNITED ENVELOPE prefix, CNFNI) | 9,670,000 |
| ENVMEAPEX9X12PFC | 3 variants | 652,500 |
| ENVMERIDGEN14NI11/08 | 3 variants | 668,000 |
| ENVMERIDGE9X12NI11/08 | 2 variants | 65,000 |
| ENVCONRIDGE9X12DW | 2 variants | 80,000 |
| Tax Form Envelopes (1099/1099-R) | 5 variants | 210,000 |
| Tax Form Envelopes (1042/IRA) | 1 variant | 8,000 |

**Audit findings (efficiency assessment):**
1. **Erratic ordering (CV = 0.61)** — purchases range 70K–1.8M/month vs ~430K avg usage
2. **2022 structural failure** — Q1-Q3 purchased only 51-67% of usage during Ridge/Penson migration volume spike
3. **Wastage discrepancy** — Brandon claimed 10-15%, data shows ~0.3% spoils (likely waste embedded in "used" not separately tracked)
4. **2025 markup anomaly** — 31.8% markup appeared vs ~12% expected per contract amendment
5. **33% of quarters in deficit** — 8 of 24 quarters had purchases below usage

**Corrected Key Findings:**
| Metric | Previous | Corrected |
|--------|----------|-----------|
| Total Purchased | 31,133,500 | 30,065,500 |
| Non-Apex Excluded | — | 1,068,000 |
| Net Variance | +202,412 (+0.7%) | (865,588) (-2.9%) |
| Total Cost | $2,116,463 | $2,094,631 |
| Total Invoiced | $2,211,855 | $2,190,023 |

**Next Steps:**
- [ ] Obtain 3-5 vendor invoices to validate Receipt Amount composition (wastage embedded or separate)
- [ ] Clarify markup structure discrepancy between 2020-2024 (no visible markup) and 2025 (31.8% vs expected 12%)
- [ ] Consider drafting formal letter to Broadridge re: contract variance and inventory management concerns
- [ ] Request actual Jun-20 purchase report from Broadridge (currently using consolidated file as fill-in)

### 2026-02-23 (session 7)

**Accomplished:**
- Fixed critical 2022 double-counting bug — YTD file (`Apex YTD Envelope Usage (2).xlsx`) duplicated Jan-Aug 2022 volume and postage data already covered by individual billing workbooks, inflating 2022 usage by 3,716,075
- Removed Data Quality section from HTML report; replaced with a one-line alert banner that only appears when months have missing source data
- Removed all cost/invoiced columns and Contract Audit section from HTML report — report now focuses purely on purchased vs. used reconciliation
- Regenerated Excel and HTML outputs with corrected data

**Bug Fix 15 — YTD file double-counting 2022 volume and postage:**
- `read_ytd_usage()` read Volume Data and Postage Data tabs from the YTD file, which contained identical records to the Jan-Aug 2022 individual billing workbooks
- Every month from Jan-Aug 2022 was exactly doubled (e.g., Jan-22: 511,349 billing + 511,349 YTD = 1,022,698)
- Total duplication: 3,716,075 volume records + 3,716,075 postage records
- Fix: Removed `read_ytd_usage()` call from `main()` — function is entirely redundant with individual billing workbooks
- Impact: 2022 variance corrected from -42.0% (2.7M deficit) to +14.9% (972K surplus)

**Report changes:**
- Removed: Data Quality section (full collapsible section with table)
- Added: One-line alert banner at top of content area, only visible when months are missing
- Removed: Total Invoiced KPI card, Contract Overcharge KPI card
- Removed: Total Cost and Total Invoiced columns from Annual Summary table
- Removed: Purchase Cost and Invoiced columns from Monthly Detail table
- Removed: Total Cost and Avg Unit Price columns from Purchases by Envelope Type table
- Removed: Contract Audit section entirely (terms grid, audit stats, top 10 discrepancies)
- Removed: Data Quality and Contract Audit nav links
- Removed: Unused CSS classes (`.terms-grid`, `.term-card`, `.audit-stats`, `.audit-stat`, `.flag-over`)
- Removed: Unused Python functions (`fmt_money`, `fmt_money_always`, `fmt_money_parens`, `build_top10_audit_rows`, `build_dq_rows`)
- Net reduction: 199 lines removed, 26 added (file went from 760 to 587 lines)

**Corrected Key Findings (supersedes session 6):**
| Metric | Session 6 | Session 7 (corrected) |
|--------|-----------|----------------------|
| Total Purchased | 30,065,500 | 30,065,500 |
| Total Used (Volume) | 30,931,088 | 27,215,013 |
| Total Mailed (Postage) | — | 27,210,696 |
| Total Spoils | — | 84,822 |
| Net Variance | (865,588) (-2.9%) | +2,850,487 (+9.5%) |
| 2022 Variance | (2,744,001) (-42.0%) | +972,074 (+14.9%) |

### 2026-02-23 (session 8)

**Accomplished:**
- Consolidated Annual Summary and Monthly Detail into a single table with annual subtotal rows after each December and a grand total row at the bottom; removed separate Annual Summary section
- Merged "Purchases by Envelope Type" and "Usage by Product" into a single side-by-side section ("Purchases & usage breakdown"); removed separate Usage by Product section
- Verified annual/monthly data integrity — all 6 years match exactly between the two views
- Added `.subtotal-row` CSS styling and excluded subtotal rows from column sorting

**Report structure (current):**
1. Executive Summary (4 KPI cards)
2. Monthly Trend (SVG bar chart)
3. Monthly Detail (72 months + 6 annual subtotals + grand total)
4. Purchases & Usage Breakdown (side-by-side tables)
5. Envelope Specifications (7 types with physical specs)

### 2026-02-23 (session 9)

**Accomplished:**
- Added **usage by envelope type** — extracted `Flat_Fold` and `Address_Type` fields from billing workbook Volume Data sheets to map each usage record to a canonical envelope type
- Created `map_usage_to_envelope_type()` function based on Brandon Koebel's Sep 2023 mail-type mapping
- Added "Usage by Envelope Type" tab to Excel output
- Built combined **Purchases & Usage by Envelope Type** table in HTML report, grouped by physical envelope size (not postage imprint) for meaningful comparison
- Located the **pass-through paper dispute settlement** document — `GTO Proxy and BPS Term Sheet` (June 2022) in `C:\Users\smendoza\OneDrive - Apex Clearing\Broadridge Billing\Broadridge Contracts and Schedules`
- Added settlement context info box to Executive Summary

**Usage-to-envelope-type mapping:**
| Product Category | Flat/Fold | Address Type | Envelope Type |
|---|---|---|---|
| STATEMENT | FOLD/MIXED/BULK | DOMESTIC/MIXED/other | ENVMEAPEXN14PFC |
| STATEMENT | FLAT | DOMESTIC/MIXED/other | ENVMEAPEX9X12PFC |
| STATEMENT | FOLD/MIXED/BULK | FOREIGN | ENVMERIDGEN14NI11/08 |
| STATEMENT | FLAT | FOREIGN | ENVMERIDGE9X12NI11/08 |
| CONFIRM/LETTER/CHECK | FOLD/MIXED/BULK | DOMESTIC/MIXED/other | ENVAPXN10 Confirms+Letters (PFC) |
| CONFIRM/LETTER/CHECK | FLAT | any | ENVCONRIDGE9X12DW |
| CONFIRM/LETTER/CHECK | FOLD/MIXED/BULK | FOREIGN | ENVCONPFSN10NI |
| TAX DOCUMENT | any | any | Tax Form Envelopes (1099/1099-R) |

**Combined envelope type table (grouped by physical size):**
| Envelope Type | Purchased | Used | Variance | Var% |
|---|---|---|---|---|
| N14 Fold Statements | 10,038,000 | 9,467,708 | +570,292 | +5.7% |
| 9x12 Flat Statements | 717,500 | 517,871 | +199,629 | +27.8% |
| #10 Confirms + Letters | 19,012,000 | 16,970,875 | +2,041,125 | +10.7% |
| 9x12 Flat Confirms | 80,000 | 198,707 | (118,707) | -148.4% |
| Tax Form Envelopes | 218,000 | 59,852 | +158,148 | +72.5% |
| **Total** | **30,065,500** | **27,215,013** | **+2,850,487** | **+9.5%** |

**Pass-through paper dispute settlement (key finding):**
- Source: GTO Proxy & BPS Early Renewal Term Sheet (June 2022)
- Broadridge agreed to internalize **$643,457.92** in accumulated paper and envelope costs prior to March 1, 2022
- Apex began paying for paper and envelopes per contract terms effective March 1, 2022
- This explains why pre-2022 purchase data shows no markup — Broadridge was absorbing the cost
- Settlement was part of broader Proxy & BPS early renewal negotiation

**Report structure (current):**
1. Executive Summary (4 KPI cards + settlement info box)
2. Monthly Trend (SVG bar chart)
3. Monthly Detail (72 months + 6 annual subtotals + grand total)
4. Purchases & Usage by Envelope Type (combined table + per-SKU/per-product detail)
5. Envelope Specifications (7 types with physical specs)

### 2026-02-23 (session 9 continued)

**Accomplished:**
- Searched `S:\Departments\Accounting\Private\Broadridge` exhaustively for ~$200K envelope/paper settlement document — not found in that directory
- Located the settlement in the **GTO Proxy & BPS Early Renewal Term Sheet** (June 2022) in `C:\Users\smendoza\OneDrive - Apex Clearing\Broadridge Billing\Broadridge Contracts and Schedules`
- Added **post-settlement inventory analysis** to Executive Summary covering Mar 2022 – Dec 2025

**Pass-through paper dispute settlement:**
- Source: GTO Proxy & BPS Early Renewal Term Sheet, ADDITIONAL section item #2
- **$643,457.92** in accumulated paper and envelope costs prior to March 1, 2022 — Broadridge agreed to internalize (absorb) all costs
- Apex began paying for paper and envelopes per contract terms effective March 1, 2022
- Also in term sheet: $345,649 BPO Temp Worker Dispute — Apex to pay $222,982, Broadridge forgives remainder
- Also: Apex agrees to reduce print as % of volumes by at least 30% by January 1, 2024

**Post-settlement analysis (Mar 2022 – Dec 2025):**
| Year | Purchased | Used | Variance | Var% | Avg Mo Used |
|------|-----------|------|----------|------|-------------|
| 2022 (Mar-Dec) | 5,807,000 | 4,629,923 | +1,177,077 | +20.3% | 462,992 |
| 2023 | 6,571,000 | 6,081,272 | +489,728 | +7.5% | 506,773 |
| 2024 | 4,621,000 | 4,348,349 | +272,651 | +5.9% | 362,362 |
| 2025 | 4,040,500 | 3,410,405 | +630,095 | +15.6% | 284,200 |
| **Total** | **21,039,500** | **18,469,949** | **+2,569,551** | **+12.2%** | |

**Key findings:**
- Implied inventory of 2,569,551 envelopes = **8.7 months buffer stock** at trailing 6-mo usage of 294,965/mo (Broadridge policy: 2-3 months)
- Average monthly usage declined **39%** from 462,992/mo (2022) to 284,200/mo (2025) — consistent with 30% print reduction target in term sheet
- All 4 years show surplus (smallest: 2024 at +5.9%)
- Running balance never goes negative after a brief dip in Apr 2022 (-64K)

**Report structure (current):**
1. Executive Summary (4 full-period KPI cards + settlement info box + 4 post-settlement KPI cards + year-by-year table + key findings)
2. Monthly Trend (SVG bar chart)
3. Monthly Detail (72 months + 6 annual subtotals + grand total)
4. Purchases & Usage by Envelope Type (combined table + per-SKU/per-product detail)
5. Envelope Specifications (7 types with physical specs)

**Next Steps:**
- [ ] Obtain 3-5 vendor invoices to validate Receipt Amount composition (wastage embedded or separate)
- [ ] Clarify markup structure discrepancy between 2020-2024 (no visible markup) and 2025 (31.8% vs expected 12%)
- [ ] Consider drafting formal letter to Broadridge re: contract variance and inventory management concerns
- [ ] Request actual Jun-20 purchase report from Broadridge (currently using consolidated file as fill-in)
- [x] ~~Investigate why implied inventory exceeds Broadridge's stated 2-3 month buffer policy~~ — Now 8.7 months; see excess inventory next step

### 2026-02-23 (session 10)

**Accomplished:**
- CEO-lens review and full report restructure for executive readability
- Bottom Line callout, post-settlement KPIs promoted, full-period demoted to context line
- Monthly Detail and Reference sections collapsed by default; nav links auto-expand
- Added inventory gauge, SKU recon table, usage trend sparkline, usage-by-product table
- Changed avg monthly usage from trailing 12 to trailing 6 months (rolling average for ongoing recon)

### 2026-02-23 (session 11)

**Accomplished:**
- Completed all HTML template wiring from session 10 (inventory gauge, SKU recon, usage trend)
- Analytical review of report — identified and fixed 7 issues

**Bug Fix 16 — `fmt_pct` percentage formatting:**
- Any variance ratio exceeding +/-1.0 (i.e., >100%) displayed as raw ratio instead of percentage
- Example: -148.4% showed as "-1.5%" because `abs(val) < 1` branch multiplied by 100, else branch did not
- Fix: always multiply by 100 (`return f"{val * 100:.1f}%"`)
- Affected: 9x12 Flat Confirms (-148.4%), Tax Forms (100.0%), and several monthly detail rows

**Bug Fix 17 — SKU usage mapping not date-aware:**
- `map_usage_to_envelope_type()` in `build_recon_from_source.py` was static — always mapped domestic #10 confirms/letters to ENVAPXN10 (PFC)
- Before Oct 2022, those envelopes were physically ENVCONPFSN10NI (NI version)
- All pre-Oct 2022 domestic confirm/letter usage (~7.1M) was incorrectly attributed to the PFC SKU
- Fix: added `month_key` parameter; domestic #10 usage before `2022-10` maps to ENVCONPFSN10NI, after maps to ENVAPXN10 (PFC)
- Result: ENVAPXN10 (PFC) corrected from (7.3M) deficit to (182K) at -2.0%; ENVCONPFSN10NI corrected from 9.4M surplus to 2.2M at +23.0%

**Report scoped to post-settlement (Mar 2022 – Dec 2025):**
- Header subtitle, SVG bar chart, monthly detail table, rolling usage trend all filtered to post-settlement
- Removed Mailed and Spoils columns from monthly detail (noise — Mailed within 4K of Used every month, Spoils <2K most months)
- By Type section retains full-period data with info box noting "~70% is post-settlement" (type-level monthly breakdown not available)
- Pre-settlement context preserved in gray context line at bottom of Executive Summary
- 2022 subtotal labeled "2022 (Mar–Dec)" for clarity

**Spoilage confirmed:**
- Post-settlement spoils: 55,733 of 18,469,949 used = 0.30%
- Well within 10% contractual wastage limit
- Added to Used KPI card sub-text and key findings bullet
- Note: actual wastage (10-15% per Brandon Koebel) is embedded in "Used" figure, not separately reported as spoils

**Other report improvements:**
- SKU transition note: explains ENVCONPFSN10NI → ENVAPXN10 (PFC) Oct 2022 change
- May-25 anomaly footnote: 752K usage with 0 purchases flagged for verification
- Product-to-envelope mapping note: connects product names (Address Verification Letters, Monthly Statements, etc.) to physical envelope types
- Full-period scope note on By Type section (info box)

**Corrected SKU-level recon (after date-aware mapping fix):**
| SKU | Purchased | Used | Variance | Var% |
|-----|-----------|------|----------|------|
| ENVAPXN10 Confirms+Letters (PFC) | 9,342,000 | 9,524,224 | (182,224) | -2.0% |
| ENVMEAPEXN14PFC | 9,370,000 | 9,117,954 | +252,046 | +2.7% |
| ENVCONPFSN10NI | 9,670,000 | 7,446,651 | +2,223,349 | +23.0% |
| ENVMEAPEX9X12PFC | 652,500 | 500,829 | +151,671 | +23.2% |
| ENVMERIDGEN14NI11/08 | 668,000 | 349,754 | +318,246 | +47.6% |
| ENVCONRIDGE9X12DW | 80,000 | 198,707 | (118,707) | -148.4% |
| Tax Form (1099/1099-R) | 210,000 | 59,852 | +150,148 | +71.5% |
| ENVMERIDGE9X12NI11/08 | 65,000 | 17,042 | +47,958 | +73.8% |
| Tax Form (1042/IRA) | 8,000 | 0 | +8,000 | +100.0% |
| **Total** | **30,065,500** | **27,215,013** | **+2,850,487** | **+9.5%** |

**Current report structure:**
1. Executive Summary (bottom line + 4 KPIs + inventory gauge + key findings + year-by-year + context line)
2. Monthly Trend (SVG bar chart — post-settlement + rolling 6-mo avg sparkline)
3. Purchases & Usage by Envelope Type (grouped table + SKU recon + usage by product)
4. Monthly Detail (collapsed, post-settlement, 6 columns)
5. Reference (collapsed, envelope specs only)

**How to refresh outputs:**
```bash
py -3 "C:\Users\smendoza\Projects\Broadridge Envelopes\build_recon_from_source.py"
py -3 "C:\Users\smendoza\Projects\Broadridge Envelopes\generate_html_report.py"
```

**Next Steps:**
- [ ] Obtain 3-5 vendor invoices to validate Receipt Amount composition (wastage embedded or separate)
- [ ] Clarify markup structure discrepancy between 2020-2024 (no visible markup) and 2025 (31.8% vs expected 12%)
- [ ] Consider drafting formal letter to Broadridge re: excess inventory (8.7 months buffer vs 2-3 month policy)
- [ ] Investigate ENVCONRIDGE9X12DW deficit — 80K purchased vs 199K used (-148%) suggests missing purchase records or misclassified usage
- [ ] Verify May-25 752K usage spike — possible billing consolidation from multiple periods
- [ ] Request actual Jun-20 purchase report from Broadridge (currently using consolidated file as fill-in)

### 2026-02-23 (session 12)

**Accomplished:**
- Comprehensive Outlook email search across all folders — 7 keyword/sender searches covering envelope, United Envelope, wastage/spoilage, print and mail, @broadridge.com senders, postage, and commodities
- Read and analyzed 9 email threads not previously reviewed (beyond the 130 Koebel/Denci emails from session 4)

**Key New Findings from Outlook:**

1. **Envelope wastage confirmed at 5% (Aug 23, 2023)**
   - Christopher Denci formally answered Terry Ray's renewal pricing questions
   - *"The current agreement reflects Inventory Cost Plus 10% margin. Materials are billed at cost plus wastage for generic stock. Specifically, the wastage charge is 10% for any generic paper stock and 5% for generic envelope stock."*
   - Generic stock (envelopes): unit rate billed based on **usage**
   - Client-specific stock: unit rate billed based on **receipt**
   - Pre-2024: 5% wastage on envelopes; Post-2024 Amendment: 2% wastage + 10% margin

2. **Materials NOT included in annual fee (May 31, 2023)**
   - Sameer Jobanputra confirmed: Section 4 "Compensation" states materials (paper, envelopes, inserts) charged separately from the $475K annual fee

3. **Interim Extension Agreement executed Dec 27, 2023**
   - File: `apex extension 12-27 fully executed.pdf`
   - Negotiated by John Russ (Broadridge GRM), signed before formal Amendment No. 1

4. **PostEdge reconciliation limitations (Aug 23, 2023)**
   - 500K+ line items per month, over an hour to download
   - No domestic vs. foreign postage breakout, no image volume data
   - Matthew Burkavage (Broadridge) offered alternative reconciliation tools

5. **PostEdge ICS credits — $63K expected (Jan 4, 2024)**
   - Andy Graf projected ~$650K in ICS credits; your calculation estimated ~$63,475.92
   - Credits expected as credit memo, not cash payment

6. **Insert SOW pricing details (Sep 12, 2023)**
   - Brandon Koebel: insert estimates add 10-20% for waste and new accounts
   - Insert paper: 24# stock (= 60# text), thicker to avoid machine jams
   - SOWs are estimates; actual billing per-unit
   - Logan Jones: Apex not currently upcharging inserts to correspondents

**Broadridge Contact Map (compiled from all threads):**
| Person | Role | Handles |
|--------|------|---------|
| Christopher Denci | ICS Account Manager | Post-sale, pricing definitions |
| Sameer Jobanputra | BRCC Account Manager | BRCC mailing, invoices |
| Brandon Koebel | Sr. Client Relationship Manager | Envelope ops, inserts, billing data |
| Josh Edelstein | GRM | Renewal proposals |
| John Russ | Global Relationship Management | Contract execution |
| Brian O'Toole | Account Management | A/R, billing disputes |
| Matthew Burkavage | Supervisor, Post Sale Client Services | PostEdge access |
| Michael Schnupp | Director, Account Mgmt - GTO | Escalation, postage detail |
| Kimberly Rookwood | GTO Relationship | Initial contact routing |
| Lynnette Kappler | GTO AR | Outstanding invoices |
| Gary Stuart | GTO AR | Aged A/R |
| Woodie Cheu | (unknown) | Outstanding items |

**Email threads reviewed this session:**
| Thread | Dates | Key Content |
|--------|-------|-------------|
| Request for Definitions of Charges | Aug 2023 | Terry Ray → Denci/Edelstein; formal charge definitions including wastage rates |
| Print and Mailing Services | May-Aug 2023 | Sameer Jobanputra; materials separate from annual fee |
| Interim Agreement | Dec 2023-Jan 2024 | Extension letter executed 12/27/23; renewal timeline |
| Apex/Broadridge Renewals | Feb 2024 | Terry Ray forward; references spoilage in context |
| May-2021 Confirm Postage Detail | Aug 2021 | Michael Schnupp; Address Letter Volume attachment |
| PostEdge Access | Aug 2023 | Matthew Burkavage; PostEdge limitations documented |
| BR Postedge (ICS credits) | Jan 2024 | Andy Graf; $63K credit estimate |
| Insert Grps Production Table | Apr-Sep 2023 | Insert pricing, SOW process, waste estimates |
| Ally Mass Mailing | Dec 2023 | United Envelope mentioned in context |

**Next Steps:**
- [ ] Obtain 3-5 vendor invoices to validate Receipt Amount composition (wastage embedded or separate)
- [x] ~~Clarify markup structure: pre-2024 = cost + 5% wastage + 10% margin vs 2025 showing 31.8%~~ — Resolved in session 13
- [ ] Consider drafting formal letter to Broadridge re: excess inventory (8.7 months buffer vs 2-3 month policy)
- [ ] Investigate ENVCONRIDGE9X12DW deficit — 80K purchased vs 199K used (-148%)
- [ ] Verify May-25 752K usage spike
- [ ] Request actual Jun-20 purchase report from Broadridge

### 2026-02-23 (session 13)

**Accomplished:**
- Investigated and resolved the reported "31.8% markup discrepancy" in 2025 purchase data
- Read both contracts (original Jan 2019, Amendment No. 1 Jan 2024) and extracted exact pricing formulas
- Examined raw purchase report files across all three format eras (old .xlsm, new .xlsx)
- Identified and fixed two bugs in `build_recon_from_source.py` that caused incorrect cost/invoiced calculations
- Rebuilt Excel and HTML outputs with corrected data

**Contract pricing formulas (extracted from contract language):**
| Period | Formula | Effective Rate |
|--------|---------|----------------|
| Jan 2019 – Dec 2023 | Vendor price + 5% wastage | 5.0% over vendor |
| Jan 2024 – present | (Vendor price + 2% wastage) × 1.10 margin | 12.2% over vendor |

**Bug Fix 18 — Old format (.xlsm) vendor cost lost:**
- `read_standard_purchase()` line 325: `total_cost = invoiced if invoiced > 0 else receipt_amt`
- The "Mark up %" column contains the invoiced total in dollars (vendor cost × 1.10), not a percentage
- Code set both `total_cost` and `invoiced_amount` to the invoiced value, losing the vendor cost
- Fix: `total_cost = receipt_amt` (vendor cost from "Receipt Amount" column)
- Effect: Jan 2024 – Sep 2025 POs now correctly show 10% markup instead of 0%

**Bug Fix 19 — New format (.xlsx) double-counting vendor cost:**
- `read_new_format_purchase()` line 396: `invoiced = total_cost + markup_total`
- The "Markup Total" column contains the invoiced total (vendor cost × 1.10), not just the markup delta
- Code added vendor cost + invoiced total = 2.1× vendor cost → appeared as 110% markup
- Fix: `invoiced = markup_total` (already the full invoiced amount)
- Effect: Oct–Dec 2025 POs now correctly show 10% markup instead of 110%

**Root cause of "31.8%" figure:**
- Jan–Sep 2025 (old format): 0% apparent markup (both columns = invoiced)
- Oct–Dec 2025 (new format): 110% apparent markup (double-counted vendor cost)
- Blended average ≈ 31.8% — entirely a data parsing artifact

**Actual vs contract markup:**
| Component | Contract (Amendment) | Broadridge actual | Delta |
|-----------|---------------------|-------------------|-------|
| Wastage | +2% on vendor price | Not separately applied | (2%) |
| Margin | +10% on inventory cost | +10% on vendor price | — |
| **Effective rate** | **12.2%** | **10.0%** | **-2.2% (favors Apex)** |

**Corrected cost totals (supersedes session 11/12):**
| Metric | Before fix | After fix |
|--------|-----------|-----------|
| Total Purchase Cost (vendor) | $2,116,463 | $1,994,141 |
| Total Invoiced Amount | $2,211,855 | $2,103,303 |
| Effective blended markup | Mixed (0%/110%) | ~5.5% (0% pre-2024, 10% post-2024) |

**Key conclusion:** No markup discrepancy to pursue. Broadridge charges 10% margin on vendor price without the additional 2% wastage, resulting in 10% effective markup vs the 12.2% the contract permits. This favors Apex.

**Next Steps:**
- [ ] Obtain 3-5 vendor invoices to validate Receipt Amount composition (wastage embedded or separate)
- [ ] Consider drafting formal letter to Broadridge re: excess inventory (8.7 months buffer vs 2-3 month policy)
- [x] ~~Investigate ENVCONRIDGE9X12DW deficit — 80K purchased vs 199K used (-148%)~~ — Resolved in session 13 (continued)
- [ ] Verify May-25 752K usage spike
- [ ] Request actual Jun-20 purchase report from Broadridge

### 2026-02-23 (session 13 continued)

**Accomplished:**
- Investigated ENVCONRIDGE9X12DW deficit (80K purchased vs 199K used, -148%)
- Extracted and analyzed 5,042 FLAT confirm/letter/check volume records from all billing sources
- Cross-referenced with Address Verification Letters and Apex MTC (New) Flat_Fold distributions

**ENVCONRIDGE9X12DW deficit — Root cause: 2022 production surge**

Four months in mid-2022 account for 88% of all-time flat confirm/letter usage:

| Month | FLAT Envelopes | Driver |
|-------|---------------|--------|
| Jan 2022 | 21,463 | MTC 16K + AVL 5K |
| Aug 2022 | 45,075 | MTC 23K + AVL 22K |
| Sep 2022 | 64,651 | MTC 46K + AVL 18K |
| Nov 2022 | 45,594 | MTC 36K + AVL 9K |
| **Subtotal** | **176,783** | **88% of 199K total** |

Outside those 4 months, FLAT confirms average ~350/month.

**Why it spiked:** Coincides with Ridge/Penson migration and peak volumes. Apex MTC (New) confirms went from 0.2% flat (Jul-22) to 19.9% flat (Sep-22) to 27.9% flat (Nov-22), then back to 0% flat (Dec-22). Broadridge temporarily routed a portion of confirms through flat production (9x12 envelopes) during the volume surge — likely due to inserter capacity or batch composition.

**Why the deficit is not actionable:**
1. Mapping is correct per Brandon Koebel's Sep 2023 email (ENVCONRIDGE9X12DW = flat confirms)
2. Pre-existing buffer stock covers the deficit (same pattern as overall recon)
3. 199K flat envelopes = 0.7% of 27.2M total; 119K deficit is noise in context
4. Classification is legitimate Broadridge production system data, not a mapping error

**Address Verification Letters Flat_Fold distribution (full period):**
| Flat_Fold | Envelopes | % of AVL Total |
|-----------|-----------|----------------|
| MIXED | 6,971,688 | 65.1% |
| (blank) | 3,599,316 | 33.6% |
| FOLD | 92,696 | 0.9% |
| FLAT | 53,568 | 0.5% |
| BULK | 5,718 | 0.1% |

**Apex MTC (New) Flat_Fold distribution (full period):**
| Flat_Fold | Envelopes | % of MTC Total |
|-----------|-----------|----------------|
| FOLD | 7,091,927 | 78.0% |
| (blank) | 1,754,557 | 19.3% |
| FLAT | 146,099 | 1.6% |
| MIXED | 99,124 | 1.1% |
| BULK | 52,812 | 0.6% |

**Next Steps:**
- [ ] Obtain 3-5 vendor invoices to validate Receipt Amount composition (wastage embedded or separate)
- [ ] Consider drafting formal letter to Broadridge re: excess inventory (8.7 months buffer vs 2-3 month policy)
- [x] ~~Verify May-25 752K usage spike~~ — Resolved in session 13 (triple-counted billing data)
- [ ] Request actual Jun-20 purchase report from Broadridge

### 2026-02-23 (session 13 — part 3)

**Accomplished:**
- Investigated May-25 752K usage spike — confirmed as **triple-counted billing data**
- Three files in 2025 folder (`5 Apex May-22 Billing.xlsx`, `Updated V2`, `Updated V3`) all contained identical May 2025 data (195 APEX records, 250,802 envelopes each); 250,802 × 3 = 752,406 exact match
- Added `peek_billing_month()` function that reads the actual billing month from inside each file before dedup, replacing unreliable filename-based month parsing
- Discovered and recovered **Nov 2023 data** that was being incorrectly skipped — `Apex Billing November.xlsx` in the 2024 folder contains Nov 2023 data (misfiled), not a duplicate of `Apex November Billing.xlsx` (Nov 2024)
- Also correctly deduplicates `Apex Sept Billing Sheet Revised.xlsx` (identical to non-revised version, Sep 2024)

**Bug Fix 20 — Triple-counted May 2025 billing:**
- Three versions of the May 2025 billing file existed with different filenames
- No dedup logic existed for multiple non-master files covering the same billing month
- Added data-aware dedup: `peek_billing_month()` reads the first APEX record's Billing_Month/Billing_Year, groups files by actual billing month, keeps only the last file alphabetically per month
- Result: May-25 usage corrected from 752,406 to 250,802

**Bug Fix 21 — Nov 2023 data incorrectly skipped:**
- Initial filename-based dedup grouped `Apex Billing November.xlsx` (Nov 2023) with `Apex November Billing.xlsx` (Nov 2024) because both filenames contain "november"
- Data-aware dedup correctly identifies them as different months and processes both
- Result: Nov-23 usage recovered (326,113 envelopes, 396K purchased)

**Corrected totals (supersedes all prior sessions):**
| Metric | Value |
|--------|-------|
| Reconciliation Period | Jan 2020 – Dec 2025 |
| Purchase files processed | 63 |
| Billing files processed | 50 |
| Total Purchased | 30,065,500 |
| Total Used (Volume) | 26,373,417 |
| Net Variance | +3,692,083 (+12.3%) |
| Total Purchase Cost (vendor) | $1,994,141 |
| Total Invoiced Amount | $2,103,303 |

**All data quality issues resolved:**
| Issue | Status |
|-------|--------|
| 2025 markup discrepancy (31.8%) | Fixed — two parsing bugs; actual markup 10% |
| ENVCONRIDGE9X12DW deficit (-148%) | Explained — 2022 production surge, buffer stock |
| May-25 752K usage spike | Fixed — triple-counted billing file |
| Nov-23 missing data | Fixed — misfiled in 2024 folder, recovered |

**Next Steps:**
- [ ] Obtain 3-5 vendor invoices to validate Receipt Amount composition (wastage embedded or separate)
- [ ] Consider drafting formal letter to Broadridge re: excess inventory
- [ ] Request actual Jun-20 purchase report from Broadridge

### 2026-02-23 (session 14)

**Accomplished:**
- Fixed stale May-25 footnote in `generate_html_report.py` (line 951) — removed reference to "752K" triple-counted number
- Created `generate_broadridge_report.py` — new script that generates a clean, neutral, data-focused HTML report for sharing with Broadridge
- Output file: `Broadridge Envelope Reconciliation - For Review.html` (51.6 KB)

**Bug Fix 22 — Stale May-25 footnote:**
- Internal report still referenced "May-25 usage of 752K is 2.3x the trailing average" — the old triple-counted number from before Bug Fix 20
- Removed the stale investigation language; footnote now reads: "May & Jun 2025: Zero purchases confirmed (not missing data)."

**Broadridge-facing report structure:**
1. Summary (4 KPI cards: Purchased, Used, Variance, Months Covered + year-by-year table)
2. Monthly Trend (SVG bar chart — purchased vs used, post-settlement)
3. Monthly Detail (expanded by default — 46 months + 4 annual subtotals + grand total, with running balance)
4. Purchases & Usage by Envelope Type (grouped table + SKU-level breakdown)
5. Envelope Specifications (collapsed — WMS codes, sizes, mail types, abbreviation legend)

**Content excluded from Broadridge report (internal-only):**
- Bottom line assessment and recommendations
- Inventory gauge (buffer stock vs policy comparison)
- Rolling 6-month usage trend sparkline
- Key findings bullets
- Pre-settlement context line and $643K settlement reference
- Usage by product table
- Scope notes and investigation footnotes
- Avg Monthly Usage KPI and column
- Spoilage statistics

**New files:**
| File | Description |
|------|-------------|
| `generate_broadridge_report.py` | Reads `Envelope Reconciliation - Source Data.xlsx`, generates external-facing HTML |
| `Broadridge Envelope Reconciliation - For Review.html` | Self-contained HTML report for Broadridge review |

**How to refresh outputs:**
```bash
# Internal report
py -3 "C:\Users\smendoza\Projects\Broadridge Envelopes\generate_html_report.py"

# Broadridge-facing report
py -3 "C:\Users\smendoza\Projects\Broadridge Envelopes\generate_broadridge_report.py"
```

**Next Steps:**
- [ ] Obtain 3-5 vendor invoices to validate Receipt Amount composition (wastage embedded or separate)
- [ ] Consider drafting formal letter to Broadridge re: excess inventory
- [ ] Request actual Jun-20 purchase report from Broadridge
- [ ] Send Broadridge report to Broadridge contacts for data validation

### 2026-02-23 (session 14 continued)

**Accomplished:**
- Fixed variance mismatch between sections in Broadridge report — By Type section was showing full-period totals (30M/26.4M/+3.7M) while Summary and Monthly Detail showed post-settlement (21M/18.5M/+2.6M)
- Added monthly granularity to `By Envelope Type` and `Usage by Envelope Type` Excel tabs in `build_recon_from_source.py` — previously these were full-period aggregates with no month column
- Broadridge report now filters type data to post-settlement (Mar 2022+) before aggregating
- Internal report continues to use full-period type data (unchanged behavior)
- All sections in Broadridge report now consistently show: **21,039,500 purchased / 18,469,949 used / +2,569,551 variance (+12.2%)**

**Schema change — `By Envelope Type` tab:**
- Before: `Envelope Type | Total Purchased | Total Cost | Avg Unit Price | First Purchase | Last Purchase` (9 rows)
- After: `Month | Envelope Type | Purchased | Total Cost` (~350 rows, monthly detail)

**Schema change — `Usage by Envelope Type` tab:**
- Before: `Envelope Type | Total Envelopes Used` (8 rows)
- After: `Month | Envelope Type | Envelopes Used` (~400 rows, monthly detail)

**How to refresh outputs:**
```bash
# Step 1: Rebuild Excel (required if source data changes)
py -3 "C:\Users\smendoza\Projects\Broadridge Envelopes\build_recon_from_source.py"

# Step 2: Regenerate reports
py -3 "C:\Users\smendoza\Projects\Broadridge Envelopes\generate_html_report.py"
py -3 "C:\Users\smendoza\Projects\Broadridge Envelopes\generate_broadridge_report.py"
```

**Next Steps:**
- [ ] Obtain 3-5 vendor invoices to validate Receipt Amount composition (wastage embedded or separate)
- [ ] Consider drafting formal letter to Broadridge re: excess inventory
- [ ] Request actual Jun-20 purchase report from Broadridge
- [ ] Send Broadridge report to Broadridge contacts for data validation
