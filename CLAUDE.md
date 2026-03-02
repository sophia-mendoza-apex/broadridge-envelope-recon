# Broadridge Envelope Reconciliation Project

## Project Overview

Reconciliation of Apex Clearing's envelope purchases vs. usage from Broadridge Print & Mail Services, covering January 2020 through December 2025. Post-settlement scope (March 2022+) is the primary analysis window.

## Key Files

### Outputs (deliverables)
| File | Description |
|------|-------------|
| `Envelope Reconciliation Report.html` | **Internal** - Dark-theme executive dashboard with recommendations, wastage/billing analysis, inventory gauge |
| `Broadridge Envelope Reconciliation - For Review.html` | **External** - Print-ready light-theme report for Broadridge review and data validation |
| `Envelope Reconciliation - Source Data.xlsx` | 7-tab Excel workbook (Monthly Summary, Annual Summary, By Envelope Type, Usage by Envelope Type, Purchase Detail, plus others) |

### Scripts (rerunnable)
| File | Description |
|------|-------------|
| `build_recon_from_source.py` | Reads ~63 Purchase Reports + ~50 Billing Workbooks/Masters, consolidates into Source Data xlsx |
| `generate_html_report.py` | Reads Source Data xlsx, generates internal HTML report |
| `generate_broadridge_report.py` | Reads Source Data xlsx, generates Broadridge-facing HTML report |

### Source Data (read-only, do not modify)
| File | Location |
|------|----------|
| Purchase Reports (63 files) | `...\Print & Mail Reports\Envelope Purchase Orders\` (2022-2025) and `...\Purchase Reports (from email)\Purchase Reports\FY'20\` / `FY'21\` (2020-2021) |
| Billing Workbooks (50 files) | `...\Print & Mail Reports\Postage and Volume Report\{2022,2023,2024,2025}\` |
| Billing Master 2020-2021 | `Billing Master - 2020 - 2021.xlsx` in project directory |
| Consolidated PO history | `2019-2022 Purchases.xlsx` — supplementary source for gap-filling (Jun 2020) |
| Contracts | `GTO Print and Mail Services Schedule_Effective Jan-2019.pdf`, `Amendment No.1_Effective Jan-2024.pdf` |

### Superseded files (kept for reference)
| File | Description |
|------|-------------|
| `build_envelope_recon.py` | Earlier script that built recon from `P&M Postage and Material Recon.xlsx` |
| `Envelope Reconciliation Mar2022-Current.xlsx` | Earlier version built from summary workbook |
| `audit_script.py` | Contract compliance audit script (findings now in internal report) |

## How to Refresh Outputs

```bash
# Step 1: Rebuild Excel from source reports
py -3 "C:\Users\smendoza\Projects\Broadridge Envelopes\build_recon_from_source.py"

# Step 2: Regenerate reports
py -3 "C:\Users\smendoza\Projects\Broadridge Envelopes\generate_html_report.py"
py -3 "C:\Users\smendoza\Projects\Broadridge Envelopes\generate_broadridge_report.py"
```

## Current State

### Final Reconciliation Numbers (audited, all sections consistent)

**Post-settlement (Mar 2022 - Dec 2025):**
| Metric | Value |
|--------|-------|
| Purchased | 21,039,500 |
| Used | 18,469,949 |
| Wastage (Est., contractual max) | 690,713 |
| Adj. Variance | +1,878,838 (+8.9%) |
| Total Invoiced | $1,575,143 |
| Buffer stock | ~6.5 months (policy: 2-3 months) |

**Full period (Jan 2020 - Dec 2025):**
| Metric | Value |
|--------|-------|
| Purchased | 30,065,500 |
| Used | 26,373,417 |
| Vendor Cost | $1,994,141 |
| Invoiced | $2,103,303 |

### Year-by-Year (post-settlement)
| Year | Purchased | Used | Variance | Var% |
|------|-----------|------|----------|------|
| 2022 (Mar-Dec) | 5,807,000 | 4,629,923 | +1,177,077 | +20.3% |
| 2023 | 6,571,000 | 6,081,272 | +489,728 | +7.5% |
| 2024 | 4,621,000 | 4,348,349 | +272,651 | +5.9% |
| 2025 | 4,040,500 | 3,410,405 | +630,095 | +15.6% |

### Contract Rates (verified from source PDFs)

**Original Schedule (Jan 2019 - Dec 2023):**
- Materials billed at cost + wastage (no margin)
- Wastage: 10% generic paper, 5% generic envelopes

**Amendment No. 1 (Jan 2024 - present):**
- Materials billed at inventory cost + 10% margin
- Generic inventory cost = vendor price + wastage (10% continuous, 3% cutsheet, 2% envelopes)
- Client-specific inventory cost = vendor price (no wastage)
- Generic stock: billed on usage. Client-specific stock: billed on receipt.

### Key Findings
- Usage declined 39% from 462,992/mo (2022) to 284,200/mo (2025)
- Broadridge admits 10-15% operational wastage vs 2% contract limit for envelopes
- Broadridge classifies our envelopes as **client-specific** (receipt-based billing). Our position: unbranded standard envelopes should be **generic** (usage-based). Classification determines whether $192K billing difference is a discrepancy.
- Under client-specific classification, 10% margin with no wastage surcharge is contractually correct
- Post-settlement spoils: 55,733 of 18,469,949 = 0.30%
- Data confidence: ~95% overall totals, ~75% per-SKU breakdown

### Internal Report Structure
1. Executive Summary (bottom line, recommendations, KPIs, gauge, wastage/billing callouts, year-by-year, 2026 projection, pre-settlement context)
2. Buffer Stock by Envelope Type (overstock callout, per-SKU table, NI/PFC context)
3. Monthly Detail (interactive, collapsed, filterable by SKU)
4. Reference (contract terms, wastage quotes, summary table, envelope specs)

### Broadridge Report Structure
1. Summary (intro + 4 KPIs + year-by-year with wastage & invoiced)
2. Items for review (wastage observation + excess inventory + 2 confirmation items)
3. Purchases & usage by envelope type (4 groups + 9x12 footnote + NI/PFC context + 8 SKUs)
4. Reference (data sources + contract quotes + contract summary + Koebel quotes + pre-settlement context + generic stock classification)

## Next Steps

- [ ] **Send reply to Denci** re: client-specific vs generic classification (draft ready, pending review)
- [ ] Get Denci's answer on what makes unbranded envelopes client-specific
- [ ] Follow up on Edgewood materials audit timeline
- [ ] Confirm 2026 ordering adjusted for usage decline
- [ ] Obtain 3-5 vendor invoices to validate Receipt Amount composition
- [ ] Request actual Jun-20 purchase report from Broadridge
- [x] ~~Send Broadridge report to Broadridge contacts for data validation~~ (done, 2026-03-02)

## Reference Files

Detailed supporting documentation moved from session logs:

| File | Contents |
|------|----------|
| [`references/session-log.md`](references/session-log.md) | Full session-by-session history (sessions 1-19), variance progression |
| [`references/bug-fixes.md`](references/bug-fixes.md) | All 24 bug fixes with technical details and root causes |
| [`references/envelope-types.md`](references/envelope-types.md) | WMS codes, canonical type mapping, usage mapping, specs, format eras |
| [`references/email-analysis.md`](references/email-analysis.md) | Koebel/Denci findings, Broadridge contact map, operational facts |
| [`references/contract-terms.md`](references/contract-terms.md) | Pricing formulas, audit rights, settlement details, billing/wastage analysis |
