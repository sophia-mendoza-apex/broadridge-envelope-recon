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
- Broadridge classifies our envelopes as **client-specific** (receipt-based billing). Our position: unbranded standard envelopes should be **generic** (usage-based).
- **Classification dispute (Mar 9):** Denci gave four shifting answers, ultimately conceding "yes they are standard envelopes" (Mar 6 PM). Sophia sent email Mar 9 pinning concession to contract language, asking to confirm generic classification.
- **Financial impact of misclassification: $225,870 (14.3%)** -- computed comparing actual invoiced (receipt-based) vs generic terms (usage-based). Added to Broadridge report but NOT yet cited in emails (strategic: establish principle first, bring dollars later).
- **2023 unauthorized margin: $44,218** -- 10% markup applied all of 2023 before Amendment authorized it (Jan 2024). Separate issue from classification, not yet raised.
- **CPI does not apply to materials** -- Amendment explicitly excludes materials from CPI adjustments. No CPI escalation found in envelope unit rates.
- Post-settlement spoils: 55,733 of 18,469,949 = 0.30%
- Data confidence: ~95% overall totals, ~75% per-SKU breakdown

### Internal Report Structure
1. Executive Summary (bottom line, recommendations, KPIs, gauge, wastage/billing callouts, year-by-year, 2026 projection, pre-settlement context)
2. Buffer Stock by Envelope Type (overstock callout, per-SKU table, NI/PFC context)
3. Monthly Detail (interactive, collapsed, filterable by SKU)
4. Reference (contract terms, wastage quotes, summary table, envelope specs)

### Broadridge Report Structure
1. Summary (intro + 4 KPIs + year-by-year with wastage & invoiced)
2. Items for review (wastage observation + excess inventory + classification billing impact + 4 confirmation items)
3. Purchases & usage by envelope type (4 groups + 9x12 footnote + NI/PFC context + 8 SKUs)
4. Reference (data sources + contract quotes + contract summary + Koebel quotes + pre-settlement context + generic stock classification)

## Next Steps

- [ ] **Reply to Denci's Mar 9 email** — address the die line sample, separate #10 LTR (single-window) from all other types (double-window)
- [ ] **Confirm whether N10 LTR return address is vendor-preprinted or runtime-printed** — critical question for classification of that one SKU
- [ ] **Demand volume split** between N10 LTR (letters) vs N10 CON (confirms) — confirms are provably double-window/generic
- [ ] **Awaiting Edgewood audit completion date** — asked four times (Mar 6 AM, Mar 6 PM, Mar 9 AM, Mar 9 PM), no answer
- [ ] If Denci pushes back on classification: deploy Move 2 (financial impact $225,870 + 2023 margin $44,218 + his Aug 2023 "generic stock" email)
- [ ] Raise 2023 unauthorized margin ($44,218) as separate issue from classification
- [ ] Confirm 2026 ordering adjusted for usage decline (unanswered)
- [ ] Obtain 3-5 vendor invoices (envelopes + paper) to validate Receipt Amount composition
- [ ] Request actual Jun-20 purchase report from Broadridge
- [x] ~~Send Mar 9 email pinning "standard envelopes" concession, asking to confirm generic classification~~
- [x] ~~Update Broadridge report with classification billing impact section~~
- [x] ~~Send Broadridge report to Broadridge contacts for data validation~~ (done, 2026-03-02)

## Session Log

### 2026-03-09 — Session 20 (Envelope Classification Deep Dive)

**Accomplished:**
- **Analyzed Denci's Mar 9 reply** — two new arguments: (1) Apex code printed on envelopes, (2) operational machine setup procedure
- **Analyzed both envelope samples Denci sent:**
  - Die line (N10 LTR PFC 4/24): single-window, postage indicia "PAID APEX", NO return address on die line
  - Mar 3 finished piece: single-window with "Apex Clearing Corporation, PO BOX 9007" return address
- **Searched Outlook (Graph API)** for prior correspondence about envelope types, classification, and purchase history
- **Discovered critical distinction: CONFIRM vs LETTER envelopes**
  - All 7 existing spec files (from `Products and Envelope Samples` folder) show DOUBLE-WINDOW construction — return address shows through window from document, nothing client-specific on envelope
  - N10 LTR PFC (4/24) is a separate LETTER variant with SINGLE-WINDOW — return address either pre-printed or runtime-printed
- **Found Koebel Aug 2025 emails** confirming generic NI envelopes purchased and sprayed with indicia, then relabeled as APX SKUs
- **Found Koebel Apr 24, 2025 email** about IMB barcode truncation in envelope window — planning new revision (4/25) with expanded window, treats as routine operational update, no mention of client-specific classification
- **Found Denci Aug 17, 2023 email** where he described envelope pricing as "cost plus wastage for generic stock — specifically for envelopes that is 5%" — directly contradicts his Mar 2026 client-specific position
- **Corrected email-analysis.md** — Aug 2023 email was FROM Denci (not Terry Ray), making the generic stock characterization even more significant
- **Strategic assessment:** Even if N10 LTR is legitimately client-specific (pre-printed return address), 7 of 8 envelope types are provably generic based on manufacturer die lines. Best approach: concede LTR if needed, force reclassification on everything else.

**Key Evidence Catalog:**
| Evidence | Source | Impact |
|---|---|---|
| 7 envelope die lines — all double-window, no client info | `Products and Envelope Samples` folder | Proves generic for confirms/statements |
| N10 LTR die line — no return address shown | Denci Mar 9 attachment | Suggests return address is runtime-printed |
| Denci Aug 2023: "generic stock — 5% for envelopes" | Outlook search | Contradicts his Mar 2026 position |
| Koebel Aug 2025: NI envelopes sprayed and relabeled | Outlook search | Shows "APX" codes are internal labels |
| Koebel Apr 2025: window expansion = routine update | Outlook search | Treats envelope changes as operational |
| Denci Feb 2026: envelopes = "paper, boxes, etc." | Outlook search | Treats envelopes as materials |
| Koebel Nov 2024: "3 months inventory" target | Outlook search | Generic stock management practice |

**Open Question:**
- Is the N10 LTR return address vendor-preprinted or runtime-printed? This determines classification for that one SKU only. All other types are provably generic regardless.

**Confidence Assessment:** 75-80% overall. Strong on 7 of 8 envelope types. The N10 LTR is the only one with ambiguity.

## Reference Files

Detailed supporting documentation moved from session logs:

| File | Contents |
|------|----------|
| [`references/session-log.md`](references/session-log.md) | Full session-by-session history (sessions 1-19), variance progression |
| [`references/bug-fixes.md`](references/bug-fixes.md) | All 24 bug fixes with technical details and root causes |
| [`references/envelope-types.md`](references/envelope-types.md) | WMS codes, canonical type mapping, usage mapping, specs, format eras |
| [`references/email-analysis.md`](references/email-analysis.md) | Koebel/Denci findings, Broadridge contact map, operational facts |
| [`references/contract-terms.md`](references/contract-terms.md) | Pricing formulas, audit rights, settlement details, billing/wastage analysis |
