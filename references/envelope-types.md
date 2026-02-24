# Envelope Types Reference

## Active Envelope Types (7 SKUs)

| WMS Code | Mail Type | Category | Size |
|----------|-----------|----------|------|
| ENVMEAPEXN14PFC | Domestic Fold Statement | Statements | #14 |
| ENVMEAPEX9X12PFC | Domestic Flat Statement | Statements | 9x12 |
| ENVMERIDGEN14NI11/08 | Foreign Fold Statement | Statements | #14 |
| ENVMERIDGE9X12NI11/08 | Foreign Flat Statement | Statements | 9x12 |
| ENVCONRIDGE9X12DW | Flat Confirms (domestic + foreign) | Confirms | 9x12 |
| ENVAPXN10PFSCONN10IND(10/22) | Domestic Fold Confirms + Domestic Letters | Confirms/Letters | #10 |
| ENVCONPFSN10NI | Foreign Fold Confirms + Foreign Letters | Confirms/Letters | #10 |

## NI/PFC SKU Transition (Oct 2022)
- ENVAPXN10PFSCONN10IND(10/22) replaced ENVCONPFSN10NI in Oct 2022 for domestic mail
- Reason: postal permit update for better postage rate
- NI envelopes not retired — still used for foreign mail
- These two SKUs are interchangeable physically; combined inventory covers both

## Canonical Type Consolidation (35 variants -> 9 types)

| Canonical Type | Variants Merged | Total Purchased (full period) |
|----------------|----------------|-------------------------------|
| ENVMEAPEXN14PFC | 3 variants (incl. SUPPLIER ID, STMTPFC9/24) | 9,370,000 |
| ENVAPXN10 Confirms+Letters (PFC) | 9 variants (IND 10/22, 4/25, CNFPFC, LTRPFC) | 9,342,000 |
| ENVCONPFSN10NI | 4 variants (incl. SUPPLIER ID, UNITED ENVELOPE prefix, CNFNI) | 9,670,000 |
| ENVMEAPEX9X12PFC | 3 variants | 652,500 |
| ENVMERIDGEN14NI11/08 | 3 variants | 668,000 |
| ENVMERIDGE9X12NI11/08 | 2 variants | 65,000 |
| ENVCONRIDGE9X12DW | 2 variants | 80,000 |
| Tax Form Envelopes (1099/1099-R) | 5 variants | 210,000 |
| Tax Form Envelopes (1042/IRA) | 1 variant | 8,000 |

## Usage-to-Envelope-Type Mapping

Maps billing workbook fields (`Product`, `Flat_Fold`, `Address_Type`) to canonical envelope type. Date-aware: domestic #10 maps to NI before Oct 2022, PFC after.

| Product Category | Flat/Fold | Address Type | Envelope Type |
|---|---|---|---|
| STATEMENT | FOLD/MIXED/BULK | DOMESTIC/MIXED/other | ENVMEAPEXN14PFC |
| STATEMENT | FLAT | DOMESTIC/MIXED/other | ENVMEAPEX9X12PFC |
| STATEMENT | FOLD/MIXED/BULK | FOREIGN | ENVMERIDGEN14NI11/08 |
| STATEMENT | FLAT | FOREIGN | ENVMERIDGE9X12NI11/08 |
| CONFIRM/LETTER/CHECK | FOLD/MIXED/BULK | DOMESTIC/MIXED/other | ENVAPXN10 or ENVCONPFSN10NI (date-dependent) |
| CONFIRM/LETTER/CHECK | FLAT | any | ENVCONRIDGE9X12DW |
| CONFIRM/LETTER/CHECK | FOLD/MIXED/BULK | FOREIGN | ENVCONPFSN10NI |
| TAX DOCUMENT | any | any | Tax Form Envelopes (1099/1099-R) |

## Generic Stock Classification

All 7 Apex envelope types are standard double-window envelopes with no company logos, branding, or custom design:
- PFC (Pre-Sorted First-Class) indicia is a functional USPS postage marking, not client branding
- NI (No Imprint) envelopes are completely blank — usable by any Broadridge client
- All use 24WW paper, black ink, crosshatch/wood grain security tint
- Supplier: United Envelope LLC, Mt. Pocono, PA
- Classification as generic stock confirms: usage-based billing and lower wastage rate (5%/2%) apply per contract

## Envelope Spec Source Files

| File | Order # | WMS Code |
|------|---------|----------|
| `926131_APEX14PFC 712_rp.pdf` | 926131 | ENVMEAPEXN14PFC |
| `851251_APEX9X12PFC 712.pdf` | 830851 | ENVMEAPEX9X12PFC |
| `942095_RIDGEPLN14 (11-08)_rp.pdf` | 942095 | ENVMERIDGEN14NI11/08 |
| `823804_RIDGEPLN9X12_11_08..pdf` | 823804 | ENVMERIDGE9X12NI11/08 |
| `992124_PFS CON N10 IND (1022)_sp.pdf` | 992124 | ENVAPXN10PFSCONN10IND(10/22) |
| `856743_PFS CON N10 (0210).pdf` | 856743 | ENVCONPFSN10NI |
| `893283_RIDHE 9x12.pdf` | 818105 | ENVCONRIDGE9X12DW |

## Abbreviation Legend

| Code | Meaning |
|------|---------|
| PFC | Pre-Sorted First-Class (USPS postage imprint) |
| NI | No Imprint (blank, no postage marking) |
| DW | Double Window |
| IND | Indicia (postal permit number printed on envelope) |
| ME | Mailing Envelope |
| CON | Confirm envelope |

## ENVCONRIDGE9X12DW Deficit Explanation

80K purchased vs 199K used (-148%). Root cause: mid-2022 production surge during Ridge/Penson migration. Four months (Jan, Aug, Sep, Nov 2022) account for 88% of all-time flat confirm/letter usage (176,783 of 199K). Outside those months, FLAT confirms average ~350/month. Not actionable — 0.7% of total volume, covered by pre-existing buffer stock.

## Purchase Report Format Eras

| Era | Years | Format | Key Sheet | Notes |
|-----|-------|--------|-----------|-------|
| 0 | 2020-2021 | .xlsx | Month-named tab | Nested FY dirs, no Client column in some files, `Order Date`/`Qty`/`Total Price` headers |
| 1 | 2022 | .xlsx | Month-named tab | Simple flat table, `Markup%` is dollar amount |
| 2 | 2023 | .xlsm | `Final Data` | Added `Owner`, `Mark % 1` columns, 3-4 sheets |
| 3 | 2024 | .xlsm | `Final Data` | Same as Era 2, PO prefix changed to SPSPO |
| 4 | 2025 | .xlsx/.xlsm | `Purchase Report` or `Final Data` | Restructured columns, text quantities with commas, `Unit Cost`/`Total Cost` replace `Unit Price`/`Receipt Amount` |
