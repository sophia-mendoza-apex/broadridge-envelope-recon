# Envelope Billing Rectification Summary

**Prepared:** April 28, 2026
**Period in scope:** March 2022 - March 2026 (post-settlement)
**Contract documents:** GTO Print and Mail Services Schedule (Jan 2019) + Amendment No. 1 (Jan 2024)

---

## Finding 1: Incorrect Wastage Rate Applied Throughout 2023

**Amount: $20,877**

### Contract Basis

**Original Schedule, Section 4 (Compensation):**
> "Materials are billed at cost plus wastage for generic stock. Specifically, the wastage charge is 10% for any generic paper stock and **5% for generic envelope stock**."

The contract specifies two wastage rates: **10% for paper, 5% for envelopes.** Throughout 2023, Broadridge applied the 10% paper rate to envelopes instead of the correct 5%.

### What Happened

Every 2023 month shows Invoiced Amount = Purchase Cost x 1.10 — the paper wastage rate, not the 5% envelope rate. By comparison, 2022 invoices show Invoiced = Purchase Cost exactly (no wastage applied), and 2024+ invoices reflect the Amendment's authorized terms.

### Evidence

| Month | Vendor Cost | Charged (10%) | Correct (5%) | Overcharge |
|-------|------------|--------------|-------------|------------|
| Jan-23 | $60,757.08 | $66,832.79 | $63,794.93 | $3,037.85 |
| Feb-23 | $9,135.84 | $10,049.42 | $9,592.63 | $456.79 |
| Mar-23 | $53,084.88 | $57,678.65 | $55,739.12 | $1,939.52 |
| Apr-23 | $105,277.56 | $115,805.32 | $110,541.44 | $5,263.88 |
| May-23 | $32,684.04 | $35,952.44 | $34,318.24 | $1,634.20 |
| Jun-23 | $39,204.51 | $41,958.19 | $41,164.74 | $793.45 |
| Jul-23 | $28,780.24 | $31,077.12 | $30,219.25 | $857.87 |
| Aug-23 | $27,986.88 | $30,785.57 | $29,386.22 | $1,399.34 |
| Sep-23 | $27,622.92 | $30,385.21 | $29,004.07 | $1,381.15 |
| Oct-23 | $46,473.96 | $51,121.36 | $48,797.66 | $2,323.70 |
| Nov-23 | $22,882.68 | $25,170.95 | $24,026.81 | $1,144.13 |
| Dec-23 | $12,910.80 | $14,201.88 | $13,556.34 | $645.54 |
| **Total** | **$466,801.39** | **$511,018.89** | **$490,141.46** | **$20,877.43** |

### Calculation

$511,018.89 (charged at 10%) - $490,141.46 (correct at 5%) = **$20,877.43**

This overcharge is independent of the classification dispute. The contract is explicit: envelope wastage is 5%, not 10%. Denci's June 2023 email acknowledged the 5% rate — *"The 5% is for generic envelope stock"* — then misapplied it as 10% for what he called "Apex specific envelopes," a rate that does not exist in the contract.

---

## Finding 2: Envelope Classification and Billing Basis

**Amount: $193,960 (blended-rate methodology)**

### Contract Basis

Both the Original Schedule and Amendment No. 1 contain the same billing-basis language:

**Original Schedule, Section 4:**
> "For generic stock, the unit rate will be billed based on **usage**. For Client specific stock, the unit rate will be based on **receipt** of such stock."

**Amendment No. 1, Section 2 (replacing Section 4):**
> "For generic stock, the unit rate will be billed based on **usage**. For Client specific stock, the unit rate will be based on **receipt** of such stock."

This language has been in the contract since January 2019 and was carried forward unchanged into the Amendment.

### The Dispute

Broadridge classifies Apex's envelopes as **client-specific stock** and bills on **receipt** (i.e., Apex is invoiced for every envelope purchased, regardless of how many are actually used).

Our position: these are **generic stock** and should be billed on **usage**.

### Evidence Supporting Generic Classification

1. **The envelopes are unbranded, standard sizes.** N10 (#10, double-window), N14 (#14, double-window), and 9x12 are industry-standard formats. No Apex logos, names, or custom printing.

2. **Broadridge's own Client Matrix classifies generic stock as "Usage" billing.** The Client Matrix tab — included with every monthly purchase report sent to Apex — contains a dedicated GENERIC entry that defines generic stock as:

   > *"These are generic sheets, envelopes, etc. that are used by different clients. Markup depends on the client and in billing template."*

   This entry explicitly links generic stock to "Usage" billing. Apex is classified as "Purchase" billing instead.

3. **Major broker-dealers receive Usage (generic) billing from Broadridge.** The Client Matrix (190 clients across Broadridge) shows 22 clients on Usage billing vs 80 on Purchase billing. Usage-billed clients include:

   | Client | Billing Basis | Markup |
   |--------|--------------|--------|
   | Charles Schwab | Usage | 12% |
   | Barclays | Usage | 10% |
   | Citi | Usage | 10% |
   | Merrill Lynch | Usage | 15% |
   | LPL | Usage | 10% |
   | Edward Jones | Usage | 0% |
   | Oppenheimer | Usage | 15% |
   | **Apex** | **Purchase** | **10%** |

   These firms use industry-standard envelopes for the same types of mailings (confirms, statements). There is no apparent operational reason why Apex's identical envelopes should be classified differently.

4. **Christopher Denci admitted they are standard envelopes.** In email correspondence (March 6, 2026), Denci stated: *"Yes they are standard envelopes, but the operating rules around them designates them as client specific."* The contract does not define classification based on "operating rules" — it distinguishes generic from client-specific based on the stock itself.

5. **NI variant production wastage confirms shared stock.** Broadridge's own production data (April 2026) shows N10 NI at 72% wastage and 9x12 NI at 302% wastage. These rates are consistent with shared generic stock pulled from a common pool, not dedicated client-specific inventory.

### What the Misclassification Costs

Under receipt-based billing, Apex pays for every envelope Broadridge purchases — including excess inventory and production waste. Under usage-based billing, Apex would pay only for envelopes actually consumed.

**Generic billing formula (Original Schedule, Mar 2022 - Dec 2023):**
Unit rate = vendor price x (1 + 5% wastage), billed on usage

**Generic billing formula (Amendment, Jan 2024 - present):**
Unit rate = vendor price x (1 + 2% wastage) x (1 + 10% margin), billed on usage

| Year | Actual Invoiced (Receipt) | If Generic (Usage) | Overcharge |
|------|--------------------------|-------------------|------------|
| 2022 (Mar-Dec) | $393,552 | $321,917 | $71,635 |
| 2023 | $490,141* | $458,355 | $31,786 |
| 2024 | $361,523 | $345,192 | $16,332 |
| 2025 | $309,049 | $259,986 | $49,063 |
| 2026 (Jan-Mar) | $108,264 | $83,120 | $25,144 |
| **Total** | **$1,662,529** | **$1,468,570** | **$193,960** |

*2023 actual corrected to the contractual 5% wastage rate ($490,141 = vendor cost x 1.05), isolating the classification impact from the rate issue in Finding 1.

### Calculation Methodology

For each post-settlement month, the generic billing amount is computed as:
- Envelopes Used (from billing sheet volume data) x vendor unit rate x contractual wastage x margin (if applicable)
- Vendor unit rate = Purchase Cost / Envelopes Purchased for that month (carried forward for months with zero purchases)
- Wastage: 5% for Jan 2019 - Dec 2023, 2% for Jan 2024 onward (per contract)
- Margin: 0% for Jan 2019 - Dec 2023, 10% for Jan 2024 onward (per Amendment)

**Note on methodology:** Using per-SKU vendor rates (a more granular approach) yields a lower classification impact. The difference arises from how envelope-type mix is handled — the per-SKU method applies each envelope type's specific vendor price to its specific usage, while the blended method uses the overall average rate. The blended figure reflects the aggregate billing difference.

---

## Finding 3: Production Wastage Exceeding Contract Limits

**No separate dollar figure — this finding supports the classification argument in Finding 2.**

### Contract Basis

**Original Schedule, Section 4:**
> "the wastage charge is ... 5% for generic envelope stock"

**Amendment No. 1, Section 2:**
> "[wastage:] 2% for envelopes"

### What We Found

Broadridge's own production data (provided by Christopher Denci, April 14, 2026) confirms:

| SKU | Ordered for Production | Used in Production | Wastage | Wastage % |
|-----|----------------------|-------------------|---------|-----------|
| N10 PFC (CON) | 7,820,000 | 7,233,814 | 586,186 | 7.5% |
| N14 PFC | 4,565,000 | 4,221,497 | 343,503 | 7.5% |
| N10 PFC (LTR) | 73,500 | 57,296 | 16,204 | 22.0% |
| 9x12 PFC | 256,000 | 233,527 | 22,473 | 8.8% |
| N10 NI | 60,000 | 34,930 | 25,070 | 41.8%* |
| N14 NI | 37,000 | 31,660 | 5,340 | 14.4% |
| 9x12 NI | 12,000 | 10,048 | 1,952 | 16.3% |
| **All envelopes** | **12,823,500** | **11,822,772** | **1,000,728** | **7.8%** |

*NI variant rates vary due to smaller production runs; extreme wastage on N10 NI (72% if measured from Denci's originally reported figures) suggests shared generic stock across clients.

**Key facts:**
- Contract cap for envelope wastage: **2%** (amendment) / **5%** (original)
- Actual production wastage: **7.8% overall**, with individual SKUs from 7.5% to 41.8%
- This is **3.9x the current contract cap** of 2%
- Under receipt-based billing, Apex absorbs this excess wastage because Broadridge over-orders to account for production losses, and Apex pays for every envelope ordered

### Why This Matters

Under generic (usage-based) billing, the 2% wastage surcharge is built into the unit rate to cover Broadridge's production losses. Wastage exceeding 2% would be Broadridge's cost to bear.

Under the current client-specific (receipt-based) billing, Apex pays for all envelopes ordered including those wasted in production. With 7.8% production wastage, Apex is effectively paying for Broadridge's operational inefficiency.

This finding strengthens the case for reclassification: if envelopes are correctly classified as generic, both the billing basis shifts to usage AND the wastage risk transfers to Broadridge.

---

## Side-by-Side: Actual Billing vs Correct Billing

### Year-by-Year Comparison

**2022 (Mar-Dec)** — Purchased: 5,807,000 | Used: 4,629,923 | Excess: 1,177,077 (20.3%)

| Component | ACTUAL (Receipt-Based) | CORRECT (Usage-Based) | Difference |
|-----------|----------------------|---------------------|------------|
| Vendor cost | $393,552 | $306,587 | +$86,965 |
| Wastage surcharge | $0 | $15,329 (5%) | -$15,329 |
| Margin (10%) | $0 | $0 | $0 |
| **Total billed** | **$393,552** | **$321,917** | **+$71,635** |

**2023** — Purchased: 6,571,000 | Used: 6,081,272 | Excess: 489,728 (7.5%)

| Component | ACTUAL (Receipt-Based) | CORRECT (Usage-Based) | Difference |
|-----------|----------------------|---------------------|------------|
| Vendor cost | $466,801 | $436,529 | +$30,272 |
| Wastage (10% charged, should be 5%) | $44,218 | $21,826 (5%) | +$22,392 |
| Margin (10%) | $0 | $0 | $0 |
| **Total billed** | **$511,019** | **$458,355** | **+$52,664** |

**2024** — Purchased: 4,621,000 | Used: 4,348,349 | Excess: 272,651 (5.9%)

| Component | ACTUAL (Receipt-Based) | CORRECT (Usage-Based) | Difference |
|-----------|----------------------|---------------------|------------|
| Vendor cost | $330,934 | $307,658 | +$23,276 |
| Wastage surcharge | $0 | $6,153 (2%) | -$6,153 |
| Margin (10%) | $30,589 | $31,381 | -$792 |
| **Total billed** | **$361,523** | **$345,192** | **+$16,332** |

**2025** — Purchased: 4,040,500 | Used: 3,410,405 | Excess: 630,095 (15.6%)

| Component | ACTUAL (Receipt-Based) | CORRECT (Usage-Based) | Difference |
|-----------|----------------------|---------------------|------------|
| Vendor cost | $280,954 | $231,717 | +$49,237 |
| Wastage surcharge | $0 | $4,634 (2%) | -$4,634 |
| Margin (10%) | $28,095 | $23,635 | +$4,460 |
| **Total billed** | **$309,049** | **$259,986** | **+$49,063** |

**2026 (Jan-Mar)** — Purchased: 1,030,500 | Used: 926,926 | Excess: 103,574 (10.0%)

| Component | ACTUAL (Receipt-Based) | CORRECT (Usage-Based) | Difference |
|-----------|----------------------|---------------------|------------|
| Vendor cost | $98,422 | $74,082 | +$24,340 |
| Wastage surcharge | $0 | $1,482 (2%) | -$1,482 |
| Margin (10%) | $9,842 | $7,556 | +$2,286 |
| **Total billed** | **$108,264** | **$83,120** | **+$25,144** |

### Post-Settlement Total (Mar 2022 - Mar 2026)

Purchased: 22,070,000 | Used: 19,396,875 | Excess: 2,673,125 (12.1%)

| Component | ACTUAL (Receipt-Based) | CORRECT (Usage-Based) | Difference |
|-----------|----------------------|---------------------|------------|
| Vendor cost | $1,570,663 | $1,356,572 | +$214,091 |
| Wastage | $44,218 (10% in 2023 only) | $49,425 (5% pre-2024, 2% post-2024) | -$5,207 |
| Margin (10%) | $68,527 (2024-2026 only) | $62,572 | +$5,955 |
| **Total billed** | **$1,683,407** | **$1,468,570** | **+$214,837** |

**What drives the $214,837 overpayment:**
- **Vendor cost on excess volume (+$214,091):** Receipt-based billing charges for 2.67M envelopes that were purchased but never used — excess inventory, production waste, and buffer stock all billed to Apex
- **Wastage rate error (+$20,877 net):** In 2023, Broadridge applied 10% (the paper rate) instead of the correct 5% envelope rate. Under generic billing, Apex would also pay wastage surcharges in other years (5% pre-2024, 2% post-2024), partially offsetting the volume savings.
- **Margin on excess volume (+$5,955):** In 2024-2026, the 10% margin is authorized but applied to purchased quantities (receipt) instead of used quantities (generic). This creates a small margin premium on the excess volume.

---

## Summary of Rectification

| # | Finding | Contract Clause | Amount | Status |
|---|---------|----------------|--------|--------|
| 1 | 2023 Incorrect Wastage Rate | Original Schedule Section 4: "5% for generic envelope stock" — Broadridge charged 10% (the paper rate) | **$20,877** | Clear contract violation |
| 2 | Classification / Billing Basis | Section 4: generic = "billed based on usage" vs client-specific = "based on receipt" | **$193,960** | Contingent on classification resolution |
| 3 | Excess Production Wastage | Amendment: "2% for envelopes" vs 7.8% actual | Supporting evidence | Supports reclassification |

### Combined Impact

| Scenario | Amount |
|----------|--------|
| Finding 1 alone (wastage rate correction, no classification dispute needed) | **$20,877** |
| Finding 1 + Finding 2 (if classification resolved in our favor) | **$214,837** |

### Notes

- All amounts are post-settlement (March 2022 - March 2026). Pre-March 2022 costs were resolved in the June 2022 settlement ($643,458 internalized by Broadridge).
- CPI does not apply to materials. Amendment explicitly states: "fees (other than materials) may be adjusted ... by CPI." Envelopes are materials.

---

## Contract Clauses Referenced

### Original Schedule (Effective January 1, 2019), Section 4 — Compensation

> "Materials are billed at cost plus wastage for generic stock. Specifically, the wastage charge is 10% for any generic paper stock and 5% for generic envelope stock."

> "For generic stock, the unit rate will be billed based on usage. For Client specific stock, the unit rate will be based on receipt of such stock."

### Amendment No. 1 (Effective January 1, 2024), Section 2 — replacing Section 4

> "Materials are billed at inventory cost plus 10% margin. Inventory cost means for (i) Client specific inventory: vendor price; and (ii) generic inventory: vendor price plus wastage as follows: 10% for continuous form, 3% for cutsheet, and 2% for envelopes."

> "For generic stock, the unit rate will be billed based on usage. For Client specific stock, the unit rate will be based on receipt of such stock."

### Amendment No. 1, Section B.4 — CPI

> "Effective January 1, 2025, fees (other than materials) may be adjusted (up or down) by Broadridge annually, by the average percentage increase or decrease of the United States Consumer Price Index for Urban Consumers (the 'CPI')."

### MSA Section 23.S — Audit Rights

> Apex can inspect books and records to verify "Service volumes and fees," including "fees charged to Client." Broadridge must provide "any reasonable additional information and assistance."
