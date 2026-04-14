# Meeting Prep: Denci Session — April 15, 2026

**Subject:** RE: Envelop Recon
**Attendees:** Christopher Denci, Fritz Louis-Charles, Brandon Koebel, Cameron Canter, Brian OToole (CC)
**Context:** Denci sent data table (Apr 14) — "APEX Item Usage and Purchases JAN 1 2024 - MAR 31 2026" with Production Usage, Billed Qty, Purchases, Inventory, and Delta columns

---

## 1. OPENING

Thank Denci for providing the data. Acknowledge that this is helpful and aligns with what we've been asking for (inventory counts in particular).

**Tone:** Collaborative, data-focused. We're "trying to reconcile" and "make sure we're on the same page." Not adversarial.

---

## 2. CLARIFYING QUESTIONS (ask these first — they're genuine unknowns)

### Column Definitions
- **"Production Usage"** — Is this the count of individual envelopes pulled from warehouse to the production floor? Does it include envelopes pulled for setup/calibration runs that are discarded?
- **"Billed Qty"** — This is Brandon's monthly billing report volume data, correct? Same data source as the billing workbooks we've been reconciling against?
- **"Delta"** — Is this calculated as (Production Usage - Billed Qty) / Billed Qty? Just want to confirm the formula.
  - *Why ask:* Confirming Production Usage > Billed Qty = wastage. Gets Denci to confirm the gap on record.

### Inventory
- The inventory column — is that across **all** warehouse locations (Edgewood, South Windsor, Coppell, Kansas City, El Dorado Hills, Canada)?
- Is this a point-in-time WMS snapshot as of March 31, or a physical count?
- **Total inventory is 735,036 envelopes** — at current usage rates that's about 2.3 months of buffer. That's within the 2-3 month policy Brandon described. Is that the target going forward?

### New People
- **Cameron Canter** is new to this thread. What's Cameron's role? (Understanding his function helps us know what Broadridge is bringing to the table.)

---

## 3. DATA VALIDATION (collaborative — "let's make sure we match")

### Purchases — We Match on 6 of 8 SKUs

| SKU | Denci (Jan 24–Mar 26) | Our Data (Jan 24–Dec 25) | Gap = Q1 2026? |
|-----|---:|---:|---|
| N10 CON PFC | 5,124,000 | 4,878,000 | +246K — reasonable |
| N10 LTR PFC | 12,000 | 12,000 | Exact match |
| N10 NI | 672,000 | 514,000 | +158K — reasonable |
| 9x12 DW | 10,000 | 10,000 | Exact match |
| 9x12 Stmt PFC | 187,500 | 142,500 | +45K — reasonable |
| 9x12 NI | 15,000 | 15,000 | Exact match |
| N14 PFC | 2,842,000 | 2,842,000 | Exact match |
| **N14 NI** | **240,000** | **248,000** | **We show 8K MORE** |

**To raise:** "On the N14 NI — we show 248,000 in purchases for Jan 24 through Dec 25, but your total through Mar 26 is 240,000. We have a Sep 2024 purchase order for 8,000 envelopes under the updated WMS code ENVAPXN14APEXN14STMTNI9/24. Was that captured in your data?" (This is non-confrontational — just a data reconciliation question.)

### Billed Qty — Generally Consistent
Our "Envelopes Used (Volume)" from Brandon's billing workbooks should match Denci's "Billed Qty" for the overlapping period. Denci's data includes Q1 2026 which we don't have. The totals look directionally consistent.

**Don't challenge the billed qty numbers.** These are Brandon's reports — both sides are using the same source.

---

## 4. STRATEGIC QUESTIONS (advance our position without showing our hand)

### Production Usage / Wastage

The delta column shows Production Usage exceeds Billed Qty by varying amounts. This is the actual operational wastage.

**Key question:** *"The delta between production usage and billed quantity — is that what you'd call operational wastage? Envelopes that are pulled from inventory but don't make it into a mailpiece?"*
- Gets Denci to define it on record.
- If yes, the blended wastage rate across all SKUs is **9.2%** — well above the 2% contractual wastage for envelopes.

**Follow-up:** *"The N10 NI shows 72% delta and the 9x12 NI shows a similar pattern. Are those shared across multiple clients on the production floor? Could that explain the higher pull-through rate?"*
- **Why this matters:** If NI envelopes are pulled for multiple clients' jobs, that confirms they're generic shared stock — exactly our classification argument. Denci may not realize the implication.

**Follow-up (9x12 PFC):** *"The 9x12 PFC shows a negative delta — billed quantity is higher than production usage. What would cause that? Is there a timing lag between warehouse pulls and billing?"*
- This is a genuine anomaly. Billed 166,616 but only 126,369 pulled. We should understand it.

### Time Period

**Do NOT challenge the Jan 2024 start date directly.** Instead, if it comes up naturally:
- *"This is helpful for 2024 forward. We've been looking at the full post-settlement period from March 2022 to get the complete picture. Would you be able to pull this same view for the earlier period?"*
- Positions it as a data completeness request, not an accusation.

### Classification / Billing Basis

**Don't raise classification proactively in this meeting.** Denci offered to switch to usage-based on Mar 10. Terry/Fritz haven't responded yet. Let Denci bring it up. If he does:
- *"We're still discussing that internally. This data is very helpful context for that conversation."*
- Does not commit to anything. Keeps the door open.

If Denci pushes on "finishing off the Apex-coded stock" before switching:
- *"How much Apex-coded stock remains? Is the 735,000 in inventory all Apex-coded?"*
- Gets him to quantify what "finishing off" means. At current usage, 735K = 2.3 months. That's a defined timeline.

---

## 5. DO NOT MENTION (reserve ammunition)

| Item | Why Hold Back |
|------|---------------|
| $225,870 total overcharge calculation | Strategic reserve — establish principle before showing dollars |
| $44,218 unauthorized 2023 margin | Separate issue; his data conveniently starts Jan 2024 and skips this entirely |
| Denci's Aug 2023 email ("generic stock — 5% for envelopes") | Strongest contradiction of his position — save for escalation |
| Denci's Jun 2023 email (10% rate that doesn't exist for client-specific) | Same — reserve |
| The specific overcharge by year | Don't break it down for them |
| Our scenario analysis (actual vs client-specific vs generic) | Don't share methodology |
| MSA Section 23.S audit rights | Formal lever — don't invoke unless negotiation stalls |

---

## 6. THINGS TO LISTEN FOR

| If Denci says... | It means... | Your move |
|------------------|-------------|-----------|
| "Let's focus on going forward" | He wants to avoid retroactive adjustment | Note it but don't agree. "We want to make sure the historical period is clean too." |
| "The production wastage is normal for our operation" | He's acknowledging 9.2% wastage as standard | Ask: "But the contract specifies 2% for envelopes — is that a number that should be updated?" Gets him to confirm the gap on record. |
| "These envelopes are coded to Apex in WMS" | WMS codes are internal labels, not contract definitions | Don't argue — just note it. Koebel's Aug 2025 emails already show NI envelopes are purchased generic and relabeled. |
| "Cameron handles production operations" | Explains why Canter was added | Good — means they're bringing operational expertise, suggests they're taking this seriously. |
| "We can switch to usage-based now" | Faster than expected | Ask about retroactive adjustment: "Would that apply to the historical period as well, or just going forward?" |
| Any mention of ordering/purchasing adjustments | Shows they know the cadence is off | Positive signal. Note the commitment. |

---

## 7. DESIRED OUTCOMES FOR THIS MEETING

**Minimum:**
1. Confirm column definitions (especially Production Usage = warehouse pulls)
2. Confirm inventory is all-location, not just one warehouse
3. Get Denci to acknowledge the Production Usage > Billed Qty gap as wastage
4. Understand the N10 NI / 9x12 NI extreme deltas (shared stock?)
5. Understand 9x12 PFC negative delta

**Stretch:**
6. Get Denci to explain the Jan 2024 start date and offer to extend the view
7. Get a timeline for the usage-based switch
8. Learn what "finishing off Apex-coded stock" means in months
9. Get Cameron Canter's role and perspective on production operations

---

## 8. QUICK REFERENCE — OUR NUMBERS

**Full post-settlement (Mar 2022 – Dec 2025):**
- Purchased: 21,039,500 | Used: 18,469,949 | Variance: +1,878,838 (+8.9%)
- Total invoiced: $1,575,143
- If billed as generic (usage-based): $1,349,273 → saves $225,870

**Jan 2024 – Dec 2025 only (Denci's overlapping period):**
- Purchased: 8,661,500 | Used: 7,758,754 | Variance: +902,746 (+10.4%)
- Total invoiced: $611,887 (vendor cost; invoiced amounts higher with margin)

**Contract wastage: 2% for envelopes (Amendment)**
**Denci's data shows: 9.2% blended actual wastage**

**Denci's five shifting answers on classification:**
1. Mar 2: "Client-specific" (no explanation)
2. Mar 3: "Apex branding" (dropped when challenged)
3. Mar 6 AM: Operational segregation (admitted same envelopes used for next client)
4. Mar 6 PM: "Yes they are standard envelopes" but operating rules make them client-specific
5. Mar 10: "It doesn't matter what envelope we are using or what it looks like" — purely operational handling argument

---

## 9. AFTER THE MEETING

- Document any commitments Denci makes (especially on timelines, classification, data requests)
- If he provides additional data, cross-reference against our full post-settlement analysis
- Update Terry/Fritz on outcomes
- If Denci agrees to extend the data view to Mar 2022, request it in writing
- Do NOT send any financial impact numbers until Terry/Fritz approve the approach
