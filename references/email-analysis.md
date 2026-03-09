# Email Analysis & Broadridge Contacts

## Broadridge Contact Map

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

## Key Operational Facts from Brandon Koebel

### Buffer Stock & Inventory Policy
- Brandon (Nov 1, 2022): *"We typically try to keep a 2-3 month supply of envelopes in stock."*
- At ~700K envelopes/month (2022 levels), that's 1.4M-2.1M envelopes on hand

### "Usage" Definition
- Brandon (Oct 31, 2022): *"The usage report shows what was ordered from our inventory system and brought to the machines processing Apex jobs. For example, the first line shows 336,000 used but only 282,497 confirms mailed, which means aside from ~10% waste, the remainder of the envelopes are on the floor queued up to be used, also called 'surplus'."*
- Usage = envelopes pulled from warehouse to production floor, including waste and floor surplus

### Wastage Rates
- Brandon (Sep 29, 2022): *"did not account for any waste or spoilage (typically 10-15%)"*
- Brandon (Nov 7, 2022): *"Wastage is roughly 10%... This includes envelopes that are damaged, need to be reprinted and reinserted, etc."*

### Envelope Classification Dispute (Client-Specific vs Generic)

**Jun 26, 2023 — "Materials Follow Up" thread (Denci <-> Sophia):**
- Denci (initial): *"I confirmed that the Envelopes are billed on order and Paper is billed on usage. Paper and Envelopes are billed at cost plus %10 and 5% respectively per section 4 of the contract."*
- Sophia (pushback): *"we are charged by inventory purchase, not usage and we are being charged at a 10% markup not 5% per our contract per section 4."*
- Denci (response): *"The 5% is for generic envelope stock, however we are using **Apex specific envelopes** which run at 10%."*
- Note: Under the original schedule, client-specific stock = vendor price with NO wastage surcharge. The "10%" rate for "Apex specific envelopes" has no contractual basis. Denci appears to be applying the generic paper stock rate (10%) to envelopes.

**Aug 17, 2023 — "Re: Request for Definitions of Charges on Renewal Proposal" (Denci -> Terry Ray):**
- **Denci himself** (not Terry Ray as previously noted) provided pricing definitions in reply to Terry Ray's questions
- Section 8b (Paper markup): Quoted full contract language including both generic and client-specific billing models
- Section 11c (Envelopes - standard markup): *"Cost plus wastage for generic stock - specifically for envelopes that is 5%."* — described envelope pricing using ONLY the generic stock framework, no mention of client-specific classification
- Bottom note: *"Materials are billed at cost plus wastage for generic stock."*
- **Significance:** Denci applied generic stock pricing to our envelopes in Aug 2023, directly contradicting his Mar 2026 position that they are client-specific

**Mar 2, 2026 — "RE: Envelope reconciliation for review" (Denci -> Sophia):**
- Denci: *"Your envelopes are client specific inventory and are billed under that model."*
- Recommended switching to generic stock billing model, claiming 8-13% savings
- Quoted Amendment No. 1 language verbatim for billing definitions

**Summary:** Denci has consistently classified our envelopes as "Apex specific" / "client specific" since at least June 2023. However, his explanations have been inconsistent (citing a 10% rate that doesn't exist for client-specific stock under the contract). Meanwhile, our envelopes are unbranded, standard-size, double-window envelopes with no custom specs, which is the typical definition of generic stock.

### Wastage Pricing (Formal)
- Original schedule (Jan 2019): wastage is 10% for generic paper stock, 5% for generic envelope stock. No margin.
- Amendment No. 1 (Jan 2024): wastage is 10% continuous form, 3% cutsheet, 2% envelopes (for generic). Plus 10% margin on inventory cost. Client-specific = vendor price only (no wastage).
- Generic stock: unit rate billed based on **usage**
- Client-specific stock: unit rate billed based on **receipt**

## Envelope Operations & Logistics

| Source | Fact |
|--------|------|
| Sep 28, 2022 | Supplier: United Envelope LLC, Mt. Pocono PA. 1-2 trucks daily to 300 Executive Drive. Mailing at 51 Mercedes Way. |
| Sep 28, 2022 | Reorder on 1.5-month (small vol) or 3-month (high vol) trend. Yellow/Red inventory alerts. 20 envelope types in WMS. |
| Mar 31, 2023 | ENVAPXN10PFSCONN10IND(10/22) replaced ENVCONPFSN10NI in Oct 2022 — postal permit update for better postage rate |
| Apr 2, 2023 | Jan 2023 purchase report was missing one Ridge line item (legacy Apex name) |
| Apr 14, 2023 | PO#688124 cancelled. WMS-104050 was billed for 180,000 (not estimated qty) |
| May 5, 2023 | "Copy of Envelope Mapping.xlsx" attachment — maps envelope types to product names |
| May 12, 2023 | NI envelopes not retired, still used for foreign mail |
| Aug 28, 2023 | "Postage_No Volume Support" items = spoils (damaged pieces reprinted in separate jobs) |
| Sep 12, 2023 | Insert waste estimate: 10-20% added to projected monthly statement count |
| Sep 27, 2023 | Hand insertion cutoff: <200 envelopes = hand, >=200 = machine. Manifested mail used for statements (USPS discount, quicker). |
| Oct 26, 2023 | "Any months missing from the purchase reports means there were no applicable purchases in that month" |
| Nov 8, 2023 | 2018 data only available as QBR totals (no detailed backup) |
| Nov 30, 2023 | Account-level detail only stored for 60 days |
| Dec 11, 2023 | All envelopes are double-window (return address visible through window) |
| Dec 20, 2023 | ADS letters SOW signed late 2022, development mid-2023, live Sep 2023 |
| Jan 22, 2024 | Account-level report shows print accounts, pages, images, job names. Each line = one envelope. No reports tie individual account to postage. |
| Feb 13, 2024 | Volumes tracked in "mailing database" — envelopes, images, sheets, postage all sourced from it |

## Additional Email Findings (Session 12)

1. **Materials NOT included in annual fee** (May 31, 2023) — Sameer Jobanputra confirmed: Section 4 "Compensation" states materials (paper, envelopes, inserts) charged separately from the $475K annual fee

2. **Interim Extension Agreement executed Dec 27, 2023** — File: `apex extension 12-27 fully executed.pdf`. Negotiated by John Russ (Broadridge GRM), signed before formal Amendment No. 1

3. **PostEdge reconciliation limitations** (Aug 23, 2023) — 500K+ line items per month, over an hour to download. No domestic vs. foreign postage breakout, no image volume data. Matthew Burkavage (Broadridge) offered alternative reconciliation tools.

4. **PostEdge ICS credits - $63K expected** (Jan 4, 2024) — Andy Graf projected ~$650K in ICS credits; calculation estimated ~$63,475.92. Credits expected as credit memo, not cash payment.

5. **Insert SOW pricing details** (Sep 12, 2023) — Brandon Koebel: insert estimates add 10-20% for waste and new accounts. Insert paper: 24# stock (= 60# text), thicker to avoid machine jams. SOWs are estimates; actual billing per-unit. Logan Jones: Apex not currently upcharging inserts to correspondents.

## Email Threads Reviewed

| Thread | Dates | Key Content |
|--------|-------|-------------|
| Materials Follow Up | Jun 2023 | Denci <-> Sophia; Denci says "Apex specific envelopes" at 10%, Sophia challenges billing basis and rate |
| Request for Definitions of Charges on Renewal Proposal | Aug 2023 | Terry Ray -> Denci/Edelstein; formal charge definitions, describes materials as generic stock |
| Print and Mailing Services | May-Aug 2023 | Sameer Jobanputra; materials separate from annual fee |
| Interim Agreement | Dec 2023-Jan 2024 | Extension letter executed 12/27/23; renewal timeline |
| Apex/Broadridge Renewals | Feb 2024 | Terry Ray forward; references spoilage in context |
| May-2021 Confirm Postage Detail | Aug 2021 | Michael Schnupp; Address Letter Volume attachment |
| PostEdge Access | Aug 2023 | Matthew Burkavage; PostEdge limitations documented |
| BR Postedge (ICS credits) | Jan 2024 | Andy Graf; $63K credit estimate |
| Insert Grps Production Table | Apr-Sep 2023 | Insert pricing, SOW process, waste estimates |
| Ally Mass Mailing | Dec 2023 | United Envelope mentioned in context |

## Outlook Search Summary
- 130 emails found from Brandon Koebel and Christopher Denci
- No purchase reports or billing workbooks for 2025 found in Outlook from any sender
- `Purchase Reports.zip` from Oct 2023 contains FY2019-2021 only
- Most recent billing attachments from these senders: Jan 2024 (December 2023 billing)

## Client Names Matched in Reconciliation
- Purchase side: APEX, Apex, APEX/RIDGE, RIDGE/APEX, RIDGE, Ridge Clearing, Pension Financial, Pension financial, APEX BCC006392
- Usage side: APEX CLEARING (all 13,762 records)
