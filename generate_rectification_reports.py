#!/usr/bin/env python3
"""
Generates two HTML rectification reports from Source Data xlsx:
  1. Internal (dark theme) — full detail including cross-client evidence, strategic notes
  2. Broadridge (light theme) — security-filtered, framed as items for review
"""

import os
import sys
import io
import pandas as pd
from collections import defaultdict

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, "Envelope Reconciliation - Source Data.xlsx")
INTERNAL_HTML = os.path.join(BASE_DIR, "Print & Mail Billing Rectification Summary - Internal.html")
BROADRIDGE_HTML = os.path.join(BASE_DIR, "Print & Mail Billing Review - For Broadridge Review.html")

# ---------------------------------------------------------------------------
# Load data
# ---------------------------------------------------------------------------
monthly = pd.read_excel(EXCEL_PATH, sheet_name="Monthly Summary")
by_type_monthly = pd.read_excel(EXCEL_PATH, sheet_name="By Envelope Type")
usage_by_env_type_monthly = pd.read_excel(EXCEL_PATH, sheet_name="Usage by Envelope Type")
purchase_detail = pd.read_excel(EXCEL_PATH, sheet_name="Purchase Detail")

# Broadridge-provided production data
try:
    br_provided = pd.read_excel(EXCEL_PATH, sheet_name="Broadridge Provided (Apr 2026)")
except Exception:
    br_provided = None

print("Data loaded successfully.")

# ---------------------------------------------------------------------------
# Formatting helpers
# ---------------------------------------------------------------------------
DASH = "—"

def fmt_num(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return DASH
    try:
        return f"{int(v):,}"
    except (ValueError, TypeError):
        return str(v)

def fmt_money(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return DASH
    try:
        val = float(v)
        if val < 0:
            return f"(${abs(val):,.0f})"
        return f"${val:,.0f}"
    except (ValueError, TypeError):
        return str(v)

def fmt_pct(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return DASH
    try:
        val = float(v)
        if val == 0:
            return DASH
        return f"{val * 100:.1f}%"
    except (ValueError, TypeError):
        return str(v)

def safe(v, default=0):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return default
    try:
        return float(v)
    except (ValueError, TypeError):
        return default

# ---------------------------------------------------------------------------
# Data computation — post-settlement scope (Mar 2022 onward)
# ---------------------------------------------------------------------------
month_order = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]

def month_label_to_sortkey(label):
    parts = label.split('-')
    mi = month_order.index(parts[0])
    yi = int(parts[1])
    return (yi, mi)

SETTLEMENT_DATE = "Mar-22"
settlement_key = month_label_to_sortkey(SETTLEMENT_DATE)
post_mask = monthly["Month"].apply(lambda x: month_label_to_sortkey(x) >= settlement_key)
post = monthly[post_mask].copy()

post_purchased = int(post["Envelopes Purchased"].sum())
post_used = int(post["Envelopes Used (Volume)"].sum())

WASTAGE_CUTOVER = (24, 0)  # Jan-24
WASTAGE_PRE_2024 = 0.05
WASTAGE_POST_2024 = 0.02

def get_wastage_rate(month_label):
    return WASTAGE_POST_2024 if month_label_to_sortkey(month_label) >= WASTAGE_CUTOVER else WASTAGE_PRE_2024

post_wastage = sum(int(safe(r["Envelopes Used (Volume)"]) * get_wastage_rate(r["Month"])) for _, r in post.iterrows())
post_invoiced = post["Invoiced Amount"].sum()
post_cost = post["Purchase Cost"].sum()

# Year-by-year
post_yearly = defaultdict(lambda: {"purchased": 0, "used": 0, "months": 0,
                                    "cost": 0, "invoiced": 0, "wastage": 0})
for _, r in post.iterrows():
    yr = 2000 + int(r["Month"].split('-')[1])
    d = post_yearly[yr]
    d["purchased"] += safe(r["Envelopes Purchased"])
    d["used"] += safe(r["Envelopes Used (Volume)"])
    d["months"] += 1
    d["cost"] += safe(r["Purchase Cost"])
    d["invoiced"] += safe(r["Invoiced Amount"])
    d["wastage"] += int(safe(r["Envelopes Used (Volume)"]) * get_wastage_rate(r["Month"]))

# ---------------------------------------------------------------------------
# Finding 1: 2023 wastage rate (10% charged vs 5% correct)
# ---------------------------------------------------------------------------
post_2023 = post[post["Month"].str.endswith("-23")]
f1_months = []
f1_total_vendor = 0
f1_total_charged = 0
f1_total_correct = 0

for _, r in post_2023.iterrows():
    vendor = safe(r["Purchase Cost"])
    charged = safe(r["Invoiced Amount"])
    correct = vendor * 1.05
    overcharge = charged - correct
    f1_months.append({
        "month": r["Month"],
        "vendor": vendor,
        "charged": charged,
        "correct": correct,
        "overcharge": overcharge,
    })
    f1_total_vendor += vendor
    f1_total_charged += charged
    f1_total_correct += correct

f1_total_overcharge = f1_total_charged - f1_total_correct

# ---------------------------------------------------------------------------
# Finding 2: Classification / billing basis
# ---------------------------------------------------------------------------
generic_by_year = {}
actual_by_year = {}
generic_total = 0
_last_vendor_rate = 0

for _, r in post.iterrows():
    yr = 2000 + int(r["Month"].split('-')[1])
    purchased = safe(r["Envelopes Purchased"])
    used = safe(r["Envelopes Used (Volume)"])
    cost = safe(r["Purchase Cost"])
    invoiced = safe(r["Invoiced Amount"])

    if purchased > 0:
        _last_vendor_rate = cost / purchased

    if used > 0 and _last_vendor_rate > 0:
        if yr <= 2023:
            generic_bill = used * _last_vendor_rate * 1.05
        else:
            generic_bill = used * _last_vendor_rate * 1.02 * 1.10
    else:
        generic_bill = 0

    generic_by_year[yr] = generic_by_year.get(yr, 0) + generic_bill
    actual_by_year[yr] = actual_by_year.get(yr, 0) + invoiced
    generic_total += generic_bill

# For 2023, the "actual" for classification purposes uses the correct 5% rate
# to isolate classification impact from rate impact
actual_by_year_corrected = dict(actual_by_year)
actual_by_year_corrected[2023] = f1_total_correct  # vendor * 1.05

classification_total = sum(actual_by_year_corrected.get(yr, 0) for yr in actual_by_year) - generic_total
# Per-year classification
classification_by_year = {}
for yr in sorted(actual_by_year.keys()):
    if yr == 2023:
        classification_by_year[yr] = f1_total_correct - generic_by_year.get(yr, 0)
    else:
        classification_by_year[yr] = actual_by_year.get(yr, 0) - generic_by_year.get(yr, 0)

f2_total = sum(classification_by_year.values())

# ---------------------------------------------------------------------------
# Finding 3: CPI applied to materials (paper) — from D17 invoice analysis
# ---------------------------------------------------------------------------
# Amendment: "fees (other than materials) may be adjusted by CPI, max 4%"
# Paper is a material. CPI should NOT apply. But Broadridge applied CPI to paper
# rates starting Jan 2025: $0.0109 x 1.0289 CPI = $0.0112 (exact formula match).
f3_cpi_total = 4256  # $4,255.75 rounded — D17 paper items Jan 2025 - Mar 2026 (incl Rebranding invoice)
f3_cpi_by_year = {2025: 3101, 2026: 1155}
f3_cpi_rate_charged = 0.0112
f3_cpi_rate_correct = 0.0109
f3_cpi_pages = 12_475_407
f3_cpi_months = 15  # Jan 2025 through Mar 2026

combined_total = f1_total_overcharge + f2_total + f3_cpi_total

# ---------------------------------------------------------------------------
# Finding 4: Production wastage (from Broadridge-provided data)
# ---------------------------------------------------------------------------
production_rows = []
production_total_ordered = 0
production_total_used = 0
production_total_wastage = 0

if br_provided is not None:
    for _, r in br_provided.iterrows():
        desc = str(r.get("Description", ""))
        if desc == "nan" or desc.strip() == "":
            continue
        prod_usage = safe(r.get("Production Usage", 0))
        billed = safe(r.get("Billed Qty", 0))
        if prod_usage > 0:
            wastage = prod_usage - billed
            wastage_pct = wastage / billed if billed > 0 else 0
            production_rows.append({
                "sku": desc,
                "ordered": prod_usage,
                "used": billed,
                "wastage": wastage,
                "wastage_pct": wastage_pct,
            })
            production_total_ordered += prod_usage
            production_total_used += billed
            production_total_wastage += wastage

production_rows.sort(key=lambda x: -x["ordered"])
# Wastage % relative to billed quantity (matches contract convention: 2% surcharge on usage)
production_overall_pct = production_total_wastage / production_total_used if production_total_used else 0

# ---------------------------------------------------------------------------
# Side-by-side: Actual vs Correct billing
# ---------------------------------------------------------------------------
side_by_side = []
for yr in sorted(actual_by_year.keys()):
    d = post_yearly[yr]
    purchased = d["purchased"]
    used = d["used"]
    excess = purchased - used

    actual_vendor = d["cost"]
    actual_invoiced = d["invoiced"]
    actual_wastage_surcharge = actual_invoiced - actual_vendor if yr == 2023 else 0
    actual_margin = actual_invoiced - actual_vendor if yr >= 2024 else 0

    # Correct (generic, usage-based)
    if purchased > 0:
        vendor_rate = actual_vendor / purchased
    else:
        vendor_rate = 0
    correct_vendor = used * vendor_rate
    if yr <= 2023:
        correct_wastage = correct_vendor * 0.05
        correct_margin = 0
    else:
        correct_wastage = correct_vendor * 0.02
        correct_margin = (correct_vendor + correct_wastage) * 0.10
    correct_total = correct_vendor + correct_wastage + correct_margin

    side_by_side.append({
        "year": yr,
        "purchased": purchased,
        "used": used,
        "excess": excess,
        "excess_pct": excess / purchased * 100 if purchased else 0,
        "actual_vendor": actual_vendor,
        "actual_wastage": actual_wastage_surcharge,
        "actual_margin": actual_margin,
        "actual_total": actual_invoiced,
        "correct_vendor": correct_vendor,
        "correct_wastage": correct_wastage,
        "correct_margin": correct_margin,
        "correct_total": correct_total,
        "difference": actual_invoiced - correct_total,
    })

# Totals
sbs_totals = {
    "actual_vendor": sum(s["actual_vendor"] for s in side_by_side),
    "actual_wastage": sum(s["actual_wastage"] for s in side_by_side),
    "actual_margin": sum(s["actual_margin"] for s in side_by_side),
    "actual_total": sum(s["actual_total"] for s in side_by_side),
    "correct_vendor": sum(s["correct_vendor"] for s in side_by_side),
    "correct_wastage": sum(s["correct_wastage"] for s in side_by_side),
    "correct_margin": sum(s["correct_margin"] for s in side_by_side),
    "correct_total": sum(s["correct_total"] for s in side_by_side),
    "difference": sum(s["difference"] for s in side_by_side),
}


# =====================================================================
#  CSS THEMES
# =====================================================================

DARK_CSS = """*, *::before, *::after { box-sizing: border-box; }
body {
    margin: 0; padding: 0;
    font-family: "Helvetica Neue", Arial, sans-serif;
    font-size: 13px; line-height: 1.6; color: #E0E1E6; background: #131416;
}
.header {
    background: linear-gradient(135deg, #0A1A4A, #1A3A8F);
    color: #FFFFFF; padding: 44px 40px 36px;
}
.header h1 { margin: 0 0 6px; font-size: 28px; font-weight: 600; letter-spacing: -0.5px; }
.header .subtitle { font-size: 15px; opacity: 0.85; margin: 0 0 4px; }
.header .meta { font-size: 12px; opacity: 0.6; margin: 0; }
.content { max-width: 1200px; margin: 0 auto; padding: 32px 40px 60px; }
.section { margin-bottom: 36px; }
.section h2 {
    margin: 0 0 16px; font-size: 19px; font-weight: 600; color: #82B4FF;
    border-bottom: 2px solid #2A2B30; padding-bottom: 8px;
}
h3 { color: #9CC5FF; font-size: 15px; margin: 24px 0 12px; font-weight: 600; }
p { margin: 0 0 10px; }
blockquote {
    margin: 12px 0; padding: 12px 16px;
    background: #1E1F23; border-left: 3px solid #5B9BF7;
    border-radius: 0 6px 6px 0; font-size: 12px; line-height: 1.7; color: #C8C9CE;
}
.kpi-grid { display: flex; flex-wrap: wrap; gap: 16px; margin-bottom: 20px; }
.kpi-card {
    flex: 1 1 200px; background: #1E1F23; border-radius: 10px; padding: 22px;
    border: 1px solid #2A2B30; min-width: 180px;
}
.kpi-card .kpi-label {
    font-size: 11px; font-weight: 600; text-transform: uppercase;
    letter-spacing: 0.5px; color: #9A9BA0; margin: 0 0 6px;
}
.kpi-card .kpi-value { font-size: 26px; font-weight: 700; margin: 0; line-height: 1.2; }
.kpi-card .kpi-sub { font-size: 11px; color: #9A9BA0; margin: 5px 0 0; }
.table-wrap { border-radius: 6px; border: 1px solid #2A2B30; overflow: hidden; margin-bottom: 14px; }
table { width: 100%; border-collapse: collapse; font-size: 12px; background: #1E1F23; }
table th {
    background: #252630; color: #9CC5FF; padding: 9px 14px; text-align: left;
    font-weight: 600; font-size: 11px; text-transform: uppercase; letter-spacing: 0.3px;
    white-space: nowrap;
}
table th.num { text-align: right; }
table td { padding: 7px 14px; border-bottom: 1px solid #2A2B30; white-space: nowrap; }
table .num { text-align: right; font-variant-numeric: tabular-nums; }
table tbody tr:nth-child(even) { background: #1A1B1F; }
.total-row { background: #252630 !important; border-top: 2px solid #3A3B45; }
.callout {
    background: #1E1F23; border-radius: 8px; padding: 18px 22px;
    border: 1px solid #2A2B30; margin-bottom: 16px;
}
.callout-title {
    font-size: 11px; font-weight: 700; text-transform: uppercase;
    letter-spacing: 0.4px; color: #5B9BF7; margin: 0 0 10px;
}
.strategic-note {
    background: #1A1520; border: 1px solid #3A2A50; border-radius: 8px;
    padding: 18px 22px; margin-bottom: 16px;
}
.strategic-note .callout-title { color: #B388FF; }
.footer {
    text-align: center; padding: 24px 40px; font-size: 11px;
    color: #6A6B70; border-top: 1px solid #2A2B30;
}
.green { color: #4CAF79; } .red { color: #EF5350; } .amber { color: #FFB74D; }
.muted { color: #9A9BA0; }
@media print {
    body { background: #FFF; color: #333; }
    .header { padding: 24px 0 16px; }
    .content { padding: 16px 0; }
    .strategic-note { display: none; }
    table { font-size: 11px; }
    table th { background: #E0E0E0; color: #333; }
    table td { border-color: #CCC; }
    .kpi-card { border: 1px solid #CCC; background: #F5F5F5; }
}
"""

LIGHT_CSS = """*, *::before, *::after { box-sizing: border-box; }
body {
    margin: 0; padding: 0;
    font-family: "Helvetica Neue", Arial, sans-serif;
    font-size: 12px; line-height: 1.5; color: #333; background: #FFFFFF;
}
.header {
    background: linear-gradient(135deg, #052390, #1A3A8F);
    color: #FFFFFF; padding: 40px 40px 32px;
}
.header h1 { margin: 0 0 6px; font-size: 26px; font-weight: 600; letter-spacing: -0.5px; }
.header .subtitle { font-size: 14px; opacity: 0.85; margin: 0 0 4px; }
.header .meta { font-size: 12px; opacity: 0.65; margin: 0; }
.content { max-width: 1100px; margin: 0 auto; padding: 28px 40px 40px; }
.section { margin-bottom: 32px; page-break-inside: avoid; }
.section h2 {
    margin: 0 0 14px; font-size: 18px; font-weight: 600; color: #052390;
    border-bottom: 2px solid #052390; padding-bottom: 6px;
}
h3 { color: #052390; font-size: 14px; margin: 20px 0 10px; font-weight: 600; }
p { margin: 0 0 10px; }
blockquote {
    margin: 10px 0; padding: 10px 14px;
    background: #FFFFFF; border-left: 3px solid #052390;
    border-radius: 0 4px 4px 0; font-size: 12px; line-height: 1.7; color: #333;
}
.kpi-grid { display: flex; flex-wrap: wrap; gap: 14px; margin-bottom: 16px; }
.kpi-card {
    flex: 1 1 180px; background: #F5F5F7; border-radius: 8px; padding: 18px 20px;
    border: 1px solid #E2E2E2; min-width: 170px;
}
.kpi-card .kpi-label {
    font-size: 11px; font-weight: 600; text-transform: uppercase;
    letter-spacing: 0.5px; color: #6D6E71; margin: 0 0 6px;
}
.kpi-card .kpi-value { font-size: 24px; font-weight: 700; margin: 0; line-height: 1.2; }
.kpi-card .kpi-sub { font-size: 11px; color: #6D6E71; margin: 4px 0 0; }
.table-wrap { border-radius: 4px; border: 1px solid #E2E2E2; overflow: hidden; margin-bottom: 12px; }
table { width: 100%; border-collapse: collapse; font-size: 12px; background: #FFFFFF; }
table th {
    background: #052390; color: #FFFFFF; padding: 8px 12px; text-align: left;
    font-weight: 600; font-size: 11px; text-transform: uppercase; letter-spacing: 0.3px;
    white-space: nowrap;
}
table th.num { text-align: right; }
table td { padding: 6px 12px; border-bottom: 1px solid #E2E2E2; white-space: nowrap; color: #333; }
table .num { text-align: right; font-variant-numeric: tabular-nums; }
table tbody tr:nth-child(even) { background: #FAFAFA; }
.total-row { background: #E8E8ED !important; border-top: 2px solid #D0D0D8; }
.callout {
    background: #F5F5F7; border-radius: 8px; padding: 18px 22px;
    border: 1px solid #E2E2E2; margin-bottom: 16px;
}
.callout-title {
    font-size: 11px; font-weight: 700; text-transform: uppercase;
    letter-spacing: 0.4px; color: #052390; margin: 0 0 10px;
}
.footer {
    text-align: center; padding: 24px 40px; font-size: 11px;
    color: #6D6E71; border-top: 1px solid #E2E2E2;
}
.green { color: #186741; } .red { color: #9D1526; } .amber { color: #B8860B; }
.muted { color: #6D6E71; }
@media print {
    .header { padding: 28px 0 20px; }
    .content { padding: 16px 0 20px; }
    .kpi-card { box-shadow: none; }
    .table-wrap { box-shadow: none; border: 1px solid #CCC; page-break-inside: avoid; }
    blockquote, .kpi-grid, h3 { page-break-inside: avoid; }
    h3 { page-break-after: avoid; }
}
"""

# =====================================================================
#  Color helpers per theme
# =====================================================================
def var_color_dark(v):
    try:
        val = float(v)
        if val > 0: return "#4CAF79"
        if val < 0: return "#EF5350"
    except (ValueError, TypeError): pass
    return "#9A9BA0"

def var_color_light(v):
    try:
        val = float(v)
        if val > 0: return "#186741"
        if val < 0: return "#9D1526"
    except (ValueError, TypeError): pass
    return "#6D6E71"


# =====================================================================
#  REPORT 1: INTERNAL (Dark Theme — Full Detail)
# =====================================================================
def build_internal_report():
    vc = var_color_light
    h = '<!DOCTYPE html>\n<html lang="en">\n<head>\n'
    h += '<meta charset="UTF-8">\n'
    h += '<meta name="viewport" content="width=device-width, initial-scale=1.0">\n'
    h += '<title>Print & Mail Billing Rectification Summary - Internal</title>\n'
    h += f'<style>{LIGHT_CSS}</style>\n'
    h += '</head>\n<body>\n'

    # Header
    h += '<div class="header">\n'
    h += '  <div style="display:flex;justify-content:space-between;align-items:flex-start;">\n'
    h += '    <div>\n'
    h += '      <h1>Print &amp; mail billing rectification summary</h1>\n'
    h += '      <p class="subtitle">Post-settlement scope: March 2022 &ndash; March 2026</p>\n'
    h += f'      <p class="meta">Generated {pd.Timestamp.now().strftime("%B %d, %Y")}</p>\n'
    h += '    </div>\n'
    h += '    <div style="background:rgba(255,255,255,0.2);border:2px solid #FFFFFF;border-radius:6px;'
    h += 'padding:8px 18px;text-align:center;margin-top:4px;">\n'
    h += '      <p style="margin:0;font-size:13px;font-weight:700;color:#FFFFFF;letter-spacing:0.5px;">INTERNAL USE ONLY</p>\n'
    h += '    </div>\n'
    h += '  </div>\n'
    h += '</div>\n'
    h += '<div class="content">\n'

    # Internal-only banner
    h += '<div style="background:#FDE8E8;border:1px solid #9D1526;border-radius:6px;padding:10px 18px;'
    h += 'margin-bottom:24px;text-align:center;">\n'
    h += '  <p style="margin:0;font-size:12px;color:#9D1526;font-weight:600;">'
    h += 'FOR INTERNAL APEX USE ONLY</p>\n'
    h += '</div>\n'

    # KPI cards
    h += '<div class="section">\n'
    h += '  <div class="kpi-grid">\n'
    h += f'    <div class="kpi-card"><p class="kpi-label">Wrong wastage rate (2023)</p>'
    h += f'<p class="kpi-value red">{fmt_money(f1_total_overcharge)}</p>'
    h += f'<p class="kpi-sub">Charged 10% (paper rate) instead of 5% (envelope rate)</p></div>\n'
    h += f'    <div class="kpi-card"><p class="kpi-label">Wrong billing basis (2022&ndash;2025)</p>'
    h += f'<p class="kpi-value red">{fmt_money(f2_total)}</p>'
    h += f'<p class="kpi-sub">Billed on purchases instead of usage</p></div>\n'
    h += f'    <div class="kpi-card"><p class="kpi-label">CPI applied to materials</p>'
    h += f'<p class="kpi-value red">{fmt_money(f3_cpi_total)}</p>'
    h += f'<p class="kpi-sub">Paper rate inflated by CPI (Jan 2025+)</p></div>\n'
    h += f'    <div class="kpi-card"><p class="kpi-label">Total overcharge</p>'
    h += f'<p class="kpi-value" style="color:#9D1526;font-weight:800">{fmt_money(combined_total)}</p>'
    h += f'<p class="kpi-sub">Broadridge offer: $57K (Jan 2024 only)</p></div>\n'
    h += '  </div>\n'
    h += '</div>\n'

    # === Situation ===
    h += '<div class="section">\n'
    h += '  <h2>Background</h2>\n'
    h += '  <p>Broadridge has been billing our envelopes as &ldquo;client-specific&rdquo; stock, which means we pay based on '
    h += 'what they <strong>purchase</strong> from the supplier rather than what they actually <strong>use</strong> to mail our jobs. '
    h += 'We first raised this in <strong>June 2023</strong> and were told our envelopes were &ldquo;Apex specific.&rdquo; '
    h += 'They are not &mdash; they are standard, unbranded, double-window envelopes with no custom printing.</p>\n'
    h += '  <p>After months of back and forth, Denci acknowledged in March 2026 that &ldquo;yes they are standard envelopes&rdquo; '
    h += 'but maintained the classification is based on how they handle them operationally, not what the envelope actually is. '
    h += 'The contract makes no such distinction.</p>\n'
    h += '</div>\n'

    # === Finding 1 ===
    h += '<div class="section">\n'
    h += '  <h2>Finding 1: Wrong wastage rate in 2023</h2>\n'
    h += f'  <p>The contract has two wastage rates: 10% for paper, 5% for envelopes. Throughout 2023, '
    h += f'Broadridge charged the 10% paper rate on envelopes. Overcharge: <strong class="red">{fmt_money(f1_total_overcharge)}</strong>. '
    h += f'This is a straightforward contract error independent of the classification issue.</p>\n'
    h += '</div>\n'

    # === Finding 2 ===
    h += '<div class="section">\n'
    h += '  <h2>Finding 2: Billing on purchases instead of usage</h2>\n'
    h += '  <p>The contract says generic stock is billed on <strong>usage</strong> and client-specific on <strong>receipt</strong>. '
    h += 'Because Broadridge classifies our envelopes as client-specific, we pay for every envelope they order from the supplier &mdash; '
    h += f'including {fmt_num(int(post_purchased - post_used))} envelopes purchased but never used. '
    h += f'Broadridge\'s own production data shows <strong class="red">{production_overall_pct*100:.1f}%</strong> wastage '
    h += f'vs a 2% contract cap, confirming they routinely over-order.</p>\n'

    # Year-by-year table (compact)
    h += '  <div class="table-wrap"><table>\n'
    h += '    <thead><tr><th>Year</th><th class="num">Invoiced (receipt)</th>'
    h += '<th class="num">Correct (usage)</th><th class="num">Overcharge</th></tr></thead>\n'
    h += '    <tbody>\n'
    for yr in sorted(actual_by_year.keys()):
        act = f1_total_correct if yr == 2023 else actual_by_year[yr]
        gen = generic_by_year.get(yr, 0)
        diff = classification_by_year[yr]
        yr_label = f"{yr} (Mar&ndash;Dec)" if yr == 2022 else str(yr)
        h += f'      <tr><td>{yr_label}</td>'
        h += f'<td class="num">{fmt_money(act)}</td>'
        h += f'<td class="num">{fmt_money(gen)}</td>'
        h += f'<td class="num red">{fmt_money(diff)}</td></tr>\n'
    h += f'      <tr class="total-row"><td><strong>Total</strong></td>'
    h += f'<td class="num"><strong>{fmt_money(sum(f1_total_correct if yr == 2023 else actual_by_year[yr] for yr in actual_by_year))}</strong></td>'
    h += f'<td class="num"><strong>{fmt_money(generic_total)}</strong></td>'
    h += f'<td class="num red"><strong>{fmt_money(f2_total)}</strong></td></tr>\n'
    h += '    </tbody>\n'
    h += '  </table></div>\n'
    h += '</div>\n'

    # === Finding 3: CPI ===
    h += '<div class="section">\n'
    h += '  <h2>Finding 3: CPI applied to materials</h2>\n'
    h += f'  <p>The Amendment states &ldquo;fees (<strong>other than materials</strong>) may be adjusted by CPI.&rdquo; '
    h += f'Paper is a material. Starting January 2025, Broadridge applied CPI to the paper per-page rate: '
    h += f'charged ${f3_cpi_rate_charged:.4f}/page instead of ${f3_cpi_rate_correct:.4f}/page. '
    h += f'The math confirms it: $0.0109 &times; 1.0289 (CPI) = $0.0112 exactly. '
    h += f'Overcharge: <strong class="red">{fmt_money(f3_cpi_total)}</strong> across {f3_cpi_months} months '
    h += f'({fmt_num(f3_cpi_pages)} pages).</p>\n'
    h += '</div>\n'

    # === Broadridge's position vs ours ===
    h += '<div class="section">\n'
    h += '  <h2>Their offer vs our position</h2>\n'
    h += '  <div class="table-wrap"><table>\n'
    h += '    <thead><tr><th></th><th>Broadridge offer</th><th>Our position</th></tr></thead>\n'
    h += '    <tbody>\n'
    h += '      <tr><td>Period</td><td>Jan 2024 &ndash; Mar 2026</td><td>Mar 2022 &ndash; Mar 2026 (full post-settlement)</td></tr>\n'
    h += '      <tr><td>Methodology</td><td>Production wastage &times; avg cost</td><td>Billing basis correction + rate errors + CPI</td></tr>\n'
    h += '      <tr><td>What it covers</td><td>Envelopes pulled from warehouse but not mailed</td>'
    h += '<td>Difference between receipt-based and usage-based billing</td></tr>\n'
    h += '      <tr><td>What it misses</td><td>Over-ordering, 2022&ndash;2023, rate errors, CPI</td>'
    h += '<td>&mdash;</td></tr>\n'
    h += f'      <tr class="total-row"><td><strong>Credit amount</strong></td><td><strong>$57,000</strong></td>'
    h += f'<td class="red"><strong>{fmt_money(combined_total)}</strong></td></tr>\n'
    h += '    </tbody>\n'
    h += '  </table></div>\n'

    # Denci's calculation breakdown
    h += '  <div class="callout">\n'
    h += '    <p class="callout-title">How Broadridge calculated $57K</p>\n'
    h += '    <p>Denci provided production data (Jan 2024 &ndash; Mar 2026) showing envelopes pulled from warehouse '
    h += '(&ldquo;Production Usage&rdquo;) vs envelopes billed to Apex (&ldquo;Billed Qty&rdquo;). '
    h += 'His credit = the wastage gap &times; average cost per envelope:</p>\n'
    h += '    <div class="table-wrap"><table>\n'
    h += '      <thead><tr><th>Column</th><th>Denci&rsquo;s term</th><th class="num">Total</th><th>Our equivalent</th></tr></thead>\n'
    h += '      <tbody>\n'
    h += '        <tr><td>A</td><td>Purchases</td><td class="num">9,102,500</td><td>Purchased (invoice data)</td></tr>\n'
    h += '        <tr><td>B</td><td>Production Usage</td><td class="num">9,486,408</td><td>Pulled from warehouse to floor</td></tr>\n'
    h += '        <tr><td>C</td><td>Billed Qty</td><td class="num">8,685,680</td><td>Used (billing reports)</td></tr>\n'
    h += '        <tr><td>B&minus;C</td><td><strong>Wastage (his credit)</strong></td><td class="num"><strong>800,728</strong></td>'
    h += '<td>Production waste only</td></tr>\n'
    h += '        <tr><td>A&minus;C</td><td><em>Not credited</em></td><td class="num" style="color:#9D1526"><strong>416,820</strong></td>'
    h += '<td>Over-purchasing (ignored by Denci)</td></tr>\n'
    h += '      </tbody>\n'
    h += '    </table></div>\n'
    h += '    <p>His $57K covers row B&minus;C only. Our position covers the full A&minus;C gap across the entire post-settlement period, '
    h += 'plus the rate errors (Findings 1 and 3). Under usage-based billing, Apex should pay based on row C &mdash; '
    h += 'rows A and B are Broadridge&rsquo;s operational costs.</p>\n'
    h += '  </div>\n'
    h += '</div>\n'

    # === Summary ===
    h += '<div class="section">\n'
    h += '  <h2>Summary</h2>\n'
    h += '  <div class="table-wrap"><table>\n'
    h += '    <thead><tr><th>#</th><th>Finding</th><th class="num">Amount</th><th>Status</th></tr></thead>\n'
    h += '    <tbody>\n'
    h += f'      <tr><td>1</td><td>Wrong wastage rate (2023)</td>'
    h += f'<td class="num red"><strong>{fmt_money(f1_total_overcharge)}</strong></td>'
    h += f'<td>Clear contract error</td></tr>\n'
    h += f'      <tr><td>2</td><td>Billing on purchases instead of usage</td>'
    h += f'<td class="num red"><strong>{fmt_money(f2_total)}</strong></td>'
    h += f'<td>Classification established</td></tr>\n'
    h += f'      <tr><td>3</td><td>CPI applied to materials (paper)</td>'
    h += f'<td class="num red"><strong>{fmt_money(f3_cpi_total)}</strong></td>'
    h += f'<td>Clear contract exclusion</td></tr>\n'
    h += f'      <tr><td>4</td><td>Production wastage ({production_overall_pct*100:.1f}% vs 2% cap)</td>'
    h += f'<td class="num muted">&mdash;</td>'
    h += f'<td>Supporting evidence</td></tr>\n'
    h += f'      <tr class="total-row"><td></td><td><strong>Total</strong></td>'
    h += f'<td class="num" style="color:#9D1526;font-weight:700"><strong>{fmt_money(combined_total)}</strong></td>'
    h += f'<td></td></tr>\n'
    h += '    </tbody>\n'
    h += '  </table></div>\n'
    h += '</div>\n'

    # === Leverage ===
    h += '<div class="section">\n'
    h += '  <h2>Leverage</h2>\n'
    h += '  <div class="callout">\n'
    h += '    <ul style="line-height:2.0;padding-left:20px;margin:0;">\n'
    h += '      <li><strong>Denci\'s own words (Aug 2023 to Terry Ray):</strong> described envelope markup as '
    h += '&ldquo;cost plus wastage for generic stock &mdash; specifically for envelopes that is 5%.&rdquo; No mention of client-specific.</li>\n'
    h += '      <li><strong>Denci (Mar 2026):</strong> &ldquo;Yes they are standard envelopes.&rdquo;</li>\n'
    h += '      <li><strong>CPI applied to materials:</strong> Amendment states &ldquo;fees (other than materials) may be adjusted by CPI.&rdquo; '
    h += 'Materials are explicitly excluded from CPI escalation.</li>\n'
    h += '      <li><strong>MSA Section S (Records and Inspection):</strong> Apex has audit rights to inspect books and records '
    h += 'to verify service volumes and fees, including fees charged to Client.</li>\n'
    h += '    </ul>\n'
    h += '  </div>\n'
    h += '</div>\n'

    # Footer
    h += '</div><!-- content -->\n'
    h += '<div class="footer" style="background:#FDE8E8;border-top:2px solid #9D1526;">\n'
    h += '  <p style="color:#9D1526;font-weight:700;font-size:12px;margin:0 0 4px;">INTERNAL USE ONLY</p>\n'
    h += '  <p style="margin:0;">Apex Clearing Corporation</p>\n'
    h += '</div>\n'
    h += '</body>\n</html>\n'
    return h


# =====================================================================
#  REPORT 2: BROADRIDGE (Light Theme — Security Filtered)
# =====================================================================
def build_broadridge_report():
    vc = var_color_light
    h = '<!DOCTYPE html>\n<html lang="en">\n<head>\n'
    h += '<meta charset="UTF-8">\n'
    h += '<meta name="viewport" content="width=device-width, initial-scale=1.0">\n'
    h += '<title>Print & Mail Billing Review - For Broadridge Review</title>\n'
    h += f'<style>{LIGHT_CSS}</style>\n'
    h += '</head>\n<body>\n'

    # Header
    h += '<div class="header">\n'
    h += '  <h1>Print &amp; Mail Billing Review</h1>\n'
    h += '  <p class="subtitle">Post-settlement scope: March 2022 &ndash; March 2026</p>\n'
    h += f'  <p class="meta">From: Apex Clearing Corporation &mdash; {pd.Timestamp.now().strftime("%B %d, %Y")}</p>\n'
    h += '</div>\n'
    h += '<div class="content">\n'

    h += '<p style="font-size:13px;line-height:1.7;margin-bottom:20px;">This document presents billing discrepancies identified '
    h += 'during Apex\'s reconciliation of print and mail charges against the GTO Print and Mail Services Schedule (Jan 2019) '
    h += 'and Amendment No. 1 (Jan 2024).</p>\n'

    # Summary KPIs
    h += '<div class="section">\n'
    h += '  <h2>Summary</h2>\n'
    h += '  <div class="kpi-grid">\n'
    h += f'    <div class="kpi-card"><p class="kpi-label">Finding 1: Wastage rate</p>'
    h += f'<p class="kpi-value red">{fmt_money(f1_total_overcharge)}</p>'
    h += f'<p class="kpi-sub">10% paper rate charged on envelopes (2023)</p></div>\n'
    h += f'    <div class="kpi-card"><p class="kpi-label">Finding 2: Billing basis</p>'
    h += f'<p class="kpi-value red">{fmt_money(f2_total)}</p>'
    h += f'<p class="kpi-sub">Receipt-based vs usage-based (generic stock)</p></div>\n'
    h += f'    <div class="kpi-card"><p class="kpi-label">Finding 3: CPI on materials</p>'
    h += f'<p class="kpi-value red">{fmt_money(f3_cpi_total)}</p>'
    h += f'<p class="kpi-sub">CPI applied to paper (excluded by contract)</p></div>\n'
    h += f'    <div class="kpi-card"><p class="kpi-label">Combined rectification</p>'
    h += f'<p class="kpi-value" style="color:#9D1526;font-weight:800">{fmt_money(combined_total)}</p>'
    h += f'<p class="kpi-sub">Post-settlement (Mar 2022 &ndash; Mar 2026)</p></div>\n'
    h += '  </div>\n'
    h += '</div>\n'

    # === Finding 1: 2023 Wastage Rate ===
    h += '<div class="section">\n'
    h += '  <h2>Finding 1: Incorrect wastage rate applied throughout 2023</h2>\n'
    h += '  <div class="kpi-grid">\n'
    h += f'    <div class="kpi-card"><p class="kpi-label">Overcharge</p>'
    h += f'<p class="kpi-value red">{fmt_money(f1_total_overcharge)}</p>'
    h += f'<p class="kpi-sub">Difference between 10% charged and 5% contract rate</p></div>\n'
    h += f'    <div class="kpi-card"><p class="kpi-label">Contract rate</p>'
    h += f'<p class="kpi-value" style="color:#052390">5%</p>'
    h += f'<p class="kpi-sub">&ldquo;5% for generic envelope stock&rdquo;</p></div>\n'
    h += f'    <div class="kpi-card"><p class="kpi-label">Rate charged</p>'
    h += f'<p class="kpi-value red">10%</p>'
    h += f'<p class="kpi-sub">Paper rate applied to envelopes</p></div>\n'
    h += '  </div>\n'

    h += '  <div class="callout">\n'
    h += '    <p class="callout-title">Contract language</p>\n'
    h += '    <p>Original Schedule, Section 4:</p>\n'
    h += '    <blockquote>&ldquo;Materials are billed at cost plus wastage for generic stock. '
    h += 'Specifically, the wastage charge is <strong>10% for any generic paper stock</strong> and '
    h += '<strong>5% for generic envelope stock.</strong>&rdquo;</blockquote>\n'
    h += '    <p>The contract specifies two distinct rates. Throughout 2023, the 10% paper rate was applied to envelopes.</p>\n'
    h += '  </div>\n'

    # Month-by-month table
    h += '  <div class="table-wrap"><table>\n'
    h += '    <thead><tr><th>Month</th><th class="num">Vendor cost</th><th class="num">Charged (10%)</th>'
    h += '<th class="num">Correct (5%)</th><th class="num">Difference</th></tr></thead>\n'
    h += '    <tbody>\n'
    for m in f1_months:
        h += f'      <tr><td>{m["month"]}</td>'
        h += f'<td class="num">{fmt_money(m["vendor"])}</td>'
        h += f'<td class="num">{fmt_money(m["charged"])}</td>'
        h += f'<td class="num">{fmt_money(m["correct"])}</td>'
        h += f'<td class="num red">{fmt_money(m["overcharge"])}</td></tr>\n'
    h += f'      <tr class="total-row"><td><strong>Total</strong></td>'
    h += f'<td class="num"><strong>{fmt_money(f1_total_vendor)}</strong></td>'
    h += f'<td class="num"><strong>{fmt_money(f1_total_charged)}</strong></td>'
    h += f'<td class="num"><strong>{fmt_money(f1_total_correct)}</strong></td>'
    h += f'<td class="num red"><strong>{fmt_money(f1_total_overcharge)}</strong></td></tr>\n'
    h += '    </tbody>\n'
    h += '  </table></div>\n'
    h += f'  <p class="muted" style="font-size:11px;">We request a credit of <strong>{fmt_money(f1_total_overcharge)}</strong> '
    h += 'representing the difference between the paper rate (10%) and the contractual envelope rate (5%) for 2023.</p>\n'
    h += '</div>\n'

    # === Finding 2: Classification / billing basis ===
    h += '<div class="section">\n'
    h += '  <h2>Finding 2: Envelope classification and billing basis</h2>\n'

    h += '  <div class="callout">\n'
    h += '    <p class="callout-title">Contract language</p>\n'
    h += '    <p>Both the Original Schedule and Amendment No. 1 state:</p>\n'
    h += '    <blockquote>&ldquo;For <strong>generic stock</strong>, the unit rate will be billed based on <strong>usage</strong>. '
    h += 'For <strong>Client specific stock</strong>, the unit rate will be based on <strong>receipt</strong> of such stock.&rdquo;</blockquote>\n'
    h += '  </div>\n'

    h += '  <p>Apex\'s envelopes have been classified as client-specific stock and billed on receipt (purchase volume). '
    h += 'These are standard, unbranded envelopes that meet the contract definition of generic stock:</p>\n'

    h += '  <ul style="line-height:1.9;padding-left:20px;">\n'
    h += '    <li>Per manufacturer die line specifications on file, <strong>7 of 8 envelope types are standard double-window</strong> '
    h += 'with no logos, names, or client-specific printing.</li>\n'
    h += '    <li>The envelope sizes (#10, #14, 9x12) are <strong>industry-standard formats</strong> used across the financial services industry.</li>\n'
    h += '    <li>The single-window N10 LTR variant (first purchased May 2024) represents <strong>0.09%</strong> of #10 purchases.</li>\n'
    h += '  </ul>\n'

    h += '  <div class="callout">\n'
    h += '    <p class="callout-title">Classification history</p>\n'
    h += '    <p>This issue was first raised by Apex in <strong>June 2023</strong>. When we questioned why envelopes were being '
    h += 'charged at 10% instead of the contractual 5%, we were told our envelopes were &ldquo;Apex specific&rdquo; and not generic stock. '
    h += 'Subsequent discussions have established that Broadridge defines &ldquo;client-specific&rdquo; based on operational process '
    h += '(how envelopes are handled on the production floor), not the actual envelope itself. The contract defines classification '
    h += 'based on whether the stock is generic or custom &mdash; not based on operational handling.</p>\n'
    h += '  </div>\n'

    h += '  <p>Under generic classification, billing should be based on <strong>usage</strong> rather than receipt, '
    h += 'applying the contractual wastage surcharge and margin:</p>\n'

    # Year-by-year overcharge table
    h += '  <div class="table-wrap"><table>\n'
    h += '    <thead><tr><th>Year</th><th class="num">Actual invoiced (receipt)</th>'
    h += '<th class="num">Correct (usage-based)</th><th class="num">Difference</th></tr></thead>\n'
    h += '    <tbody>\n'
    for yr in sorted(actual_by_year.keys()):
        if yr == 2023:
            act = f1_total_correct
        else:
            act = actual_by_year[yr]
        gen = generic_by_year.get(yr, 0)
        diff = classification_by_year[yr]
        dc = vc(-diff)
        yr_label = f"{yr} (Mar&ndash;Dec)" if yr == 2022 else str(yr)
        h += f'      <tr><td>{yr_label}</td>'
        h += f'<td class="num">{fmt_money(act)}</td>'
        h += f'<td class="num">{fmt_money(gen)}</td>'
        h += f'<td class="num" style="color:{dc};font-weight:600">{fmt_money(diff)}</td></tr>\n'
    h += f'      <tr class="total-row"><td><strong>Total</strong></td>'
    corrected_actual_total = sum(f1_total_correct if yr == 2023 else actual_by_year[yr] for yr in actual_by_year)
    h += f'<td class="num"><strong>{fmt_money(corrected_actual_total)}</strong></td>'
    h += f'<td class="num"><strong>{fmt_money(generic_total)}</strong></td>'
    h += f'<td class="num red"><strong>{fmt_money(f2_total)}</strong></td></tr>\n'
    h += '    </tbody>\n'
    h += '  </table></div>\n'

    h += f'  <p class="muted" style="font-size:11px;">Generic billing computed as: usage x vendor unit rate x contractual wastage '
    h += f'(5% pre-2024, 2% post-2024) x margin (10% post-2024 per Amendment). '
    h += f'2023 actual uses corrected 5% envelope rate (see Finding 1). '
    h += f'The difference is driven primarily by billing on purchased volume vs actual usage &mdash; '
    h += f'post-settlement purchases exceeded usage by {fmt_num(post_purchased - post_used)} envelopes.</p>\n'
    h += '</div>\n'

    # === Finding 3: CPI applied to materials ===
    h += '<div class="section">\n'
    h += '  <h2>Finding 3: CPI applied to materials</h2>\n'

    h += '  <div class="callout">\n'
    h += '    <p class="callout-title">Contract language</p>\n'
    h += '    <p>Amendment No. 1, Section 2:</p>\n'
    h += '    <blockquote>&ldquo;Effective January 1, 2025, fees (<strong>other than materials</strong>) '
    h += 'may be adjusted (up or down) by Broadridge annually, by the average percentage increase or decrease '
    h += 'of the United States Consumer Price Index for Urban Consumers (the &ldquo;CPI&rdquo;)&hellip; '
    h += 'provided, however, in no case may Broadridge increase the fees by more than 4% per annum.&rdquo;</blockquote>\n'
    h += '    <p>Materials (paper, envelopes) are explicitly excluded from CPI adjustments.</p>\n'
    h += '  </div>\n'

    h += f'  <p>Starting January 2025, the paper per-page rate was increased from ${f3_cpi_rate_correct:.4f} to '
    h += f'${f3_cpi_rate_charged:.4f}. This increase matches the CPI formula exactly:</p>\n'

    h += '  <div class="table-wrap"><table>\n'
    h += '    <thead><tr><th>Period</th><th class="num">Rate charged</th><th class="num">Contract rate</th>'
    h += '<th class="num">Difference</th><th>Derivation</th></tr></thead>\n'
    h += '    <tbody>\n'
    h += '      <tr><td>2024 (baseline)</td><td class="num">$0.0109</td><td class="num">$0.0109</td>'
    h += '<td class="num">&mdash;</td><td>Vendor cost &times; 1.10 margin</td></tr>\n'
    h += '      <tr><td>Jan&ndash;Dec 2025</td><td class="num">$0.0112</td><td class="num">$0.0109</td>'
    h += '<td class="num red">+$0.0003</td><td>$0.0109 &times; 1.0289 (CPI) = $0.0112</td></tr>\n'
    h += '      <tr><td>Jan 2026</td><td class="num">$0.0115</td><td class="num">$0.0109</td>'
    h += '<td class="num red">+$0.0006</td><td>$0.0112 &times; 1.0271 (CPI) = $0.0115</td></tr>\n'
    h += '      <tr><td>Feb 2026+</td><td class="num">$0.0115</td><td class="num">$0.0112</td>'
    h += '<td class="num red">+$0.0003</td><td>New vendor cost, still CPI-inflated</td></tr>\n'
    h += '    </tbody>\n'
    h += '  </table></div>\n'

    h += f'  <p>Total overcharge across {f3_cpi_months} months ({fmt_num(f3_cpi_pages)} pages): '
    h += f'<strong class="red">{fmt_money(f3_cpi_total)}</strong>.</p>\n'
    h += f'  <p class="muted" style="font-size:11px;">As paper is a material, CPI should not be applied regardless of effective date. '
    h += f'The contract rate for 2025 should remain $0.0109 (same as 2024). '
    h += f'We request the paper rate be corrected and a credit issued for the difference.</p>\n'
    h += '</div>\n'

    # === Finding 4: Production wastage ===
    h += '<div class="section">\n'
    h += '  <h2>Finding 4: Production wastage exceeding contract limits</h2>\n'

    h += '  <div class="callout">\n'
    h += '    <p class="callout-title">Contract wastage rates</p>\n'
    h += '    <p>Original Schedule: <strong>5%</strong> for generic envelope stock.<br>'
    h += 'Amendment No. 1: <strong>2%</strong> for envelopes.</p>\n'
    h += '  </div>\n'

    h += '  <p>Broadridge-provided production data (April 14, 2026, covering Jan 2024 &ndash; Mar 2026) '
    h += 'shows the following envelope wastage:</p>\n'

    if production_rows:
        # Filter to PFC rows only (no NI variants for Broadridge version)
        pfc_rows = [pr for pr in production_rows if "NI" not in pr["sku"].upper()
                     or "CONPFS" in pr["sku"].upper()]
        # Actually show all — the data was provided by Denci, they know we have it
        h += '  <div class="table-wrap"><table>\n'
        h += '    <thead><tr><th>SKU</th><th class="num">Ordered for production</th>'
        h += '<th class="num">Used in production</th><th class="num">Wastage</th><th class="num">Wastage %</th></tr></thead>\n'
        h += '    <tbody>\n'
        for pr in production_rows:
            wpct_color = "#9D1526" if pr["wastage_pct"] > 0.05 else "#186741"
            h += f'      <tr><td>{pr["sku"]}</td>'
            h += f'<td class="num">{fmt_num(pr["ordered"])}</td>'
            h += f'<td class="num">{fmt_num(pr["used"])}</td>'
            h += f'<td class="num">{fmt_num(pr["wastage"])}</td>'
            h += f'<td class="num" style="color:{wpct_color};font-weight:600">{pr["wastage_pct"]*100:.1f}%</td></tr>\n'
        h += f'      <tr class="total-row"><td><strong>All envelopes</strong></td>'
        h += f'<td class="num"><strong>{fmt_num(production_total_ordered)}</strong></td>'
        h += f'<td class="num"><strong>{fmt_num(production_total_used)}</strong></td>'
        h += f'<td class="num"><strong>{fmt_num(production_total_wastage)}</strong></td>'
        h += f'<td class="num red"><strong>{production_overall_pct*100:.1f}%</strong></td></tr>\n'
        h += '    </tbody>\n'
        h += '  </table></div>\n'

    h += f'  <p>The blended production wastage rate of <strong>{production_overall_pct*100:.1f}%</strong> is '
    h += f'{production_overall_pct/0.02:.1f}x the current contractual cap of 2%. '
    h += 'Per Section 4, the wastage surcharge is intended to cover production losses. '
    h += 'Under the current receipt-based billing, Apex absorbs wastage beyond the contractual cap '
    h += 'because Broadridge over-orders to account for production losses and Apex pays for every envelope ordered. '
    h += 'Under generic (usage-based) billing, the 2% surcharge covers wastage and anything beyond 2% '
    h += 'is Broadridge\'s cost to bear.</p>\n'
    h += '</div>\n'

    # === Contract references ===
    h += '<div class="section">\n'
    h += '  <h2>Contract references</h2>\n'

    h += '  <h3>Original Schedule (Jan 2019), Section 4 &mdash; Compensation</h3>\n'
    h += '  <blockquote>&ldquo;Materials are billed at cost plus wastage for generic stock. '
    h += 'Specifically, the wastage charge is 10% for any generic paper stock and 5% for generic envelope stock. '
    h += 'For generic stock, the unit rate will be billed based on usage. '
    h += 'For Client specific stock, the unit rate will be based on receipt of such stock.&rdquo;</blockquote>\n'

    h += '  <h3>Amendment No. 1 (Jan 2024), Section 2 &mdash; replacing Section 4</h3>\n'
    h += '  <blockquote>&ldquo;Materials are billed at inventory cost plus 10% margin. '
    h += 'Inventory cost means for (i) Client specific inventory: vendor price; and (ii) generic inventory: '
    h += 'vendor price plus wastage as follows: 10% for continuous form, 3% for cutsheet, and 2% for envelopes. '
    h += 'For generic stock, the unit rate will be billed based on usage. '
    h += 'For Client specific stock, the unit rate will be based on receipt of such stock.&rdquo;</blockquote>\n'

    h += '  <h3>Amendment No. 1 (Jan 2024) &mdash; CPI adjustment</h3>\n'
    h += '  <blockquote>&ldquo;Effective January 1, 2025, fees (<strong>other than materials</strong>) may be adjusted '
    h += '(up or down) by Broadridge annually, by the average percentage increase or decrease of the United States '
    h += 'Consumer Price Index for Urban Consumers (the &ldquo;CPI&rdquo;)&hellip; provided, however, in no case may '
    h += 'Broadridge increase the fees by more than 4% per annum.&rdquo;</blockquote>\n'
    h += '  <p class="muted" style="font-size:11px;">Materials (paper, envelopes) are explicitly excluded from CPI adjustments.</p>\n'

    h += '  <h3>MSA, Section S &mdash; Records and Inspection</h3>\n'
    h += '  <blockquote>&ldquo;Broadridge shall maintain such books and records as are (a) necessary to demonstrate '
    h += 'Broadridge\'s compliance with its obligations under this Agreement, (b) necessary to verify Service volumes '
    h += 'and fees, and (c) necessary to comply with all applicable laws. [Broadridge shall provide access] for the '
    h += 'purposes of performing assessments and inspections of (i) Broadridge\'s compliance with the provisions of '
    h += 'this Agreement, including, without limitation, the fees charged to Client&hellip;&rdquo;</blockquote>\n'

    h += '</div>\n'

    # Footer
    h += '</div><!-- content -->\n'
    h += '<div class="footer">\n'
    h += '  Apex Clearing Corporation<br>\n'
    h += '  Prepared for discussion with Broadridge Financial Solutions\n'
    h += '</div>\n'
    h += '</body>\n</html>\n'
    return h


# =====================================================================
#  Generate both reports
# =====================================================================
print("Generating internal report...")
internal_html = build_internal_report()
with open(INTERNAL_HTML, "w", encoding="utf-8") as f:
    f.write(internal_html)
print(f"  -> {INTERNAL_HTML} ({len(internal_html):,} bytes)")

print("Generating Broadridge report...")
broadridge_html = build_broadridge_report()
with open(BROADRIDGE_HTML, "w", encoding="utf-8") as f:
    f.write(broadridge_html)
print(f"  -> {BROADRIDGE_HTML} ({len(broadridge_html):,} bytes)")

# Verification summary
print(f"\n--- Verification ---")
print(f"Finding 1 (wastage rate):     {fmt_money(f1_total_overcharge)}")
print(f"Finding 2 (classification):   {fmt_money(f2_total)}")
print(f"Combined:                     {fmt_money(combined_total)}")
print(f"Production wastage:           {production_overall_pct*100:.1f}%")
print(f"Post-settlement invoiced:     {fmt_money(post_invoiced)}")
print(f"Generic total (computed):     {fmt_money(generic_total)}")
print("Done.")
