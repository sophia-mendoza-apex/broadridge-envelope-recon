#!/usr/bin/env python3
"""
Broadridge Envelope Reconciliation - External Report Generator
Reads source data from Excel and generates a clean, data-focused HTML report
suitable for sharing with Broadridge for reconciliation review.
"""

import os
import pandas as pd

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, "Envelope Reconciliation - Source Data.xlsx")
HTML_PATH = os.path.join(BASE_DIR, "Broadridge Envelope Reconciliation - For Review.html")

monthly = pd.read_excel(EXCEL_PATH, sheet_name="Monthly Summary")
by_type_monthly = pd.read_excel(EXCEL_PATH, sheet_name="By Envelope Type")
usage_by_env_type_monthly = pd.read_excel(EXCEL_PATH, sheet_name="Usage by Envelope Type")

print("Data loaded successfully.")

# ---------------------------------------------------------------------------
# Formatting helpers
# ---------------------------------------------------------------------------
DASH = "\u2014"

def fmt_num(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return DASH
    try:
        return f"{int(v):,}"
    except (ValueError, TypeError):
        return str(v)

def fmt_num_parens(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return DASH
    try:
        val = int(v)
        if val < 0:
            return f"({abs(val):,})"
        return f"{val:,}"
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

def fmt_buffer_months(variance, avg_monthly_usage):
    """Format buffer months from variance / avg monthly usage."""
    if avg_monthly_usage == 0:
        return DASH
    mo = variance / avg_monthly_usage
    if mo >= 0:
        return f"{mo:.1f}"
    return f"({abs(mo):.1f})"

def buffer_color(variance, avg_monthly_usage):
    """Color buffer months: red if negative, orange if >3 months (Broadridge policy), green if 0-3."""
    if avg_monthly_usage == 0:
        return "#6D6E71"
    mo = variance / avg_monthly_usage
    if mo < 0:
        return "#9D1526"
    if mo > 3:
        return "#B8860B"
    return "#186741"

def var_color(v):
    try:
        val = float(v)
        if val > 0:
            return "#186741"
        if val < 0:
            return "#9D1526"
    except (ValueError, TypeError):
        pass
    return "#6D6E71"

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

def safe(v, default=0):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return default
    try:
        return float(v)
    except (ValueError, TypeError):
        return default

# ---------------------------------------------------------------------------
# Data — post-settlement scope (Mar 2022 onward)
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
post = monthly[post_mask]

post_purchased = int(post["Envelopes Purchased"].sum())
post_used = int(post["Envelopes Used (Volume)"].sum())
post_months = len(post)

# Wastage rates per contract
WASTAGE_CUTOVER = (24, 0)  # Jan-24
WASTAGE_PRE_2024 = 0.05
WASTAGE_POST_2024 = 0.02

def get_wastage_rate(month_label):
    return WASTAGE_POST_2024 if month_label_to_sortkey(month_label) >= WASTAGE_CUTOVER else WASTAGE_PRE_2024

# Compute total wastage allowance
post_wastage = sum(int(safe(r["Envelopes Used (Volume)"]) * get_wastage_rate(r["Month"])) for _, r in post.iterrows())
post_adj_used = post_used + post_wastage
post_variance = post_purchased - post_adj_used  # wastage-adjusted

# Year-by-year post-settlement
from collections import defaultdict as _defaultdict
post_yearly = _defaultdict(lambda: [0, 0, 0, 0, 0, 0])  # purchased, used, month_count, cost, invoiced, wastage
for _, r in post.iterrows():
    yr = 2000 + int(r["Month"].split('-')[1])
    u = safe(r["Envelopes Purchased"])
    v = safe(r["Envelopes Used (Volume)"])
    w = int(v * get_wastage_rate(r["Month"]))
    post_yearly[yr][0] += u
    post_yearly[yr][1] += v
    post_yearly[yr][2] += 1
    post_yearly[yr][3] += safe(r["Purchase Cost"])
    post_yearly[yr][4] += safe(r["Invoiced Amount"])
    post_yearly[yr][5] += w

# Post-settlement cost totals
post_invoiced = post["Invoiced Amount"].sum()
post_cost = post["Purchase Cost"].sum()

# Full-period totals (for context)
total_purchased = int(monthly["Envelopes Purchased"].sum())
total_used = int(monthly["Envelopes Used (Volume)"].sum())
total_months = len(monthly)
net_variance = total_purchased - total_used

# Pre-settlement totals
pre_purchased = total_purchased - post_purchased
pre_used = total_used - post_used

# Source file counts (from build_recon_from_source.py output)
purchase_files = 63
billing_files = 50

# Wastage analysis — contract max vs Broadridge-confirmed operational rates
_pre24_used = sum(post_yearly[y][1] for y in post_yearly if y < 2024)
_post24_used = sum(post_yearly[y][1] for y in post_yearly if y >= 2024)
_contract_waste = int(_pre24_used * 0.05 + _post24_used * 0.02)
_actual_waste_lo = int(post_used * 0.10)
_actual_waste_hi = int(post_used * 0.15)
_excess_lo = _actual_waste_lo - _contract_waste
_excess_hi = _actual_waste_hi - _contract_waste

# Buffer stock / excess inventory analysis
# Trailing 12-month average usage for buffer coverage (captures full seasonal cycle)
_last_12 = post.tail(12)
_trailing_avg = int(_last_12["Envelopes Used (Volume)"].mean())
_buffer_months = post_variance / _trailing_avg if _trailing_avg else 0

# Usage decline calculation (2022 avg vs 2025 avg)
_2022_avg = post_yearly[2022][1] / post_yearly[2022][2] if post_yearly[2022][2] else 0
_2025_avg = post_yearly[2025][1] / post_yearly[2025][2] if post_yearly[2025][2] else 0
_usage_decline_pct = (_2022_avg - _2025_avg) / _2022_avg if _2022_avg else 0

# Envelope type lookups — filtered to post-settlement
post_type_mask = by_type_monthly["Month"].apply(lambda x: month_label_to_sortkey(x) >= settlement_key)
by_type = by_type_monthly[post_type_mask].groupby("Envelope Type", as_index=False).agg({"Purchased": "sum"})
by_type.rename(columns={"Purchased": "Total Purchased"}, inplace=True)

post_usage_type_mask = usage_by_env_type_monthly["Month"].apply(lambda x: month_label_to_sortkey(x) >= settlement_key)
post_usage_type_monthly = usage_by_env_type_monthly[post_usage_type_mask]
usage_by_env_type = post_usage_type_monthly.groupby("Envelope Type", as_index=False).agg({"Envelopes Used": "sum"})
usage_by_env_type.rename(columns={"Envelopes Used": "Total Envelopes Used"}, inplace=True)

# Wastage by envelope type (apply per-month rate to each type's monthly usage)
wastage_by_type_dict = {}
for _, r in post_usage_type_monthly.iterrows():
    t = r["Envelope Type"]
    w = int(safe(r["Envelopes Used"]) * get_wastage_rate(r["Month"]))
    wastage_by_type_dict[t] = wastage_by_type_dict.get(t, 0) + w

# Trailing 12-month usage by envelope type (for buffer month calculation)
_last_12_months = sorted(post["Month"].unique(), key=month_label_to_sortkey)[-12:]
_t12_usage_mask = post_usage_type_monthly["Month"].isin(_last_12_months)
_t12_usage_by_type = post_usage_type_monthly[_t12_usage_mask].groupby("Envelope Type", as_index=False).agg({"Envelopes Used": "sum"})
trailing12_by_type_dict = dict(zip(_t12_usage_by_type["Envelope Type"], _t12_usage_by_type["Envelopes Used"]))

by_type_sorted = by_type.sort_values("Total Purchased", ascending=False)
env_type_total = by_type["Total Purchased"].sum()
purchase_by_type = dict(zip(by_type["Envelope Type"], by_type["Total Purchased"]))
usage_by_type_dict = dict(zip(usage_by_env_type["Envelope Type"], usage_by_env_type["Total Envelopes Used"]))

ENVELOPE_GROUPS = [
    {
        "label": "N14 Fold Statements",
        "purchase_types": ["ENVMEAPEXN14PFC", "ENVMERIDGEN14NI11/08"],
        "usage_types": ["ENVMEAPEXN14PFC", "ENVMERIDGEN14NI11/08"],
    },
    {
        "label": "9x12 Flat Statements",
        "purchase_types": ["ENVMEAPEX9X12PFC", "ENVMERIDGE9X12NI11/08"],
        "usage_types": ["ENVMEAPEX9X12PFC", "ENVMERIDGE9X12NI11/08"],
    },
    {
        "label": "#10 Confirms + Letters",
        "purchase_types": ["ENVAPXN10 Confirms+Letters (PFC)", "ENVCONPFSN10NI"],
        "usage_types": ["ENVAPXN10 Confirms+Letters (PFC)", "ENVCONPFSN10NI"],
    },
    {
        "label": "9x12 Flat Confirms",
        "purchase_types": ["ENVCONRIDGE9X12DW"],
        "usage_types": ["ENVCONRIDGE9X12DW"],
    },
    {
        "label": "Tax Form Envelopes",
        "purchase_types": ["Tax Form Envelopes (1099/1099-R)", "Tax Form Envelopes (1042/IRA)"],
        "usage_types": ["Tax Form Envelopes (1099/1099-R)"],
    },
]

# ---------------------------------------------------------------------------
# Combined envelope group rows (purchased + used + variance)
# ---------------------------------------------------------------------------
def build_combined_env_rows():
    rows = []
    grand_p = grand_u = grand_w = 0
    for g in ENVELOPE_GROUPS:
        p = sum(safe(purchase_by_type.get(t, 0)) for t in g["purchase_types"])
        u = sum(safe(usage_by_type_dict.get(t, 0)) for t in g["usage_types"])
        w = sum(safe(wastage_by_type_dict.get(t, 0)) for t in g["usage_types"])
        if p == 0 and u <= 1:
            continue
        au = u + w
        v = p - au
        grand_p += p
        grand_u += u
        grand_w += w
        vc = var_color(v)
        t12_u = sum(safe(trailing12_by_type_dict.get(t, 0)) for t in g["usage_types"])
        avg_mo = t12_u / 12 if t12_u else 0
        bc = buffer_color(v, avg_mo)
        rows.append(
            '<tr>'
            + f'<td class="env-name">{g["label"]}</td>'
            + f'<td class="num">{fmt_num(p)}</td>'
            + f'<td class="num">{fmt_num(u)}</td>'
            + f'<td class="num" style="color:#6D6E71;">{fmt_num(w)}</td>'
            + f'<td class="num" style="color:#6D6E71;">{fmt_pct(w / u if u else 0)}</td>'
            + f'<td class="num" style="color:{vc};font-weight:600">{fmt_num_parens(v)}</td>'
            + f'<td class="num" style="color:{bc};font-weight:600">{fmt_buffer_months(v, avg_mo)}</td>'
            + '</tr>'
        )
    gau = grand_u + grand_w
    gv = grand_p - gau
    gvc = var_color(gv)
    gbc = buffer_color(gv, _trailing_avg)
    rows.append(
        '<tr class="total-row">'
        + f'<td><strong>Total</strong></td>'
        + f'<td class="num"><strong>{fmt_num(grand_p)}</strong></td>'
        + f'<td class="num"><strong>{fmt_num(grand_u)}</strong></td>'
        + f'<td class="num" style="color:#6D6E71;"><strong>{fmt_num(grand_w)}</strong></td>'
        + f'<td class="num" style="color:#6D6E71;"><strong>{fmt_pct(grand_w / grand_u if grand_u else 0)}</strong></td>'
        + f'<td class="num" style="color:{gvc};font-weight:700"><strong>{fmt_num_parens(gv)}</strong></td>'
        + f'<td class="num" style="color:{gbc};font-weight:700"><strong>{fmt_buffer_months(gv, _trailing_avg)}</strong></td>'
        + '</tr>'
    )
    return "\n".join(rows)

# ---------------------------------------------------------------------------
# SKU-level recon rows
# ---------------------------------------------------------------------------
def build_sku_recon_rows():
    all_types = sorted(
        set(list(purchase_by_type.keys()) + list(usage_by_type_dict.keys())),
        key=lambda t: -(safe(purchase_by_type.get(t, 0)) + safe(usage_by_type_dict.get(t, 0)))
    )
    rows = []
    grand_p = grand_u = grand_w = 0
    for t in all_types:
        p = safe(purchase_by_type.get(t, 0))
        u = safe(usage_by_type_dict.get(t, 0))
        w = safe(wastage_by_type_dict.get(t, 0))
        if p == 0 and u == 0:
            continue
        au = u + w
        v = p - au
        grand_p += p
        grand_u += u
        grand_w += w
        vc = var_color(v)
        t12_u = safe(trailing12_by_type_dict.get(t, 0))
        avg_mo = t12_u / 12 if t12_u else 0
        bc = buffer_color(v, avg_mo)
        rows.append(
            '<tr>'
            + f'<td class="env-name">{t}</td>'
            + f'<td class="num">{fmt_num(p)}</td>'
            + f'<td class="num">{fmt_num(u)}</td>'
            + f'<td class="num" style="color:#6D6E71;">{fmt_num(w)}</td>'
            + f'<td class="num" style="color:#6D6E71;">{fmt_pct(w / u if u else 0)}</td>'
            + f'<td class="num" style="color:{vc};font-weight:600">{fmt_num_parens(v)}</td>'
            + f'<td class="num" style="color:{bc};font-weight:600">{fmt_buffer_months(v, avg_mo)}</td>'
            + '</tr>'
        )
    gau = grand_u + grand_w
    gv = grand_p - gau
    gvc = var_color(gv)
    gbc = buffer_color(gv, _trailing_avg)
    rows.append(
        '<tr class="total-row">'
        + f'<td><strong>Total</strong></td>'
        + f'<td class="num"><strong>{fmt_num(grand_p)}</strong></td>'
        + f'<td class="num"><strong>{fmt_num(grand_u)}</strong></td>'
        + f'<td class="num" style="color:#6D6E71;"><strong>{fmt_num(grand_w)}</strong></td>'
        + f'<td class="num" style="color:#6D6E71;"><strong>{fmt_pct(grand_w / grand_u if grand_u else 0)}</strong></td>'
        + f'<td class="num" style="color:{gvc};font-weight:700"><strong>{fmt_num_parens(gv)}</strong></td>'
        + f'<td class="num" style="color:{gbc};font-weight:700"><strong>{fmt_buffer_months(gv, _trailing_avg)}</strong></td>'
        + '</tr>'
    )
    return "\n".join(rows)

# ---------------------------------------------------------------------------
# Envelope Specifications
# ---------------------------------------------------------------------------
# Build all data components
# ---------------------------------------------------------------------------
combined_env_rows = build_combined_env_rows()
sku_recon_rows = build_sku_recon_rows()

post_var_color = "#186741" if post_variance >= 0 else "#9D1526"
post_var_pct = post_variance / post_purchased if post_purchased else 0

# ---------------------------------------------------------------------------
# CSS (same Apex brand design system as internal report)
# ---------------------------------------------------------------------------
CSS = """*, *::before, *::after { box-sizing: border-box; }
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
.header .from-line { font-size: 13px; opacity: 0.75; margin: 0 0 4px; }
.header .generated { font-size: 12px; opacity: 0.65; margin: 0; }
.content { max-width: 1100px; margin: 0 auto; padding: 28px 40px 40px; }
.section { margin-bottom: 32px; page-break-inside: avoid; }
.section-header { margin-bottom: 14px; }
.section-header h2 { margin: 0; font-size: 18px; font-weight: 600; color: #052390; border-bottom: 2px solid #052390; padding-bottom: 6px; }
.kpi-grid { display: flex; flex-wrap: wrap; gap: 16px; margin-bottom: 10px; }
.kpi-card {
    flex: 1 1 180px; background: #F5F5F7; border-radius: 8px; padding: 18px 20px;
    min-width: 170px; border: 1px solid #E2E2E2;
}
.kpi-card .kpi-label {
    font-size: 11px; font-weight: 600; text-transform: uppercase;
    letter-spacing: 0.5px; color: #6D6E71; margin: 0 0 6px;
}
.kpi-card .kpi-value { font-size: 24px; font-weight: 700; margin: 0; line-height: 1.2; }
.kpi-card .kpi-sub { font-size: 11px; color: #6D6E71; margin: 4px 0 0; }
.table-wrap { border-radius: 4px; border: 1px solid #E2E2E2; }
table { width: 100%; border-collapse: collapse; font-size: 12px; background: #FFFFFF; }
table th {
    background: #052390; color: #FFFFFF; padding: 8px 12px; text-align: left;
    font-weight: 600; font-size: 11px; text-transform: uppercase; letter-spacing: 0.3px;
    white-space: nowrap;
}
table td { padding: 6px 12px; border-bottom: 1px solid #E2E2E2; white-space: nowrap; color: #333; }
table .num { text-align: right; font-variant-numeric: tabular-nums; }
table .env-name { max-width: 260px; white-space: normal; word-break: break-word; }
.subtotal-row { background: #F0F0F5 !important; border-top: 2px solid #D0D0D8; }
.total-row { background: #E8E8ED !important; border-top: 2px solid #D0D0D8; }
table tbody tr:nth-child(even) { background: #FAFAFA; }
.footer {
    text-align: center; padding: 24px 40px; font-size: 11px;
    color: #6D6E71; border-top: 1px solid #E2E2E2;
}
@media print {
    .header { padding: 28px 0 20px; }
    .content { padding: 16px 0 20px; }
    .section { page-break-inside: avoid; margin-bottom: 20px; }
    .kpi-card { box-shadow: none; }
    .table-wrap { box-shadow: none; border: 1px solid #CCC; }
}
"""

def th_row(headers):
    return "".join(f'<th>{h}</th>' for h in headers)

# ---------------------------------------------------------------------------
# Build HTML
# ---------------------------------------------------------------------------
html = '<!DOCTYPE html>\n<html lang="en">\n<head>\n'
html += '<meta charset="UTF-8">\n'
html += '<meta name="viewport" content="width=device-width, initial-scale=1.0">\n'
html += '<title>Envelope Reconciliation — For Review</title>\n'
html += f'<style>{CSS}</style>\n'
html += '</head>\n<body>\n\n'

# --- Header ---
html += '<div class="header">\n'
html += '    <h1>Envelope reconciliation &mdash; for review</h1>\n'
html += '    <p class="subtitle">March 2022 &ndash; December 2025</p>\n'
html += '    <p class="from-line">From: Apex Clearing Corporation</p>\n'
html += f'    <p class="generated">Generated {pd.Timestamp.now().strftime("%B %d, %Y")}</p>\n'
html += '</div>\n'

html += '<div class="content">\n'

# ===== SECTION 1: SUMMARY =====
html += '<div class="section" id="summary">\n'
html += '    <div class="section-header">\n'
html += '        <h2>Summary</h2>\n'
html += '    </div>\n'
html += '    <div class="section-body">\n'
html += '        <p style="font-size:13px;color:#333;margin:0 0 18px;line-height:1.6;">This report presents Apex&rsquo;s reconciliation of envelope purchases versus usage across 46 months of post-settlement activity. We are sharing this for Broadridge&rsquo;s review and validation.</p>\n'

# KPI cards
html += '        <div class="kpi-grid">\n'
html += f'            <div class="kpi-card"><p class="kpi-label">Purchased</p><p class="kpi-value" style="color:#052390">{fmt_num(post_purchased)}</p><p class="kpi-sub">Total envelopes ordered</p></div>\n'
html += f'            <div class="kpi-card"><p class="kpi-label">Used + Wastage (Est.)</p><p class="kpi-value" style="color:#052390">{fmt_num(post_adj_used)}</p><p class="kpi-sub">{fmt_num(post_used)} used + {fmt_num(post_wastage)} est. wastage</p></div>\n'
html += f'            <div class="kpi-card"><p class="kpi-label">Variance</p><p class="kpi-value" style="color:{post_var_color}">{fmt_num_parens(post_variance)}</p><p class="kpi-sub">{_buffer_months:.1f} months of buffer stock</p></div>\n'
html += f'            <div class="kpi-card"><p class="kpi-label">Months Covered</p><p class="kpi-value" style="color:#052390">{post_months}</p><p class="kpi-sub">Mar 2022 &ndash; Dec 2025</p></div>\n'
html += '        </div>\n'
html += '        <p style="font-size:10px;color:#6D6E71;margin:6px 0 4px;line-height:1.5;"><strong>Wastage (Est.)</strong> = contractual maximum wastage allowance per Section 4: 5% of usage (Jan 2019&ndash;Dec 2023) and 2% of usage (Jan 2024+). Actual operational wastage may differ. See <em>Items for Review</em>.</p>\n'
html += '        <p style="font-size:10px;color:#6D6E71;margin:0 0 12px;line-height:1.5;"><strong>Buffer (Mo.)</strong> = Adj. Variance &divide; Avg Monthly Usage, where Adj. Variance = Purchased &minus; Used &minus; Wastage (Est.), and Avg Monthly Usage = trailing 12-month average (captures full seasonal cycle including quarter-end and year-end peaks). Year rows use that year&rsquo;s own month count and usage. Broadridge&rsquo;s stated policy is 2&ndash;3 months of supply (Brandon Koebel, Nov 2022).</p>\n'

# Year-by-year table
html += '        <div class="table-wrap" style="margin-top:20px;"><table>\n'
html += '            <thead><tr><th>Year</th><th>Purchased</th><th>Used</th><th>Wastage (Est.)</th><th>Wastage %</th><th>Adj. Variance</th><th>Buffer (Mo.)</th><th>Invoiced</th></tr></thead>\n'
html += '            <tbody>\n'
for yr in sorted(post_yearly.keys()):
    d = post_yearly[yr]
    yp, yu, mc, yc, yi, yw = d
    yau = yu + yw
    yv = yp - yau
    vc = var_color(yv)
    y_avg_mo = yu / mc if mc else 0
    ybc = buffer_color(yv, y_avg_mo)
    yr_label = f"{yr} (Mar&ndash;Dec)" if yr == 2022 else str(yr)
    html += f'            <tr><td>{yr_label}</td>'
    html += f'<td class="num">{fmt_num(yp)}</td>'
    html += f'<td class="num">{fmt_num(yu)}</td>'
    html += f'<td class="num" style="color:#6D6E71;">{fmt_num(yw)}</td>'
    wpct = yw / yu if yu else 0
    html += f'<td class="num" style="color:#6D6E71;">{fmt_pct(wpct)}</td>'
    html += f'<td class="num" style="color:{vc};font-weight:600">{fmt_num_parens(yv)}</td>'
    html += f'<td class="num" style="color:{ybc};font-weight:600">{fmt_buffer_months(yv, y_avg_mo)}</td>'
    html += f'<td class="num">{fmt_money(yi)}</td></tr>\n'
# Total row
html += f'            <tr class="total-row"><td><strong>Total</strong></td>'
html += f'<td class="num"><strong>{fmt_num(post_purchased)}</strong></td>'
html += f'<td class="num"><strong>{fmt_num(post_used)}</strong></td>'
html += f'<td class="num" style="color:#6D6E71;"><strong>{fmt_num(post_wastage)}</strong></td>'
total_wpct = post_wastage / post_used if post_used else 0
html += f'<td class="num" style="color:#6D6E71;"><strong>{fmt_pct(total_wpct)}</strong></td>'
html += f'<td class="num" style="color:{post_var_color};font-weight:700"><strong>{fmt_num_parens(post_variance)}</strong></td>'
_total_bc = buffer_color(post_variance, _trailing_avg)
html += f'<td class="num" style="color:{_total_bc};font-weight:700"><strong>{fmt_buffer_months(post_variance, _trailing_avg)}</strong></td>'
html += f'<td class="num"><strong>{fmt_money(post_invoiced)}</strong></td></tr>\n'
html += '            </tbody>\n'
html += '        </table></div>\n'

html += '    </div>\n'
html += '</div>\n\n'

# ===== SECTION 2: ITEMS FOR REVIEW =====
html += '<div class="section" id="review">\n'
html += '    <div class="section-header">\n'
html += '        <h2>Items for review</h2>\n'
html += '    </div>\n'
html += '    <div class="section-body">\n'

# Wastage observation
html += '        <div style="background:#F5F5F7;border-radius:8px;padding:18px 22px;margin-bottom:16px;border:1px solid #E2E2E2;">\n'
html += '            <p style="font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:0.3px;color:#052390;margin:0 0 10px;">Wastage &mdash; contractual vs. operational</p>\n'
html += f'            <p style="font-size:12px;line-height:1.7;color:#333;margin:0 0 8px;">The contract caps the wastage charge at <strong>5%</strong> (original, through Dec 2023) '
html += f'and <strong>2%</strong> (Amendment No. 1, Jan 2024+). Broadridge has separately confirmed that actual operational wastage runs 10&ndash;15%.</p>\n'
html += '            <table style="margin-top:8px;width:auto;background:transparent;font-size:12px;">\n'
html += f'                <tr><td style="border:none;padding:4px 16px 4px 0;color:#6D6E71;">Contract max wastage (5%/2% blended)</td>'
html += f'<td style="border:none;padding:4px 0;color:#333;font-weight:600;">{fmt_num(_contract_waste)} envelopes</td></tr>\n'
html += f'                <tr><td style="border:none;padding:4px 16px 4px 0;color:#6D6E71;">Operational wastage at 10%</td>'
html += f'<td style="border:none;padding:4px 0;color:#333;font-weight:600;">{fmt_num(_actual_waste_lo)} envelopes</td></tr>\n'
html += f'                <tr><td style="border:none;padding:4px 16px 4px 0;color:#6D6E71;">Operational wastage at 15%</td>'
html += f'<td style="border:none;padding:4px 0;color:#333;font-weight:600;">{fmt_num(_actual_waste_hi)} envelopes</td></tr>\n'
html += f'                <tr><td style="border:none;padding:4px 16px 4px 0;color:#6D6E71;">Excess beyond contract allowance</td>'
html += f'<td style="border:none;padding:4px 0;color:#052390;font-weight:600;">{fmt_num(_excess_lo)}&ndash;{fmt_num(_excess_hi)} envelopes</td></tr>\n'
html += '            </table>\n'
html += f'            <p style="font-size:11px;color:#6D6E71;margin:8px 0 0;">Per Section 4, excess wastage beyond the contractual rate is Broadridge&rsquo;s responsibility. We would like to confirm the current operational wastage rate and how it is reflected in invoicing.</p>\n'
html += '        </div>\n'

# Excess inventory observation
html += '        <div style="background:#F5F5F7;border-radius:8px;padding:18px 22px;margin-bottom:16px;border:1px solid #E2E2E2;">\n'
html += '            <p style="font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:0.3px;color:#052390;margin:0 0 10px;">Excess inventory</p>\n'
html += f'            <p style="font-size:12px;line-height:1.7;color:#333;margin:0 0 8px;">After accounting for contractual wastage, post-settlement purchases exceed adjusted usage by <strong>{fmt_num(post_variance)} envelopes</strong>. '
html += f'At the trailing 12-month average usage of {fmt_num(_trailing_avg)}/month, this represents approximately <strong>{_buffer_months:.1f} months</strong> of buffer stock.</p>\n'
html += f'            <p style="font-size:12px;line-height:1.7;color:#333;margin:0 0 8px;">Average monthly usage has declined <strong>{_usage_decline_pct * 100:.0f}%</strong> since 2022 '
html += f'(from {fmt_num(int(_2022_avg))}/month to {fmt_num(int(_2025_avg))}/month in 2025). '
html += f'Broadridge has stated a policy of maintaining 2&ndash;3 months of envelope supply (Brandon Koebel, Nov 2022). '
html += f'Section 8 requires Broadridge to be &ldquo;responsible for procuring and maintaining sufficient quantity of Materials, based on average volumes.&rdquo;</p>\n'
html += f'            <p style="font-size:11px;color:#6D6E71;margin:8px 0 0;">We would like to understand the current inventory position and whether ordering volumes are being adjusted to reflect the {_usage_decline_pct * 100:.0f}% decline in average monthly usage.</p>\n'
html += '        </div>\n'

# Confirmation questions
html += '        <div style="background:#FFFFFF;border-radius:8px;padding:18px 22px;margin-bottom:16px;border:2px solid #052390;">\n'
html += '            <p style="font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:0.3px;color:#052390;margin:0 0 12px;">Outstanding items for Broadridge confirmation</p>\n'
html += '            <ol style="font-size:12px;line-height:1.8;color:#333;margin:0;padding-left:20px;">\n'
html += '                <li><strong>Data validation:</strong> Please review the purchased and used totals in this report and confirm they align with Broadridge&rsquo;s records.</li>\n'
html += '                <li><strong>Wastage rate applied:</strong> What wastage rate is currently being applied to envelope charges? The contract specifies 5% (original) / 2% (amendment) for generic envelope stock. Operational wastage has been confirmed at 10&ndash;15% &mdash; please clarify whether the invoiced rate reflects the contractual rate or the operational rate.</li>\n'
html += '                <li><strong>Unit rate breakdown:</strong> Please provide a sample breakdown of the invoiced unit rate for envelopes showing the vendor price, wastage component, and margin (if applicable) for a recent month.</li>\n'
html += '                <li><strong>Inventory position:</strong> Please provide the current envelope inventory levels and confirm whether ordering volumes are being adjusted to reflect the decline in average monthly usage.</li>\n'
html += '            </ol>\n'
html += '        </div>\n'

html += '    </div>\n'
html += '</div>\n\n'

# ===== SECTION 3: PURCHASES & USAGE BY ENVELOPE TYPE =====
html += '<div class="section" id="envelope-types">\n'
html += '    <div class="section-header">\n'
html += '        <h2>Purchases &amp; usage by envelope type</h2>\n'
html += '    </div>\n'
html += '    <div class="section-body">\n'
html += '        <div class="table-wrap"><table>\n'
html += '            <thead><tr>' + th_row(["Envelope Type","Purchased","Used","Wastage (Est.)","Wastage %","Adj. Variance","Buffer (Mo.)"]) + '</tr></thead>\n'
html += '            <tbody>\n' + combined_env_rows + '\n            </tbody>\n'
html += '        </table></div>\n'
html += '        <p style="font-size:11px;color:#6D6E71;margin:10px 0 0;">Related envelope SKUs grouped by physical size. Usage mapped from billing data via product category, flat/fold, and address type.</p>\n'
html += '        <p style="font-size:11px;color:#6D6E71;margin:4px 0 0;"><strong style="color:#333;">Wastage % by type:</strong> The per-type wastage rate reflects each SKU&rsquo;s blend of pre-2024 usage (5% contract rate) and post-2024 usage (2% amended rate). SKUs with usage concentrated before 2024 show rates closer to 5%; those with more recent usage trend toward 2%. The overall blended rate across all types is 3.7%.</p>\n'
html += '        <p style="font-size:11px;color:#6D6E71;margin:4px 0 0;"><strong style="color:#333;">9x12 Flat Confirms:</strong> The deficit reflects a temporary production surge in mid-2022 (Ridge/Penson migration) when Broadridge routed a portion of confirms through flat production. Pre-existing buffer stock covered the shortfall. This category represents 0.7% of total volume.</p>\n'

# NI/PFC transition context
html += '        <div style="background:#F5F5F7;border-radius:8px;padding:14px 20px;margin-top:14px;border:1px solid #E2E2E2;font-size:12px;line-height:1.7;color:#333;">\n'
html += '            <p style="font-weight:700;color:#052390;margin:0 0 6px;">NI / PFC transition (Oct 2022)</p>\n'
html += '            <p style="margin:0 0 6px;"><strong>Before Oct 2022:</strong> All #10 fold confirms, letters, and checks (domestic + foreign) used ENVCONPFSN10NI (No Imprint).</p>\n'
html += '            <p style="margin:0 0 6px;"><strong>After Oct 2022:</strong> Domestic mail switched to ENVAPXN10PFSCONN10IND(10/22) (Pre-Sorted First-Class); foreign mail remained on ENVCONPFSN10NI at ~8,000/month.</p>\n'
html += '            <p style="margin:0;">The same transition applies to N14 statements (PFC for domestic, NI for foreign) and 9x12 flats. NI and PFC versions are physically interchangeable &mdash; NI envelopes can be used for domestic mail by adding indicia at time of mailing.</p>\n'
html += '        </div>\n'

# SKU-level breakdown
html += '        <h3 style="color:#052390;font-size:14px;margin:24px 0 10px;">By SKU</h3>\n'
html += '        <div class="table-wrap"><table>\n'
html += '            <thead><tr>' + th_row(["SKU","Purchased","Used","Wastage (Est.)","Wastage %","Adj. Variance","Buffer (Mo.)"]) + '</tr></thead>\n'
html += '            <tbody>\n' + sku_recon_rows + '\n            </tbody>\n'
html += '        </table></div>\n'

html += '    </div>\n</div>\n\n'

# ===== SECTION 4: REFERENCE =====
html += '<div class="section" id="reference">\n'
html += '    <div class="section-header">\n'
html += '        <h2>Reference</h2>\n'
html += '    </div>\n'
html += '    <div class="section-body">\n'

# Data sources & methodology
html += '        <h3 style="color:#052390;font-size:14px;margin:0 0 12px;">Data sources &amp; methodology</h3>\n'
html += '        <div style="background:#F5F5F7;border-radius:8px;padding:16px 20px;margin-bottom:16px;border:1px solid #E2E2E2;">\n'
html += '            <table style="width:auto;background:transparent;font-size:12px;">\n'
html += f'                <tr><td style="border:none;padding:4px 16px 4px 0;color:#6D6E71;white-space:nowrap;">Reconciliation period</td><td style="border:none;padding:4px 0;color:#333;font-weight:600;">January 2020 &ndash; December 2025 (post-settlement focus: March 2022+)</td></tr>\n'
html += f'                <tr><td style="border:none;padding:4px 16px 4px 0;color:#6D6E71;white-space:nowrap;">Purchase reports processed</td><td style="border:none;padding:4px 0;color:#333;font-weight:600;">{purchase_files} files (monthly Broadridge Purchase Reports, FY&rsquo;20&ndash;FY&rsquo;25)</td></tr>\n'
html += f'                <tr><td style="border:none;padding:4px 16px 4px 0;color:#6D6E71;white-space:nowrap;">Billing workbooks processed</td><td style="border:none;padding:4px 0;color:#333;font-weight:600;">{billing_files} files (monthly Billing Workbooks + Billing Master 2020&ndash;2021)</td></tr>\n'
html += f'                <tr><td style="border:none;padding:4px 16px 4px 0;color:#6D6E71;white-space:nowrap;">Usage-to-envelope mapping</td><td style="border:none;padding:4px 0;color:#333;">Product category + Flat_Fold + Address_Type fields from Volume Data (per Brandon Koebel, Sep 2023)</td></tr>\n'
html += f'                <tr><td style="border:none;padding:4px 16px 4px 0;color:#6D6E71;white-space:nowrap;">NI/PFC usage cutover</td><td style="border:none;padding:4px 0;color:#333;">Domestic #10 usage mapped to ENVCONPFSN10NI before Oct 2022, ENVAPXN10 (PFC) after</td></tr>\n'
html += f'                <tr><td style="border:none;padding:4px 16px 4px 0;color:#6D6E71;white-space:nowrap;">&ldquo;Used&rdquo; definition</td><td style="border:none;padding:4px 0;color:#333;">Envelopes pulled from warehouse inventory to the production floor, including waste and floor surplus &mdash; not limited to envelopes actually mailed (per Brandon Koebel, Oct 2022)</td></tr>\n'
html += '            </table>\n'
html += '        </div>\n'

# Contract terms
html += '        <h3 style="color:#052390;font-size:14px;margin:24px 0 12px;">Contract terms &mdash; envelope materials</h3>\n'

# Original contract
html += '        <div style="background:#F5F5F7;border-radius:8px;padding:16px 20px;margin-bottom:12px;border:1px solid #E2E2E2;">\n'
html += '            <p style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:0.3px;color:#052390;margin:0 0 6px;">Original Contract &mdash; Section 4, Compensation</p>\n'
html += '            <p style="font-size:11px;color:#6D6E71;margin:0 0 6px;">GTO Print and Mail Services Schedule, effective January 1, 2019.</p>\n'
html += '            <blockquote style="margin:0;padding:10px 14px;background:#FFFFFF;border-left:3px solid #052390;border-radius:0 4px 4px 0;font-size:12px;line-height:1.7;color:#333;">\n'
html += '                &ldquo;Materials (such as paper, envelopes, and inserts) and postage, presort and insert related fees are not included in the Annual Fee and will be charged separately. Materials are billed at cost plus wastage for generic stock. <strong>Specifically, the wastage charge is 10% for any generic paper stock and 5% for generic envelope stock.</strong> For generic stock, the unit rate will be billed based on usage. For Client specific stock, the unit rate will be based on receipt of such stock.&rdquo;\n'
html += '            </blockquote>\n'
html += '        </div>\n'

# Amendment
html += '        <div style="background:#F5F5F7;border-radius:8px;padding:16px 20px;margin-bottom:12px;border:1px solid #E2E2E2;">\n'
html += '            <p style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:0.3px;color:#052390;margin:0 0 6px;">Amendment No. 1 &mdash; Section 4, Compensation (replaces original)</p>\n'
html += '            <p style="font-size:11px;color:#6D6E71;margin:0 0 6px;">GTO Print and Mail Services Schedule Amendment No. 1, effective January 1, 2024. Term extended through December 31, 2028.</p>\n'
html += '            <blockquote style="margin:0;padding:10px 14px;background:#FFFFFF;border-left:3px solid #052390;border-radius:0 4px 4px 0;font-size:12px;line-height:1.7;color:#333;">\n'
html += '                &ldquo;Materials are billed at inventory cost plus 10% margin. Inventory cost means for (i) Client specific inventory: vendor price; and (ii) generic inventory: vendor price plus wastage as follows: <strong>10% for continuous form, 3% for cutsheet, and 2% for envelopes.</strong> For generic stock, the unit rate will be billed based on usage. For Client specific stock, the unit rate will be based on receipt of such stock.&rdquo;\n'
html += '            </blockquote>\n'
html += '        </div>\n'

# Materials obligation
html += '        <div style="background:#F5F5F7;border-radius:8px;padding:16px 20px;margin-bottom:12px;border:1px solid #E2E2E2;">\n'
html += '            <p style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:0.3px;color:#052390;margin:0 0 6px;">Original Contract &mdash; Section 8, Materials</p>\n'
html += '            <blockquote style="margin:0;padding:10px 14px;background:#FFFFFF;border-left:3px solid #052390;border-radius:0 4px 4px 0;font-size:12px;line-height:1.7;color:#333;">\n'
html += '                &ldquo;Client will specify the materials (paper, inserts and envelope) (&ldquo;Materials&rdquo;) to be used; provided, however, that materials specified must conform to Broadridge print and insert equipment specifications. <strong>Broadridge shall be responsible for procuring and maintaining sufficient quantity of Materials, based on average volumes</strong>, except where Client is responsible for Materials as specified in Section 3.&rdquo;\n'
html += '            </blockquote>\n'
html += '        </div>\n'

# Contract summary table
html += '        <div class="table-wrap" style="margin-bottom:24px;"><table>\n'
html += '            <thead><tr><th>Period</th><th>Envelope Wastage</th><th>Margin</th><th>Effective Rate</th><th>Billing Basis (Contract)</th></tr></thead>\n'
html += '            <tbody>\n'
html += '            <tr><td>Jan 2019 &ndash; Dec 2023</td><td>5%</td><td>&mdash;</td><td>5.0% over vendor</td><td>Usage</td></tr>\n'
html += '            <tr><td>Jan 2024 &ndash; Dec 2028</td><td>2%</td><td>10%</td><td>12.2% over vendor</td><td>Usage</td></tr>\n'
html += '            </tbody>\n'
html += '        </table></div>\n'

# Broadridge confirmation — wastage rates (Koebel emails)
html += '        <div style="background:#F5F5F7;border-radius:8px;padding:16px 20px;margin-bottom:12px;border:1px solid #E2E2E2;">\n'
html += '            <p style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:0.3px;color:#052390;margin:0 0 6px;">Broadridge confirmation &mdash; operational wastage rates</p>\n'
html += '            <p style="font-size:11px;color:#6D6E71;margin:0 0 6px;">Brandon Koebel (Sr. Client Relationship Manager), emails Sep&ndash;Nov 2022.</p>\n'
html += '            <blockquote style="margin:0;padding:10px 14px;background:#FFFFFF;border-left:3px solid #052390;border-radius:0 4px 4px 0;font-size:12px;line-height:1.7;color:#333;">\n'
html += '                &ldquo;Wastage is roughly 10%&hellip; This includes envelopes that are damaged, need to be reprinted and reinserted, etc.&rdquo; (Nov 7, 2022)<br>\n'
html += '                &ldquo;Did not account for any waste or spoilage (<strong>typically 10&ndash;15%</strong>).&rdquo; (Sep 29, 2022)\n'
html += '            </blockquote>\n'
html += '            <p style="font-size:11px;color:#6D6E71;margin:6px 0 0;"><strong style="color:#333;">Note:</strong> Contractual wastage (billed to Apex) is 5% pre-2024 / 2% post-amendment. Operational wastage (10&ndash;15%) is higher but embedded in the &ldquo;Used&rdquo; figure &mdash; not separately charged.</p>\n'
html += '        </div>\n'

# Pre-settlement context
html += f'        <div style="background:#F5F5F7;border-radius:8px;padding:14px 20px;margin-bottom:12px;border:1px solid #E2E2E2;font-size:12px;color:#6D6E71;line-height:1.7;">\n'
html += f'            <p style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:0.3px;color:#052390;margin:0 0 6px;">Pre-settlement context</p>\n'
html += f'            <strong style="color:#333;">Jan 2020 &ndash; Feb 2022:</strong> {fmt_num(pre_purchased)} purchased, {fmt_num(pre_used)} used. '
html += f'These costs were absorbed by Broadridge per the pass-through paper and envelope dispute settlement. '
html += f'<strong style="color:#333;">Full period (Jan 2020 &ndash; Dec 2025):</strong> {fmt_num(total_purchased)} purchased, {fmt_num(total_used)} used, '
html += f'{"+" if net_variance >= 0 else ""}{fmt_num_parens(net_variance)} variance.\n'
html += '        </div>\n'

# Generic stock classification
html += '        <h3 style="color:#052390;font-size:14px;margin:24px 0 12px;">Generic stock classification</h3>\n'
html += '        <div style="background:#F5F5F7;border-radius:8px;padding:16px 20px;margin-bottom:12px;border:1px solid #E2E2E2;">\n'
html += '            <p style="font-size:12px;line-height:1.7;color:#333;margin:0 0 10px;">Both the original contract and Amendment No. 1 distinguish between <strong>generic stock</strong> (billed on usage, lower wastage) and <strong>client-specific stock</strong> (billed on receipt). All Apex envelope types are generic stock:</p>\n'
html += '            <div class="table-wrap" style="margin:0;"><table style="font-size:12px;">\n'
html += '                <thead><tr><th>Envelope</th><th>Size</th><th>Windows</th><th>Paper</th><th>Ink</th><th>Security Tint</th><th>Indicia</th><th>Client Branding</th></tr></thead>\n'
html += '                <tbody>\n'
html += '                <tr><td>ENVMEAPEXN14PFC</td><td>4&frac34; &times; 11&frac716;</td><td>Double</td><td>24WW</td><td>Black</td><td>Crosshatch</td><td>PFC</td><td>None</td></tr>\n'
html += '                <tr><td>ENVMEAPEX9X12PFC</td><td>9 &times; 12</td><td>Double</td><td>24WW</td><td>Black</td><td>Crosshatch</td><td>PFC</td><td>None</td></tr>\n'
html += '                <tr><td>ENVAPXN10PFSCONN10IND</td><td>4&frac18; &times; 9&frac12;</td><td>Double</td><td>24WW</td><td>Black</td><td>Crosshatch</td><td>PFC</td><td>None</td></tr>\n'
html += '                <tr><td>ENVCONPFSN10NI</td><td>4&frac18; &times; 9&frac12;</td><td>Double</td><td>24WW</td><td>Black</td><td>Crosshatch</td><td>None (NI)</td><td>None</td></tr>\n'
html += '                <tr><td>ENVMERIDGEN14NI11/08</td><td>4&frac34; &times; 11&frac716;</td><td>Double</td><td>24WW</td><td>Black</td><td>Crosshatch</td><td>None (NI)</td><td>None</td></tr>\n'
html += '                <tr><td>ENVMERIDGE9X12NI11/08</td><td>9 &times; 12</td><td>Double</td><td>24WW</td><td>Black</td><td>Crosshatch</td><td>None (NI)</td><td>None</td></tr>\n'
html += '                <tr><td>ENVCONRIDGE9X12DW</td><td>9 &times; 12</td><td>Double (vert.)</td><td>24WW</td><td>Black</td><td>Wood grain</td><td>None</td><td>None</td></tr>\n'
html += '                </tbody>\n'
html += '            </table></div>\n'
html += '            <p style="font-size:11px;color:#6D6E71;margin:10px 0 0;">All envelopes are standard double-window envelopes with no company logos, branding, or custom design elements. Return and recipient addresses are visible through the windows from the printed content inside. PFC (Pre-Sorted First-Class) indicia is a functional USPS postage marking, not client branding. NI (No Imprint) envelopes are completely blank. Supplier: United Envelope LLC, Mt. Pocono, PA.</p>\n'
html += '        </div>\n'

html += '    </div>\n</div>\n\n'

html += '</div><!-- end .content -->\n\n'

# Footer
html += '<div class="footer">\n'
html += '    Confidential &mdash; Apex Clearing Corporation<br>\n'
html += '    Prepared for reconciliation review with Broadridge Financial Solutions\n'
html += '</div>\n'

html += '</body>\n</html>\n'

with open(HTML_PATH, "w", encoding="utf-8") as f:
    f.write(html)

file_size = os.path.getsize(HTML_PATH)
if file_size > 1048576:
    print(f"Report generated: {HTML_PATH}")
    print(f"File size: {file_size / 1048576:.2f} MB")
else:
    print(f"Report generated: {HTML_PATH}")
    print(f"File size: {file_size / 1024:.1f} KB")
