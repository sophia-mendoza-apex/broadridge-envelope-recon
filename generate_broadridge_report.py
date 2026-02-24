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

def var_color(v):
    try:
        val = float(v)
        if val > 0:
            return "#4CAF79"
        if val < 0:
            return "#EF5350"
    except (ValueError, TypeError):
        pass
    return "#9A9BA0"

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
        vpct = v / p if p else 0
        rows.append(
            '<tr>'
            + f'<td class="env-name">{g["label"]}</td>'
            + f'<td class="num">{fmt_num(p)}</td>'
            + f'<td class="num">{fmt_num(u)}</td>'
            + f'<td class="num" style="color:#9A9BA0;">{fmt_num(w)}</td>'
            + f'<td class="num" style="color:#9A9BA0;">{fmt_pct(w / u if u else 0)}</td>'
            + f'<td class="num" style="color:{vc};font-weight:600">{fmt_num_parens(v)}</td>'
            + f'<td class="num" style="color:{vc}">{fmt_pct(vpct)}</td>'
            + '</tr>'
        )
    gau = grand_u + grand_w
    gv = grand_p - gau
    gvc = var_color(gv)
    gpct = gv / grand_p if grand_p else 0
    rows.append(
        '<tr class="total-row">'
        + f'<td><strong>Total</strong></td>'
        + f'<td class="num"><strong>{fmt_num(grand_p)}</strong></td>'
        + f'<td class="num"><strong>{fmt_num(grand_u)}</strong></td>'
        + f'<td class="num" style="color:#9A9BA0;"><strong>{fmt_num(grand_w)}</strong></td>'
        + f'<td class="num" style="color:#9A9BA0;"><strong>{fmt_pct(grand_w / grand_u if grand_u else 0)}</strong></td>'
        + f'<td class="num" style="color:{gvc};font-weight:700"><strong>{fmt_num_parens(gv)}</strong></td>'
        + f'<td class="num" style="color:{gvc};font-weight:700"><strong>{fmt_pct(gpct)}</strong></td>'
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
        vpct = v / p if p else 0
        rows.append(
            '<tr>'
            + f'<td class="env-name">{t}</td>'
            + f'<td class="num">{fmt_num(p)}</td>'
            + f'<td class="num">{fmt_num(u)}</td>'
            + f'<td class="num" style="color:#9A9BA0;">{fmt_num(w)}</td>'
            + f'<td class="num" style="color:#9A9BA0;">{fmt_pct(w / u if u else 0)}</td>'
            + f'<td class="num" style="color:{vc};font-weight:600">{fmt_num_parens(v)}</td>'
            + f'<td class="num" style="color:{vc}">{fmt_pct(vpct)}</td>'
            + '</tr>'
        )
    gau = grand_u + grand_w
    gv = grand_p - gau
    gvc = var_color(gv)
    gpct = gv / grand_p if grand_p else 0
    rows.append(
        '<tr class="total-row">'
        + f'<td><strong>Total</strong></td>'
        + f'<td class="num"><strong>{fmt_num(grand_p)}</strong></td>'
        + f'<td class="num"><strong>{fmt_num(grand_u)}</strong></td>'
        + f'<td class="num" style="color:#9A9BA0;"><strong>{fmt_num(grand_w)}</strong></td>'
        + f'<td class="num" style="color:#9A9BA0;"><strong>{fmt_pct(grand_w / grand_u if grand_u else 0)}</strong></td>'
        + f'<td class="num" style="color:{gvc};font-weight:700"><strong>{fmt_num_parens(gv)}</strong></td>'
        + f'<td class="num" style="color:{gvc};font-weight:700"><strong>{fmt_pct(gpct)}</strong></td>'
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

post_var_color = "#4CAF79" if post_variance >= 0 else "#EF5350"
post_var_pct = post_variance / post_purchased if post_purchased else 0

# ---------------------------------------------------------------------------
# CSS (same Apex brand design system as internal report)
# ---------------------------------------------------------------------------
CSS = """*, *::before, *::after { box-sizing: border-box; }
html { scroll-behavior: smooth; }
body {
    margin: 0; padding: 0;
    font-family: "Helvetica Neue", Arial, sans-serif;
    font-size: 14px; line-height: 1.5; color: #E0E1E6; background: #131416;
}
.header {
    background: linear-gradient(135deg, #0A1A4A, #1A3A8F);
    color: #FFFFFF; padding: 48px 40px 40px;
}
.header h1 { margin: 0 0 8px; font-size: 32px; font-weight: 600; letter-spacing: -0.5px; }
.header .subtitle { font-size: 16px; opacity: 0.85; margin: 0 0 4px; }
.header .from-line { font-size: 14px; opacity: 0.75; margin: 0 0 4px; }
.header .generated { font-size: 13px; opacity: 0.65; margin: 0; }
.nav {
    position: sticky; top: 0; z-index: 100;
    background: #1E1F23; border-bottom: 2px solid #3A3B40;
    padding: 0 40px; display: flex; gap: 0; overflow-x: auto;
}
.nav a {
    color: #5B9BF7; text-decoration: none; font-size: 13px; font-weight: 500;
    padding: 12px 16px; white-space: nowrap;
    border-bottom: 3px solid transparent; transition: border-color 0.2s, color 0.2s;
}
.nav a:hover { color: #82B4FF; border-bottom-color: #5B9BF7; }
.content { max-width: 1340px; margin: 0 auto; padding: 32px 40px 60px; }
.section { margin-bottom: 40px; }
.section-header {
    display: flex; align-items: center; justify-content: space-between;
    cursor: pointer; user-select: none; margin-bottom: 16px;
}
.section-header h2 { margin: 0; font-size: 20px; font-weight: 600; color: #82B4FF; }
.section-header .toggle {
    font-size: 18px; color: #9A9BA0; width: 28px; height: 28px;
    display: flex; align-items: center; justify-content: center;
    border-radius: 50%; transition: background 0.2s;
}
.section-header:hover .toggle { background: #2A2B30; }
.section-body { transition: max-height 0.3s ease; overflow: hidden; }
.section-body.collapsed { max-height: 0 !important; overflow: hidden; }
.kpi-grid { display: flex; flex-wrap: wrap; gap: 20px; margin-bottom: 10px; }
.kpi-card {
    flex: 1 1 200px; background: #1E1F23; border-radius: 12px; padding: 24px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.3), 0 2px 12px rgba(0,0,0,0.2);
    min-width: 190px; border: 1px solid #2A2B30;
}
.kpi-card .kpi-label {
    font-size: 12px; font-weight: 600; text-transform: uppercase;
    letter-spacing: 0.5px; color: #9A9BA0; margin: 0 0 8px;
}
.kpi-card .kpi-value { font-size: 28px; font-weight: 700; margin: 0; line-height: 1.2; }
.kpi-card .kpi-sub { font-size: 12px; color: #9A9BA0; margin: 6px 0 0; }
.table-wrap { overflow-x: auto; border-radius: 8px; box-shadow: 0 1px 4px rgba(0,0,0,0.3); }
table { width: 100%; border-collapse: collapse; font-size: 13px; background: #1E1F23; }
table th {
    background: #0A1A4A; color: #FFFFFF; padding: 10px 14px; text-align: left;
    font-weight: 600; font-size: 12px; text-transform: uppercase; letter-spacing: 0.3px;
    cursor: pointer; white-space: nowrap; user-select: none;
}
table th:hover { background: #1A3A8F; }
table th .sort-arrow { display: inline-block; margin-left: 4px; font-size: 10px; opacity: 0.6; }
table td { padding: 8px 14px; border-bottom: 1px solid #2A2B30; white-space: nowrap; color: #E0E1E6; }
table .num { text-align: right; font-variant-numeric: tabular-nums; }
table .env-name { max-width: 260px; white-space: normal; word-break: break-word; }
table tbody tr:hover { background: #1A2A3D !important; }
.subtotal-row { background: #252830 !important; border-top: 2px solid #3A3D48; }
.total-row { background: #2A2D35 !important; border-top: 2px solid #3A3D48; }
.bar-container { display: flex; align-items: center; gap: 8px; min-width: 160px; }
.bar-fill {
    height: 18px; background: linear-gradient(90deg, #2954F0, #5B9BF7);
    border-radius: 9px; min-width: 3px;
}
.bar-label { font-size: 12px; font-weight: 500; color: #9A9BA0; white-space: nowrap; }
.flag-under {
    background: #1A3A2A; color: #4CAF79; padding: 3px 10px; border-radius: 12px;
    font-size: 11px; font-weight: 600; text-transform: uppercase;
}
.flag-ok {
    background: #252629; color: #9A9BA0; padding: 3px 10px; border-radius: 12px;
    font-size: 11px; font-weight: 600; text-transform: uppercase;
}
.footer {
    text-align: center; padding: 32px 40px; font-size: 12px;
    color: #9A9BA0; border-top: 1px solid #2A2B30;
}
@media print {
    .nav { display: none; }
    .section-header .toggle { display: none; }
    .section-body.collapsed { max-height: none !important; }
    body { background: #FFFFFF !important; color: #333 !important; font-size: 11px; }
    .kpi-card, .bottom-line, .info-box, .context-line, .table-wrap { background: #FFFFFF !important; border-color: #E2E2E2 !important; box-shadow: none !important; }
    .bottom-line p, .info-box p, .kpi-card .kpi-sub { color: #333 !important; }
    .bottom-line .bl-heading, .kpi-card .kpi-label, .section-header h2 { color: #052390 !important; }
    table { background: #FFFFFF !important; }
    table td { color: #333 !important; border-bottom-color: #E2E2E2 !important; }
    table th { background: #052390 !important; }
    .subtotal-row { background: #F0F0F5 !important; }
    .total-row { background: #F5F5F7 !important; }
    .context-line { background: #F5F5F7 !important; }
    .footer { color: #6D6E71 !important; border-top-color: #E2E2E2 !important; }
}
@media (max-width: 768px) {
    .header { padding: 32px 20px 28px; }
    .header h1 { font-size: 24px; }
    .content { padding: 20px; }
    .nav { padding: 0 16px; }
    .kpi-card .kpi-value { font-size: 22px; }
}
"""

# ---------------------------------------------------------------------------
# JS
# ---------------------------------------------------------------------------
JS = """
function toggleSection(header) {
    var body = header.nextElementSibling;
    var arrow = header.querySelector('.toggle');
    if (body.classList.contains('collapsed')) {
        body.classList.remove('collapsed');
        body.style.maxHeight = body.scrollHeight + 'px';
        arrow.innerHTML = '&#9660;';
        setTimeout(function() { body.style.maxHeight = 'none'; }, 300);
    } else {
        body.style.maxHeight = body.scrollHeight + 'px';
        body.offsetHeight;
        body.style.maxHeight = '0';
        body.classList.add('collapsed');
        arrow.innerHTML = '&#9654;';
    }
}
function sortTable(th, colIdx) {
    var table = th.closest('table');
    var tbody = table.querySelector('tbody');
    var rows = Array.from(tbody.querySelectorAll('tr')).filter(function(r) {
        return !r.classList.contains('total-row') && !r.classList.contains('subtotal-row');
    });
    var totalRow = tbody.querySelector('.total-row');
    var asc = th.getAttribute('data-sort-dir') !== 'asc';
    var ths = table.querySelectorAll('th');
    for (var i = 0; i < ths.length; i++) { ths[i].removeAttribute('data-sort-dir'); }
    th.setAttribute('data-sort-dir', asc ? 'asc' : 'desc');
    rows.sort(function(a, b) {
        var av = a.cells[colIdx] ? a.cells[colIdx].textContent.trim() : '';
        var bv = b.cells[colIdx] ? b.cells[colIdx].textContent.trim() : '';
        var an = parseFloat(av.replace(/[$,%()]/g, '').replace(/,/g, ''));
        var bn = parseFloat(bv.replace(/[$,%()]/g, '').replace(/,/g, ''));
        if (av.indexOf('(') === 0) an = -an;
        if (bv.indexOf('(') === 0) bn = -bn;
        if (!isNaN(an) && !isNaN(bn)) { return asc ? an - bn : bn - an; }
        return asc ? av.localeCompare(bv) : bv.localeCompare(av);
    });
    for (var j = 0; j < rows.length; j++) { tbody.appendChild(rows[j]); }
    if (totalRow) { tbody.appendChild(totalRow); }
}
document.querySelectorAll('.nav a').forEach(function(link) {
    link.addEventListener('click', function(e) {
        var id = this.getAttribute('href').substring(1);
        var el = document.getElementById(id);
        if (el) {
            e.preventDefault();
            var body = el.querySelector('.section-body');
            if (body && body.classList.contains('collapsed')) {
                var header = el.querySelector('.section-header');
                if (header) toggleSection(header);
            }
            el.scrollIntoView({ behavior: 'smooth', block: 'start' });
        }
    });
});
document.querySelectorAll('.section-body').forEach(function(body) {
    if (!body.classList.contains('collapsed')) {
        body.style.maxHeight = 'none';
    }
});
"""

SA = '<span class="sort-arrow">&#9650;&#9660;</span>'

def th_row(headers):
    return "".join(f'<th onclick="sortTable(this,{i})">{h} {SA}</th>' for i, h in enumerate(headers))

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

# --- Nav ---
html += '<nav class="nav">\n'
html += '    <a href="#summary">Summary</a>\n'
html += '    <a href="#review">Items for Review</a>\n'
html += '    <a href="#envelope-types">By Type</a>\n'
html += '    <a href="#reference">Reference</a>\n'
html += '</nav>\n'
html += '<div class="content">\n'

# ===== SECTION 1: SUMMARY =====
html += '<div class="section" id="summary">\n'
html += '    <div class="section-header" onclick="toggleSection(this)">\n'
html += '        <h2>Summary</h2>\n'
html += '        <span class="toggle">&#9660;</span>\n'
html += '    </div>\n'
html += '    <div class="section-body">\n'
html += '        <p style="font-size:14px;color:#E0E1E6;margin:0 0 20px;line-height:1.6;">This report presents Apex&rsquo;s reconciliation of envelope purchases versus usage across 46 months of post-settlement activity. We are sharing this for Broadridge&rsquo;s review and validation.</p>\n'

# KPI cards
html += '        <div class="kpi-grid">\n'
html += f'            <div class="kpi-card"><p class="kpi-label">Purchased</p><p class="kpi-value" style="color:#5B9BF7">{fmt_num(post_purchased)}</p><p class="kpi-sub">Total envelopes ordered</p></div>\n'
html += f'            <div class="kpi-card"><p class="kpi-label">Used + Wastage</p><p class="kpi-value" style="color:#5B9BF7">{fmt_num(post_adj_used)}</p><p class="kpi-sub">{fmt_num(post_used)} used + {fmt_num(post_wastage)} wastage (contract max)</p></div>\n'
html += f'            <div class="kpi-card"><p class="kpi-label">Variance</p><p class="kpi-value" style="color:{post_var_color}">{fmt_num_parens(post_variance)}</p><p class="kpi-sub">{fmt_pct(post_var_pct)} of purchased</p></div>\n'
html += f'            <div class="kpi-card"><p class="kpi-label">Months Covered</p><p class="kpi-value" style="color:#5B9BF7">{post_months}</p><p class="kpi-sub">Mar 2022 &ndash; Dec 2025</p></div>\n'
html += '        </div>\n'

# Year-by-year table
html += '        <div class="table-wrap" style="margin-top:20px;"><table>\n'
html += '            <thead><tr><th>Year</th><th>Purchased</th><th>Used</th><th>Wastage</th><th>Wastage %</th><th>Adj. Variance</th><th>Var %</th><th>Invoiced</th></tr></thead>\n'
html += '            <tbody>\n'
for yr in sorted(post_yearly.keys()):
    d = post_yearly[yr]
    yp, yu, mc, yc, yi, yw = d
    yau = yu + yw
    yv = yp - yau
    vc = var_color(yv)
    vpct = yv / yp if yp else 0
    yr_label = f"{yr} (Mar&ndash;Dec)" if yr == 2022 else str(yr)
    html += f'            <tr><td>{yr_label}</td>'
    html += f'<td class="num">{fmt_num(yp)}</td>'
    html += f'<td class="num">{fmt_num(yu)}</td>'
    html += f'<td class="num" style="color:#9A9BA0;">{fmt_num(yw)}</td>'
    wpct = yw / yu if yu else 0
    html += f'<td class="num" style="color:#9A9BA0;">{fmt_pct(wpct)}</td>'
    html += f'<td class="num" style="color:{vc};font-weight:600">{fmt_num_parens(yv)}</td>'
    html += f'<td class="num" style="color:{vc}">{fmt_pct(vpct)}</td>'
    html += f'<td class="num">{fmt_money(yi)}</td></tr>\n'
# Total row
html += f'            <tr class="total-row"><td><strong>Total</strong></td>'
html += f'<td class="num"><strong>{fmt_num(post_purchased)}</strong></td>'
html += f'<td class="num"><strong>{fmt_num(post_used)}</strong></td>'
html += f'<td class="num" style="color:#9A9BA0;"><strong>{fmt_num(post_wastage)}</strong></td>'
total_wpct = post_wastage / post_used if post_used else 0
html += f'<td class="num" style="color:#9A9BA0;"><strong>{fmt_pct(total_wpct)}</strong></td>'
html += f'<td class="num" style="color:{post_var_color};font-weight:700"><strong>{fmt_num_parens(post_variance)}</strong></td>'
html += f'<td class="num" style="color:{post_var_color};font-weight:700"><strong>{fmt_pct(post_var_pct)}</strong></td>'
html += f'<td class="num"><strong>{fmt_money(post_invoiced)}</strong></td></tr>\n'
html += '            </tbody>\n'
html += '        </table></div>\n'

html += '    </div>\n'
html += '</div>\n\n'

# ===== SECTION 2: ITEMS FOR REVIEW =====
html += '<div class="section" id="review">\n'
html += '    <div class="section-header" onclick="toggleSection(this)">\n'
html += '        <h2>Items for review</h2>\n'
html += '        <span class="toggle">&#9660;</span>\n'
html += '    </div>\n'
html += '    <div class="section-body">\n'

# Wastage observation
html += '        <div style="background:#1E1F23;border-radius:12px;padding:20px 24px;box-shadow:0 1px 4px rgba(0,0,0,0.3);margin-bottom:16px;border:1px solid #2A2B30;">\n'
html += '            <p style="font-size:13px;font-weight:700;text-transform:uppercase;letter-spacing:0.3px;color:#82B4FF;margin:0 0 12px;">Wastage &mdash; contractual vs. operational</p>\n'
html += f'            <p style="font-size:13px;line-height:1.7;color:#E0E1E6;margin:0 0 8px;">The contract caps the wastage charge at <strong>5%</strong> (original, through Dec 2023) '
html += f'and <strong>2%</strong> (Amendment No. 1, Jan 2024+). Broadridge has separately confirmed that actual operational wastage runs 10&ndash;15%.</p>\n'
html += '            <table style="margin-top:8px;width:auto;background:transparent;font-size:13px;">\n'
html += f'                <tr><td style="border:none;padding:4px 16px 4px 0;color:#9A9BA0;">Contract max wastage (5%/2% blended)</td>'
html += f'<td style="border:none;padding:4px 0;color:#E0E1E6;font-weight:600;">{fmt_num(_contract_waste)} envelopes</td></tr>\n'
html += f'                <tr><td style="border:none;padding:4px 16px 4px 0;color:#9A9BA0;">Operational wastage at 10%</td>'
html += f'<td style="border:none;padding:4px 0;color:#E0E1E6;font-weight:600;">{fmt_num(_actual_waste_lo)} envelopes</td></tr>\n'
html += f'                <tr><td style="border:none;padding:4px 16px 4px 0;color:#9A9BA0;">Operational wastage at 15%</td>'
html += f'<td style="border:none;padding:4px 0;color:#E0E1E6;font-weight:600;">{fmt_num(_actual_waste_hi)} envelopes</td></tr>\n'
html += f'                <tr><td style="border:none;padding:4px 16px 4px 0;color:#9A9BA0;">Excess beyond contract allowance</td>'
html += f'<td style="border:none;padding:4px 0;color:#82B4FF;font-weight:600;">{fmt_num(_excess_lo)}&ndash;{fmt_num(_excess_hi)} envelopes</td></tr>\n'
html += '            </table>\n'
html += f'            <p style="font-size:12px;color:#9A9BA0;margin:8px 0 0;">Per Section 4, excess wastage beyond the contractual rate is Broadridge&rsquo;s responsibility. We would like to confirm the current operational wastage rate and how it is reflected in invoicing.</p>\n'
html += '        </div>\n'

html += '    </div>\n'
html += '</div>\n\n'

# ===== SECTION 3: PURCHASES & USAGE BY ENVELOPE TYPE =====
html += '<div class="section" id="envelope-types">\n'
html += '    <div class="section-header" onclick="toggleSection(this)">\n'
html += '        <h2>Purchases &amp; usage by envelope type</h2>\n'
html += '        <span class="toggle">&#9660;</span>\n'
html += '    </div>\n'
html += '    <div class="section-body">\n'
html += '        <div class="table-wrap"><table class="sortable">\n'
html += '            <thead><tr>' + th_row(["Envelope Type","Purchased","Used","Wastage","Wastage %","Adj. Variance","Variance %"]) + '</tr></thead>\n'
html += '            <tbody>\n' + combined_env_rows + '\n            </tbody>\n'
html += '        </table></div>\n'
html += '        <p style="font-size:12px;color:#9A9BA0;margin:12px 0 0;">Related envelope SKUs grouped by physical size. Usage mapped from billing data via product category, flat/fold, and address type.</p>\n'
html += '        <p style="font-size:12px;color:#9A9BA0;margin:4px 0 0;"><strong style="color:#E0E1E6;">9x12 Flat Confirms:</strong> The deficit reflects a temporary production surge in mid-2022 (Ridge/Penson migration) when Broadridge routed a portion of confirms through flat production. Pre-existing buffer stock covered the shortfall. This category represents 0.7% of total volume.</p>\n'

# NI/PFC transition context
html += '        <div style="background:#1E1F23;border-radius:12px;padding:16px 24px;margin-top:16px;border:1px solid #2A2B30;font-size:13px;line-height:1.7;color:#E0E1E6;">\n'
html += '            <p style="font-weight:700;color:#82B4FF;margin:0 0 8px;">NI / PFC transition (Oct 2022)</p>\n'
html += '            <p style="margin:0 0 8px;"><strong>Before Oct 2022:</strong> All #10 fold confirms, letters, and checks (domestic + foreign) used ENVCONPFSN10NI (No Imprint).</p>\n'
html += '            <p style="margin:0 0 8px;"><strong>After Oct 2022:</strong> Domestic mail switched to ENVAPXN10PFSCONN10IND(10/22) (Pre-Sorted First-Class); foreign mail remained on ENVCONPFSN10NI at ~8,000/month.</p>\n'
html += '            <p style="margin:0;">The same transition applies to N14 statements (PFC for domestic, NI for foreign) and 9x12 flats. NI and PFC versions are physically interchangeable &mdash; NI envelopes can be used for domestic mail by adding indicia at time of mailing.</p>\n'
html += '        </div>\n'

# SKU-level breakdown
html += '        <h3 style="color:#82B4FF;font-size:16px;margin:32px 0 12px;">By SKU</h3>\n'
html += '        <div class="table-wrap"><table class="sortable">\n'
html += '            <thead><tr>' + th_row(["SKU","Purchased","Used","Wastage","Wastage %","Adj. Variance","Variance %"]) + '</tr></thead>\n'
html += '            <tbody>\n' + sku_recon_rows + '\n            </tbody>\n'
html += '        </table></div>\n'

html += '    </div>\n</div>\n\n'

# ===== SECTION 5: REFERENCE =====
html += '<div class="section" id="reference">\n'
html += '    <div class="section-header" onclick="toggleSection(this)">\n'
html += '        <h2>Reference</h2>\n'
html += '        <span class="toggle">&#9660;</span>\n'
html += '    </div>\n'
html += '    <div class="section-body">\n'

# Data sources & methodology
html += '        <h3 style="color:#82B4FF;font-size:16px;margin:0 0 16px;">Data sources &amp; methodology</h3>\n'
html += '        <div style="background:#1E1F23;border-radius:12px;padding:20px 24px;box-shadow:0 1px 4px rgba(0,0,0,0.3);margin-bottom:16px;border:1px solid #2A2B30;">\n'
html += '            <table style="width:auto;background:transparent;font-size:13px;">\n'
html += f'                <tr><td style="border:none;padding:4px 16px 4px 0;color:#9A9BA0;white-space:nowrap;">Reconciliation period</td><td style="border:none;padding:4px 0;color:#E0E1E6;font-weight:600;">January 2020 &ndash; December 2025 (post-settlement focus: March 2022+)</td></tr>\n'
html += f'                <tr><td style="border:none;padding:4px 16px 4px 0;color:#9A9BA0;white-space:nowrap;">Purchase reports processed</td><td style="border:none;padding:4px 0;color:#E0E1E6;font-weight:600;">{purchase_files} files (monthly Broadridge Purchase Reports, FY&rsquo;20&ndash;FY&rsquo;25)</td></tr>\n'
html += f'                <tr><td style="border:none;padding:4px 16px 4px 0;color:#9A9BA0;white-space:nowrap;">Billing workbooks processed</td><td style="border:none;padding:4px 0;color:#E0E1E6;font-weight:600;">{billing_files} files (monthly Billing Workbooks + Billing Master 2020&ndash;2021)</td></tr>\n'
html += f'                <tr><td style="border:none;padding:4px 16px 4px 0;color:#9A9BA0;white-space:nowrap;">Usage-to-envelope mapping</td><td style="border:none;padding:4px 0;color:#E0E1E6;">Product category + Flat_Fold + Address_Type fields from Volume Data (per Brandon Koebel, Sep 2023)</td></tr>\n'
html += f'                <tr><td style="border:none;padding:4px 16px 4px 0;color:#9A9BA0;white-space:nowrap;">NI/PFC usage cutover</td><td style="border:none;padding:4px 0;color:#E0E1E6;">Domestic #10 usage mapped to ENVCONPFSN10NI before Oct 2022, ENVAPXN10 (PFC) after</td></tr>\n'
html += '            </table>\n'
html += '        </div>\n'

# Contract terms
html += '        <h3 style="color:#82B4FF;font-size:16px;margin:24px 0 16px;">Contract terms &mdash; envelope materials</h3>\n'

# Original contract
html += '        <div style="background:#1E1F23;border-radius:12px;padding:20px 24px;box-shadow:0 1px 4px rgba(0,0,0,0.3);margin-bottom:16px;border:1px solid #2A2B30;">\n'
html += '            <p style="font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:0.3px;color:#82B4FF;margin:0 0 8px;">Original Contract &mdash; Section 4, Compensation</p>\n'
html += '            <p style="font-size:12px;color:#9A9BA0;margin:0 0 8px;">GTO Print and Mail Services Schedule, effective January 1, 2019.</p>\n'
html += '            <blockquote style="margin:0;padding:12px 16px;background:#252629;border-left:3px solid #5B9BF7;border-radius:0 8px 8px 0;font-size:13px;line-height:1.7;color:#E0E1E6;">\n'
html += '                &ldquo;Materials (such as paper, envelopes, and inserts) and postage, presort and insert related fees are not included in the Annual Fee and will be charged separately. Materials are billed at cost plus wastage for generic stock. <strong>Specifically, the wastage charge is 10% for any generic paper stock and 5% for generic envelope stock.</strong> For generic stock, the unit rate will be billed based on usage. For Client specific stock, the unit rate will be based on receipt of such stock.&rdquo;\n'
html += '            </blockquote>\n'
html += '        </div>\n'

# Amendment
html += '        <div style="background:#1E1F23;border-radius:12px;padding:20px 24px;box-shadow:0 1px 4px rgba(0,0,0,0.3);margin-bottom:16px;border:1px solid #2A2B30;">\n'
html += '            <p style="font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:0.3px;color:#82B4FF;margin:0 0 8px;">Amendment No. 1 &mdash; Section 4, Compensation (replaces original)</p>\n'
html += '            <p style="font-size:12px;color:#9A9BA0;margin:0 0 8px;">GTO Print and Mail Services Schedule Amendment No. 1, effective January 1, 2024. Term extended through December 31, 2028.</p>\n'
html += '            <blockquote style="margin:0;padding:12px 16px;background:#252629;border-left:3px solid #5B9BF7;border-radius:0 8px 8px 0;font-size:13px;line-height:1.7;color:#E0E1E6;">\n'
html += '                &ldquo;Materials are billed at inventory cost plus 10% margin. Inventory cost means for (i) Client specific inventory: vendor price; and (ii) generic inventory: vendor price plus wastage as follows: <strong>10% for continuous form, 3% for cutsheet, and 2% for envelopes.</strong> For generic stock, the unit rate will be billed based on usage. For Client specific stock, the unit rate will be based on receipt of such stock.&rdquo;\n'
html += '            </blockquote>\n'
html += '        </div>\n'

# Materials obligation
html += '        <div style="background:#1E1F23;border-radius:12px;padding:20px 24px;box-shadow:0 1px 4px rgba(0,0,0,0.3);margin-bottom:16px;border:1px solid #2A2B30;">\n'
html += '            <p style="font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:0.3px;color:#82B4FF;margin:0 0 8px;">Original Contract &mdash; Section 8, Materials</p>\n'
html += '            <blockquote style="margin:0;padding:12px 16px;background:#252629;border-left:3px solid #5B9BF7;border-radius:0 8px 8px 0;font-size:13px;line-height:1.7;color:#E0E1E6;">\n'
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
html += '        <div style="background:#1E1F23;border-radius:12px;padding:20px 24px;box-shadow:0 1px 4px rgba(0,0,0,0.3);margin-bottom:16px;border:1px solid #2A2B30;">\n'
html += '            <p style="font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:0.3px;color:#82B4FF;margin:0 0 8px;">Broadridge confirmation &mdash; operational wastage rates</p>\n'
html += '            <p style="font-size:12px;color:#9A9BA0;margin:0 0 8px;">Brandon Koebel (Sr. Client Relationship Manager), emails Sep&ndash;Nov 2022.</p>\n'
html += '            <blockquote style="margin:0;padding:12px 16px;background:#252629;border-left:3px solid #5B9BF7;border-radius:0 8px 8px 0;font-size:13px;line-height:1.7;color:#E0E1E6;">\n'
html += '                &ldquo;Wastage is roughly 10%&hellip; This includes envelopes that are damaged, need to be reprinted and reinserted, etc.&rdquo; (Nov 7, 2022)<br>\n'
html += '                &ldquo;Did not account for any waste or spoilage (<strong>typically 10&ndash;15%</strong>).&rdquo; (Sep 29, 2022)\n'
html += '            </blockquote>\n'
html += '            <p style="font-size:12px;color:#9A9BA0;margin:8px 0 0;"><strong style="color:#E0E1E6;">Note:</strong> Contractual wastage (billed to Apex) is 5% pre-2024 / 2% post-amendment. Operational wastage (10&ndash;15%) is higher but embedded in the &ldquo;Used&rdquo; figure &mdash; not separately charged.</p>\n'
html += '        </div>\n'

# Pre-settlement context
html += f'        <div style="background:#1E1F23;border-radius:12px;padding:16px 24px;margin-bottom:16px;border:1px solid #2A2B30;font-size:13px;color:#9A9BA0;line-height:1.7;">\n'
html += f'            <p style="font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:0.3px;color:#82B4FF;margin:0 0 8px;">Pre-settlement context</p>\n'
html += f'            <strong style="color:#E0E1E6;">Jan 2020 &ndash; Feb 2022:</strong> {fmt_num(pre_purchased)} purchased, {fmt_num(pre_used)} used. '
html += f'These costs were absorbed by Broadridge per the pass-through paper and envelope dispute settlement. '
html += f'<strong style="color:#E0E1E6;">Full period (Jan 2020 &ndash; Dec 2025):</strong> {fmt_num(total_purchased)} purchased, {fmt_num(total_used)} used, '
html += f'{"+" if net_variance >= 0 else ""}{fmt_num_parens(net_variance)} variance.\n'
html += '        </div>\n'

# Generic stock classification
html += '        <h3 style="color:#82B4FF;font-size:16px;margin:24px 0 16px;">Generic stock classification</h3>\n'
html += '        <div style="background:#1E1F23;border-radius:12px;padding:20px 24px;box-shadow:0 1px 4px rgba(0,0,0,0.3);margin-bottom:16px;border:1px solid #2A2B30;">\n'
html += '            <p style="font-size:13px;line-height:1.7;color:#E0E1E6;margin:0 0 12px;">Both the original contract and Amendment No. 1 distinguish between <strong>generic stock</strong> (billed on usage, lower wastage) and <strong>client-specific stock</strong> (billed on receipt). All Apex envelope types are generic stock:</p>\n'
html += '            <div class="table-wrap" style="margin:0;"><table style="font-size:13px;">\n'
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
html += '            <p style="font-size:12px;color:#9A9BA0;margin:12px 0 0;">All envelopes are standard double-window envelopes with no company logos, branding, or custom design elements. Return and recipient addresses are visible through the windows from the printed content inside. PFC (Pre-Sorted First-Class) indicia is a functional USPS postage marking, not client branding. NI (No Imprint) envelopes are completely blank. Supplier: United Envelope LLC, Mt. Pocono, PA.</p>\n'
html += '        </div>\n'

html += '    </div>\n</div>\n\n'

html += '</div><!-- end .content -->\n\n'

# Footer
html += '<div class="footer">\n'
html += '    Confidential &mdash; Apex Clearing Corporation<br>\n'
html += '    Prepared for reconciliation review with Broadridge Financial Solutions\n'
html += '</div>\n'

# JS
html += f'<script>\n{JS}\n</script>\n'
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
