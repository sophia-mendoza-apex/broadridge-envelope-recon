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
            return "#186741"
        if val < 0:
            return "#9D1526"
    except (ValueError, TypeError):
        pass
    return "#6D6E71"

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
post_variance = post_purchased - post_used
post_months = len(post)

# Year-by-year post-settlement
from collections import defaultdict as _defaultdict
post_yearly = _defaultdict(lambda: [0, 0, 0])  # purchased, used, month_count
for _, r in post.iterrows():
    yr = 2000 + int(r["Month"].split('-')[1])
    post_yearly[yr][0] += safe(r["Envelopes Purchased"])
    post_yearly[yr][1] += safe(r["Envelopes Used (Volume)"])
    post_yearly[yr][2] += 1

# Envelope type lookups — filtered to post-settlement
post_type_mask = by_type_monthly["Month"].apply(lambda x: month_label_to_sortkey(x) >= settlement_key)
by_type = by_type_monthly[post_type_mask].groupby("Envelope Type", as_index=False).agg({"Purchased": "sum"})
by_type.rename(columns={"Purchased": "Total Purchased"}, inplace=True)

post_usage_type_mask = usage_by_env_type_monthly["Month"].apply(lambda x: month_label_to_sortkey(x) >= settlement_key)
usage_by_env_type = usage_by_env_type_monthly[post_usage_type_mask].groupby("Envelope Type", as_index=False).agg({"Envelopes Used": "sum"})
usage_by_env_type.rename(columns={"Envelopes Used": "Total Envelopes Used"}, inplace=True)

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
# SVG bar chart (purchased vs used, post-settlement)
# ---------------------------------------------------------------------------
def build_svg_chart():
    months = post["Month"].tolist()
    purchased = post["Envelopes Purchased"].tolist()
    used = post["Envelopes Used (Volume)"].tolist()
    n = len(months)
    if n == 0:
        return ""
    cw, ch = 1200, 400
    ml, mr, mt, mb = 80, 30, 50, 80
    pw = cw - ml - mr
    ph = ch - mt - mb
    max_val = max(max(purchased), max(used)) * 1.1
    bgw = pw / n
    bw = bgw * 0.35
    L = []
    L.append(f'<svg viewBox="0 0 {cw} {ch}" style="width:100%;max-width:{cw}px;height:auto;" xmlns="http://www.w3.org/2000/svg">')
    L.append(f'<rect width="{cw}" height="{ch}" fill="#FFFFFF" rx="12"/>')
    for i in range(6):
        yv = max_val * i / 5
        yp = mt + ph - (ph * i / 5)
        L.append(f'<line x1="{ml}" y1="{yp:.1f}" x2="{cw - mr}" y2="{yp:.1f}" stroke="#E2E2E2" stroke-width="1"/>')
        lb = f"{yv / 1000000:.1f}M" if yv >= 1000000 else f"{yv / 1000:.0f}K"
        L.append(f'<text x="{ml - 8}" y="{yp + 4:.1f}" text-anchor="end" fill="#6D6E71" font-size="11" font-family="Helvetica Neue,Arial,sans-serif">{lb}</text>')
    for i in range(n):
        xc = ml + (i + 0.5) * bgw
        hp = (purchased[i] / max_val) * ph if max_val > 0 else 0
        xp = xc - bw - 1
        yp = mt + ph - hp
        L.append(f'<rect x="{xp:.1f}" y="{yp:.1f}" width="{bw:.1f}" height="{hp:.1f}" fill="#2954F0" rx="2"><title>{months[i]} Purchased: {int(purchased[i]):,}</title></rect>')
        hu = (used[i] / max_val) * ph if max_val > 0 else 0
        xu = xc + 1
        yu = mt + ph - hu
        L.append(f'<rect x="{xu:.1f}" y="{yu:.1f}" width="{bw:.1f}" height="{hu:.1f}" fill="#3F8EFC" rx="2"><title>{months[i]} Used: {int(used[i]):,}</title></rect>')
        if i % 3 == 0:
            ly = mt + ph + 20
            L.append(f'<text x="{xc:.1f}" y="{ly}" text-anchor="middle" fill="#6D6E71" font-size="10" font-family="Helvetica Neue,Arial,sans-serif" transform="rotate(-45 {xc:.1f} {ly})">{months[i]}</text>')
    lx = ml + 10
    ly2 = mt - 25
    L.append(f'<rect x="{lx}" y="{ly2}" width="14" height="14" fill="#2954F0" rx="2"/>')
    L.append(f'<text x="{lx + 20}" y="{ly2 + 12}" fill="#052390" font-size="12" font-family="Helvetica Neue,Arial,sans-serif" font-weight="500">Purchased</text>')
    L.append(f'<rect x="{lx + 110}" y="{ly2}" width="14" height="14" fill="#3F8EFC" rx="2"/>')
    L.append(f'<text x="{lx + 130}" y="{ly2 + 12}" fill="#052390" font-size="12" font-family="Helvetica Neue,Arial,sans-serif" font-weight="500">Used (Volume)</text>')
    L.append("</svg>")
    return "\n".join(L)

# ---------------------------------------------------------------------------
# Monthly detail rows (post-settlement) with annual subtotals and grand total
# ---------------------------------------------------------------------------
def build_monthly_rows():
    rows = []
    running_balance = 0
    yr_p = yr_u = 0
    prev_year = None

    def _year_from_label(label):
        return 2000 + int(label.split('-')[1])

    def _subtotal_row(year, yp, yu, rb):
        yv = yp - yu
        vc = var_color(yv)
        rbc = var_color(rb)
        vpct = yv / yp if yp else 0
        yr_label = f"{year} (Mar&ndash;Dec)" if year == 2022 else str(year)
        return (
            '<tr class="subtotal-row">'
            + f'<td><strong>{yr_label} Total</strong></td>'
            + f'<td class="num"><strong>{fmt_num(yp)}</strong></td>'
            + f'<td class="num"><strong>{fmt_num(yu)}</strong></td>'
            + f'<td class="num" style="color:{vc};font-weight:700"><strong>{fmt_num_parens(yv)}</strong></td>'
            + f'<td class="num" style="color:{vc};font-weight:700"><strong>{fmt_pct(vpct)}</strong></td>'
            + f'<td class="num" style="color:{rbc};font-weight:700"><strong>{fmt_num_parens(rb)}</strong></td>'
            + '</tr>'
        )

    for i, (_, r) in enumerate(post.iterrows()):
        cur_year = _year_from_label(r["Month"])

        if prev_year is not None and cur_year != prev_year:
            rows.append(_subtotal_row(prev_year, yr_p, yr_u, running_balance))
            yr_p = yr_u = 0

        prev_year = cur_year
        p = safe(r["Envelopes Purchased"])
        u = safe(r["Envelopes Used (Volume)"])
        vv = p - u
        running_balance += vv
        yr_p += p; yr_u += u

        vc = var_color(vv)
        rbc = var_color(running_balance)
        bg = "#F5F5F7" if i % 2 == 1 else "#FFFFFF"
        rows.append(
            f'<tr style="background:{bg}">'
            + f'<td>{r["Month"]}</td>'
            + f'<td class="num">{fmt_num(p)}</td>'
            + f'<td class="num">{fmt_num(u)}</td>'
            + f'<td class="num" style="color:{vc};font-weight:600">{fmt_num_parens(vv)}</td>'
            + f'<td class="num" style="color:{vc}">{fmt_pct(r["Variance %"])}</td>'
            + f'<td class="num" style="color:{rbc};font-weight:600">{fmt_num_parens(running_balance)}</td>'
            + '</tr>'
        )

    # Final year subtotal
    if prev_year is not None:
        rows.append(_subtotal_row(prev_year, yr_p, yr_u, running_balance))

    # Grand total
    gv = post_purchased - post_used
    vc = var_color(gv)
    rbc = var_color(running_balance)
    gpct = gv / post_purchased if post_purchased else 0
    rows.append(
        '<tr class="total-row">'
        + f'<td><strong>Grand Total</strong></td>'
        + f'<td class="num"><strong>{fmt_num(post_purchased)}</strong></td>'
        + f'<td class="num"><strong>{fmt_num(post_used)}</strong></td>'
        + f'<td class="num" style="color:{vc};font-weight:700"><strong>{fmt_num_parens(gv)}</strong></td>'
        + f'<td class="num" style="color:{vc};font-weight:700"><strong>{fmt_pct(gpct)}</strong></td>'
        + f'<td class="num" style="color:{vc};font-weight:700"><strong>{fmt_num_parens(running_balance)}</strong></td>'
        + '</tr>'
    )
    return "\n".join(rows)

# ---------------------------------------------------------------------------
# Combined envelope group rows (purchased + used + variance)
# ---------------------------------------------------------------------------
def build_combined_env_rows():
    rows = []
    grand_p = grand_u = 0
    for g in ENVELOPE_GROUPS:
        p = sum(safe(purchase_by_type.get(t, 0)) for t in g["purchase_types"])
        u = sum(safe(usage_by_type_dict.get(t, 0)) for t in g["usage_types"])
        v = p - u
        grand_p += p
        grand_u += u
        vc = var_color(v)
        vpct = v / p if p else 0
        rows.append(
            '<tr>'
            + f'<td class="env-name">{g["label"]}</td>'
            + f'<td class="num">{fmt_num(p)}</td>'
            + f'<td class="num">{fmt_num(u)}</td>'
            + f'<td class="num" style="color:{vc};font-weight:600">{fmt_num_parens(v)}</td>'
            + f'<td class="num" style="color:{vc}">{fmt_pct(vpct)}</td>'
            + '</tr>'
        )
    gv = grand_p - grand_u
    gvc = var_color(gv)
    gpct = gv / grand_p if grand_p else 0
    rows.append(
        '<tr class="total-row">'
        + f'<td><strong>Total</strong></td>'
        + f'<td class="num"><strong>{fmt_num(grand_p)}</strong></td>'
        + f'<td class="num"><strong>{fmt_num(grand_u)}</strong></td>'
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
    grand_p = grand_u = 0
    for t in all_types:
        p = safe(purchase_by_type.get(t, 0))
        u = safe(usage_by_type_dict.get(t, 0))
        if p == 0 and u == 0:
            continue
        v = p - u
        grand_p += p
        grand_u += u
        vc = var_color(v)
        vpct = v / p if p else 0
        rows.append(
            '<tr>'
            + f'<td class="env-name">{t}</td>'
            + f'<td class="num">{fmt_num(p)}</td>'
            + f'<td class="num">{fmt_num(u)}</td>'
            + f'<td class="num" style="color:{vc};font-weight:600">{fmt_num_parens(v)}</td>'
            + f'<td class="num" style="color:{vc}">{fmt_pct(vpct)}</td>'
            + '</tr>'
        )
    gv = grand_p - grand_u
    gvc = var_color(gv)
    gpct = gv / grand_p if grand_p else 0
    rows.append(
        '<tr class="total-row">'
        + f'<td><strong>Total</strong></td>'
        + f'<td class="num"><strong>{fmt_num(grand_p)}</strong></td>'
        + f'<td class="num"><strong>{fmt_num(grand_u)}</strong></td>'
        + f'<td class="num" style="color:{gvc};font-weight:700"><strong>{fmt_num_parens(gv)}</strong></td>'
        + f'<td class="num" style="color:{gvc};font-weight:700"><strong>{fmt_pct(gpct)}</strong></td>'
        + '</tr>'
    )
    return "\n".join(rows)

# ---------------------------------------------------------------------------
# Envelope Specifications
# ---------------------------------------------------------------------------
ENVELOPE_SPECS = [
    ("ENVMEAPEXN14PFC", "926131", '4&frac34; x 11<sup>7</sup>&frasl;<sub>16</sub>', "N14 Booklet", "Domestic Fold Statement", "Pre-Sorted First-Class", "24WW", "All domestic monthly/quarterly account statements that fold into the envelope"),
    ("ENVMEAPEX9X12PFC", "830851", "9 x 12", "Flat", "Domestic Flat Statement", "Pre-Sorted First-Class", "24WW", "Domestic statements too large to fold (high page count)"),
    ("ENVMERIDGEN14NI11/08", "942095", '4&frac34; x 11<sup>7</sup>&frasl;<sub>16</sub>', "N14 Booklet", "Foreign Fold Statement", "No Imprint", "24WW", "Foreign statements &mdash; postage applied at mailing"),
    ("ENVMERIDGE9X12NI11/08", "823804", "9 x 12", "Flat", "Foreign Flat Statement", "No Imprint", "24WW", "Foreign flat statements &mdash; postage applied at mailing"),
    ("ENVAPXN10PFSCONN10IND(10/22)", "992124", '4&frac18; x 9&frac12;', "#10 Booklet", "Domestic Confirms + Letters", "Pre-Sorted First-Class", "24WW", "Replaced ENVCONPFSN10NI in Oct 2022 &mdash; updated postal permit"),
    ("ENVCONPFSN10NI", "856743", '4&frac18; x 9&frac12;', "#10 Booklet", "Foreign Confirms + Letters", "No Imprint", "24WW", "Foreign confirms and letters &mdash; postage applied at mailing"),
    ("ENVCONRIDGE9X12DW", "818105", "9 x 12", "Flat", "Flat Confirms (Dom. + Foreign)", "No Imprint", "24WW", "Oversize confirms that cannot fold into #10 envelope"),
]

def build_envelope_spec_rows():
    rows = []
    for wms, order, size, style, mail_type, postage, paper, notes in ENVELOPE_SPECS:
        postage_cls = "flag-under" if "Pre-Sorted" in postage else "flag-ok"
        rows.append(
            '<tr>'
            + f'<td class="env-name" style="font-weight:600;font-size:12px;font-family:monospace">{wms}</td>'
            + f'<td>{mail_type}</td>'
            + f'<td>{size}</td>'
            + f'<td>{style}</td>'
            + f'<td><span class="{postage_cls}">{postage}</span></td>'
            + f'<td style="font-size:12px;white-space:normal;max-width:240px">{notes}</td>'
            + '</tr>'
        )
    return "\n".join(rows)

# ---------------------------------------------------------------------------
# Build all data components
# ---------------------------------------------------------------------------
svg_chart = build_svg_chart()
monthly_rows = build_monthly_rows()
combined_env_rows = build_combined_env_rows()
sku_recon_rows = build_sku_recon_rows()
envelope_spec_rows = build_envelope_spec_rows()

post_var_color = "#186741" if post_variance >= 0 else "#9D1526"
post_var_pct = post_variance / post_purchased if post_purchased else 0

# ---------------------------------------------------------------------------
# CSS (same Apex brand design system as internal report)
# ---------------------------------------------------------------------------
CSS = """*, *::before, *::after { box-sizing: border-box; }
html { scroll-behavior: smooth; }
body {
    margin: 0; padding: 0;
    font-family: "Helvetica Neue", Arial, sans-serif;
    font-size: 14px; line-height: 1.5; color: #333; background: #FAFAFA;
}
.header {
    background: linear-gradient(135deg, #052390, #2954F0);
    color: #FFFFFF; padding: 48px 40px 40px;
}
.header h1 { margin: 0 0 8px; font-size: 32px; font-weight: 600; letter-spacing: -0.5px; }
.header .subtitle { font-size: 16px; opacity: 0.85; margin: 0 0 4px; }
.header .from-line { font-size: 14px; opacity: 0.75; margin: 0 0 4px; }
.header .generated { font-size: 13px; opacity: 0.65; margin: 0; }
.nav {
    position: sticky; top: 0; z-index: 100;
    background: #FFFFFF; border-bottom: 2px solid #E2E2E2;
    padding: 0 40px; display: flex; gap: 0; overflow-x: auto;
}
.nav a {
    color: #3F8EFC; text-decoration: none; font-size: 13px; font-weight: 500;
    padding: 12px 16px; white-space: nowrap;
    border-bottom: 3px solid transparent; transition: border-color 0.2s, color 0.2s;
}
.nav a:hover { color: #2954F0; border-bottom-color: #2954F0; }
.content { max-width: 1340px; margin: 0 auto; padding: 32px 40px 60px; }
.section { margin-bottom: 40px; }
.section-header {
    display: flex; align-items: center; justify-content: space-between;
    cursor: pointer; user-select: none; margin-bottom: 16px;
}
.section-header h2 { margin: 0; font-size: 20px; font-weight: 600; color: #052390; }
.section-header .toggle {
    font-size: 18px; color: #6D6E71; width: 28px; height: 28px;
    display: flex; align-items: center; justify-content: center;
    border-radius: 50%; transition: background 0.2s;
}
.section-header:hover .toggle { background: #F5F5F7; }
.section-body { transition: max-height 0.3s ease; overflow: hidden; }
.section-body.collapsed { max-height: 0 !important; overflow: hidden; }
.kpi-grid { display: flex; flex-wrap: wrap; gap: 20px; margin-bottom: 10px; }
.kpi-card {
    flex: 1 1 200px; background: #FFFFFF; border-radius: 12px; padding: 24px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.06), 0 2px 12px rgba(0,0,0,0.04);
    min-width: 190px;
}
.kpi-card .kpi-label {
    font-size: 12px; font-weight: 600; text-transform: uppercase;
    letter-spacing: 0.5px; color: #6D6E71; margin: 0 0 8px;
}
.kpi-card .kpi-value { font-size: 28px; font-weight: 700; margin: 0; line-height: 1.2; }
.kpi-card .kpi-sub { font-size: 12px; color: #6D6E71; margin: 6px 0 0; }
.table-wrap { overflow-x: auto; border-radius: 8px; box-shadow: 0 1px 4px rgba(0,0,0,0.06); }
table { width: 100%; border-collapse: collapse; font-size: 13px; background: #FFFFFF; }
table th {
    background: #052390; color: #FFFFFF; padding: 10px 14px; text-align: left;
    font-weight: 600; font-size: 12px; text-transform: uppercase; letter-spacing: 0.3px;
    cursor: pointer; white-space: nowrap; user-select: none;
}
table th:hover { background: #2954F0; }
table th .sort-arrow { display: inline-block; margin-left: 4px; font-size: 10px; opacity: 0.6; }
table td { padding: 8px 14px; border-bottom: 1px solid #F0F0F0; white-space: nowrap; }
table .num { text-align: right; font-variant-numeric: tabular-nums; }
table .env-name { max-width: 260px; white-space: normal; word-break: break-word; }
table tbody tr:hover { background: #EBF7FF !important; }
.subtotal-row { background: #EDF0F7 !important; border-top: 2px solid #C5CAD6; }
.total-row { background: #F5F5F7 !important; border-top: 2px solid #E2E2E2; }
.bar-container { display: flex; align-items: center; gap: 8px; min-width: 160px; }
.bar-fill {
    height: 18px; background: linear-gradient(90deg, #2954F0, #3F8EFC);
    border-radius: 9px; min-width: 3px;
}
.bar-label { font-size: 12px; font-weight: 500; color: #6D6E71; white-space: nowrap; }
.flag-under {
    background: #E8F5E9; color: #186741; padding: 3px 10px; border-radius: 12px;
    font-size: 11px; font-weight: 600; text-transform: uppercase;
}
.flag-ok {
    background: #F5F5F7; color: #6D6E71; padding: 3px 10px; border-radius: 12px;
    font-size: 11px; font-weight: 600; text-transform: uppercase;
}
.footer {
    text-align: center; padding: 32px 40px; font-size: 12px;
    color: #6D6E71; border-top: 1px solid #E2E2E2;
}
@media print {
    .nav { display: none; }
    .section-header .toggle { display: none; }
    .section-body.collapsed { max-height: none !important; }
    body { font-size: 11px; }
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
html += '    <a href="#monthly-trend">Trend</a>\n'
html += '    <a href="#monthly-detail">Monthly Detail</a>\n'
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

# KPI cards
html += '        <div class="kpi-grid">\n'
html += f'            <div class="kpi-card"><p class="kpi-label">Purchased</p><p class="kpi-value" style="color:#2954F0">{fmt_num(post_purchased)}</p><p class="kpi-sub">Total envelopes ordered</p></div>\n'
html += f'            <div class="kpi-card"><p class="kpi-label">Used</p><p class="kpi-value" style="color:#2954F0">{fmt_num(post_used)}</p><p class="kpi-sub">Total envelopes used (volume)</p></div>\n'
html += f'            <div class="kpi-card"><p class="kpi-label">Variance</p><p class="kpi-value" style="color:{post_var_color}">{fmt_num_parens(post_variance)}</p><p class="kpi-sub">{fmt_pct(post_var_pct)} of purchased</p></div>\n'
html += f'            <div class="kpi-card"><p class="kpi-label">Months Covered</p><p class="kpi-value" style="color:#2954F0">{post_months}</p><p class="kpi-sub">Mar 2022 &ndash; Dec 2025</p></div>\n'
html += '        </div>\n'

# Year-by-year table
html += '        <div class="table-wrap" style="margin-top:20px;"><table>\n'
html += '            <thead><tr><th>Year</th><th>Purchased</th><th>Used</th><th>Variance</th><th>Var %</th></tr></thead>\n'
html += '            <tbody>\n'
for yr in sorted(post_yearly.keys()):
    d = post_yearly[yr]
    yp, yu, mc = d
    yv = yp - yu
    vc = var_color(yv)
    vpct = yv / yp if yp else 0
    yr_label = f"{yr} (Mar&ndash;Dec)" if yr == 2022 else str(yr)
    html += f'            <tr><td>{yr_label}</td>'
    html += f'<td class="num">{fmt_num(yp)}</td>'
    html += f'<td class="num">{fmt_num(yu)}</td>'
    html += f'<td class="num" style="color:{vc};font-weight:600">{fmt_num_parens(yv)}</td>'
    html += f'<td class="num" style="color:{vc}">{fmt_pct(vpct)}</td></tr>\n'
# Total row
html += f'            <tr class="total-row"><td><strong>Total</strong></td>'
html += f'<td class="num"><strong>{fmt_num(post_purchased)}</strong></td>'
html += f'<td class="num"><strong>{fmt_num(post_used)}</strong></td>'
html += f'<td class="num" style="color:{post_var_color};font-weight:700"><strong>{fmt_num_parens(post_variance)}</strong></td>'
html += f'<td class="num" style="color:{post_var_color};font-weight:700"><strong>{fmt_pct(post_var_pct)}</strong></td></tr>\n'
html += '            </tbody>\n'
html += '        </table></div>\n'

html += '    </div>\n'
html += '</div>\n\n'

# ===== SECTION 2: MONTHLY TREND =====
html += '<div class="section" id="monthly-trend">\n'
html += '    <div class="section-header" onclick="toggleSection(this)">\n'
html += '        <h2>Monthly trend</h2>\n'
html += '        <span class="toggle">&#9660;</span>\n'
html += '    </div>\n'
html += '    <div class="section-body">\n'
html += '        <div style="background:#FFFFFF;border-radius:12px;padding:24px;box-shadow:0 1px 4px rgba(0,0,0,0.06);">\n'
html += svg_chart
html += '\n        </div>\n'
html += '    </div>\n</div>\n\n'

# ===== SECTION 3: MONTHLY DETAIL (expanded by default) =====
html += '<div class="section" id="monthly-detail">\n'
html += '    <div class="section-header" onclick="toggleSection(this)">\n'
html += '        <h2>Monthly detail</h2>\n'
html += '        <span class="toggle">&#9660;</span>\n'
html += '    </div>\n'
html += '    <div class="section-body">\n'
html += '        <div class="table-wrap"><table class="sortable">\n'
html += '            <thead><tr>' + th_row(["Month","Purchased","Used","Net Variance","Variance %","Running Balance"]) + '</tr></thead>\n'
html += '            <tbody>\n' + monthly_rows + '\n            </tbody>\n'
html += '        </table></div>\n'
html += '    </div>\n</div>\n\n'

# ===== SECTION 4: PURCHASES & USAGE BY ENVELOPE TYPE =====
html += '<div class="section" id="envelope-types">\n'
html += '    <div class="section-header" onclick="toggleSection(this)">\n'
html += '        <h2>Purchases &amp; usage by envelope type</h2>\n'
html += '        <span class="toggle">&#9660;</span>\n'
html += '    </div>\n'
html += '    <div class="section-body">\n'
html += '        <div class="table-wrap"><table class="sortable">\n'
html += '            <thead><tr>' + th_row(["Envelope Type","Purchased","Used","Variance","Variance %"]) + '</tr></thead>\n'
html += '            <tbody>\n' + combined_env_rows + '\n            </tbody>\n'
html += '        </table></div>\n'
html += '        <p style="font-size:12px;color:#6D6E71;margin:12px 0 0;">Related envelope SKUs grouped by physical size. Usage mapped from billing data via product category, flat/fold, and address type.</p>\n'

# SKU-level breakdown
html += '        <h3 style="color:#052390;font-size:16px;margin:32px 0 12px;">By SKU</h3>\n'
html += '        <div class="table-wrap"><table class="sortable">\n'
html += '            <thead><tr>' + th_row(["SKU","Purchased","Used","Variance","Variance %"]) + '</tr></thead>\n'
html += '            <tbody>\n' + sku_recon_rows + '\n            </tbody>\n'
html += '        </table></div>\n'

html += '    </div>\n</div>\n\n'

# ===== SECTION 5: REFERENCE (collapsed by default) =====
html += '<div class="section" id="reference">\n'
html += '    <div class="section-header" onclick="toggleSection(this)">\n'
html += '        <h2>Envelope specifications</h2>\n'
html += '        <span class="toggle">&#9654;</span>\n'
html += '    </div>\n'
html += '    <div class="section-body collapsed">\n'
html += '        <p style="font-size:13px;color:#6D6E71;margin:0 0 16px;">All envelopes are double-window, 24WW paper, black ink with crosshatch black inside tint. Supplier: United Envelope LLC, Mt. Pocono, PA.</p>\n'
html += '        <div class="table-wrap"><table>\n'
html += '            <thead><tr><th>WMS Code</th><th>Mail Type</th><th>Size</th><th>Style</th><th>Postage</th><th>Notes</th></tr></thead>\n'
html += '            <tbody>\n' + envelope_spec_rows + '\n            </tbody>\n'
html += '        </table></div>\n'
html += '        <div style="margin-top:12px;display:flex;gap:12px;flex-wrap:wrap;font-size:12px;color:#6D6E71;">\n'
html += '            <span><strong>PFC</strong> = Pre-printed First-Class permit (domestic)</span>\n'
html += '            <span><strong>NI</strong> = No Imprint (foreign &mdash; postage applied at mailing)</span>\n'
html += '            <span><strong>DW</strong> = Double Window</span>\n'
html += '            <span><strong>IND</strong> = Individual (Oct 2022 revision)</span>\n'
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
