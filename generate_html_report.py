#!/usr/bin/env python3
"""
Broadridge Envelope Reconciliation - HTML Report Generator
Reads source data from Excel and generates a self-contained HTML report.
"""

import os
import pandas as pd

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, "Envelope Reconciliation - Source Data.xlsx")
HTML_PATH = os.path.join(BASE_DIR, "Envelope Reconciliation Report.html")

monthly = pd.read_excel(EXCEL_PATH, sheet_name="Monthly Summary")
by_type_monthly = pd.read_excel(EXCEL_PATH, sheet_name="By Envelope Type")
usage_by_product = pd.read_excel(EXCEL_PATH, sheet_name="Usage by Product")
usage_by_env_type_monthly = pd.read_excel(EXCEL_PATH, sheet_name="Usage by Envelope Type")

# Aggregate type-level data to full-period totals (internal report uses full period)
by_type = by_type_monthly.groupby("Envelope Type", as_index=False).agg({"Purchased": "sum", "Total Cost": "sum"})
by_type.rename(columns={"Purchased": "Total Purchased"}, inplace=True)
usage_by_env_type = usage_by_env_type_monthly.groupby("Envelope Type", as_index=False).agg({"Envelopes Used": "sum"})
usage_by_env_type.rename(columns={"Envelopes Used": "Total Envelopes Used"}, inplace=True)

print("Data loaded successfully.")

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

total_purchased = int(monthly["Envelopes Purchased"].sum())
total_used = int(monthly["Envelopes Used (Volume)"].sum())
total_mailed = int(monthly["Envelopes Mailed (Postage)"].sum())
total_spoils = int(monthly["Spoils"].sum())
net_variance = total_purchased - total_used

usage_by_product_sorted = usage_by_product.sort_values("Total Envelopes Used", ascending=False)
usage_product_total = usage_by_product["Total Envelopes Used"].sum()

# Post-settlement analysis (Mar 2022 onward — when Apex started paying for envelopes)
SETTLEMENT_DATE = "Mar-22"
month_order = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]

def month_label_to_sortkey(label):
    parts = label.split('-')
    mi = month_order.index(parts[0])
    yi = int(parts[1])
    return (yi, mi)

settlement_key = month_label_to_sortkey(SETTLEMENT_DATE)
post_mask = monthly["Month"].apply(lambda x: month_label_to_sortkey(x) >= settlement_key)
post = monthly[post_mask]

post_purchased = int(post["Envelopes Purchased"].sum())
post_used = int(post["Envelopes Used (Volume)"].sum())
post_mailed = int(post["Envelopes Mailed (Postage)"].sum())
post_spoils = int(post["Spoils"].sum())
post_variance = post_purchased - post_used
post_months = len(post)
spoilage_rate = post_spoils / post_used * 100 if post_used else 0

# Year-by-year post-settlement
from collections import defaultdict as _defaultdict
post_yearly = _defaultdict(lambda: [0, 0, 0, 0, 0])  # purchased, used, mailed, spoils, month_count
for _, r in post.iterrows():
    yr = 2000 + int(r["Month"].split('-')[1])
    post_yearly[yr][0] += safe(r["Envelopes Purchased"])
    post_yearly[yr][1] += safe(r["Envelopes Used (Volume)"])
    post_yearly[yr][2] += safe(r["Envelopes Mailed (Postage)"])
    post_yearly[yr][3] += safe(r["Spoils"])
    post_yearly[yr][4] += 1

# Rolling avg monthly usage (last 6 months) for buffer stock calculation
recent_6 = monthly.tail(6)
avg_monthly_usage = int(recent_6["Envelopes Used (Volume)"].mean())

by_type_sorted = by_type.sort_values("Total Purchased", ascending=False)
env_type_total = by_type["Total Purchased"].sum()

# Build combined envelope group table — groups related purchase SKUs and usage types
# by physical envelope shape/size (not postage imprint) for meaningful comparison.
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

# Build lookup dicts
purchase_by_type = dict(zip(by_type["Envelope Type"], by_type["Total Purchased"]))
usage_by_type_dict = dict(zip(usage_by_env_type["Envelope Type"], usage_by_env_type["Total Envelopes Used"]))

# Check for missing months (shown as alert banner if any gaps exist)
def find_missing_months():
    expected = []
    for y in range(2020, 2026):
        for m in range(1, 13):
            label = f'{["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"][m-1]}-{y % 100:02d}'
            expected.append(label)
    present = set(monthly["Month"].tolist())
    return [label for label in expected if label not in present]

missing_months = find_missing_months()

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

def build_monthly_rows():
    """Build monthly rows (post-settlement scope) with annual subtotals and grand total."""
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

def build_env_type_rows():
    rows = []
    for _, r in by_type_sorted.iterrows():
        tp = safe(r["Total Purchased"])
        pct = (tp / env_type_total * 100) if env_type_total else 0
        bw = max(pct, 0.5)
        rows.append(
            '<tr>'
            + f'<td class="env-name">{r["Envelope Type"]}</td>'
            + f'<td class="num">{fmt_num(tp)}</td>'
            + f'<td><div class="bar-container"><div class="bar-fill" style="width:{bw:.1f}%"></div><span class="bar-label">{pct:.1f}%</span></div></td>'
            + '</tr>'
        )
    return "\n".join(rows)

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
    # Grand total
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

def build_usage_by_product_rows():
    rows = []
    for _, r in usage_by_product_sorted.iterrows():
        used = safe(r["Total Envelopes Used"])
        pct = (used / usage_product_total * 100) if usage_product_total else 0
        bw = max(pct, 0.5)
        rows.append(
            '<tr>'
            + f'<td class="env-name">{r["Product Name"]}</td>'
            + f'<td class="num">{fmt_num(used)}</td>'
            + f'<td><div class="bar-container"><div class="bar-fill" style="width:{bw:.1f}%"></div><span class="bar-label">{pct:.1f}%</span></div></td>'
            + '</tr>'
        )
    return "\n".join(rows)

def build_sku_recon_rows():
    """SKU-level recon: purchased, used, variance, variance %."""
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

def build_inventory_gauge():
    """SVG gauge comparing actual inventory vs Broadridge 2-3 month policy."""
    policy_min = avg_monthly_usage * 2
    policy_max = avg_monthly_usage * 3
    actual = post_variance
    max_val = max(actual, policy_max) * 1.3
    excess = max(0, actual - policy_max)

    w, h = 700, 94
    ml, mr = 50, 50
    bar_w = w - ml - mr
    bar_y, bar_h = 36, 22

    def xp(val):
        return ml + (val / max_val) * bar_w

    s = []
    s.append(f'<svg viewBox="0 0 {w} {h}" style="width:100%;max-width:{w}px;height:auto;" xmlns="http://www.w3.org/2000/svg">')
    # Track
    s.append(f'<rect x="{ml}" y="{bar_y}" width="{bar_w}" height="{bar_h}" fill="#F0F0F0" rx="4"/>')
    # Under-target zone (0 to policy_min) — amber
    s.append(f'<rect x="{ml}" y="{bar_y}" width="{xp(policy_min)-ml:.1f}" height="{bar_h}" fill="#FFF3E0" rx="4"/>')
    # Policy zone (min to max) — green
    s.append(f'<rect x="{xp(policy_min):.1f}" y="{bar_y}" width="{xp(policy_max)-xp(policy_min):.1f}" height="{bar_h}" fill="#C8E6C9"/>')
    # Over-target zone (max to end) — light red
    s.append(f'<rect x="{xp(policy_max):.1f}" y="{bar_y}" width="{xp(max_val)-xp(policy_max):.1f}" height="{bar_h}" fill="#FFCDD2" rx="4"/>')
    # Policy range borders
    s.append(f'<line x1="{xp(policy_min):.1f}" y1="{bar_y}" x2="{xp(policy_min):.1f}" y2="{bar_y+bar_h}" stroke="#186741" stroke-width="1.5"/>')
    s.append(f'<line x1="{xp(policy_max):.1f}" y1="{bar_y}" x2="{xp(policy_max):.1f}" y2="{bar_y+bar_h}" stroke="#186741" stroke-width="1.5"/>')
    # Actual marker
    ax = xp(actual)
    s.append(f'<line x1="{ax:.1f}" y1="{bar_y-6}" x2="{ax:.1f}" y2="{bar_y+bar_h+6}" stroke="#052390" stroke-width="3"/>')
    s.append(f'<circle cx="{ax:.1f}" cy="{bar_y+bar_h/2:.1f}" r="8" fill="#052390"/>')
    s.append(f'<circle cx="{ax:.1f}" cy="{bar_y+bar_h/2:.1f}" r="4" fill="#FFFFFF"/>')
    # Labels — above bar
    s.append(f'<text x="{ax:.1f}" y="{bar_y-12}" text-anchor="middle" font-size="13" fill="#052390" font-weight="700" font-family="Helvetica Neue,Arial,sans-serif">Actual: {fmt_num(actual)}</text>')
    # Labels — below bar
    s.append(f'<text x="{xp(policy_min):.1f}" y="{bar_y+bar_h+16}" text-anchor="middle" font-size="11" fill="#186741" font-family="Helvetica Neue,Arial,sans-serif">{fmt_num(policy_min)}</text>')
    s.append(f'<text x="{xp(policy_min):.1f}" y="{bar_y+bar_h+28}" text-anchor="middle" font-size="10" fill="#186741" font-family="Helvetica Neue,Arial,sans-serif">2-mo min</text>')
    s.append(f'<text x="{xp(policy_max):.1f}" y="{bar_y+bar_h+16}" text-anchor="middle" font-size="11" fill="#186741" font-family="Helvetica Neue,Arial,sans-serif">{fmt_num(policy_max)}</text>')
    s.append(f'<text x="{xp(policy_max):.1f}" y="{bar_y+bar_h+28}" text-anchor="middle" font-size="10" fill="#186741" font-family="Helvetica Neue,Arial,sans-serif">3-mo max</text>')
    s.append('</svg>')
    return '\n'.join(s)

def build_monthly_usage_trend():
    """SVG sparkline showing rolling 6-month avg usage over time."""
    n = len(post)
    if n < 6:
        return ""
    # Compute 6-month rolling average
    usage_vals = post["Envelopes Used (Volume)"].tolist()
    labels = post["Month"].tolist()
    rolling = []
    for i in range(5, n):
        avg = sum(usage_vals[i-5:i+1]) / 6
        rolling.append((labels[i], avg))

    cw, ch = 700, 160
    ml, mr, mt, mb = 60, 30, 30, 50
    pw = cw - ml - mr
    ph = ch - mt - mb
    nr = len(rolling)
    max_val = max(v for _, v in rolling) * 1.15
    min_val = min(v for _, v in rolling) * 0.85

    L = [f'<svg viewBox="0 0 {cw} {ch}" style="width:100%;max-width:{cw}px;height:auto;" xmlns="http://www.w3.org/2000/svg">']
    L.append(f'<rect width="{cw}" height="{ch}" fill="#FFFFFF" rx="8"/>')

    # Y-axis gridlines
    for i in range(5):
        yv = min_val + (max_val - min_val) * i / 4
        yp = mt + ph - (ph * i / 4)
        L.append(f'<line x1="{ml}" y1="{yp:.1f}" x2="{cw-mr}" y2="{yp:.1f}" stroke="#F0F0F0" stroke-width="1"/>')
        lb = f"{yv/1000:.0f}K"
        L.append(f'<text x="{ml-6}" y="{yp+4:.1f}" text-anchor="end" fill="#6D6E71" font-size="10" font-family="Helvetica Neue,Arial,sans-serif">{lb}</text>')

    # Line path
    pts = []
    for i, (lbl, val) in enumerate(rolling):
        x = ml + (i / (nr - 1)) * pw
        y = mt + ph - ((val - min_val) / (max_val - min_val)) * ph
        pts.append(f"{x:.1f},{y:.1f}")
    L.append(f'<polyline points="{" ".join(pts)}" fill="none" stroke="#3F8EFC" stroke-width="2.5" stroke-linejoin="round"/>')

    # Start/end dots + labels
    sx, sy = pts[0].split(',')
    ex, ey = pts[-1].split(',')
    L.append(f'<circle cx="{sx}" cy="{sy}" r="4" fill="#3F8EFC"/>')
    L.append(f'<circle cx="{ex}" cy="{ey}" r="5" fill="#052390"/>')
    L.append(f'<text x="{sx}" y="{float(sy)-10}" text-anchor="start" font-size="11" fill="#3F8EFC" font-weight="600" font-family="Helvetica Neue,Arial,sans-serif">{rolling[0][1]/1000:.0f}K</text>')
    L.append(f'<text x="{ex}" y="{float(ey)-10}" text-anchor="end" font-size="11" fill="#052390" font-weight="600" font-family="Helvetica Neue,Arial,sans-serif">{rolling[-1][1]/1000:.0f}K</text>')

    # X-axis labels (every 12th)
    for i, (lbl, val) in enumerate(rolling):
        if i % 12 == 0 or i == nr - 1:
            x = ml + (i / (nr - 1)) * pw
            ly = mt + ph + 16
            L.append(f'<text x="{x:.1f}" y="{ly}" text-anchor="middle" fill="#6D6E71" font-size="10" font-family="Helvetica Neue,Arial,sans-serif">{lbl}</text>')

    L.append('</svg>')
    return '\n'.join(L)

svg_chart = build_svg_chart()
monthly_rows = build_monthly_rows()
env_type_rows = build_env_type_rows()
combined_env_rows = build_combined_env_rows()
usage_product_rows = build_usage_by_product_rows()
sku_recon_rows = build_sku_recon_rows()
inventory_gauge = build_inventory_gauge()
usage_trend_svg = build_monthly_usage_trend()

kpi_var_color = "#9D1526" if net_variance < 0 else "#186741"

# --- Post-settlement derived values for Executive Summary ---
post_var_color = "#186741" if post_variance >= 0 else "#9D1526"
buffer_months = post_variance / avg_monthly_usage if avg_monthly_usage else 0

usage_2022 = post_yearly[2022][1] / post_yearly[2022][4] if post_yearly[2022][4] else 0
usage_2025 = post_yearly[2025][1] / post_yearly[2025][4] if post_yearly[2025][4] else 0
usage_decline = (1 - usage_2025 / usage_2022) * 100 if usage_2022 else 0

# Recent 2 years variance for trajectory assessment
recent_yrs = sorted(post_yearly.keys())[-2:]
recent_p = sum(post_yearly[y][0] for y in recent_yrs)
recent_u = sum(post_yearly[y][1] for y in recent_yrs)
recent_var_pct = (recent_p - recent_u) / recent_p * 100 if recent_p else 0

# ---------------------------------------------------------------------------
# CSS
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
.bottom-line {
    background: #FFFFFF; border-radius: 12px; padding: 24px 28px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.06), 0 2px 12px rgba(0,0,0,0.04);
    border-left: 5px solid #052390; margin-bottom: 24px;
}
.bottom-line .bl-heading {
    font-size: 13px; font-weight: 700; text-transform: uppercase;
    letter-spacing: 0.5px; color: #052390; margin: 0 0 8px;
}
.bottom-line p { margin: 0; font-size: 14px; line-height: 1.7; color: #333; }
.bottom-line strong { color: #052390; }
.alert-box {
    background: rgba(252, 94, 23, 0.06); border-left: 4px solid #FC5E17;
    border-radius: 0 8px 8px 0; padding: 16px 24px; margin-bottom: 20px;
}
.alert-box p { margin: 0; font-size: 13px; color: #333; }
.alert-box strong { color: #FC5E17; }
.info-box {
    background: rgba(41, 84, 240, 0.05); border-left: 4px solid #2954F0;
    border-radius: 0 8px 8px 0; padding: 16px 24px; margin-top: 20px;
}
.info-box p { margin: 0; font-size: 13px; color: #333; line-height: 1.6; }
.info-box strong { color: #052390; }
.context-line {
    font-size: 13px; color: #6D6E71; margin: 20px 0 0;
    padding: 12px 16px; background: #F5F5F7; border-radius: 8px; line-height: 1.6;
}
.context-line strong { color: #333; }
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
            // Expand the section if collapsed
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
# Envelope Specifications (data, used in Reference section)
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

envelope_spec_rows = build_envelope_spec_rows()

# ---------------------------------------------------------------------------
# Build HTML
# ---------------------------------------------------------------------------
html = '<!DOCTYPE html>\n<html lang="en">\n<head>\n'
html += '<meta charset="UTF-8">\n'
html += '<meta name="viewport" content="width=device-width, initial-scale=1.0">\n'
html += '<title>Broadridge Envelope Reconciliation Report</title>\n'
html += f'<style>{CSS}</style>\n'
html += '</head>\n<body>\n\n'

# --- Header ---
html += '<div class="header">\n'
html += '    <h1>Broadridge envelope reconciliation</h1>\n'
html += '    <p class="subtitle">Purchase vs. usage analysis &nbsp;|&nbsp; March 2022 &ndash; December 2025</p>\n'
html += f'    <p class="generated">Generated {pd.Timestamp.now().strftime("%B %d, %Y")}</p>\n'
html += '</div>\n'

# --- Nav ---
html += '<nav class="nav">\n'
html += '    <a href="#executive-summary">Summary</a>\n'
html += '    <a href="#monthly-trend">Trend</a>\n'
html += '    <a href="#envelope-types">By Type</a>\n'
html += '    <a href="#monthly-detail">Monthly Detail</a>\n'
html += '    <a href="#reference">Reference</a>\n'
html += '</nav>\n'
html += '<div class="content">\n'

# Missing data alert (shown only if gaps exist)
if missing_months:
    html += '    <div class="alert-box">\n'
    html += f'        <p><strong>&#9888; Missing data:</strong> {", ".join(missing_months)} &mdash; '
    html += 'missing source files affect variance calculations.</p>\n'
    html += '    </div>\n\n'

# ===== EXECUTIVE SUMMARY =====
html += '<div class="section" id="executive-summary">\n'
html += '    <div class="section-header" onclick="toggleSection(this)">\n'
html += '        <h2>Executive summary</h2>\n'
html += '        <span class="toggle">&#9660;</span>\n'
html += '    </div>\n'
html += '    <div class="section-body">\n'

# -- Bottom line --
if buffer_months > 3:
    bl_assessment = (
        f'Broadridge is holding <strong>{buffer_months:.1f} months</strong> of envelope buffer stock '
        f'&mdash; more than double their stated 2&ndash;3 month policy. '
        f'While recent purchasing has corrected to within {abs(recent_var_pct):.0f}% of usage (2024&ndash;2025), '
        f'the accumulated surplus from 2022 over-purchasing remains. '
        f'Consider requesting Broadridge reduce purchase orders until inventory aligns with their 2&ndash;3 month buffer target.'
    )
elif buffer_months < 2:
    bl_assessment = (
        f'Buffer stock is at <strong>{buffer_months:.1f} months</strong>, below Broadridge&rsquo;s 2&ndash;3 month policy. '
        f'Monitor upcoming purchase orders to ensure adequate supply.'
    )
else:
    bl_assessment = (
        f'Buffer stock is at <strong>{buffer_months:.1f} months</strong>, within Broadridge&rsquo;s 2&ndash;3 month target. '
        f'Purchasing is well-calibrated to usage.'
    )

html += '        <div class="bottom-line">\n'
html += '            <p class="bl-heading">Bottom line</p>\n'
html += f'            <p>{bl_assessment}</p>\n'
html += '        </div>\n'

# -- Post-settlement KPIs (what Apex is paying for — the numbers that matter) --
html += '        <div class="kpi-grid">\n'
html += f'            <div class="kpi-card"><p class="kpi-label">Purchased</p><p class="kpi-value" style="color:#2954F0">{fmt_num(post_purchased)}</p><p class="kpi-sub">Mar 2022 &ndash; Dec 2025 ({post_months} mo)</p></div>\n'
html += f'            <div class="kpi-card"><p class="kpi-label">Used</p><p class="kpi-value" style="color:#2954F0">{fmt_num(post_used)}</p><p class="kpi-sub">Spoilage: {spoilage_rate:.1f}% ({fmt_num(post_spoils)} of {fmt_num(post_used)})</p></div>\n'
html += f'            <div class="kpi-card"><p class="kpi-label">Implied Inventory</p><p class="kpi-value" style="color:{post_var_color}">{fmt_num(post_variance)}</p><p class="kpi-sub">{buffer_months:.1f} months at current usage</p></div>\n'
html += f'            <div class="kpi-card"><p class="kpi-label">Avg Monthly Usage</p><p class="kpi-value" style="color:#2954F0">{fmt_num(avg_monthly_usage)}</p><p class="kpi-sub">Trailing 6 months</p></div>\n'
html += '        </div>\n'

# -- Inventory gauge: actual vs Broadridge 2-3 month policy --
html += '        <div style="background:#FFFFFF;border-radius:12px;padding:20px 24px;box-shadow:0 1px 4px rgba(0,0,0,0.06);margin-top:12px;">\n'
html += '            <p style="font-size:13px;font-weight:600;color:#052390;margin:0 0 8px;text-transform:uppercase;letter-spacing:0.3px;">Buffer stock vs. Broadridge 2&ndash;3 month policy</p>\n'
html += inventory_gauge + '\n'
html += f'            <p style="font-size:12px;color:#6D6E71;margin:8px 0 0;">Based on trailing 6-month average usage of {fmt_num(avg_monthly_usage)}/month. Policy range: {fmt_num(avg_monthly_usage*2)} (2-mo) to {fmt_num(avg_monthly_usage*3)} (3-mo).</p>\n'
html += '        </div>\n'

# -- Key findings (promoted to top) --
html += '        <div style="margin-top:16px;font-size:14px;color:#333;line-height:1.8;">\n'
html += '            <ul style="margin:0;padding-left:20px;">\n'
html += f'            <li>Implied inventory of <strong>{fmt_num(post_variance)}</strong> envelopes = <strong>{buffer_months:.1f} months</strong> of buffer stock at current usage rates (Broadridge policy: 2&ndash;3 months).</li>\n'
html += f'            <li>Average monthly usage declined <strong>{usage_decline:.0f}%</strong> from {fmt_num(usage_2022)}/mo (2022) to {fmt_num(usage_2025)}/mo (2025), consistent with the 30% print reduction target in the 2022 renewal term sheet.</li>\n'
html += f'            <li>Purchasing trajectory has corrected: 2022 over-purchased by 20.3% building initial buffer; 2024&ndash;2025 are within 3% of usage.</li>\n'
html += f'            <li>Reported spoilage: <strong>{fmt_num(post_spoils)}</strong> envelopes ({spoilage_rate:.1f}% of usage) &mdash; well within the 10% contractual wastage limit. Note: actual wastage (10&ndash;15% per Broadridge) is embedded in the &ldquo;Used&rdquo; figure, not separately reported.</li>\n'
html += '            </ul>\n'
html += '        </div>\n'

# -- Year-by-year table --
html += '        <div class="table-wrap" style="margin-top:20px;"><table>\n'
html += '            <thead><tr><th>Year</th><th>Purchased</th><th>Used</th><th>Variance</th><th>Var %</th><th>Avg Mo Used</th></tr></thead>\n'
html += '            <tbody>\n'
for yr in sorted(post_yearly.keys()):
    d = post_yearly[yr]
    yp, yu, ym, ys, mc = d
    yv = yp - yu
    vc = var_color(yv)
    vpct = yv / yp if yp else 0
    avg_u = yu / mc if mc else 0
    yr_label = f"{yr} (Mar&ndash;Dec)" if yr == 2022 else str(yr)
    html += f'            <tr><td>{yr_label}</td>'
    html += f'<td class="num">{fmt_num(yp)}</td>'
    html += f'<td class="num">{fmt_num(yu)}</td>'
    html += f'<td class="num" style="color:{vc};font-weight:600">{fmt_num_parens(yv)}</td>'
    html += f'<td class="num" style="color:{vc}">{fmt_pct(vpct)}</td>'
    html += f'<td class="num">{fmt_num(avg_u)}</td></tr>\n'
# Total row
html += f'            <tr class="total-row"><td><strong>Total</strong></td>'
html += f'<td class="num"><strong>{fmt_num(post_purchased)}</strong></td>'
html += f'<td class="num"><strong>{fmt_num(post_used)}</strong></td>'
html += f'<td class="num" style="color:{post_var_color};font-weight:700"><strong>{fmt_num_parens(post_variance)}</strong></td>'
html += f'<td class="num" style="color:{post_var_color};font-weight:700"><strong>{fmt_pct(post_variance/post_purchased if post_purchased else 0)}</strong></td>'
html += f'<td class="num"><strong>{fmt_num(post_used/post_months if post_months else 0)}</strong></td></tr>\n'
html += '            </tbody>\n'
html += '        </table></div>\n'

# -- Full-period context (secondary, not KPI cards) --
full_var_pct = net_variance / total_purchased * 100 if total_purchased else 0
pre_purchased = total_purchased - post_purchased
pre_used = total_used - post_used
html += f'        <div class="context-line">\n'
html += f'            <strong>Pre-settlement (Jan 2020 &ndash; Feb 2022):</strong> {fmt_num(pre_purchased)} purchased, {fmt_num(pre_used)} used. '
html += f'These costs were absorbed by Broadridge per the $643,458 pass-through paper dispute settlement (June 2022 renewal term sheet). '
html += f'<strong>Full period (Jan 2020 &ndash; Dec 2025):</strong> {fmt_num(total_purchased)} purchased, {fmt_num(total_used)} used, '
html += f'{"+" if net_variance >= 0 else ""}{fmt_num_parens(net_variance)} ({full_var_pct:+.1f}%).\n'
html += '        </div>\n'

html += '    </div>\n'
html += '</div>\n\n'

# ===== MONTHLY TREND =====
html += '<div class="section" id="monthly-trend">\n'
html += '    <div class="section-header" onclick="toggleSection(this)">\n'
html += '        <h2>Monthly trend</h2>\n'
html += '        <span class="toggle">&#9660;</span>\n'
html += '    </div>\n'
html += '    <div class="section-body">\n'
html += '        <div style="background:#FFFFFF;border-radius:12px;padding:24px;box-shadow:0 1px 4px rgba(0,0,0,0.06);">\n'
html += svg_chart
html += '\n        </div>\n'

# -- Rolling 6-month average usage trend --
html += '        <div style="background:#FFFFFF;border-radius:12px;padding:20px 24px;box-shadow:0 1px 4px rgba(0,0,0,0.06);margin-top:20px;">\n'
html += '            <p style="font-size:13px;font-weight:600;color:#052390;margin:0 0 8px;text-transform:uppercase;letter-spacing:0.3px;">Rolling 6-month average usage</p>\n'
html += usage_trend_svg + '\n'
html += '        </div>\n'

html += '    </div>\n</div>\n\n'

# ===== PURCHASES & USAGE BY ENVELOPE TYPE =====
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
html += '        <div class="info-box" style="margin-top:16px;"><p><strong>Scope note:</strong> Envelope type breakdown covers the full contract period (Jan 2020 &ndash; Dec 2025) because type-level monthly data is not available in the source. Post-settlement purchases (Mar 2022+) account for ~70% of the totals shown.</p></div>\n'

# -- SKU-level recon (purchased + used + variance + variance %) --
html += '        <h3 style="color:#052390;font-size:16px;margin:32px 0 12px;">By SKU</h3>\n'
html += '        <p style="font-size:12px;color:#6D6E71;margin:0 0 12px;">ENVCONPFSN10NI was replaced by ENVAPXN10&hellip;IND(10/22) in Oct 2022 (postal permit update). Usage shifted but both SKUs remain in purchase history. See the grouped table above for the combined view.</p>\n'
html += '        <div class="table-wrap"><table class="sortable">\n'
html += '            <thead><tr>' + th_row(["SKU","Purchased","Used","Variance","Variance %"]) + '</tr></thead>\n'
html += '            <tbody>\n' + sku_recon_rows + '\n            </tbody>\n'
html += '        </table></div>\n'

# -- Usage by product --
html += '        <h3 style="color:#052390;font-size:16px;margin:32px 0 12px;">Usage by product</h3>\n'
html += '        <p style="font-size:12px;color:#6D6E71;margin:0 0 12px;">Products map to envelopes: <strong>Monthly Statements / Efail Statements</strong> &rarr; N14 Fold or 9x12 Flat (domestic/foreign). <strong>Address Verification Letters / Apex MTC / Apex Checks / Disbursement Letters</strong> &rarr; #10 Confirms+Letters. <strong>Daily Confirms</strong> &rarr; #10 or 9x12 Flat Confirms. <strong>1099 / 1042 / tax forms</strong> &rarr; Tax Form Envelopes.</p>\n'
html += '        <div class="table-wrap"><table class="sortable">\n'
html += '            <thead><tr>' + th_row(["Product Name","Total Used"]) + '<th>% of Total</th></tr></thead>\n'
html += '            <tbody>\n' + usage_product_rows + '\n            </tbody>\n'
html += '        </table></div>\n'

html += '    </div>\n</div>\n\n'

# ===== MONTHLY DETAIL (collapsed by default) =====
html += '<div class="section" id="monthly-detail">\n'
html += '    <div class="section-header" onclick="toggleSection(this)">\n'
html += '        <h2>Monthly detail</h2>\n'
html += '        <span class="toggle">&#9654;</span>\n'
html += '    </div>\n'
html += '    <div class="section-body collapsed">\n'
html += '        <div class="table-wrap"><table class="sortable">\n'
html += '            <thead><tr>' + th_row(["Month","Purchased","Used","Net Variance","Variance %","Running Balance"]) + '</tr></thead>\n'
html += '            <tbody>\n' + monthly_rows + '\n            </tbody>\n'
html += '        </table></div>\n'
html += '        <p style="font-size:12px;color:#6D6E71;margin:12px 0 0;"><strong>May &amp; Jun 2025:</strong> Zero purchases confirmed (not missing data).</p>\n'
html += '    </div>\n</div>\n\n'

# ===== REFERENCE (collapsed by default) =====
html += '<div class="section" id="reference">\n'
html += '    <div class="section-header" onclick="toggleSection(this)">\n'
html += '        <h2>Reference</h2>\n'
html += '        <span class="toggle">&#9654;</span>\n'
html += '    </div>\n'
html += '    <div class="section-body collapsed">\n'

# Envelope Specifications
html += '        <h3 style="color:#052390;font-size:16px;margin:0 0 12px;">Envelope specifications</h3>\n'
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
html += '<div class="footer">\n    Confidential &mdash; Apex Clearing Corporation\n</div>\n'

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
