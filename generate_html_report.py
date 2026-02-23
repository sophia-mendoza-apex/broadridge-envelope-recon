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
annual = pd.read_excel(EXCEL_PATH, sheet_name="Annual Summary")
by_type = pd.read_excel(EXCEL_PATH, sheet_name="By Envelope Type")
detail = pd.read_excel(EXCEL_PATH, sheet_name="Purchase Detail")
audit = pd.read_excel(EXCEL_PATH, sheet_name="Contract Audit")
usage_by_product = pd.read_excel(EXCEL_PATH, sheet_name="Usage by Product")

print("Data loaded successfully.")

DASH = "\u2014"

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
        if val == 0:
            return DASH
        return f"${val:,.2f}"
    except (ValueError, TypeError):
        return str(v)

def fmt_money_always(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return DASH
    try:
        return f"${float(v):,.2f}"
    except (ValueError, TypeError):
        return str(v)

def fmt_money_parens(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return DASH
    try:
        val = float(v)
        if val < 0:
            return f"(${abs(val):,.2f})"
        return f"${val:,.2f}"
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
        if abs(val) < 1:
            return f"{val * 100:.1f}%"
        return f"{val:.1f}%"
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
total_cost = monthly["Purchase Cost"].sum()
total_invoiced = monthly["Invoiced Amount"].sum()

audit_over = int((audit["Flag"] == "OVER").sum())
audit_under = int((audit["Flag"] == "UNDER").sum())
audit_ok = int((audit["Flag"] == "OK").sum())
audit_net_diff = audit["Difference"].sum()

usage_by_product_sorted = usage_by_product.sort_values("Total Envelopes Used", ascending=False)
usage_product_total = usage_by_product["Total Envelopes Used"].sum()

by_type_sorted = by_type.sort_values("Total Purchased", ascending=False)
env_type_total = by_type["Total Purchased"].sum()

audit_sorted = audit.copy()
audit_sorted["AbsDiff"] = audit_sorted["Difference"].abs()
audit_sorted = audit_sorted.sort_values("AbsDiff", ascending=False)
top10_audit = audit_sorted.head(10)

## zero_purchase / zero_usage removed — 0 values mean no activity, not missing data

def build_svg_chart():
    months = monthly["Month"].tolist()
    purchased = monthly["Envelopes Purchased"].tolist()
    used = monthly["Envelopes Used (Volume)"].tolist()
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

def build_annual_rows():
    rows = []
    running_balance = 0
    for _, r in annual.iterrows():
        vv = safe(r["Net Variance"])
        vc = var_color(vv)
        running_balance += vv
        rbc = var_color(running_balance)
        rows.append(
            '<tr>'
            + f'<td>{int(r["Year"])}</td>'
            + f'<td class="num">{fmt_num(r["Envelopes Purchased"])}</td>'
            + f'<td class="num">{fmt_num(r["Envelopes Used (Volume)"])}</td>'
            + f'<td class="num">{fmt_num(r["Envelopes Mailed (Postage)"])}</td>'
            + f'<td class="num">{fmt_num(r["Spoils"])}</td>'
            + f'<td class="num" style="color:{vc};font-weight:600">{fmt_num_parens(vv)}</td>'
            + f'<td class="num" style="color:{rbc};font-weight:600">{fmt_num_parens(running_balance)}</td>'
            + f'<td class="num">{fmt_money(r["Total Cost"])}</td>'
            + f'<td class="num">{fmt_money(r["Total Invoiced"])}</td>'
            + '</tr>'
        )
    gp = int(annual["Envelopes Purchased"].sum())
    gu = int(annual["Envelopes Used (Volume)"].sum())
    gm = int(annual["Envelopes Mailed (Postage)"].sum())
    gs = int(annual["Spoils"].sum())
    gv = gp - gu
    gc = annual["Total Cost"].sum()
    gi = annual["Total Invoiced"].sum()
    vc = var_color(gv)
    rows.append(
        '<tr class="total-row">'
        + f'<td><strong>Grand Total</strong></td>'
        + f'<td class="num"><strong>{fmt_num(gp)}</strong></td>'
        + f'<td class="num"><strong>{fmt_num(gu)}</strong></td>'
        + f'<td class="num"><strong>{fmt_num(gm)}</strong></td>'
        + f'<td class="num"><strong>{fmt_num(gs)}</strong></td>'
        + f'<td class="num" style="color:{vc};font-weight:700"><strong>{fmt_num_parens(gv)}</strong></td>'
        + f'<td class="num" style="color:{vc};font-weight:700"><strong>{fmt_num_parens(running_balance)}</strong></td>'
        + f'<td class="num"><strong>{fmt_money(gc)}</strong></td>'
        + f'<td class="num"><strong>{fmt_money(gi)}</strong></td>'
        + '</tr>'
    )
    return "\n".join(rows)

def build_monthly_rows():
    rows = []
    running_balance = 0
    for i, (_, r) in enumerate(monthly.iterrows()):
        vv = safe(r["Net Variance (Purchased - Used)"])
        vc = var_color(vv)
        running_balance += vv
        rbc = var_color(running_balance)
        bg = "#F5F5F7" if i % 2 == 1 else "#FFFFFF"
        rows.append(
            f'<tr style="background:{bg}">'
            + f'<td>{r["Month"]}</td>'
            + f'<td class="num">{fmt_num(r["Envelopes Purchased"])}</td>'
            + f'<td class="num">{fmt_num(r["Envelopes Used (Volume)"])}</td>'
            + f'<td class="num">{fmt_num(r["Envelopes Mailed (Postage)"])}</td>'
            + f'<td class="num">{fmt_num(r["Spoils"])}</td>'
            + f'<td class="num" style="color:{vc};font-weight:600">{fmt_num_parens(vv)}</td>'
            + f'<td class="num" style="color:{vc}">{fmt_pct(r["Variance %"])}</td>'
            + f'<td class="num" style="color:{rbc};font-weight:600">{fmt_num_parens(running_balance)}</td>'
            + f'<td class="num">{fmt_money(r["Purchase Cost"])}</td>'
            + f'<td class="num">{fmt_money(r["Invoiced Amount"])}</td>'
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
            + f'<td class="num">{fmt_money(r["Total Cost"])}</td>'
            + f'<td class="num">{fmt_money_always(r["Avg Unit Price"])}</td>'
            + f'<td><div class="bar-container"><div class="bar-fill" style="width:{bw:.1f}%"></div><span class="bar-label">{pct:.1f}%</span></div></td>'
            + '</tr>'
        )
    return "\n".join(rows)

def build_top10_audit_rows():
    rows = []
    for _, r in top10_audit.iterrows():
        diff = safe(r["Difference"])
        vc = "#9D1526" if diff > 0 else "#186741"
        flag = r["Flag"]
        fc = "flag-over" if flag == "OVER" else ("flag-under" if flag == "UNDER" else "flag-ok")
        rows.append(
            '<tr>'
            + f'<td>{r["Month"]}</td>'
            + f'<td class="env-name">{r["Description"]}</td>'
            + f'<td class="num">{fmt_num(r["Qty Ordered"])}</td>'
            + f'<td class="num">{fmt_money_always(r["Unit Price"])}</td>'
            + f'<td class="num">{fmt_money_always(r["Expected Invoiced"])}</td>'
            + f'<td class="num">{fmt_money_always(r["Actual Invoiced"])}</td>'
            + f'<td class="num" style="color:{vc};font-weight:600">{fmt_money_parens(diff)}</td>'
            + f'<td><span class="{fc}">{flag}</span></td>'
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
            + f'<td>{r["First Month"]}</td>'
            + f'<td>{r["Last Month"]}</td>'
            + f'<td><div class="bar-container"><div class="bar-fill" style="width:{bw:.1f}%"></div><span class="bar-label">{pct:.1f}%</span></div></td>'
            + '</tr>'
        )
    return "\n".join(rows)

def build_dq_rows():
    rows = []
    # Flag months entirely missing from the data (no row at all)
    # Build expected month range Mar-22 to Dec-25
    expected = []
    for y in range(2020, 2026):
        start_m = 1
        for m in range(start_m, 13):
            label = f'{["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"][m-1]}-{y % 100:02d}'
            expected.append(label)
    present = set(monthly["Month"].tolist())
    for label in expected:
        if label not in present:
            rows.append(
                '<tr>'
                + f'<td>{label}</td>'
                + '<td>No data &mdash; missing both purchase report and billing workbook</td>'
                + '<td class="num">0</td>'
                + '<td class="num">0</td>'
                + '<td><span class="flag-over">Request from Broadridge</span></td>'
                + '</tr>'
            )
    return "\n".join(rows)

svg_chart = build_svg_chart()
annual_rows = build_annual_rows()
monthly_rows = build_monthly_rows()
env_type_rows = build_env_type_rows()
usage_product_rows = build_usage_by_product_rows()
top10_rows = build_top10_audit_rows()
dq_rows = build_dq_rows()

kpi_var_color = "#9D1526" if net_variance < 0 else "#186741"
kpi_over_color = "#9D1526" if audit_net_diff > 0 else "#186741"

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
.alert-box {
    background: rgba(252, 94, 23, 0.06); border-left: 4px solid #FC5E17;
    border-radius: 0 8px 8px 0; padding: 20px 24px; margin-bottom: 10px;
}
.alert-box h3 { margin: 0 0 10px; font-size: 15px; font-weight: 600; color: #FC5E17; }
.alert-box p { margin: 0 0 10px; font-size: 13px; color: #333; }
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
.total-row { background: #F5F5F7 !important; border-top: 2px solid #E2E2E2; }
.zero-inv { color: #9D1526; opacity: 0.6; }
.bar-container { display: flex; align-items: center; gap: 8px; min-width: 160px; }
.bar-fill {
    height: 18px; background: linear-gradient(90deg, #2954F0, #3F8EFC);
    border-radius: 9px; min-width: 3px;
}
.bar-label { font-size: 12px; font-weight: 500; color: #6D6E71; white-space: nowrap; }
.flag-over {
    background: #FDEAEC; color: #9D1526; padding: 3px 10px; border-radius: 12px;
    font-size: 11px; font-weight: 600; text-transform: uppercase;
}
.flag-under {
    background: #E8F5E9; color: #186741; padding: 3px 10px; border-radius: 12px;
    font-size: 11px; font-weight: 600; text-transform: uppercase;
}
.flag-ok {
    background: #F5F5F7; color: #6D6E71; padding: 3px 10px; border-radius: 12px;
    font-size: 11px; font-weight: 600; text-transform: uppercase;
}
.terms-grid { display: flex; gap: 24px; flex-wrap: wrap; margin-bottom: 24px; }
.term-card {
    flex: 1 1 280px; background: #FFFFFF; border-radius: 10px; padding: 20px 24px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.06); border-top: 3px solid #2954F0;
}
.term-card h4 { margin: 0 0 12px; font-size: 14px; color: #052390; font-weight: 600; }
.term-card table { box-shadow: none; }
.term-card table th { position: static; background: #F5F5F7; color: #052390; }
.audit-stats { display: flex; gap: 20px; flex-wrap: wrap; margin-bottom: 24px; }
.audit-stat {
    flex: 1 1 160px; background: #FFFFFF; border-radius: 10px; padding: 18px 22px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.06); text-align: center;
}
.audit-stat .stat-val { font-size: 28px; font-weight: 700; margin: 0; }
.audit-stat .stat-label {
    font-size: 12px; font-weight: 600; color: #6D6E71;
    text-transform: uppercase; margin: 4px 0 0;
}
.footer {
    text-align: center; padding: 32px 40px; font-size: 12px;
    color: #6D6E71; border-top: 1px solid #E2E2E2;
}
@media print {
    .nav { display: none; }
    .section-header .toggle { display: none; }
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
        return !r.classList.contains('total-row');
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
        if (el) { e.preventDefault(); el.scrollIntoView({ behavior: 'smooth', block: 'start' }); }
    });
});
document.querySelectorAll('.section-body').forEach(function(body) {
    body.style.maxHeight = 'none';
});
"""

SA = '<span class="sort-arrow">&#9650;&#9660;</span>'

def th_row(headers):
    return "".join(f'<th onclick="sortTable(this,{i})">{h} {SA}</th>' for i, h in enumerate(headers))

html = '<!DOCTYPE html>\n<html lang="en">\n<head>\n'
html += '<meta charset="UTF-8">\n'
html += '<meta name="viewport" content="width=device-width, initial-scale=1.0">\n'
html += '<title>Broadridge Envelope Reconciliation Report</title>\n'
html += f'<style>{CSS}</style>\n'
html += '</head>\n<body>\n\n'

html += '<div class="header">\n'
html += '    <h1>Broadridge envelope reconciliation</h1>\n'
html += '    <p class="subtitle">Purchase vs. usage analysis &nbsp;|&nbsp; January 2020 &ndash; December 2025</p>\n'
html += f'    <p class="generated">Generated {pd.Timestamp.now().strftime("%B %d, %Y")}</p>\n'
html += '</div>\n'

html += '<nav class="nav">\n'
html += '    <a href="#executive-summary">Summary</a>\n'
html += '    <a href="#data-quality">Data Quality</a>\n'
html += '    <a href="#annual-summary">Annual</a>\n'
html += '    <a href="#monthly-trend">Trend</a>\n'
html += '    <a href="#monthly-detail">Monthly Detail</a>\n'
html += '    <a href="#envelope-types">Purchases by Type</a>\n'
html += '    <a href="#envelope-specs">Envelope Specs</a>\n'
html += '    <a href="#usage-by-product">Usage by Product</a>\n'
html += '    <a href="#contract-audit">Contract Audit</a>\n'
html += '</nav>\n'
html += '<div class="content">\n'

# Executive Summary
html += '<div class="section" id="executive-summary">\n'
html += '    <div class="section-header" onclick="toggleSection(this)">\n'
html += '        <h2>Executive summary</h2>\n'
html += '        <span class="toggle">&#9660;</span>\n'
html += '    </div>\n'
html += '    <div class="section-body">\n'
html += '        <div class="kpi-grid">\n'

html += f'            <div class="kpi-card"><p class="kpi-label">Total Purchased</p><p class="kpi-value" style="color:#2954F0">{fmt_num(total_purchased)}</p><p class="kpi-sub">Jan 2020 &ndash; Dec 2025</p></div>\n'
html += f'            <div class="kpi-card"><p class="kpi-label">Total Used (Volume)</p><p class="kpi-value" style="color:#2954F0">{fmt_num(total_used)}</p><p class="kpi-sub">{fmt_num(total_mailed)} mailed &middot; {fmt_num(total_spoils)} spoils</p></div>\n'
html += f'            <div class="kpi-card"><p class="kpi-label">Net Variance</p><p class="kpi-value" style="color:{kpi_var_color}">{fmt_num_parens(net_variance)}</p><p class="kpi-sub">Purchased minus used</p></div>\n'
html += f'            <div class="kpi-card"><p class="kpi-label">Total Invoiced</p><p class="kpi-value" style="color:#2954F0">{fmt_money_always(total_invoiced)}</p><p class="kpi-sub">Total cost: {fmt_money_always(total_cost)}</p></div>\n'
html += f'            <div class="kpi-card"><p class="kpi-label">Contract Overcharge</p><p class="kpi-value" style="color:{kpi_over_color}">{fmt_money_parens(audit_net_diff)}</p><p class="kpi-sub">{audit_over} over &middot; {audit_under} under &middot; {audit_ok} OK</p></div>\n'
html += f'            <div class="kpi-card"><p class="kpi-label">Variance %</p><p class="kpi-value" style="color:{kpi_var_color}">{net_variance/total_purchased*100 if total_purchased else 0:.1f}%</p><p class="kpi-sub">{"Surplus" if net_variance >= 0 else "Deficit"}</p></div>\n'

html += '        </div>\n'

html += '    </div>\n'
html += '</div>\n\n'

# Data Quality
html += '<div class="section" id="data-quality">\n'
html += '    <div class="section-header" onclick="toggleSection(this)">\n'
html += '        <h2>Data quality notes</h2>\n'
html += '        <span class="toggle">&#9660;</span>\n'
html += '    </div>\n'
html += '    <div class="section-body">\n'
if dq_rows.strip():
    html += '        <div class="alert-box">\n'
    html += '            <h3>&#9888; Outstanding data gaps</h3>\n'
    html += '            <p>The following months have missing source files that need to be requested from Broadridge. '
    html += 'These gaps affect the running balance and net variance calculations. '
    html += 'The reconciliation period is January 2020 through December 2025.</p>\n'
    html += '        </div>\n'
    html += '        <div class="table-wrap"><table>\n'
    html += '            <thead><tr><th>Month</th><th>Issue</th><th>Usage</th><th>Purchases</th><th>Action</th></tr></thead>\n'
    html += '            <tbody>\n' + dq_rows + '\n            </tbody>\n'
    html += '        </table></div>\n'
else:
    html += '        <p style="color:#186741;font-weight:600;font-size:14px;">&#10003; All months from January 2020 through December 2025 have source data. No outstanding gaps.</p>\n'
html += '    </div>\n</div>\n\n'

# Annual Summary
html += '<div class="section" id="annual-summary">\n'
html += '    <div class="section-header" onclick="toggleSection(this)">\n'
html += '        <h2>Annual summary</h2>\n'
html += '        <span class="toggle">&#9660;</span>\n'
html += '    </div>\n'
html += '    <div class="section-body">\n'
html += '        <div class="table-wrap"><table class="sortable">\n'
html += '            <thead><tr>' + th_row(["Year","Purchased","Used","Mailed","Spoils","Net Variance","Running Balance","Total Cost","Total Invoiced"]) + '</tr></thead>\n'
html += '            <tbody>\n' + annual_rows + '\n            </tbody>\n'
html += '        </table></div>\n'
html += '    </div>\n</div>\n\n'

# Monthly Trend
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

# Monthly Detail
html += '<div class="section" id="monthly-detail">\n'
html += '    <div class="section-header" onclick="toggleSection(this)">\n'
html += '        <h2>Monthly detail</h2>\n'
html += '        <span class="toggle">&#9660;</span>\n'
html += '    </div>\n'
html += '    <div class="section-body">\n'
html += '        <div class="table-wrap"><table class="sortable">\n'
html += '            <thead><tr>' + th_row(["Month","Purchased","Used","Mailed","Spoils","Net Variance","Variance %","Running Balance","Purchase Cost","Invoiced"]) + '</tr></thead>\n'
html += '            <tbody>\n' + monthly_rows + '\n            </tbody>\n'
html += '        </table></div>\n'
html += '    </div>\n</div>\n\n'

# Envelope Types
html += '<div class="section" id="envelope-types">\n'
html += '    <div class="section-header" onclick="toggleSection(this)">\n'
html += '        <h2>Purchases by envelope type</h2>\n'
html += '        <span class="toggle">&#9660;</span>\n'
html += '    </div>\n'
html += '    <div class="section-body">\n'
html += '        <div class="table-wrap"><table class="sortable">\n'
html += '            <thead><tr>' + th_row(["Envelope Type","Total Purchased","Total Cost","Avg Unit Price"]) + '<th>% of Total</th></tr></thead>\n'
html += '            <tbody>\n' + env_type_rows + '\n            </tbody>\n'
html += '        </table></div>\n'
html += '    </div>\n</div>\n\n'

# Envelope Specifications
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

html += '<div class="section" id="envelope-specs">\n'
html += '    <div class="section-header" onclick="toggleSection(this)">\n'
html += '        <h2>Envelope specifications</h2>\n'
html += '        <span class="toggle">&#9660;</span>\n'
html += '    </div>\n'
html += '    <div class="section-body">\n'
html += '        <p style="font-size:13px;color:#6D6E71;margin:0 0 16px;">All envelopes are double-window, 24WW paper, black ink with crosshatch black inside tint. Supplier: United Envelope LLC, Mt. Pocono, PA.</p>\n'
html += '        <div class="table-wrap"><table>\n'
html += '            <thead><tr><th>WMS Code</th><th>Mail Type</th><th>Size</th><th>Style</th><th>Postage</th><th>Notes</th></tr></thead>\n'
html += '            <tbody>\n' + envelope_spec_rows + '\n            </tbody>\n'
html += '        </table></div>\n'
html += '        <div style="margin-top:16px;display:flex;gap:12px;flex-wrap:wrap;font-size:12px;color:#6D6E71;">\n'
html += '            <span><strong>PFC</strong> = Pre-printed First-Class permit (domestic)</span>\n'
html += '            <span><strong>NI</strong> = No Imprint (foreign &mdash; postage applied at mailing)</span>\n'
html += '            <span><strong>DW</strong> = Double Window</span>\n'
html += '            <span><strong>IND</strong> = Individual (Oct 2022 revision)</span>\n'
html += '        </div>\n'
html += '    </div>\n</div>\n\n'

# Usage by Product
html += '<div class="section" id="usage-by-product">\n'
html += '    <div class="section-header" onclick="toggleSection(this)">\n'
html += '        <h2>Usage by product</h2>\n'
html += '        <span class="toggle">&#9660;</span>\n'
html += '    </div>\n'
html += '    <div class="section-body">\n'
html += '        <div class="table-wrap"><table class="sortable">\n'
html += '            <thead><tr>' + th_row(["Product Name","Total Used","First Month","Last Month"]) + '<th>% of Total</th></tr></thead>\n'
html += '            <tbody>\n' + usage_product_rows + '\n            </tbody>\n'
html += '        </table></div>\n'
html += '    </div>\n</div>\n\n'

# Contract Audit
html += '<div class="section" id="contract-audit">\n'
html += '    <div class="section-header" onclick="toggleSection(this)">\n'
html += '        <h2>Contract audit summary</h2>\n'
html += '        <span class="toggle">&#9660;</span>\n'
html += '    </div>\n'
html += '    <div class="section-body">\n'

# Contract terms
html += '        <div class="terms-grid">\n'
html += '            <div class="term-card"><h4>Original contract (Jan 2019)</h4>\n'
html += '                <table><thead><tr><th>Term</th><th>Value</th></tr></thead>\n'
html += '                <tbody><tr><td>Markup</td><td>15%</td></tr>\n'
html += '                <tr><td>Effective Date</td><td>January 2019</td></tr>\n'
html += '                <tr><td>Scope</td><td>Print &amp; mail services</td></tr></tbody></table></div>\n'
html += '            <div class="term-card"><h4>Amendment No. 1 (Jan 2024)</h4>\n'
html += '                <table><thead><tr><th>Term</th><th>Value</th></tr></thead>\n'
html += '                <tbody><tr><td>Markup</td><td>Reduced / revised</td></tr>\n'
html += '                <tr><td>Effective Date</td><td>January 2024</td></tr>\n'
html += '                <tr><td>Scope</td><td>Updated pricing schedule</td></tr></tbody></table></div>\n'
html += '        </div>\n'

# Audit stats
html += '        <div class="audit-stats">\n'
html += f'            <div class="audit-stat"><p class="stat-val" style="color:#9D1526">{audit_over}</p><p class="stat-label">Over-billed lines</p></div>\n'
html += f'            <div class="audit-stat"><p class="stat-val" style="color:#186741">{audit_under}</p><p class="stat-label">Under-billed lines</p></div>\n'
html += f'            <div class="audit-stat"><p class="stat-val" style="color:#6D6E71">{audit_ok}</p><p class="stat-label">Correctly billed</p></div>\n'
html += f'            <div class="audit-stat"><p class="stat-val" style="color:{kpi_over_color}">{fmt_money_parens(audit_net_diff)}</p><p class="stat-label">Net overcharge</p></div>\n'
html += '        </div>\n'

# Top 10
html += '        <h3 style="color:#052390;font-size:16px;margin:0 0 12px;">Top 10 largest discrepancies</h3>\n'
html += '        <div class="table-wrap"><table class="sortable">\n'
html += '            <thead><tr>' + th_row(["Month","Description","Qty","Unit Price","Expected","Actual","Difference","Flag"]) + '</tr></thead>\n'
html += '            <tbody>\n' + top10_rows + '\n            </tbody>\n'
html += '        </table></div>\n'
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
