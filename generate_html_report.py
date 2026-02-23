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
by_type = pd.read_excel(EXCEL_PATH, sheet_name="By Envelope Type")
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

usage_by_product_sorted = usage_by_product.sort_values("Total Envelopes Used", ascending=False)
usage_product_total = usage_by_product["Total Envelopes Used"].sum()

by_type_sorted = by_type.sort_values("Total Purchased", ascending=False)
env_type_total = by_type["Total Purchased"].sum()

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

def build_monthly_rows():
    """Build monthly rows with annual subtotals after each December and a grand total."""
    rows = []
    running_balance = 0
    # Year accumulators
    yr_p = yr_u = yr_m = yr_s = 0
    prev_year = None

    def _year_from_label(label):
        return 2000 + int(label.split('-')[1])

    def _subtotal_row(year, yp, yu, ym, ys, rb):
        yv = yp - yu
        vc = var_color(yv)
        rbc = var_color(rb)
        vpct = yv / yp if yp else 0
        return (
            '<tr class="subtotal-row">'
            + f'<td><strong>{year} Total</strong></td>'
            + f'<td class="num"><strong>{fmt_num(yp)}</strong></td>'
            + f'<td class="num"><strong>{fmt_num(yu)}</strong></td>'
            + f'<td class="num"><strong>{fmt_num(ym)}</strong></td>'
            + f'<td class="num"><strong>{fmt_num(ys)}</strong></td>'
            + f'<td class="num" style="color:{vc};font-weight:700"><strong>{fmt_num_parens(yv)}</strong></td>'
            + f'<td class="num" style="color:{vc};font-weight:700"><strong>{fmt_pct(vpct)}</strong></td>'
            + f'<td class="num" style="color:{rbc};font-weight:700"><strong>{fmt_num_parens(rb)}</strong></td>'
            + '</tr>'
        )

    for i, (_, r) in enumerate(monthly.iterrows()):
        cur_year = _year_from_label(r["Month"])

        # Insert subtotal when year changes (not on the first row)
        if prev_year is not None and cur_year != prev_year:
            rows.append(_subtotal_row(prev_year, yr_p, yr_u, yr_m, yr_s, running_balance))
            yr_p = yr_u = yr_m = yr_s = 0

        prev_year = cur_year
        p = safe(r["Envelopes Purchased"])
        u = safe(r["Envelopes Used (Volume)"])
        m = safe(r["Envelopes Mailed (Postage)"])
        s = safe(r["Spoils"])
        vv = p - u
        running_balance += vv
        yr_p += p; yr_u += u; yr_m += m; yr_s += s

        vc = var_color(vv)
        rbc = var_color(running_balance)
        bg = "#F5F5F7" if i % 2 == 1 else "#FFFFFF"
        rows.append(
            f'<tr style="background:{bg}">'
            + f'<td>{r["Month"]}</td>'
            + f'<td class="num">{fmt_num(p)}</td>'
            + f'<td class="num">{fmt_num(u)}</td>'
            + f'<td class="num">{fmt_num(m)}</td>'
            + f'<td class="num">{fmt_num(s)}</td>'
            + f'<td class="num" style="color:{vc};font-weight:600">{fmt_num_parens(vv)}</td>'
            + f'<td class="num" style="color:{vc}">{fmt_pct(r["Variance %"])}</td>'
            + f'<td class="num" style="color:{rbc};font-weight:600">{fmt_num_parens(running_balance)}</td>'
            + '</tr>'
        )

    # Final year subtotal
    if prev_year is not None:
        rows.append(_subtotal_row(prev_year, yr_p, yr_u, yr_m, yr_s, running_balance))

    # Grand total
    gv = total_purchased - total_used
    vc = var_color(gv)
    rbc = var_color(running_balance)
    gpct = gv / total_purchased if total_purchased else 0
    rows.append(
        '<tr class="total-row">'
        + f'<td><strong>Grand Total</strong></td>'
        + f'<td class="num"><strong>{fmt_num(total_purchased)}</strong></td>'
        + f'<td class="num"><strong>{fmt_num(total_used)}</strong></td>'
        + f'<td class="num"><strong>{fmt_num(total_mailed)}</strong></td>'
        + f'<td class="num"><strong>{fmt_num(total_spoils)}</strong></td>'
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

svg_chart = build_svg_chart()
monthly_rows = build_monthly_rows()
env_type_rows = build_env_type_rows()
usage_product_rows = build_usage_by_product_rows()

kpi_var_color = "#9D1526" if net_variance < 0 else "#186741"

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
    border-radius: 0 8px 8px 0; padding: 16px 24px; margin-bottom: 20px;
}
.alert-box p { margin: 0; font-size: 13px; color: #333; }
.alert-box strong { color: #FC5E17; }
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
html += '    <a href="#monthly-trend">Trend</a>\n'
html += '    <a href="#monthly-detail">Monthly Detail</a>\n'
html += '    <a href="#envelope-types">Purchases &amp; Usage</a>\n'
html += '    <a href="#envelope-specs">Envelope Specs</a>\n'
html += '</nav>\n'
html += '<div class="content">\n'

# Missing data alert (shown only if gaps exist)
if missing_months:
    html += '    <div class="alert-box">\n'
    html += f'        <p><strong>&#9888; Missing data:</strong> {", ".join(missing_months)} &mdash; '
    html += 'missing source files affect variance calculations.</p>\n'
    html += '    </div>\n\n'

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
html += f'            <div class="kpi-card"><p class="kpi-label">Variance %</p><p class="kpi-value" style="color:{kpi_var_color}">{net_variance/total_purchased*100 if total_purchased else 0:.1f}%</p><p class="kpi-sub">{"Surplus" if net_variance >= 0 else "Deficit"}</p></div>\n'

html += '        </div>\n'

html += '    </div>\n'
html += '</div>\n\n'

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
html += '            <thead><tr>' + th_row(["Month","Purchased","Used","Mailed","Spoils","Net Variance","Variance %","Running Balance"]) + '</tr></thead>\n'
html += '            <tbody>\n' + monthly_rows + '\n            </tbody>\n'
html += '        </table></div>\n'
html += '    </div>\n</div>\n\n'

# Envelope Types (purchases + usage side by side)
html += '<div class="section" id="envelope-types">\n'
html += '    <div class="section-header" onclick="toggleSection(this)">\n'
html += '        <h2>Purchases &amp; usage breakdown</h2>\n'
html += '        <span class="toggle">&#9660;</span>\n'
html += '    </div>\n'
html += '    <div class="section-body">\n'
html += '        <div style="display:flex;gap:24px;flex-wrap:wrap;">\n'
html += '            <div style="flex:1 1 400px;">\n'
html += '                <h3 style="color:#052390;font-size:15px;margin:0 0 12px;">Purchased by envelope type</h3>\n'
html += '                <div class="table-wrap"><table class="sortable">\n'
html += '                    <thead><tr>' + th_row(["Envelope Type","Total Purchased"]) + '<th>% of Total</th></tr></thead>\n'
html += '                    <tbody>\n' + env_type_rows + '\n                    </tbody>\n'
html += '                </table></div>\n'
html += '            </div>\n'
html += '            <div style="flex:1 1 400px;">\n'
html += '                <h3 style="color:#052390;font-size:15px;margin:0 0 12px;">Used by product</h3>\n'
html += '                <div class="table-wrap"><table class="sortable">\n'
html += '                    <thead><tr>' + th_row(["Product Name","Total Used"]) + '<th>% of Total</th></tr></thead>\n'
html += '                    <tbody>\n' + usage_product_rows + '\n                    </tbody>\n'
html += '                </table></div>\n'
html += '            </div>\n'
html += '        </div>\n'
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
