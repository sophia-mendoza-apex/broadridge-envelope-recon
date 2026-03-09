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
usage_by_env_type_monthly = pd.read_excel(EXCEL_PATH, sheet_name="Usage by Envelope Type")
usage_by_product = pd.read_excel(EXCEL_PATH, sheet_name="Usage by Product")
purchase_detail = pd.read_excel(EXCEL_PATH, sheet_name="Purchase Detail")

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
post_yearly = _defaultdict(lambda: [0, 0, 0, 0, 0, 0, 0])  # purchased, used, mailed, spoils, month_count, cost, invoiced
for _, r in post.iterrows():
    yr = 2000 + int(r["Month"].split('-')[1])
    post_yearly[yr][0] += safe(r["Envelopes Purchased"])
    post_yearly[yr][1] += safe(r["Envelopes Used (Volume)"])
    post_yearly[yr][2] += safe(r["Envelopes Mailed (Postage)"])
    post_yearly[yr][3] += safe(r["Spoils"])
    post_yearly[yr][4] += 1
    post_yearly[yr][5] += safe(r["Purchase Cost"])
    post_yearly[yr][6] += safe(r["Invoiced Amount"])

# Post-settlement cost totals
post_cost = post["Purchase Cost"].sum()
post_invoiced = post["Invoiced Amount"].sum()
post_avg_unit_cost = post_cost / post_purchased if post_purchased else 0

# Trailing 12-month cost for projection
recent_12_cost = monthly.tail(12)["Purchase Cost"].sum()
recent_12_inv = monthly.tail(12)["Invoiced Amount"].sum()
projected_annual_cost = recent_12_cost
projected_annual_inv = recent_12_inv

# Billing basis discrepancy — contract says "usage", Broadridge bills on "receipt"
avg_inv_per_unit = post_invoiced / post_purchased if post_purchased else 0
usage_based_invoice = post_used * avg_inv_per_unit
billing_excess = post_invoiced - usage_based_invoice
# Monthly breakdown: months with 0 purchases but non-zero usage → $0 billed
zero_purchase_usage_months = len(post[(post["Envelopes Purchased"] == 0) & (post["Envelopes Used (Volume)"] > 0)])

# Rolling avg monthly usage (last 12 months) for buffer stock calculation
recent_12 = monthly.tail(12)
avg_monthly_usage = int(recent_12["Envelopes Used (Volume)"].mean())


# ---------------------------------------------------------------------------
# Post-settlement per-SKU buffer stock analysis
# ---------------------------------------------------------------------------
bt_post = by_type_monthly[by_type_monthly["Month"].apply(lambda x: month_label_to_sortkey(x) >= settlement_key)]
ut_post = usage_by_env_type_monthly[usage_by_env_type_monthly["Month"].apply(lambda x: month_label_to_sortkey(x) >= settlement_key)]

# Contractual max wastage rates (applied to usage to get total consumption)
# Original contract (Jan 2019 – Dec 2023): 5% envelope wastage
# Amendment No. 1 (Jan 2024 – present): 2% envelope wastage
WASTAGE_CUTOVER = (24, 0)  # Jan-24
WASTAGE_PRE_2024 = 0.05
WASTAGE_POST_2024 = 0.02

def get_wastage_rate(month_label):
    return WASTAGE_POST_2024 if month_label_to_sortkey(month_label) >= WASTAGE_CUTOVER else WASTAGE_PRE_2024

# Add wastage-adjusted usage column (keep float precision, round at aggregate level)
ut_post = ut_post.copy()
ut_post["Wastage Rate"] = ut_post["Month"].apply(get_wastage_rate)
ut_post["Wastage_raw"] = ut_post["Envelopes Used"] * ut_post["Wastage Rate"]  # float
ut_post["Adj Used_raw"] = ut_post["Envelopes Used"] + ut_post["Wastage_raw"]

# Trailing 12 months usage by type
last_12_usage_months = sorted(ut_post["Month"].unique(), key=month_label_to_sortkey)[-12:]
ut_trail12 = ut_post[ut_post["Month"].isin(last_12_usage_months)]

# Post-settlement lookups by SKU (raw and wastage-adjusted)
post_purchase_by_sku = bt_post.groupby("Envelope Type")["Purchased"].sum().to_dict()
post_usage_by_sku = ut_post.groupby("Envelope Type")["Envelopes Used"].sum().to_dict()
# Wastage: accumulate as floats, round at aggregate level per type
_wastage_by_sku_float = ut_post.groupby("Envelope Type")["Wastage_raw"].sum().to_dict()
post_wastage_by_sku = {k: int(round(v)) for k, v in _wastage_by_sku_float.items()}
_adj_used_by_sku_float = ut_post.groupby("Envelope Type")["Adj Used_raw"].sum().to_dict()
post_adj_usage_by_sku = {k: int(round(v)) for k, v in _adj_used_by_sku_float.items()}
trail12_usage_by_sku = {k: int(round(v)) for k, v in ut_trail12.groupby("Envelope Type")["Adj Used_raw"].sum().to_dict().items()}

# Authoritative monthly-level wastage total (matches monthly detail grand total)
total_wastage_allowance = sum(int(safe(r["Envelopes Used (Volume)"]) * get_wastage_rate(r["Month"])) for _, r in post.iterrows())

# Reconcile SKU-level wastage to authoritative monthly total
_sku_wastage_sum = sum(post_wastage_by_sku.values())
if _sku_wastage_sum != total_wastage_allowance:
    _delta = total_wastage_allowance - _sku_wastage_sum
    # Adjust the largest type's wastage so the sum equals total_wastage_allowance exactly
    _largest_sku = max(post_wastage_by_sku, key=post_wastage_by_sku.get)
    post_wastage_by_sku[_largest_sku] += _delta
    post_adj_usage_by_sku[_largest_sku] += _delta

# SKU-level buffer analysis — list of (sku, purchased, used, variance, avg_mo_usage, buffer_months, note)
SKU_DISPLAY_NAMES = {
    "ENVAPXN10 Confirms+Letters (PFC)": "#10 Confirms + Letters (PFC)",
    "ENVCONPFSN10NI": "#10 Confirms + Letters (NI)",
    "ENVMEAPEXN14PFC": "N14 Fold Statement (PFC)",
    "ENVMERIDGEN14NI11/08": "N14 Fold Statement (NI)",
    "ENVMEAPEX9X12PFC": "9x12 Flat Statement (PFC)",
    "ENVMERIDGE9X12NI11/08": "9x12 Flat Statement (NI)",
    "ENVCONRIDGE9X12DW": "9x12 Flat Confirms (DW)",
    "Tax Form Envelopes (1099/1099-R)": "Tax Forms (1099/1099-R)",
    "Tax Form Envelopes (1042/IRA)": "Tax Forms (1042/IRA)",
}

# Last purchase/usage month and unit cost per SKU
last_purchase_by_sku = bt_post.groupby("Envelope Type")["Month"].apply(lambda x: max(x, key=month_label_to_sortkey)).to_dict()
cost_agg = bt_post.groupby("Envelope Type").agg({"Total Cost": "sum", "Purchased": "sum"})
cost_agg["Unit Cost"] = cost_agg["Total Cost"] / cost_agg["Purchased"]
unit_cost_by_sku = cost_agg["Unit Cost"].to_dict()

all_sku_keys = sorted(
    set(list(post_purchase_by_sku.keys()) + list(post_usage_by_sku.keys())),
    key=lambda t: -(post_purchase_by_sku.get(t, 0) + post_usage_by_sku.get(t, 0))
)

# sku_buffer_data: (sku, display, purchased, used, wastage, adj_used, variance, avg_mo_adj,
#                   buffer_months, last_purchase, unit_cost, excess_dollars, action)
sku_buffer_data = []
for sku in all_sku_keys:
    p = post_purchase_by_sku.get(sku, 0)
    u = post_usage_by_sku.get(sku, 0)
    w = post_wastage_by_sku.get(sku, 0)
    au = post_adj_usage_by_sku.get(sku, 0)
    v = p - au  # variance against adjusted usage (incl. wastage)
    t12 = trail12_usage_by_sku.get(sku, 0)  # already adjusted
    avg_mo = t12 / 12 if t12 > 0 else 0
    buf_mo = v / avg_mo if avg_mo > 0 else (999 if v > 0 else -999 if v < 0 else 0)
    display = SKU_DISPLAY_NAMES.get(sku, sku)
    lp = last_purchase_by_sku.get(sku, None)
    uc = unit_cost_by_sku.get(sku, 0)

    # Excess dollars: surplus above 3-month policy target, valued at unit cost
    target_3mo = avg_mo * 3
    excess_units = max(0, v - target_3mo)
    excess_dollars = excess_units * uc

    # Action recommendation
    if buf_mo < 0:
        action = "Covered by prior stock"
    elif buf_mo > 100:
        action = "Stop purchasing"
    elif buf_mo > 6:
        action = "Reduce orders"
    elif buf_mo > 3:
        action = "Monitor"
    elif buf_mo >= 2:
        action = "On target"
    else:
        action = "Increase orders"

    sku_buffer_data.append((sku, display, p, u, w, au, v, avg_mo, buf_mo, lp, uc, excess_dollars, action))

# Interchangeable NI/PFC pairs: when PFC shows deficit but NI has excess,
# the combined inventory covers both — adjust action accordingly
_INTERCHANGEABLE_PAIRS = {
    "ENVAPXN10 Confirms+Letters (PFC)": "ENVCONPFSN10NI",
    "ENVMEAPEXN14PFC": "ENVMERIDGEN14NI11/08",
    "ENVMEAPEX9X12PFC": "ENVMERIDGE9X12NI11/08",
}
_sku_idx = {d[0]: i for i, d in enumerate(sku_buffer_data)}
for pfc_sku, ni_sku in _INTERCHANGEABLE_PAIRS.items():
    if pfc_sku in _sku_idx and ni_sku in _sku_idx:
        pi, ni = _sku_idx[pfc_sku], _sku_idx[ni_sku]
        pfc_d, ni_d = sku_buffer_data[pi], sku_buffer_data[ni]
        combined_var = pfc_d[6] + ni_d[6]
        combined_avg = pfc_d[7] + ni_d[7]
        combined_buf = combined_var / combined_avg if combined_avg > 0 else 0
        # If PFC shows low/deficit but combined pair has adequate buffer, override action
        if pfc_d[8] < 2 and combined_buf >= 2:
            new_action = "Covered by NI stock"
            sku_buffer_data[pi] = pfc_d[:12] + (new_action,)

# Filter out noise (tax forms with 0 purchases post-settlement)
# Tuple: (sku, display, p, u, w, au, v, avg_mo, buf_mo, lp, uc, excess_dollars, action)
#         0     1        2  3  4  5   6  7       8       9   10  11              12
sku_buffer_data = [d for d in sku_buffer_data if not (d[2] == 0 and d[5] == 0)]

# Sort by excess dollars descending (biggest exposure first)
sku_buffer_data.sort(key=lambda d: -d[11])

# Totals
total_excess_dollars = sum(d[11] for d in sku_buffer_data)

# Identify excess from retired/foreign SKUs
retired_foreign_skus = {"ENVCONPFSN10NI", "ENVMERIDGEN14NI11/08", "ENVMERIDGE9X12NI11/08"}
excess_from_retired = sum(d[11] for d in sku_buffer_data if d[0] in retired_foreign_skus)
excess_pct_of_total = excess_from_retired / total_excess_dollars * 100 if total_excess_dollars > 0 else 0

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

def build_monthly_rows():
    """Build monthly rows (post-settlement scope) with wastage, annual subtotals and grand total."""
    rows = []
    running_balance = 0
    yr_p = yr_u = yr_w = 0
    prev_year = None

    def _year_from_label(label):
        return 2000 + int(label.split('-')[1])

    def _subtotal_row(year, yp, yu, yw, rb):
        yau = yu + yw
        yv = yp - yau
        vc = var_color(yv)
        rbc = var_color(rb)
        vpct = yv / yp if yp else 0
        yr_label = f"{year} (Mar&ndash;Dec)" if year == 2022 else str(year)
        return (
            '<tr class="subtotal-row">'
            + f'<td><strong>{yr_label} Total</strong></td>'
            + f'<td class="num"><strong>{fmt_num(yp)}</strong></td>'
            + f'<td class="num"><strong>{fmt_num(yu)}</strong></td>'
            + f'<td class="num" style="color:#9A9BA0;"><strong>{fmt_num(yw)}</strong></td>'
            + f'<td class="num" style="color:{vc};font-weight:700"><strong>{fmt_num_parens(yv)}</strong></td>'
            + f'<td class="num" style="color:{vc};font-weight:700"><strong>{fmt_pct(vpct)}</strong></td>'
            + f'<td class="num" style="color:{rbc};font-weight:700"><strong>{fmt_num_parens(rb)}</strong></td>'
            + '</tr>'
        )

    for i, (_, r) in enumerate(post.iterrows()):
        cur_year = _year_from_label(r["Month"])

        if prev_year is not None and cur_year != prev_year:
            rows.append(_subtotal_row(prev_year, yr_p, yr_u, yr_w, running_balance))
            yr_p = yr_u = yr_w = 0

        prev_year = cur_year
        p = safe(r["Envelopes Purchased"])
        u = safe(r["Envelopes Used (Volume)"])
        w_rate = get_wastage_rate(r["Month"])
        w = int(u * w_rate)
        au = u + w
        vv = p - au
        running_balance += vv
        yr_p += p; yr_u += u; yr_w += w

        vc = var_color(vv)
        rbc = var_color(running_balance)
        vpct = vv / p if p else 0
        bg = "#252629" if i % 2 == 1 else "#1E1F23"
        rows.append(
            f'<tr style="background:{bg}">'
            + f'<td>{r["Month"]}</td>'
            + f'<td class="num">{fmt_num(p)}</td>'
            + f'<td class="num">{fmt_num(u)}</td>'
            + f'<td class="num" style="color:#9A9BA0;">{fmt_num(w)}</td>'
            + f'<td class="num" style="color:{vc};font-weight:600">{fmt_num_parens(vv)}</td>'
            + f'<td class="num" style="color:{vc}">{fmt_pct(vpct)}</td>'
            + f'<td class="num" style="color:{rbc};font-weight:600">{fmt_num_parens(running_balance)}</td>'
            + '</tr>'
        )

    # Final year subtotal
    if prev_year is not None:
        rows.append(_subtotal_row(prev_year, yr_p, yr_u, yr_w, running_balance))

    # Grand total
    gw = sum(int(safe(r["Envelopes Used (Volume)"]) * get_wastage_rate(r["Month"])) for _, r in post.iterrows())
    gau = post_used + gw
    gv = post_purchased - gau
    vc = var_color(gv)
    rbc = var_color(running_balance)
    gpct = gv / post_purchased if post_purchased else 0
    rows.append(
        '<tr class="total-row">'
        + f'<td><strong>Grand Total</strong></td>'
        + f'<td class="num"><strong>{fmt_num(post_purchased)}</strong></td>'
        + f'<td class="num"><strong>{fmt_num(post_used)}</strong></td>'
        + f'<td class="num" style="color:#9A9BA0;"><strong>{fmt_num(gw)}</strong></td>'
        + f'<td class="num" style="color:{vc};font-weight:700"><strong>{fmt_num_parens(gv)}</strong></td>'
        + f'<td class="num" style="color:{vc};font-weight:700"><strong>{fmt_pct(gpct)}</strong></td>'
        + f'<td class="num" style="color:{vc};font-weight:700"><strong>{fmt_num_parens(running_balance)}</strong></td>'
        + '</tr>'
    )
    return "\n".join(rows)

def _status_tag(buf_mo, action):
    """Return a styled status pill for the buffer assessment."""
    if buf_mo < 0:
        return '<span style="background:rgba(186,104,200,0.15);color:#CE93D8;padding:3px 10px;border-radius:12px;font-size:11px;font-weight:600;">DEFICIT</span>'
    elif buf_mo > 100:
        return '<span style="background:rgba(239,83,80,0.15);color:#EF5350;padding:3px 10px;border-radius:12px;font-size:11px;font-weight:600;">EXCESS</span>'
    elif buf_mo > 6:
        return '<span style="background:rgba(255,167,38,0.15);color:#FFA726;padding:3px 10px;border-radius:12px;font-size:11px;font-weight:600;">OVERSTOCKED</span>'
    elif buf_mo > 3:
        return '<span style="background:rgba(255,213,79,0.15);color:#FFD54F;padding:3px 10px;border-radius:12px;font-size:11px;font-weight:600;">HIGH</span>'
    elif buf_mo >= 2:
        return '<span style="background:rgba(76,175,121,0.15);color:#4CAF79;padding:3px 10px;border-radius:12px;font-size:11px;font-weight:600;">ON TARGET</span>'
    else:
        return '<span style="background:rgba(255,167,38,0.15);color:#FFA726;padding:3px 10px;border-radius:12px;font-size:11px;font-weight:600;">LOW</span>'

def build_sku_buffer_table():
    """HTML table showing per-SKU buffer stock with wastage, dollars, and recommended action."""
    # Tuple: (sku, display, p, u, w, au, v, avg_mo, buf_mo, lp, uc, excess_dollars, action)
    #         0     1        2  3  4  5   6  7       8       9   10  11              12
    rows = []
    for d in sku_buffer_data:
        sku, display, p, u, w, au, v, avg_mo, buf_mo, lp, uc, excess_dollars, action = d
        vc = var_color(v)
        status = _status_tag(buf_mo, action)

        # Buffer months display
        if buf_mo > 100:
            buf_display = f"{buf_mo:.0f}"
        elif buf_mo < -100:
            buf_display = DASH
        else:
            buf_display = f"{buf_mo:.1f}"

        # Last purchased display
        lp_display = lp if lp else DASH

        # Excess cost display
        exc_display = f"${excess_dollars:,.0f}" if excess_dollars > 0 else DASH

        # Action styling
        if action == "Stop purchasing":
            action_html = f'<span style="color:#EF5350;font-weight:600;">{action}</span>'
        elif action in ("Reduce orders", "Increase orders"):
            action_html = f'<span style="color:#FFA726;font-weight:600;">{action}</span>'
        elif action == "On target":
            action_html = f'<span style="color:#4CAF79;">{action}</span>'
        elif action == "Covered by NI stock":
            action_html = f'<span style="color:#5B9BF7;">{action}</span>'
        else:
            action_html = f'<span style="color:#9A9BA0;">{action}</span>'

        rows.append(
            '<tr>'
            + f'<td class="env-name" style="font-weight:600;">{display}</td>'
            + f'<td style="text-align:center;">{status}</td>'
            + f'<td class="num">{fmt_num(p)}</td>'
            + f'<td class="num">{fmt_num(u)}</td>'
            + f'<td class="num" style="color:#9A9BA0;">{fmt_num(w)}</td>'
            + f'<td class="num" style="color:{vc};font-weight:600">{fmt_num_parens(v)}</td>'
            + f'<td class="num">{buf_display}</td>'
            + f'<td class="num">{exc_display}</td>'
            + f'<td>{lp_display}</td>'
            + f'<td>{action_html}</td>'
            + '</tr>'
        )

    # Total row
    tp = sum(d[2] for d in sku_buffer_data)
    tu = sum(d[3] for d in sku_buffer_data)
    tw = sum(d[4] for d in sku_buffer_data)
    tau = sum(d[5] for d in sku_buffer_data)
    tv = tp - tau
    tvc = var_color(tv)
    adj_avg_monthly = sum(d[7] for d in sku_buffer_data)
    overall_buf = tv / adj_avg_monthly if adj_avg_monthly else 0
    rows.append(
        '<tr class="total-row">'
        + f'<td><strong>Total</strong></td>'
        + f'<td></td>'
        + f'<td class="num"><strong>{fmt_num(tp)}</strong></td>'
        + f'<td class="num"><strong>{fmt_num(tu)}</strong></td>'
        + f'<td class="num" style="color:#9A9BA0;"><strong>{fmt_num(tw)}</strong></td>'
        + f'<td class="num" style="color:{tvc};font-weight:700"><strong>{fmt_num_parens(tv)}</strong></td>'
        + f'<td class="num"><strong>{overall_buf:.1f}</strong></td>'
        + f'<td class="num"><strong>${total_excess_dollars:,.0f}</strong></td>'
        + f'<td></td>'
        + f'<td></td>'
        + '</tr>'
    )
    return "\n".join(rows)

def build_inventory_gauge():
    """SVG gauge comparing actual inventory vs Broadridge 2-3 month policy."""
    policy_min = avg_monthly_usage_adj * 2
    policy_max = avg_monthly_usage_adj * 3
    actual = post_adj_variance
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
    s.append(f'<rect x="{ml}" y="{bar_y}" width="{bar_w}" height="{bar_h}" fill="#2A2B30" rx="4"/>')
    # Under-target zone (0 to policy_min) — amber
    s.append(f'<rect x="{ml}" y="{bar_y}" width="{xp(policy_min)-ml:.1f}" height="{bar_h}" fill="rgba(255,167,38,0.2)" rx="4"/>')
    # Policy zone (min to max) — green
    s.append(f'<rect x="{xp(policy_min):.1f}" y="{bar_y}" width="{xp(policy_max)-xp(policy_min):.1f}" height="{bar_h}" fill="rgba(76,175,121,0.25)"/>')
    # Over-target zone (max to end) — light red
    s.append(f'<rect x="{xp(policy_max):.1f}" y="{bar_y}" width="{xp(max_val)-xp(policy_max):.1f}" height="{bar_h}" fill="rgba(239,83,80,0.2)" rx="4"/>')
    # Policy range borders
    s.append(f'<line x1="{xp(policy_min):.1f}" y1="{bar_y}" x2="{xp(policy_min):.1f}" y2="{bar_y+bar_h}" stroke="#4CAF79" stroke-width="1.5"/>')
    s.append(f'<line x1="{xp(policy_max):.1f}" y1="{bar_y}" x2="{xp(policy_max):.1f}" y2="{bar_y+bar_h}" stroke="#4CAF79" stroke-width="1.5"/>')
    # Actual marker
    ax = xp(actual)
    s.append(f'<line x1="{ax:.1f}" y1="{bar_y-6}" x2="{ax:.1f}" y2="{bar_y+bar_h+6}" stroke="#5B9BF7" stroke-width="3"/>')
    s.append(f'<circle cx="{ax:.1f}" cy="{bar_y+bar_h/2:.1f}" r="8" fill="#5B9BF7"/>')
    s.append(f'<circle cx="{ax:.1f}" cy="{bar_y+bar_h/2:.1f}" r="4" fill="#1E1F23"/>')
    # Labels — above bar
    s.append(f'<text x="{ax:.1f}" y="{bar_y-12}" text-anchor="middle" font-size="13" fill="#82B4FF" font-weight="700" font-family="Helvetica Neue,Arial,sans-serif">Actual: {fmt_num(actual)}</text>')
    # Labels — below bar
    s.append(f'<text x="{xp(policy_min):.1f}" y="{bar_y+bar_h+16}" text-anchor="middle" font-size="11" fill="#4CAF79" font-family="Helvetica Neue,Arial,sans-serif">{fmt_num(policy_min)}</text>')
    s.append(f'<text x="{xp(policy_min):.1f}" y="{bar_y+bar_h+28}" text-anchor="middle" font-size="10" fill="#4CAF79" font-family="Helvetica Neue,Arial,sans-serif">2-mo min</text>')
    s.append(f'<text x="{xp(policy_max):.1f}" y="{bar_y+bar_h+16}" text-anchor="middle" font-size="11" fill="#4CAF79" font-family="Helvetica Neue,Arial,sans-serif">{fmt_num(policy_max)}</text>')
    s.append(f'<text x="{xp(policy_max):.1f}" y="{bar_y+bar_h+28}" text-anchor="middle" font-size="10" fill="#4CAF79" font-family="Helvetica Neue,Arial,sans-serif">3-mo max</text>')
    s.append('</svg>')
    return '\n'.join(s)

monthly_rows = build_monthly_rows()

# Build per-SKU monthly JSON for interactive Monthly Detail
import json as _json

def build_monthly_detail_json():
    """Build JSON: {sku: [{month, purchased, used, wastage, variance, var_pct}, ...], ...}"""
    # Get sorted post-settlement months
    post_months_list = sorted(post["Month"].unique(), key=month_label_to_sortkey)

    # Build purchase lookup: (month, sku) -> purchased
    purch_lookup = {}
    for _, r in bt_post.iterrows():
        purch_lookup[(r["Month"], r["Envelope Type"])] = int(r["Purchased"])

    # Build usage lookup: (month, sku) -> used
    usage_lookup = {}
    for _, r in ut_post.iterrows():
        usage_lookup[(r["Month"], r["Envelope Type"])] = int(r["Envelopes Used"])

    # Get all SKU keys that appear in post-settlement data (exclude noise)
    all_skus = sorted(
        set(k[1] for k in purch_lookup.keys()) | set(k[1] for k in usage_lookup.keys()),
        key=lambda t: -(sum(purch_lookup.get((m, t), 0) for m in post_months_list)
                       + sum(usage_lookup.get((m, t), 0) for m in post_months_list))
    )

    data = {}
    # "ALL" — aggregate view
    all_rows = []
    for m in post_months_list:
        p = int(safe(post[post["Month"] == m]["Envelopes Purchased"].sum()))
        u = int(safe(post[post["Month"] == m]["Envelopes Used (Volume)"].sum()))
        w_rate = get_wastage_rate(m)
        w = int(u * w_rate)
        v = p - u - w
        vpct = round(v / p, 4) if p else 0
        all_rows.append({"m": m, "p": p, "u": u, "w": w, "v": v, "vp": vpct})
    data["ALL"] = all_rows

    # Per-SKU
    for sku in all_skus:
        sku_rows = []
        has_data = False
        for m in post_months_list:
            p = purch_lookup.get((m, sku), 0)
            u = usage_lookup.get((m, sku), 0)
            if p == 0 and u == 0:
                continue
            has_data = True
            w_rate = get_wastage_rate(m)
            w = int(u * w_rate)
            v = p - u - w
            vpct = round(v / p, 4) if p else 0
            sku_rows.append({"m": m, "p": p, "u": u, "w": w, "v": v, "vp": vpct})
        if has_data:
            data[sku] = sku_rows

    # Combined groups — merge related SKUs that are physically interchangeable
    COMBINED_GROUPS = {
        "GRP_10_CONFIRMS": {
            "label": "#10 Confirms + Letters (NI + PFC combined)",
            "skus": ["ENVAPXN10 Confirms+Letters (PFC)", "ENVCONPFSN10NI"],
        },
        "GRP_N14_STMTS": {
            "label": "N14 Fold Statements (PFC + NI combined)",
            "skus": ["ENVMEAPEXN14PFC", "ENVMERIDGEN14NI11/08"],
        },
        "GRP_9X12_STMTS": {
            "label": "9x12 Flat Statements (PFC + NI combined)",
            "skus": ["ENVMEAPEX9X12PFC", "ENVMERIDGE9X12NI11/08"],
        },
    }
    for grp_key, grp in COMBINED_GROUPS.items():
        grp_rows = []
        for m in post_months_list:
            p = sum(purch_lookup.get((m, s), 0) for s in grp["skus"])
            u = sum(usage_lookup.get((m, s), 0) for s in grp["skus"])
            if p == 0 and u == 0:
                continue
            w_rate = get_wastage_rate(m)
            w = int(u * w_rate)
            v = p - u - w
            vpct = round(v / p, 4) if p else 0
            grp_rows.append({"m": m, "p": p, "u": u, "w": w, "v": v, "vp": vpct})
        if grp_rows:
            data[grp_key] = grp_rows

    return data, COMBINED_GROUPS

_detail_data, _combined_groups = build_monthly_detail_json()
monthly_detail_json = _json.dumps(_detail_data)
# Build display name map for dropdown
sku_display_map = {"ALL": "All envelope types (combined)"}
# Add combined groups first (so they appear near their individual SKUs)
for gk, gv in _combined_groups.items():
    sku_display_map[gk] = gv["label"]
sku_display_map.update({d[0]: d[1] for d in sku_buffer_data})
sku_dropdown_json = _json.dumps(sku_display_map)

sku_buffer_rows = build_sku_buffer_table()

kpi_var_color = "#EF5350" if net_variance < 0 else "#4CAF79"

# --- Post-settlement derived values (wastage-adjusted) ---
post_adj_used = post_used + total_wastage_allowance
post_adj_variance = post_purchased - post_adj_used
trail12_adj_total = sum(d[7] for d in sku_buffer_data)  # sum of avg_mo (already per-month)
avg_monthly_usage_adj = int(trail12_adj_total) if trail12_adj_total > 0 else avg_monthly_usage
post_var_color = "#4CAF79" if post_adj_variance >= 0 else "#EF5350"
buffer_months = post_adj_variance / avg_monthly_usage_adj if avg_monthly_usage_adj else 0

# Build gauge AFTER adjusted values are computed
inventory_gauge = build_inventory_gauge()
usage_2022 = post_yearly[2022][1] / post_yearly[2022][4] if post_yearly[2022][4] else 0
usage_2025 = post_yearly[2025][1] / post_yearly[2025][4] if post_yearly[2025][4] else 0
usage_decline = (1 - usage_2025 / usage_2022) * 100 if usage_2022 else 0

# Recent 2 years variance for trajectory assessment
recent_yrs = sorted(post_yearly.keys())[-2:]
recent_p = sum(post_yearly[y][0] for y in recent_yrs)
recent_u = sum(post_yearly[y][1] for y in recent_yrs)
recent_var_pct = (recent_p - recent_u) / recent_p * 100 if recent_p else 0

# ---------------------------------------------------------------------------
# Product usage analysis — what mail types drive envelope consumption
# ---------------------------------------------------------------------------
# Top products by volume
product_usage_data = []
total_product_usage = usage_by_product["Total Envelopes Used"].sum()
for _, r in usage_by_product.iterrows():
    vol = safe(r["Total Envelopes Used"])
    pct = vol / total_product_usage * 100 if total_product_usage else 0
    # Determine envelope category
    name = str(r["Product Name"]).upper()
    if any(x in name for x in ("STATEMENT", "EFAIL")):
        env_cat = "N14/9x12"
    elif any(x in name for x in ("1099", "1042", "5498", "TAX")):
        env_cat = "Tax Forms"
    else:
        env_cat = "N10"
    product_usage_data.append((r["Product Name"], vol, pct, env_cat, r.get("First Month", ""), r.get("Last Month", "")))
product_usage_data.sort(key=lambda x: -x[1])

# N10 LTR vs CON purchase split (post-settlement)
n10_ltr_orders = 0
n10_ltr_qty = 0
n10_ltr_cost = 0
n10_con_orders = 0
n10_con_qty = 0
n10_con_cost = 0
n10_ltr_first = None

for _, r in purchase_detail.iterrows():
    desc = str(r.get("Description", "")).upper()
    month = str(r.get("Month", ""))
    dt_check = month_label_to_sortkey(month) if month else (0, 0)
    if dt_check < settlement_key:
        continue
    if "LTR" in desc and "N10" in desc:
        n10_ltr_orders += 1
        n10_ltr_qty += int(safe(r.get("Qty Ordered", r.get("Qty Received", 0))))
        n10_ltr_cost += safe(r.get("Total Cost", 0))
        if n10_ltr_first is None:
            n10_ltr_first = month
    elif any(x in desc for x in ("N10", "CON")) and "N14" not in desc and "9X12" not in desc and "TAX" not in desc and "STMT" not in desc and "LTR" not in desc:
        n10_con_orders += 1
        n10_con_qty += int(safe(r.get("Qty Ordered", r.get("Qty Received", 0))))
        n10_con_cost += safe(r.get("Total Cost", 0))

n10_total_qty = n10_ltr_qty + n10_con_qty
n10_ltr_pct = n10_ltr_qty / n10_total_qty * 100 if n10_total_qty else 0

# Purchase cadence analysis (post-settlement)
from collections import defaultdict as _dd2
_cadence_types = _dd2(list)
for _, r in purchase_detail.iterrows():
    month = str(r.get("Month", ""))
    dt_check = month_label_to_sortkey(month) if month else (0, 0)
    if dt_check < settlement_key:
        continue
    desc = str(r.get("Description", "")).upper()
    qty = int(safe(r.get("Qty Ordered", r.get("Qty Received", 0))))
    if "N14" in desc and "RIDGE" not in desc:
        _cadence_types["N14 Stmt"].append((month, qty))
    elif "LTR" in desc and "N10" in desc:
        _cadence_types["N10 LTR"].append((month, qty))
    elif any(x in desc for x in ("N10", "CON")) and "N14" not in desc and "9X12" not in desc and "STMT" not in desc and "LTR" not in desc:
        _cadence_types["N10 CON"].append((month, qty))
    elif "9X12" in desc and ("PFC" in desc or "ME" in desc or "STMT" in desc) and "RIDGE" not in desc and "DW" not in desc:
        _cadence_types["9x12 Stmt"].append((month, qty))

cadence_data = []
for ctype, entries in sorted(_cadence_types.items(), key=lambda x: -sum(e[1] for e in x[1])):
    total_qty = sum(e[1] for e in entries)
    n_orders = len(entries)
    avg_order = total_qty / n_orders if n_orders else 0
    unique_months = sorted(set(e[0] for e in entries), key=month_label_to_sortkey)
    if len(unique_months) > 1:
        first_sk = month_label_to_sortkey(unique_months[0])
        last_sk = month_label_to_sortkey(unique_months[-1])
        span_months = (last_sk[0] - first_sk[0]) * 12 + (last_sk[1] - first_sk[1])
        avg_gap = span_months / (len(unique_months) - 1) if len(unique_months) > 1 else 0
    else:
        avg_gap = 0
    cadence_data.append((ctype, n_orders, total_qty, avg_order, avg_gap))

# N10 usage split — letters vs confirms
_n10_letter_vol = sum(safe(r["Total Envelopes Used"]) for _, r in usage_by_product.iterrows()
                      if any(x in str(r["Product Name"]).upper() for x in ("LETTER", "CHECK", "DISBURSEMENT")))
_n10_confirm_vol = sum(safe(r["Total Envelopes Used"]) for _, r in usage_by_product.iterrows()
                       if any(x in str(r["Product Name"]).upper() for x in ("MTC", "CONFIRM", "DAILY CONFIRM"))
                       and "STATEMENT" not in str(r["Product Name"]).upper())
_n10_total_usage = _n10_letter_vol + _n10_confirm_vol

# ---------------------------------------------------------------------------
# CSS
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
.bottom-line {
    background: #1E1F23; border-radius: 12px; padding: 24px 28px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.3), 0 2px 12px rgba(0,0,0,0.2);
    border-left: 5px solid #5B9BF7; margin-bottom: 24px;
}
.bottom-line .bl-heading {
    font-size: 13px; font-weight: 700; text-transform: uppercase;
    letter-spacing: 0.5px; color: #5B9BF7; margin: 0 0 8px;
}
.bottom-line p { margin: 0; font-size: 14px; line-height: 1.7; color: #E0E1E6; }
.bottom-line strong { color: #82B4FF; }
.alert-box {
    background: rgba(252, 94, 23, 0.1); border-left: 4px solid #FC5E17;
    border-radius: 0 8px 8px 0; padding: 16px 24px; margin-bottom: 20px;
}
.alert-box p { margin: 0; font-size: 13px; color: #E0E1E6; }
.alert-box strong { color: #FF8A50; }
.info-box {
    background: rgba(91, 155, 247, 0.08); border-left: 4px solid #5B9BF7;
    border-radius: 0 8px 8px 0; padding: 16px 24px; margin-top: 20px;
}
.info-box p { margin: 0; font-size: 13px; color: #E0E1E6; line-height: 1.6; }
.info-box strong { color: #82B4FF; }
.context-line {
    font-size: 13px; color: #9A9BA0; margin: 20px 0 0;
    padding: 12px 16px; background: #1E1F23; border-radius: 8px; line-height: 1.6;
}
.context-line strong { color: #E0E1E6; }
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
    background: rgba(76, 175, 121, 0.15); color: #4CAF79; padding: 3px 10px; border-radius: 12px;
    font-size: 11px; font-weight: 600; text-transform: uppercase;
}
.flag-ok {
    background: #2A2B30; color: #9A9BA0; padding: 3px 10px; border-radius: 12px;
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
    body { font-size: 11px; background: #FFFFFF !important; color: #333 !important; -webkit-print-color-adjust: exact; }
    .header { background: #052390 !important; }
    .kpi-card, .bottom-line, .info-box, .context-line, .table-wrap,
    .alert-box { background: #FFFFFF !important; border-color: #E2E2E2 !important; box-shadow: none !important; }
    .bottom-line p, .info-box p, .alert-box p, .kpi-card .kpi-sub { color: #333 !important; }
    .bottom-line .bl-heading, .kpi-card .kpi-label { color: #052390 !important; }
    .bottom-line strong, .info-box strong { color: #052390 !important; }
    .section-header h2 { color: #052390 !important; }
    table { background: #FFFFFF !important; }
    table td { color: #333 !important; border-bottom-color: #E2E2E2 !important; }
    table th { background: #052390 !important; }
    .subtotal-row { background: #F0F0F5 !important; }
    .total-row { background: #F5F5F7 !important; }
    .context-line { background: #F5F5F7 !important; }
    .context-line strong { color: #333 !important; }
    .footer { color: #6D6E71 !important; border-top-color: #E2E2E2 !important; }
    svg rect[fill="#1E1F23"] { fill: #FFFFFF !important; }
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
html += '    <a href="#envelope-types">Buffer by Type</a>\n'
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

# -- Bottom line with structured recommendations --
# Compute NI savings potential
ni_excess_units = 0
ni_annual_cost = 0
for d in sku_buffer_data:
    if d[0] in retired_foreign_skus and d[11] > 0:
        ni_excess_units += max(0, d[6])  # variance
ni_avg_annual_purchase = sum(
    safe(post_yearly[y][0]) for y in list(sorted(post_yearly.keys()))[-2:]
) / 2
buffer_drawdown_months = max(0, buffer_months - 3)
buffer_drawdown_units = int(buffer_drawdown_months * avg_monthly_usage_adj)
buffer_drawdown_dollars = int(buffer_drawdown_units * post_avg_unit_cost)

html += '        <div class="bottom-line">\n'
html += '            <p class="bl-heading">Bottom line</p>\n'
html += f'            <p>Broadridge is holding <strong>{buffer_months:.1f} months</strong> of envelope buffer stock '
html += f'({fmt_num(post_adj_variance)} envelopes) &mdash; their contract requires 2&ndash;3 months. '
html += f'Total invoiced since settlement: <strong>{fmt_money(post_invoiced)}</strong>. '
html += f'At current run-rate, projected 2026 cost is <strong>{fmt_money(projected_annual_inv)}</strong> '
html += f'(down from {fmt_money(post_yearly[2022][6])} in 2022).</p>\n'
html += '        </div>\n'

# -- Recommendations --
html += '        <div style="background:#1E1F23;border-radius:12px;padding:20px 24px;box-shadow:0 1px 4px rgba(0,0,0,0.3);margin-top:12px;margin-bottom:20px;border:1px solid #2A2B30;">\n'
html += '            <p style="font-size:13px;font-weight:700;text-transform:uppercase;letter-spacing:0.3px;color:#82B4FF;margin:0 0 12px;">Recommended actions</p>\n'
html += '            <ol style="margin:0;padding-left:20px;font-size:14px;line-height:2.0;color:#E0E1E6;">\n'
html += f'                <li><strong>Pause new purchase orders</strong> until buffer stock draws down to 3-month target. Current excess: ~{fmt_num(buffer_drawdown_units)} envelopes ({fmt_money(buffer_drawdown_dollars)}).</li>\n'
html += f'                <li><strong>Stop purchasing retired NI envelopes</strong> where PFC replacements exist. {fmt_money(excess_from_retired)} in excess NI inventory should be consumed before new orders. Broadridge can add indicia to NI stock for domestic use.</li>\n'
html += f'                <li><strong>Request quarterly inventory reconciliation</strong> from Broadridge with physical counts vs. WMS to validate the implied {fmt_num(post_adj_variance)} buffer.</li>\n'
html += f'                <li><strong>Demand usage-based billing</strong> per contract Section 4. Current receipt-based billing has resulted in {fmt_money(billing_excess)} in excess charges since settlement. Request retroactive credit or prospective adjustment.</li>\n'
html += '            </ol>\n'
html += '        </div>\n'

# -- KPI grid: volume + cost --
html += '        <div class="kpi-grid">\n'
html += f'            <div class="kpi-card"><p class="kpi-label">Purchased</p><p class="kpi-value" style="color:#5B9BF7">{fmt_num(post_purchased)}</p><p class="kpi-sub">Mar 2022 &ndash; Dec 2025 ({post_months} mo)</p></div>\n'
html += f'            <div class="kpi-card"><p class="kpi-label">Used + Wastage</p><p class="kpi-value" style="color:#5B9BF7">{fmt_num(post_adj_used)}</p><p class="kpi-sub">{fmt_num(post_used)} used + {fmt_num(total_wastage_allowance)} wastage</p></div>\n'
html += f'            <div class="kpi-card"><p class="kpi-label">Implied Inventory</p><p class="kpi-value" style="color:{post_var_color}">{fmt_num(post_adj_variance)}</p><p class="kpi-sub">{buffer_months:.1f} months of buffer stock</p></div>\n'
html += f'            <div class="kpi-card"><p class="kpi-label">Total Invoiced</p><p class="kpi-value" style="color:#5B9BF7">{fmt_money(post_invoiced)}</p><p class="kpi-sub">Vendor cost: {fmt_money(post_cost)} + markup</p></div>\n'
html += '        </div>\n'

# -- Inventory gauge: actual vs Broadridge 2-3 month policy --
html += '        <div style="background:#1E1F23;border-radius:12px;padding:20px 24px;box-shadow:0 1px 4px rgba(0,0,0,0.3);margin-top:12px;border:1px solid #2A2B30;">\n'
html += '            <p style="font-size:13px;font-weight:600;color:#82B4FF;margin:0 0 8px;text-transform:uppercase;letter-spacing:0.3px;">Buffer stock vs. Broadridge 2&ndash;3 month policy</p>\n'
html += inventory_gauge + '\n'
html += f'            <p style="font-size:12px;color:#9A9BA0;margin:8px 0 0;">Based on trailing 12-month average usage of {fmt_num(avg_monthly_usage_adj)}/month (incl. wastage at contract max). Policy range: {fmt_num(avg_monthly_usage_adj*2)} (2-mo) to {fmt_num(avg_monthly_usage_adj*3)} (3-mo).</p>\n'
html += '        </div>\n'

# -- Wastage discrepancy callout --
# Broadridge admits 10-15% actual wastage; contract allows only 5%/2%
_contract_waste = total_wastage_allowance  # reuse authoritative monthly-level total
_actual_waste_lo = int(post_used * 0.10)
_actual_waste_hi = int(post_used * 0.15)
_excess_lo = _actual_waste_lo - _contract_waste
_excess_hi = _actual_waste_hi - _contract_waste
_excess_cost_lo = int(_excess_lo * post_avg_unit_cost)
_excess_cost_hi = int(_excess_hi * post_avg_unit_cost)

html += '        <div class="bottom-line" style="border-left-color:#EF5350;margin-top:20px;margin-bottom:20px;">\n'
html += '            <p class="bl-heading" style="color:#EF5350;">Wastage discrepancy &mdash; Broadridge exceeds contract limits</p>\n'
html += f'            <p>Broadridge personnel have confirmed in writing that actual envelope wastage runs <strong>10&ndash;15%</strong> '
html += f'(Brandon Koebel, Sep&ndash;Nov 2022). The contract caps the wastage charge at <strong>5%</strong> (original, through Dec 2023) '
html += f'and <strong>2%</strong> (Amendment No. 1, Jan 2024+).</p>\n'
html += f'            <p style="margin-top:8px;">Applied to post-settlement usage of {fmt_num(post_used)} envelopes:</p>\n'
html += '            <table style="margin-top:8px;width:auto;background:transparent;font-size:13px;">\n'
html += '                <tr><td style="border:none;padding:4px 16px 4px 0;color:#9A9BA0;">Contract max wastage (5%/2%)</td>'
html += f'<td style="border:none;padding:4px 0;color:#E0E1E6;font-weight:600;">{fmt_num(_contract_waste)} envelopes</td></tr>\n'
html += '                <tr><td style="border:none;padding:4px 16px 4px 0;color:#9A9BA0;">Actual wastage at 10%</td>'
html += f'<td style="border:none;padding:4px 0;color:#FFA726;font-weight:600;">{fmt_num(_actual_waste_lo)} envelopes</td></tr>\n'
html += '                <tr><td style="border:none;padding:4px 16px 4px 0;color:#9A9BA0;">Actual wastage at 15%</td>'
html += f'<td style="border:none;padding:4px 0;color:#EF5350;font-weight:600;">{fmt_num(_actual_waste_hi)} envelopes</td></tr>\n'
html += '                <tr><td style="border:none;padding:4px 16px 4px 0;color:#9A9BA0;">Excess beyond contract</td>'
html += f'<td style="border:none;padding:4px 0;color:#EF5350;font-weight:600;">{fmt_num(_excess_lo)}&ndash;{fmt_num(_excess_hi)} envelopes ({fmt_money(_excess_cost_lo)}&ndash;{fmt_money(_excess_cost_hi)})</td></tr>\n'
html += '            </table>\n'
html += f'            <p style="margin-top:8px;">This excess wastage is embedded in the &ldquo;Used&rdquo; figure reported by Broadridge &mdash; '
html += f'Apex is billed for wastage at the contract rate, but Broadridge consumes 2&ndash;7.5&times; more than what the contract allows. '
html += f'The cost of excess wastage falls on Broadridge per Section 4.</p>\n'
html += '        </div>\n'

# -- Billing basis discrepancy callout --
html += '        <div class="bottom-line" style="border-left-color:#EF5350;margin-top:20px;margin-bottom:20px;">\n'
html += '            <p class="bl-heading" style="color:#EF5350;">Billing basis discrepancy &mdash; Receipt vs. Usage</p>\n'
html += f'            <p>Both the original contract (Section 4) and Amendment No. 1 state that generic stock (envelopes) '
html += f'shall be billed <strong>&ldquo;based on usage&rdquo;</strong>. However, D17 invoice charges match purchase report totals exactly &mdash; '
html += f'Broadridge is billing on <strong>receipt</strong> (when envelopes are purchased/restocked), not on usage (when envelopes are consumed).</p>\n'
html += f'            <p style="margin-top:8px;">Applied to post-settlement data ({post_months} months):</p>\n'
html += '            <table style="margin-top:8px;width:auto;background:transparent;font-size:13px;">\n'
html += f'                <tr><td style="border:none;padding:4px 16px 4px 0;color:#9A9BA0;">Actual invoiced (receipt-based)</td>'
html += f'<td style="border:none;padding:4px 0;color:#E0E1E6;font-weight:600;">{fmt_money(post_invoiced)}</td></tr>\n'
html += f'                <tr><td style="border:none;padding:4px 16px 4px 0;color:#9A9BA0;">If billed on usage (contract terms)</td>'
html += f'<td style="border:none;padding:4px 0;color:#4CAF79;font-weight:600;">{fmt_money(usage_based_invoice)}</td></tr>\n'
html += f'                <tr><td style="border:none;padding:4px 16px 4px 0;color:#9A9BA0;">Excess charged to Apex</td>'
html += f'<td style="border:none;padding:4px 0;color:#EF5350;font-weight:600;">{fmt_money(billing_excess)}</td></tr>\n'
html += '            </table>\n'
html += f'            <p style="margin-top:8px;">Evidence: {zero_purchase_usage_months} months had zero purchases but significant usage, '
html += f'yet Apex was invoiced $0 for those months. Under usage-based billing, charges would be spread across all months with consumption.</p>\n'
html += '        </div>\n'

# -- Product usage breakdown --
html += '        <div class="bottom-line" style="border-left-color:#5B9BF7;margin-top:20px;margin-bottom:20px;">\n'
html += '            <p class="bl-heading" style="color:#5B9BF7;">Envelope usage by mail product</p>\n'
html += '            <p>Envelope consumption is driven by three products that account for 94% of all usage:</p>\n'
html += '            <table style="margin-top:8px;width:auto;background:transparent;font-size:13px;">\n'
html += '                <tr style="border-bottom:1px solid #3A3B40;"><td style="border:none;padding:6px 20px 6px 0;color:#82B4FF;font-weight:600;">Product</td><td style="border:none;padding:6px 16px;color:#82B4FF;font-weight:600;text-align:right;">Volume</td><td style="border:none;padding:6px 16px;color:#82B4FF;font-weight:600;text-align:right;">Share</td><td style="border:none;padding:6px 0;color:#82B4FF;font-weight:600;">Envelope</td></tr>\n'
for name, vol, pct, env_cat, first, last in product_usage_data[:5]:
    html += f'                <tr><td style="border:none;padding:4px 20px 4px 0;color:#E0E1E6;">{name}</td>'
    html += f'<td style="border:none;padding:4px 16px;color:#E0E1E6;text-align:right;">{fmt_num(vol)}</td>'
    html += f'<td style="border:none;padding:4px 16px;color:#9A9BA0;text-align:right;">{pct:.1f}%</td>'
    html += f'<td style="border:none;padding:4px 0;color:#9A9BA0;">{env_cat}</td></tr>\n'
html += '            </table>\n'
html += f'            <p style="margin-top:10px;"><strong>Key insight:</strong> Address Verification Letters (8.5M) and Confirms (8.1M) both use #10 envelopes, '
html += f'yet only <strong>{fmt_num(n10_ltr_qty)}</strong> of {fmt_num(n10_total_qty)} N10 purchases ({n10_ltr_pct:.2f}%) are the single-window LTR variant. '
html += f'For 4+ years (Jan 2020 &ndash; Apr 2024), all letters used the same double-window confirm envelope.</p>\n'
html += '        </div>\n'

# -- N10 LTR vs CON and purchase cadence --
html += '        <div class="bottom-line" style="border-left-color:#FFA726;margin-top:20px;margin-bottom:20px;">\n'
html += '            <p class="bl-heading" style="color:#FFA726;">Purchase cadence not tracking usage decline</p>\n'
html += '            <table style="margin-top:8px;width:auto;background:transparent;font-size:13px;">\n'
html += '                <tr style="border-bottom:1px solid #3A3B40;"><td style="border:none;padding:6px 20px 6px 0;color:#82B4FF;font-weight:600;">Type</td><td style="border:none;padding:6px 16px;color:#82B4FF;font-weight:600;text-align:right;">Orders</td><td style="border:none;padding:6px 16px;color:#82B4FF;font-weight:600;text-align:right;">Total Qty</td><td style="border:none;padding:6px 16px;color:#82B4FF;font-weight:600;text-align:right;">Avg Order</td><td style="border:none;padding:6px 0;color:#82B4FF;font-weight:600;text-align:right;">Avg Gap</td></tr>\n'
for ctype, n_orders, total_qty, avg_order, avg_gap in cadence_data:
    gap_str = f'{avg_gap:.1f} mo' if avg_gap > 0 else 'N/A'
    html += f'                <tr><td style="border:none;padding:4px 20px 4px 0;color:#E0E1E6;">{ctype}</td>'
    html += f'<td style="border:none;padding:4px 16px;color:#E0E1E6;text-align:right;">{n_orders}</td>'
    html += f'<td style="border:none;padding:4px 16px;color:#E0E1E6;text-align:right;">{fmt_num(total_qty)}</td>'
    html += f'<td style="border:none;padding:4px 16px;color:#9A9BA0;text-align:right;">{fmt_num(avg_order)}</td>'
    html += f'<td style="border:none;padding:4px 0;color:#9A9BA0;text-align:right;">{gap_str}</td></tr>\n'
html += '            </table>\n'
html += f'            <p style="margin-top:10px;">N10 CON envelopes are ordered every <strong>~1 month</strong> at ~188K/order, '
html += f'but usage has dropped from ~300K/mo (2022) to ~200K/mo (2025). The cumulative N10 surplus has grown to <strong>~1.9M envelopes</strong>. '
html += f'Only {n10_ltr_orders} orders for the N10 LTR variant have ever been placed (first: {n10_ltr_first or "N/A"}).</p>\n'
html += '        </div>\n'

# -- Year-by-year table (with cost) --
html += '        <div class="table-wrap" style="margin-top:20px;"><table>\n'
html += '            <thead><tr><th>Year</th><th>Purchased</th><th>Used</th><th>Variance</th><th>Var %</th><th>Invoiced</th><th>Unit Cost</th></tr></thead>\n'
html += '            <tbody>\n'
for yr in sorted(post_yearly.keys()):
    d = post_yearly[yr]
    yp, yu, ym, ys, mc, yc, yi = d
    yv = yp - yu
    vc = var_color(yv)
    vpct = yv / yp if yp else 0
    y_uc = yc / yp if yp else 0
    yr_label = f"{yr} (Mar&ndash;Dec)" if yr == 2022 else str(yr)
    html += f'            <tr><td>{yr_label}</td>'
    html += f'<td class="num">{fmt_num(yp)}</td>'
    html += f'<td class="num">{fmt_num(yu)}</td>'
    html += f'<td class="num" style="color:{vc};font-weight:600">{fmt_num_parens(yv)}</td>'
    html += f'<td class="num" style="color:{vc}">{fmt_pct(vpct)}</td>'
    html += f'<td class="num">{fmt_money(yi)}</td>'
    html += f'<td class="num">${y_uc:.4f}</td></tr>\n'
# Total row
html += f'            <tr class="total-row"><td><strong>Total</strong></td>'
html += f'<td class="num"><strong>{fmt_num(post_purchased)}</strong></td>'
html += f'<td class="num"><strong>{fmt_num(post_used)}</strong></td>'
html += f'<td class="num" style="color:{post_var_color};font-weight:700"><strong>{fmt_num_parens(post_variance)}</strong></td>'
html += f'<td class="num" style="color:{post_var_color};font-weight:700"><strong>{fmt_pct(post_variance/post_purchased if post_purchased else 0)}</strong></td>'
html += f'<td class="num"><strong>{fmt_money(post_invoiced)}</strong></td>'
html += f'<td class="num"><strong>${post_avg_unit_cost:.4f}</strong></td></tr>\n'
html += '            </tbody>\n'
html += '        </table></div>\n'

# -- Forward projection --
html += '        <div style="background:#1E1F23;border-radius:12px;padding:20px 24px;box-shadow:0 1px 4px rgba(0,0,0,0.3);margin-top:20px;border:1px solid #2A2B30;">\n'
html += '            <p style="font-size:13px;font-weight:700;text-transform:uppercase;letter-spacing:0.3px;color:#82B4FF;margin:0 0 12px;">2026 projection (at current run-rate)</p>\n'
html += '            <div class="kpi-grid" style="margin-bottom:0;">\n'
html += f'                <div class="kpi-card" style="padding:16px 20px;"><p class="kpi-label">Projected Usage</p><p class="kpi-value" style="color:#5B9BF7;font-size:22px;">{fmt_num(avg_monthly_usage_adj * 12)}</p><p class="kpi-sub">{fmt_num(avg_monthly_usage_adj)}/mo &times; 12</p></div>\n'
html += f'                <div class="kpi-card" style="padding:16px 20px;"><p class="kpi-label">Projected Cost</p><p class="kpi-value" style="color:#5B9BF7;font-size:22px;">{fmt_money(projected_annual_inv)}</p><p class="kpi-sub">Based on trailing 12-mo invoiced</p></div>\n'
html += f'                <div class="kpi-card" style="padding:16px 20px;"><p class="kpi-label">vs. 2022 Cost</p><p class="kpi-value" style="color:#4CAF79;font-size:22px;">{fmt_money(projected_annual_inv - post_yearly[2022][6])}</p><p class="kpi-sub">{(projected_annual_inv - post_yearly[2022][6]) / post_yearly[2022][6] * 100:+.0f}% from 2022</p></div>\n'
html += f'                <div class="kpi-card" style="padding:16px 20px;"><p class="kpi-label">Buffer Covers</p><p class="kpi-value" style="color:{post_var_color};font-size:22px;">{buffer_months:.1f} mo</p><p class="kpi-sub">Before new purchases needed</p></div>\n'
html += '            </div>\n'
html += '        </div>\n'

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

# ===== BUFFER STOCK BY ENVELOPE TYPE =====
html += '<div class="section" id="envelope-types">\n'
html += '    <div class="section-header" onclick="toggleSection(this)">\n'
html += '        <h2>Buffer stock by envelope type</h2>\n'
html += '        <span class="toggle">&#9660;</span>\n'
html += '    </div>\n'
html += '    <div class="section-body">\n'

# -- Headline callout --
html += f'        <div class="bottom-line" style="border-left-color:#FFA726;margin-bottom:20px;">\n'
html += f'            <p class="bl-heading" style="color:#FFA726;">Where the overstock is &mdash; ${total_excess_dollars:,.0f} in excess inventory</p>\n'
html += f'            <p><strong>${excess_from_retired:,.0f}</strong> ({excess_pct_of_total:.0f}% of the excess) '
html += f'is in <strong>three retired or low-volume foreign mail envelopes</strong> '
html += f'that Broadridge continued purchasing after the Oct 2022 postal permit transition. '

# Find the most egregious SKU for the callout
# Tuple: (sku, display, p, u, w, au, v, avg_mo, buf_mo, lp, uc, excess_dollars, action)
top_sku = max(sku_buffer_data, key=lambda d: d[11])
html += f'The largest exposure is <strong>{top_sku[1]}</strong>: last purchased <strong>{top_sku[9]}</strong>, '
html += f'{top_sku[8]:.0f} months of buffer, ${top_sku[11]:,.0f} excess. '
html += f'Meanwhile, the high-volume domestic envelopes that drive day-to-day operations are near or below the 2&ndash;3 month target.</p>\n'
html += f'        </div>\n'

# -- Per-SKU buffer stock table --
html += '        <div class="table-wrap"><table>\n'
html += '            <thead><tr>'
html += '<th>Envelope Type</th>'
html += '<th>Status</th>'
html += '<th>Purchased</th>'
html += '<th>Used</th>'
html += '<th>Wastage<br><span style="font-size:10px;font-weight:400;text-transform:none;opacity:0.7;">contract max</span></th>'
html += '<th>Adj.<br>Variance</th>'
html += '<th>Buffer<br>(Months)</th>'
html += '<th>Excess $<br><span style="font-size:10px;font-weight:400;text-transform:none;opacity:0.7;">above 3-mo target</span></th>'
html += '<th>Last<br>Purchased</th>'
html += '<th>Action</th>'
html += '</tr></thead>\n'
html += '            <tbody>\n' + sku_buffer_rows + '\n            </tbody>\n'
html += '        </table></div>\n'
html += f'        <p style="font-size:12px;color:#9A9BA0;margin:12px 0 0;">Post-settlement scope (Mar 2022 &ndash; Dec 2025). Wastage applied at contractual max: <strong>5%</strong> (Jan 2019 contract, through Dec 2023) and <strong>2%</strong> (Amendment No. 1, Jan 2024+). Variance = Purchased &minus; Used &minus; Wastage. Buffer months = adjusted variance / trailing 12-month adjusted usage. Excess $ = units above 3-month target &times; avg unit cost.</p>\n'

# -- NI/PFC transition context --
html += '        <div class="info-box" style="margin-top:20px;">\n'
html += '            <p style="font-size:13px;font-weight:700;color:#82B4FF;margin:0 0 10px;">Key context: #10 Confirms + Letters &mdash; NI to PFC transition (Oct 2022)</p>\n'
html += '            <p style="font-size:13px;line-height:1.7;margin:0 0 10px;"><strong>PFC</strong> (Pre-Sorted First-Class) = postage permit pre-printed on the envelope, domestic mail only. <strong>NI</strong> (No Imprint) = no postage printed, used for foreign mail where postage is applied at mailing.</p>\n'
html += '            <p style="font-size:13px;line-height:1.7;margin:0 0 10px;"><strong>Before Oct 2022:</strong> All fold confirms, letters, and checks (domestic + foreign) used ENVCONPFSN10NI.<br>'
html += '            <strong>After Oct 2022:</strong> Domestic mail switched to ENVAPXN10PFSCONN10IND(10/22); foreign mail remained on ENVCONPFSN10NI at ~8K/month.</p>\n'

html += '            <div style="background:#252629;border-left:3px solid #5B9BF7;border-radius:0 8px 8px 0;padding:10px 14px;margin:10px 0;font-size:12px;line-height:1.7;color:#E0E1E6;">\n'
html += '                <strong>Brandon Koebel (Mar 31, 2023):</strong> &ldquo;ENVAPXN10PFSCONN10IND(10/22) is a new revision of the ENVCONPFSN10NI. It was procured in October of 2022 to replace ENVCONPFSN10NI. The new version contains a minor update to the postal permit in order to get a better postage rate.&rdquo;\n'
html += '            </div>\n'

html += '            <div style="background:#252629;border-left:3px solid #5B9BF7;border-radius:0 8px 8px 0;padding:10px 14px;margin:10px 0;font-size:12px;line-height:1.7;color:#E0E1E6;">\n'
html += '                <strong>Brandon Koebel (May 12, 2023):</strong> &ldquo;The NI were not retired, we just use those less frequently (for foreign mail). We had major spikes of letters in the last 3 weeks and already had the NI version stored at the vendor, so we ordered those and added an indicia in order to make our SLA&rsquo;s.&rdquo;\n'
html += '            </div>\n'

html += '            <p style="font-size:13px;line-height:1.7;margin:10px 0 0;"><strong>Impact:</strong> Broadridge purchased 3.6M NI envelopes post-settlement but only used 2.1M &mdash; purchasing continued at near pre-transition volumes despite domestic mail (~193K/mo) shifting to the PFC version, leaving only foreign mail (~8K/mo) on NI. This accounts for $82K of the excess inventory.</p>\n'
html += '        </div>\n'

html += '    </div>\n</div>\n\n'

# ===== MONTHLY DETAIL (interactive, collapsed by default) =====
html += '<div class="section" id="monthly-detail">\n'
html += '    <div class="section-header" onclick="toggleSection(this)">\n'
html += '        <h2>Monthly detail</h2>\n'
html += '        <span class="toggle">&#9654;</span>\n'
html += '    </div>\n'
html += '    <div class="section-body collapsed">\n'

# Filter dropdown
html += '        <div style="margin-bottom:16px;display:flex;align-items:center;gap:12px;">\n'
html += '            <label for="sku-filter" style="font-size:13px;font-weight:600;color:#82B4FF;">Filter by envelope type:</label>\n'
html += '            <select id="sku-filter" onchange="filterMonthlyDetail(this.value)" style="font-size:13px;padding:6px 12px;border:1px solid #3A3B40;border-radius:8px;background:#252629;color:#E0E1E6;min-width:300px;cursor:pointer;">\n'
html += '            </select>\n'
html += '        </div>\n'

# Table shell (populated by JS)
html += '        <div class="table-wrap"><table id="monthly-detail-table">\n'
html += '            <thead><tr>'
html += '<th>Month</th><th>Purchased</th><th>Used</th><th>Wastage</th><th>Adj. Variance</th><th>Variance %</th><th>Running Balance</th>'
html += '</tr></thead>\n'
html += '            <tbody id="monthly-detail-tbody"></tbody>\n'
html += '        </table></div>\n'
html += '        <p style="font-size:12px;color:#9A9BA0;margin:12px 0 0;"><strong>May &amp; Jun 2025:</strong> Zero purchases confirmed (not missing data). Wastage at contractual max (5% pre-2024, 2% post-2024).</p>\n'
html += '    </div>\n</div>\n\n'

# ===== REFERENCE (collapsed by default) =====
html += '<div class="section" id="reference">\n'
html += '    <div class="section-header" onclick="toggleSection(this)">\n'
html += '        <h2>Reference</h2>\n'
html += '        <span class="toggle">&#9654;</span>\n'
html += '    </div>\n'
html += '    <div class="section-body collapsed">\n'

# Contract Language
html += '        <h3 style="color:#82B4FF;font-size:16px;margin:0 0 16px;">Contract terms &mdash; envelope materials</h3>\n'

# Original contract
html += '        <div style="background:#1E1F23;border-radius:12px;padding:20px 24px;box-shadow:0 1px 4px rgba(0,0,0,0.3);margin-bottom:16px;border:1px solid #2A2B30;">\n'
html += '            <p style="font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:0.3px;color:#82B4FF;margin:0 0 8px;">Original Contract &mdash; Section 4, Compensation</p>\n'
html += '            <p style="font-size:12px;color:#9A9BA0;margin:0 0 8px;">GTO Print and Mail Services Schedule, effective January 1, 2019. Signed by William Capuzzi (CEO, Apex) and Joseph Lalli (VP, Broadridge).</p>\n'
html += '            <blockquote style="margin:0;padding:12px 16px;background:#252629;border-left:3px solid #5B9BF7;border-radius:0 8px 8px 0;font-size:13px;line-height:1.7;color:#E0E1E6;">\n'
html += '                &ldquo;Materials (such as paper, envelopes, and inserts) and postage, presort and insert related fees are not included in the Annual Fee and will be charged separately. Materials are billed at cost plus wastage for generic stock. <strong>Specifically, the wastage charge is 10% for any generic paper stock and 5% for generic envelope stock.</strong> For generic stock, the unit rate will be billed based on usage. For Client specific stock, the unit rate will be based on receipt of such stock.&rdquo;\n'
html += '            </blockquote>\n'
html += '        </div>\n'

# Amendment
html += '        <div style="background:#1E1F23;border-radius:12px;padding:20px 24px;box-shadow:0 1px 4px rgba(0,0,0,0.3);margin-bottom:16px;border:1px solid #2A2B30;">\n'
html += '            <p style="font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:0.3px;color:#82B4FF;margin:0 0 8px;">Amendment No. 1 &mdash; Section 4, Compensation (replaces original)</p>\n'
html += '            <p style="font-size:12px;color:#9A9BA0;margin:0 0 8px;">GTO Print and Mail Services Schedule Amendment No. 1, effective January 1, 2024. Signed by William Brennan (CAO, Apex) and Doug Deschutter (Co-President ICS, Broadridge). Term extended through December 31, 2028.</p>\n'
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

# Broadridge confirmation — wastage rates (Denci + Koebel emails)
html += '        <div style="background:#1E1F23;border-radius:12px;padding:20px 24px;box-shadow:0 1px 4px rgba(0,0,0,0.3);margin-bottom:16px;border:1px solid #2A2B30;">\n'
html += '            <p style="font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:0.3px;color:#82B4FF;margin:0 0 8px;">Broadridge Confirmation &mdash; Wastage Rates</p>\n'
html += '            <p style="font-size:12px;color:#9A9BA0;margin:0 0 8px;">Christopher Denci (ICS Account Manager), email to Terry Ray, August 23, 2023.</p>\n'
html += '            <blockquote style="margin:0;padding:12px 16px;background:#252629;border-left:3px solid #5B9BF7;border-radius:0 8px 8px 0;font-size:13px;line-height:1.7;color:#E0E1E6;">\n'
html += '                &ldquo;The current agreement reflects Inventory Cost Plus 10% margin. Materials are billed at cost plus wastage for generic stock. <strong>Specifically, the wastage charge is 10% for any generic paper stock and 5% for generic envelope stock.</strong> Generic stock (envelopes): unit rate billed based on usage. Client-specific stock: unit rate billed based on receipt of such stock.&rdquo;\n'
html += '            </blockquote>\n'
html += '            <p style="font-size:12px;color:#9A9BA0;margin:12px 0 8px;">Brandon Koebel (Sr. Client Relationship Manager), emails Sep&ndash;Nov 2022.</p>\n'
html += '            <blockquote style="margin:0;padding:12px 16px;background:#252629;border-left:3px solid #5B9BF7;border-radius:0 8px 8px 0;font-size:13px;line-height:1.7;color:#E0E1E6;">\n'
html += '                &ldquo;Wastage is roughly 10%&hellip; This includes envelopes that are damaged, need to be reprinted and reinserted, etc.&rdquo; (Nov 7, 2022)<br>\n'
html += '                &ldquo;Did not account for any waste or spoilage (<strong>typically 10&ndash;15%</strong>).&rdquo; (Sep 29, 2022)\n'
html += '            </blockquote>\n'
html += '            <p style="font-size:12px;color:#9A9BA0;margin:8px 0 0;"><strong style="color:#E0E1E6;">Note:</strong> Contractual wastage (billed to Apex) is 5% pre-2024 / 2% post-amendment. Operational wastage (10&ndash;15%) is higher but embedded in the &ldquo;Used&rdquo; figure &mdash; not separately charged.</p>\n'
html += '        </div>\n'

# Summary table
html += '        <div class="table-wrap" style="margin-bottom:24px;"><table>\n'
html += '            <thead><tr><th>Period</th><th>Envelope Wastage</th><th>Margin</th><th>Effective Rate</th><th>Billing Basis</th></tr></thead>\n'
html += '            <tbody>\n'
html += '            <tr><td>Jan 2019 &ndash; Dec 2023</td><td>5%</td><td>&mdash;</td><td>5.0% over vendor</td><td>Usage <span style="color:#EF5350;font-size:11px;">(actual: receipt)</span></td></tr>\n'
html += '            <tr><td>Jan 2024 &ndash; Dec 2028</td><td>2%</td><td>10%</td><td>12.2% over vendor</td><td>Usage <span style="color:#EF5350;font-size:11px;">(actual: receipt)</span></td></tr>\n'
html += '            </tbody>\n'
html += '        </table></div>\n'

# Generic stock classification
html += '        <h3 style="color:#82B4FF;font-size:16px;margin:24px 0 16px;">Generic stock classification</h3>\n'
html += '        <div style="background:#1E1F23;border-radius:12px;padding:20px 24px;box-shadow:0 1px 4px rgba(0,0,0,0.3);margin-bottom:16px;border:1px solid #2A2B30;">\n'
html += '            <p style="font-size:13px;line-height:1.7;color:#E0E1E6;margin:0 0 12px;">Both the original contract and Amendment No. 1 distinguish between <strong>generic stock</strong> (billed on usage, lower wastage) and <strong>client-specific stock</strong> (billed on receipt). 7 of 8 Apex envelope types are provably generic (double-window, no client info on envelope). The 8th (N10 LTR, single-window) represents 0.09% of N10 purchases:</p>\n'
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
html += '            <p style="font-size:12px;color:#9A9BA0;margin:12px 0 0;">7 of 8 envelope types are standard double-window envelopes with no company logos, branding, or custom design. Return and recipient addresses are visible through the windows from the printed content inside. PFC indicia is a functional USPS postage marking, not client branding. NI envelopes are completely blank. The single-window N10 LTR variant (first purchased May 2024, 12,000 units total) represents 0.09% of N10 purchases. Supplier: United Envelope LLC, Mt. Pocono, PA.</p>\n'
html += '        </div>\n'

# Envelope Specifications
html += '        <h3 style="color:#82B4FF;font-size:16px;margin:0 0 12px;">Envelope specifications</h3>\n'
html += '        <p style="font-size:13px;color:#9A9BA0;margin:0 0 16px;">All envelopes are double-window, 24WW paper, black ink with crosshatch black inside tint. Supplier: United Envelope LLC, Mt. Pocono, PA.</p>\n'
html += '        <div class="table-wrap"><table>\n'
html += '            <thead><tr><th>WMS Code</th><th>Mail Type</th><th>Size</th><th>Style</th><th>Postage</th><th>Notes</th></tr></thead>\n'
html += '            <tbody>\n' + envelope_spec_rows + '\n            </tbody>\n'
html += '        </table></div>\n'
html += '        <div style="margin-top:12px;display:flex;gap:12px;flex-wrap:wrap;font-size:12px;color:#9A9BA0;">\n'
html += '            <span><strong>PFC</strong> = Pre-printed First-Class permit (domestic)</span>\n'
html += '            <span><strong>NI</strong> = No Imprint (foreign &mdash; postage applied at mailing)</span>\n'
html += '            <span><strong>DW</strong> = Double Window</span>\n'
html += '            <span><strong>IND</strong> = Individual (Oct 2022 revision)</span>\n'
html += '        </div>\n'

html += '    </div>\n</div>\n\n'

html += '</div><!-- end .content -->\n\n'

# Footer
html += '<div class="footer">\n    Confidential &mdash; Apex Clearing Corporation\n</div>\n'

# JS — base functions
html += f'<script>\n{JS}\n</script>\n'

# JS — monthly detail interactive filter
html += '<script>\n'
html += f'var monthlyData = {monthly_detail_json};\n'
html += f'var skuNames = {sku_dropdown_json};\n'
html += """
(function() {
    // Populate dropdown with optgroups
    var sel = document.getElementById('sku-filter');
    var keys = Object.keys(monthlyData);

    // ALL option first
    var allOpt = document.createElement('option');
    allOpt.value = 'ALL';
    allOpt.textContent = skuNames['ALL'];
    sel.appendChild(allOpt);

    // Combined groups
    var grpKeys = keys.filter(function(k) { return k.indexOf('GRP_') === 0; });
    if (grpKeys.length > 0) {
        var grpGroup = document.createElement('optgroup');
        grpGroup.label = 'Combined (domestic + foreign)';
        grpKeys.sort(function(a, b) {
            return (skuNames[a] || a).localeCompare(skuNames[b] || b);
        });
        grpKeys.forEach(function(key) {
            var opt = document.createElement('option');
            opt.value = key;
            opt.textContent = skuNames[key] || key;
            grpGroup.appendChild(opt);
        });
        sel.appendChild(grpGroup);
    }

    // Individual SKUs
    var skuKeys = keys.filter(function(k) { return k !== 'ALL' && k.indexOf('GRP_') !== 0; });
    skuKeys.sort(function(a, b) {
        return (skuNames[a] || a).localeCompare(skuNames[b] || b);
    });
    var skuGroup = document.createElement('optgroup');
    skuGroup.label = 'Individual SKUs';
    skuKeys.forEach(function(key) {
        var opt = document.createElement('option');
        opt.value = key;
        opt.textContent = skuNames[key] || key;
        skuGroup.appendChild(opt);
    });
    sel.appendChild(skuGroup);

    // Initial render
    filterMonthlyDetail('ALL');
})();

function fmtNum(v) {
    if (v === null || v === undefined) return '\u2014';
    var n = Math.round(v);
    if (n < 0) return '(' + Math.abs(n).toLocaleString() + ')';
    return n.toLocaleString();
}
function fmtPct(v) {
    if (v === null || v === undefined || v === 0) return '\u2014';
    return (v * 100).toFixed(1) + '%';
}
function varColor(v) {
    if (v > 0) return '#4CAF79';
    if (v < 0) return '#EF5350';
    return '#9A9BA0';
}

function filterMonthlyDetail(sku) {
    var rows = monthlyData[sku];
    if (!rows) return;
    var tbody = document.getElementById('monthly-detail-tbody');
    tbody.innerHTML = '';

    var runBal = 0;
    var prevYear = null;
    var yrP = 0, yrU = 0, yrW = 0;

    function getYear(m) {
        return 2000 + parseInt(m.split('-')[1]);
    }

    function addSubtotal(year, yp, yu, yw, rb) {
        var yv = yp - yu - yw;
        var vpct = yp ? yv / yp : 0;
        var vc = varColor(yv);
        var rbc = varColor(rb);
        var yrLabel = (year === 2022) ? '2022 (Mar\u2013Dec) Total' : year + ' Total';
        var tr = document.createElement('tr');
        tr.className = 'subtotal-row';
        tr.innerHTML =
            '<td><strong>' + yrLabel + '</strong></td>' +
            '<td class="num"><strong>' + fmtNum(yp) + '</strong></td>' +
            '<td class="num"><strong>' + fmtNum(yu) + '</strong></td>' +
            '<td class="num" style="color:#9A9BA0"><strong>' + fmtNum(yw) + '</strong></td>' +
            '<td class="num" style="color:' + vc + ';font-weight:700"><strong>' + fmtNum(yv) + '</strong></td>' +
            '<td class="num" style="color:' + vc + ';font-weight:700"><strong>' + fmtPct(vpct) + '</strong></td>' +
            '<td class="num" style="color:' + rbc + ';font-weight:700"><strong>' + fmtNum(rb) + '</strong></td>';
        tbody.appendChild(tr);
    }

    var totalP = 0, totalU = 0, totalW = 0;

    for (var i = 0; i < rows.length; i++) {
        var r = rows[i];
        var curYear = getYear(r.m);

        if (prevYear !== null && curYear !== prevYear) {
            addSubtotal(prevYear, yrP, yrU, yrW, runBal);
            yrP = 0; yrU = 0; yrW = 0;
        }
        prevYear = curYear;

        var v = r.v;
        runBal += v;
        yrP += r.p; yrU += r.u; yrW += r.w;
        totalP += r.p; totalU += r.u; totalW += r.w;

        var vc = varColor(v);
        var rbc = varColor(runBal);
        var bg = (i % 2 === 1) ? '#252629' : '#1E1F23';
        var tr = document.createElement('tr');
        tr.style.background = bg;
        tr.innerHTML =
            '<td>' + r.m + '</td>' +
            '<td class="num">' + fmtNum(r.p) + '</td>' +
            '<td class="num">' + fmtNum(r.u) + '</td>' +
            '<td class="num" style="color:#9A9BA0">' + fmtNum(r.w) + '</td>' +
            '<td class="num" style="color:' + vc + ';font-weight:600">' + fmtNum(v) + '</td>' +
            '<td class="num" style="color:' + vc + '">' + fmtPct(r.vp) + '</td>' +
            '<td class="num" style="color:' + rbc + ';font-weight:600">' + fmtNum(runBal) + '</td>';
        tbody.appendChild(tr);
    }

    // Final year subtotal
    if (prevYear !== null) {
        addSubtotal(prevYear, yrP, yrU, yrW, runBal);
    }

    // Grand total
    var gv = totalP - totalU - totalW;
    var gvc = varColor(gv);
    var gpct = totalP ? gv / totalP : 0;
    var tr = document.createElement('tr');
    tr.className = 'total-row';
    tr.innerHTML =
        '<td><strong>Grand Total</strong></td>' +
        '<td class="num"><strong>' + fmtNum(totalP) + '</strong></td>' +
        '<td class="num"><strong>' + fmtNum(totalU) + '</strong></td>' +
        '<td class="num" style="color:#9A9BA0"><strong>' + fmtNum(totalW) + '</strong></td>' +
        '<td class="num" style="color:' + gvc + ';font-weight:700"><strong>' + fmtNum(gv) + '</strong></td>' +
        '<td class="num" style="color:' + gvc + ';font-weight:700"><strong>' + fmtPct(gpct) + '</strong></td>' +
        '<td class="num" style="color:' + gvc + ';font-weight:700"><strong>' + fmtNum(runBal) + '</strong></td>';
    tbody.appendChild(tr);
}
"""
html += '</script>\n'
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
