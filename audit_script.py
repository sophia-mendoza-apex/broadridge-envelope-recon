import pandas as pd
import numpy as np

pd.set_option('display.max_rows', 200)
pd.set_option('display.max_columns', 20)
pd.set_option('display.width', 200)

FILE = 'C:/Users/smendoza/Projects/Broadridge Envelopes/P&M Postage and Material Recon.xlsx'
env = pd.read_excel(FILE, sheet_name='Envelopes Purchased')
po = pd.read_excel(FILE, sheet_name='Purchase Orders')
vol = pd.read_excel(FILE, sheet_name='Volume Data')

env_nz = env[env['Receipt Amount'] > 0].copy()
env_nz['Period'] = np.where(env_nz['Month'] < '2024-01-01', 'Period 1', 'Period 2')
env_nz['Actual_Markup_Pct'] = env_nz['+10% Mark up'] / env_nz['Receipt Amount'] * 100

def fd(v):
    return f'${v:>14,.2f}'

def fp(v):
    return f'{v:.1f}%'

SEP = '========================================================================================================================'
DASH = '--------------------------------------------------------------------------------------------------------------'

print(SEP)
print('BROADRIDGE ENVELOPE CONTRACT COMPLIANCE AUDIT')
print('Audit Date: 2026-02-20 | Scope: March 2022 - December 2025 (full history for context)')
print(SEP)

print()
print(SEP)
print('1. MARKUP VALIDATION')
print(SEP)
print()
print('CONTRACT TERMS:')
print('  Period 1 (Jan 2019 - Dec 2023): Materials at cost + 5% wastage for generic envelopes.')
print('                                  NO separate margin or markup mentioned in original contract.')
print('  Period 2 (Jan 2024 - Dec 2028): Materials at vendor price + 2% wastage + 10% margin.')
print()
print('FINDING: Every non-zero row applies EXACTLY 10% markup on Receipt Amount,')
print('         regardless of contract period.')
print()

total_rows = len(env_nz)
rows_10 = len(env_nz[env_nz['Actual_Markup_Pct'].round(2) == 10.0])
rows_other = total_rows - rows_10
print(f'Total non-zero rows analyzed:     {total_rows}')
print(f'Rows with exactly 10.0% markup:   {rows_10}')
print(f'Rows with different markup:        {rows_other}')

p1 = env_nz[env_nz['Period'] == 'Period 1']
p2 = env_nz[env_nz['Period'] == 'Period 2']

for label, df in [('Period 1 (Pre-Amendment: Jan 2019 - Dec 2023)', p1), ('Period 2 (Post-Amendment: Jan 2024 - Dec 2025)', p2)]:
    receipt = df['Receipt Amount'].sum()
    markup = df['+10% Mark up'].sum()
    invoiced = df['Total Invoiced'].sum()
    pct = markup / receipt * 100
    print()
    print('--- ' + label + ' ---')
    print(f'  Rows with activity:   {len(df)}')
    print(f'  Total Receipt Amount: {fd(receipt)}')
    print(f'  Total Markup Charged: {fd(markup)}')
    print(f'  Total Invoiced:       {fd(invoiced)}')
    print(f'  Markup Rate Applied:  {fp(pct)}')

print()
print(SEP)
print('2. PRE-AMENDMENT OVERCHARGE ANALYSIS (Mar 2022 - Dec 2023)')
print(SEP)
print()
print('CONTRACT ISSUE:')
print('  Original contract: 5% wastage for generic envelope stock, no separate margin.')
print('  Actual billing:    10% markup applied to Receipt Amount in every row.')
print('  Potential overcharge = 10% charged - 5% contractual = 5% of Receipt Amount.')
print()

scope = env_nz[(env_nz['Month'] >= '2022-03-01') & (env_nz['Month'] < '2024-01-01')].copy()
scope['Contract_5pct'] = scope['Receipt Amount'] * 0.05
scope['Contract_Total'] = scope['Receipt Amount'] + scope['Contract_5pct']
scope['Overcharge'] = scope['+10% Mark up'] - scope['Contract_5pct']

monthly = scope.groupby(scope['Month'].dt.to_period('M')).agg(
    Receipt_Amt=('Receipt Amount', 'sum'),
    Markup_10pct=('+10% Mark up', 'sum'),
    Contract_5pct=('Contract_5pct', 'sum'),
    Overcharge=('Overcharge', 'sum'),
    Total_Invoiced=('Total Invoiced', 'sum'),
    Contract_Total=('Contract_Total', 'sum')
).reset_index()

hdr = f"{'Month':<10} {'Receipt Amt':>14} {'10% Charged':>14} {'5% Contract':>14} {'OVERCHARGE':>14} {'Invoiced':>16} {'Should Be':>16}"
print(hdr)
print(DASH)

for _, r in monthly.iterrows():
    m = str(r['Month'])
    print(f"{m:<10} {fd(r['Receipt_Amt'])} {fd(r['Markup_10pct'])} {fd(r['Contract_5pct'])} {fd(r['Overcharge'])} {fd(r['Total_Invoiced'])}   {fd(r['Contract_Total'])}")

print(DASH)
tot_r = scope['Receipt Amount'].sum()
tot_m = scope['+10% Mark up'].sum()
tot_c = scope['Contract_5pct'].sum()
tot_o = scope['Overcharge'].sum()
tot_i = scope['Total Invoiced'].sum()
tot_ct = scope['Contract_Total'].sum()
print(f"{'TOTAL':<10} {fd(tot_r)} {fd(tot_m)} {fd(tot_c)} {fd(tot_o)} {fd(tot_i)}   {fd(tot_ct)}")

print()
print(f'  >>> TOTAL POTENTIAL OVERCHARGE (Mar 2022 - Dec 2023): {fd(tot_o)}')
print(f'  >>> This is 5% of the {fd(tot_r)} Receipt Amount in this period')

print()
print(SEP)
print('3. POST-AMENDMENT PRICING VALIDATION (Jan 2024 - Dec 2025)')
print(SEP)
print()
print('CONTRACT TERMS (Amendment No. 1):')
print('  Inventory Cost = Vendor Price + 2% Wastage')
print('  Billing Amount = Inventory Cost + 10% Margin')
print('  Therefore: Total = Vendor Price x 1.02 x 1.10 = Vendor Price x 1.122')
print()
print('ANALYSIS:')
print('  If Receipt Amount = Vendor Price (base cost from supplier):')
print('    Method A (contractual compound): Receipt x 1.122   (12.2% total markup)')
print('    Method B (simple addition):      Receipt x 1.12    (12.0% total markup)')
print('    Method C (what is billed):       Receipt x 1.10    (10.0% total markup)')
print()

post = env_nz[(env_nz['Month'] >= '2024-01-01')].copy()
post['Method_A'] = post['Receipt Amount'] * 0.122
post['Method_B'] = post['Receipt Amount'] * 0.12
post['Actual'] = post['+10% Mark up']

tr = post['Receipt Amount'].sum()
ta = post['Actual'].sum()
tA = post['Method_A'].sum()
tB = post['Method_B'].sum()

print(f'Post-Amendment Total Receipt Amount:  {fd(tr)}')
print(f'Actual 10% Markup Charged:            {fd(ta)}')
print(f'Method A (compound 12.2% markup):     {fd(tA)}')
print(f'Method B (simple 12.0% markup):       {fd(tB)}')
print()
print(f'If Method A correct: Broadridge UNDERCHARGES by {fd(tA - ta)}')
print(f'If Method B correct: Broadridge UNDERCHARGES by {fd(tB - ta)}')
print()
print('INTERPRETATION:')
print('  Broadridge charges a flat 10% markup and does NOT separately add 2% wastage.')
print('  Two possible explanations:')
print('    (a) Receipt Amount already includes 2% wastage (Receipt = Vendor Price x 1.02)')
print('        and the 10% is correctly applied on top => compliant with the contract')
print('    (b) Receipt Amount = raw vendor price, and the 2% wastage is not being charged,')
print('        which would be a billing error in Apex favor')
print()
print('  RECOMMENDATION: Request vendor invoices for sample POs to verify.')

print()
print('  POST-AMENDMENT MONTHLY DETAIL:')
pm = post.groupby(post['Month'].dt.to_period('M')).agg(
    Receipt_Amt=('Receipt Amount', 'sum'),
    Actual=('+10% Mark up', 'sum'),
    Should_A=('Method_A', 'sum')
).reset_index()

hdr2 = f"  {'Month':<10} {'Receipt Amt':>14} {'10% Charged':>14} {'12.2% Should Be':>16} {'Difference':>14}"
print(hdr2)
print('  ' + '------------------------------------------------------------------------')
for _, r in pm.iterrows():
    diff = r['Actual'] - r['Should_A']
    print(f"  {str(r['Month']):<10} {fd(r['Receipt_Amt'])} {fd(r['Actual'])} {fd(r['Should_A'])}   {fd(diff)}")
print('  ' + '------------------------------------------------------------------------')
print(f"  {'TOTAL':<10} {fd(tr)} {fd(ta)} {fd(tA)}   {fd(ta - tA)}")

print()
print(SEP)
print('4. USAGE vs RECEIPT BILLING ANALYSIS')
print(SEP)
print()
print('CONTRACT TERMS:')
print('  Generic stock: billed based on USAGE (monthly consumption)')
print('  Client-specific stock: billed based on RECEIPT (when delivered to Broadridge)')
print()

monthly_purchases = env_nz.groupby(env_nz['Month'].dt.to_period('M'))['Quantity Purchased'].sum().reset_index()
monthly_purchases.columns = ['Month', 'Qty_Purchased']

monthly_usage = vol.groupby(vol['Month'].dt.to_period('M'))['Envelopes'].sum().reset_index()
monthly_usage.columns = ['Month', 'Qty_Used']

comparison = pd.merge(monthly_purchases, monthly_usage, on='Month', how='outer').fillna(0)
comparison = comparison.sort_values('Month')
comparison['Diff'] = comparison['Qty_Purchased'] - comparison['Qty_Used']
comparison['Ratio'] = np.where(comparison['Qty_Used'] > 0, comparison['Qty_Purchased'] / comparison['Qty_Used'], np.nan)

comp_scope = comparison[comparison['Month'] >= '2022-03']

hdr3 = f"{'Month':<10} {'Qty Purchased':>16} {'Qty Used':>14} {'Difference':>14} {'Ratio':>8} {'Flag':>8}"
print(hdr3)
print('---------------------------------------------------------------------------')
for _, r in comp_scope.iterrows():
    flag = ''
    if not np.isnan(r['Ratio']):
        if r['Ratio'] > 1.5:
            flag = 'HIGH'
        elif r['Ratio'] < 0.5:
            flag = 'LOW'
        ratio_str = f"{r['Ratio']:.2f}"
    else:
        ratio_str = 'N/A'
    print(f"{str(r['Month']):<10} {int(r['Qty_Purchased']):>16,} {int(r['Qty_Used']):>14,} {int(r['Diff']):>14,} {ratio_str:>8} {flag:>8}")

print()
print('INTERPRETATION:')
print('  Purchases are lumpy (bulk PO-based), while usage is relatively steady.')
print('  Months flagged HIGH/LOW show significant divergence between purchases and usage.')
print('  The contract requires generic stock billing based on USAGE.')
print('  Broadridge should reconcile monthly billing to actual envelope consumption.')

print()
print(SEP)
print('5. UNIT PRICE CONSISTENCY & PRICE CHANGE TRACKING')
print(SEP)
print()
print('Tracking vendor price per envelope over time from Purchase Orders.')
print('Contract allows up to 4% annual CPI increase on fees after Jan 2025.')
print()

po_nz = po[po['Receipt Amount'] > 0].copy()
po_nz['vendor_per_unit'] = po_nz['Receipt Amount'] / po_nz['Quantity*1000']

for desc in sorted(po_nz['Item Description'].unique()):
    subset = po_nz[po_nz['Item Description'] == desc].sort_values('Month')
    print(f'--- {desc} ---')
    prices = []
    last_price = None
    for _, r in subset.iterrows():
        price = round(r['vendor_per_unit'], 6)
        if price != last_price:
            change_pct = ((price / last_price) - 1) * 100 if last_price else 0
            prices.append({'Month': r['Month'].strftime('%Y-%m'), 'Price': price, 'Chg': change_pct})
            last_price = price
    print(f"  {'Effective':>12} {'Vendor $/Unit':>18} {'Change':>10}")
    for p in prices:
        flag = ' ** EXCEEDS 4% CPI' if p['Chg'] > 4 else ''
        print(f"  {p['Month']:>12}   ${p['Price']:>14.6f}  {p['Chg']:>+8.2f}%{flag}")
    post25 = [p for p in prices if p['Month'] >= '2025-01' and p['Chg'] > 0]
    if post25:
        for p in post25:
            if p['Chg'] > 4:
                print(f"  ** WARNING: {p['Chg']:.1f}% increase in {p['Month']} exceeds 4% CPI cap")
            else:
                print(f"     {p['Chg']:.1f}% increase in {p['Month']} is within 4% CPI cap")
    print()

print()
print(SEP)
print('6. SUMMARY FINANCIAL IMPACT')
print(SEP)

p1_scope = env_nz[(env_nz['Month'] >= '2022-03-01') & (env_nz['Month'] < '2024-01-01')]
p1_overcharge = p1_scope['Receipt Amount'].sum() * 0.05

p1_all = env_nz[env_nz['Month'] < '2024-01-01']
p1_all_overcharge = p1_all['Receipt Amount'].sum() * 0.05

p2_scope = env_nz[env_nz['Month'] >= '2024-01-01']
p2_undercharge = p2_scope['Receipt Amount'].sum() * 0.022

print()
print('+------------------------------------------------------------------------------+')
print('|                           FINANCIAL IMPACT SUMMARY                           |')
print('+------------------------------------------------------------------------------+')
print()
print('  PERIOD 1 - PRE-AMENDMENT OVERCHARGE (if 5% wastage is the correct markup):')
print()
s1r = p1_scope['Receipt Amount'].sum()
s1m = p1_scope['+10% Mark up'].sum()
s1c = s1r * 0.05
print('    Scope: Mar 2022 - Dec 2023 (audit scope)')
print(f'      Total Receipt Amount:           {fd(s1r)}')
print(f'      Markup Charged (10%):           {fd(s1m)}')
print(f'      Contract Markup (5%):           {fd(s1c)}')
print(f'      POTENTIAL OVERCHARGE:           {fd(p1_overcharge)}')
print()
s1ar = p1_all['Receipt Amount'].sum()
s1am = p1_all['+10% Mark up'].sum()
s1ac = s1ar * 0.05
print('    Extended: All of Period 1 (2019-2023) with activity')
print(f'      Total Receipt Amount:           {fd(s1ar)}')
print(f'      Markup Charged (10%):           {fd(s1am)}')
print(f'      Contract Markup (5%):           {fd(s1ac)}')
print(f'      POTENTIAL OVERCHARGE:           {fd(p1_all_overcharge)}')
print()
print('  PERIOD 2 - POST-AMENDMENT ASSESSMENT:')
print()
s2r = p2_scope['Receipt Amount'].sum()
s2m = p2_scope['+10% Mark up'].sum()
s2c = s2r * 0.122
print('    Scope: Jan 2024 - Dec 2025')
print(f'      Total Receipt Amount:           {fd(s2r)}')
print(f'      Markup Charged (10%):           {fd(s2m)}')
print(f'      Contractual (12.2% compound):   {fd(s2c)}')
print(f'      POTENTIAL UNDERCHARGE:          {fd(p2_undercharge)}')
print('      (Only if Receipt Amt = raw vendor price, not vendor + 2% wastage)')
print()
print('  NET EXPOSURE:')
print()
print(f'    Pre-Amend Overcharge (Mar 2022 - Dec 2023):  {fd(p1_overcharge)}  (overcharge to Apex)')
print(f'    Post-Amend Potential Undercharge:            -{fd(p2_undercharge)}  (favorable to Apex)')
net = p1_overcharge - p2_undercharge
print(f'    NET (audit scope, Mar 2022 - Dec 2025):      {fd(net)}')
print()
print(f'    Pre-Amend Overcharge (full Period 1):        {fd(p1_all_overcharge)}  (overcharge to Apex)')

print()
print('+------------------------------------------------------------------------------+')
print('|                        KEY FINDINGS & RECOMMENDATIONS                        |')
print('+------------------------------------------------------------------------------+')
print()
print('  1. CRITICAL: During the entire pre-amendment period (2019-2023), Broadridge')
print('     applied a 10% markup despite the contract specifying only 5% wastage for')
print('     generic envelope stock with no separate margin. The spreadsheet column')
print('     itself is labeled "+10% Mark up", confirming this was systematic.')
print()
print('  2. The 10% markup applies uniformly to ALL envelope types in ALL months,')
print('     with zero variance. This is systematic, not incidental.')
print()
print('  3. Post-amendment (Jan 2024+), the 10% markup aligns with the new contract')
print('     term for margin, but the 2% wastage component appears absent unless it')
print('     is embedded in the Receipt Amount (vendor price + 2%).')
print()
print('  4. Purchase quantities are lumpy (bulk PO-based), while usage is relatively')
print('     steady. Billing should reconcile to USAGE for generic stock per contract.')
print()
print('  5. Price increases from PO data are generally gradual and consistent with')
print('     market conditions. No single increase clearly violates a 4% CPI cap.')
print()
print('  RECOMMENDED ACTIONS:')
print('  (a) Request vendor invoices for 3-5 POs to verify whether Receipt Amount')
print('      includes or excludes the 2% wastage factor.')
print('  (b) Request Broadridge to explain the contractual basis for the 10% markup')
print('      applied during Period 1 (Jan 2019 - Dec 2023).')
print(f'  (c) If Period 1 markup was non-contractual, negotiate a credit of')
print(f'      approximately {fd(p1_overcharge)} for the Mar 2022 - Dec 2023 scope,')
print(f'      or {fd(p1_all_overcharge)} for the full Period 1.')
print('  (d) Confirm with Broadridge whether generic envelope billing is based on')
print('      monthly usage or purchase receipt, and request reconciliation support.')
print()
print(SEP)
print('END OF AUDIT REPORT')
print(SEP)