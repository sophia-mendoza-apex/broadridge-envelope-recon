"""Q1 2026 Usage-Based Envelope Billing Computation"""

jan_qty = 514500
jan_cost = 42241.05
jan_rate = jan_cost / jan_qty
jan_usage = 336248
jan_d17 = 46465.16

feb_usage = 314555
feb_rate = jan_rate  # carry forward (no purchases in Feb)
feb_d17 = 19350.38

mar_qty = 516000
mar_cost = 38589.48
mar_rate = mar_cost / mar_qty
mar_usage = 276123
mar_d17 = 42448.43

jan_ub = jan_usage * jan_rate * 1.02 * 1.10
feb_ub = feb_usage * feb_rate * 1.02 * 1.10
mar_ub = mar_usage * mar_rate * 1.02 * 1.10

total_invoiced = jan_d17 + feb_d17 + mar_d17
total_ub = jan_ub + feb_ub + mar_ub
total_usage = jan_usage + feb_usage + mar_usage
total_purchased = jan_qty + 0 + mar_qty

print("=" * 70)
print("FINAL RECONCILIATION")
print("=" * 70)
print()
print("926,926 = Denci Billed Qty (8,685,680) minus our Jan24-Dec25 (7,758,754)")
print("Per-SKU verification: 922,016 + 4,910 (LTR) = 926,926 -- CONFIRMED")
print()
print("This IS the billing sheet (billed) quantity for Q1 2026.")
print("It is NOT production usage (which would be higher by ~9.2%).")
print()
print("Therefore:")
print("  Jan billed (billing sheet verified): 336,248")
print("  Mar billed (billing sheet verified): 276,123")
print("  Feb billed (implied): 926,926 - 336,248 - 276,123 = 314,555")
print()
print("The 470,141 figure is NOT the billing sheet Jan figure.")
print("Both independent sources (Materials tab + Volume Data tab) confirm 336,248.")
print()

print("=" * 70)
print("Q1 2026 USAGE-BASED BILLING -- DEFINITIVE COMPUTATION")
print("=" * 70)
print()
print("INPUTS:")
print()
print("Month     Usage      Vendor Rate    Purchases   Purchase Cost")
print("-" * 70)
print(f"Jan      {jan_usage:>9,}   ${jan_rate:.6f}/env   {jan_qty:>9,}   ${jan_cost:>12,.2f}")
print(f"Feb      {feb_usage:>9,}   ${feb_rate:.6f}/env   {0:>9,}   ${0:>12,.2f}")
print(f"Mar      {mar_usage:>9,}   ${mar_rate:.6f}/env   {mar_qty:>9,}   ${mar_cost:>12,.2f}")
print(f"Total    {total_usage:>9,}                      {total_purchased:>9,}   ${jan_cost+mar_cost:>12,.2f}")
print()

print("COMPUTATION (Usage x Rate x 1.02 x 1.10):")
print()
print(f"Jan: {jan_usage:,} x ${jan_rate:.6f} x 1.02 x 1.10")
print(f"   = ${jan_usage * jan_rate:,.2f} x 1.02 x 1.10")
print(f"   = ${jan_usage * jan_rate * 1.02:,.2f} x 1.10")
print(f"   = ${jan_ub:,.2f}")
print()
print(f"Feb: {feb_usage:,} x ${feb_rate:.6f} x 1.02 x 1.10")
print(f"   = ${feb_usage * feb_rate:,.2f} x 1.02 x 1.10")
print(f"   = ${feb_usage * feb_rate * 1.02:,.2f} x 1.10")
print(f"   = ${feb_ub:,.2f}")
print()
print(f"Mar: {mar_usage:,} x ${mar_rate:.6f} x 1.02 x 1.10")
print(f"   = ${mar_usage * mar_rate:,.2f} x 1.02 x 1.10")
print(f"   = ${mar_usage * mar_rate * 1.02:,.2f} x 1.10")
print(f"   = ${mar_ub:,.2f}")
print()

print("RESULTS:")
print()
print(f"{'Month':<10} {'Invoiced':>14} {'Usage-Based':>14} {'Difference':>14}")
print(f"{'-'*10:<10} {'-'*14:>14} {'-'*14:>14} {'-'*14:>14}")

for label, inv, ub in [("January", jan_d17, jan_ub), ("February", feb_d17, feb_ub), ("March", mar_d17, mar_ub)]:
    diff = inv - ub
    print(f"{label:<10} ${inv:>12,.2f}  ${ub:>12,.2f}  ${diff:>12,.2f}")

diff_total = total_invoiced - total_ub
print(f"{'-'*10:<10} {'-'*14:>14} {'-'*14:>14} {'-'*14:>14}")
print(f"{'Q1 TOTAL':<10} ${total_invoiced:>12,.2f}  ${total_ub:>12,.2f}  ${diff_total:>12,.2f}")
print()

print(f"Q1 2026 Overcharge: ${diff_total:,.2f} ({diff_total/total_invoiced*100:.1f}% of invoiced)")
print()

print("NOTE ON FEBRUARY:")
print(f"  Feb usage-based (${feb_ub:,.2f}) EXCEEDS Feb D17 invoice (${feb_d17:,.2f}).")
print("  This is expected: in Feb, there were ZERO purchases (no envelope receipts),")
print("  so the receipt-based invoice was low. But Apex still used 314,555 envelopes.")
print("  Under generic billing, Apex would pay for usage regardless of purchase timing.")
print("  The aggregate Q1 overcharge still holds because total purchases (1,030,500)")
print(f"  exceed total usage ({total_usage:,}) by {total_purchased - total_usage:,} envelopes.")
print()

print("=" * 70)
print("CUMULATIVE IMPACT (Mar 2022 - Mar 2026)")
print("=" * 70)
print()
print("From Rectification Summary (Mar 2022 - Dec 2025):")
print("  Invoiced (Receipt):  $1,575,143")
print("  Correct (Usage):     $1,385,450")
print("  Overcharge:            $189,693")
print()
print("Adding Q1 2026:")
cumulative_invoiced = 1575143 + total_invoiced
cumulative_ub = 1385450 + total_ub
cumulative_overcharge = cumulative_invoiced - cumulative_ub
print(f"  Total Invoiced:      ${cumulative_invoiced:,.2f}")
print(f"  Total Usage-Based:   ${cumulative_ub:,.2f}")
print(f"  Total Overcharge:    ${cumulative_overcharge:,.2f}")
print(f"  Overcharge %:        {cumulative_overcharge/cumulative_invoiced*100:.1f}%")
