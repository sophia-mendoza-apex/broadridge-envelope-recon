import fitz, re, os, glob

BASE = r"C:\Users\smendoza\OneDrive - Apex Clearing\Broadridge Billing\Broadridge Invoices"
MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
folders = {
 2023: {str(i):f"{i}-2023" for i in range(1,13)},
 2024: {str(i):f"{i:02d}-2024" for i in range(1,13)},
 2025: {"1":"1-January","2":"2-February","3":"3-March","4":"4-April","5":"5-May","6":"6-June",
        "7":"7-July","8":"8-August","9":"9-September","10":"10-October","11":"11-November","12":"12-December"},
 2026: {"1":"1 - January","2":"2 - February","3":"3 - March"},
}

def find_pdf(year, mnum):
    d = os.path.join(BASE, str(year), folders[year][str(mnum)])
    mon = MONTHS[mnum-1]
    cands = glob.glob(os.path.join(d, f"{mon}*C10_D17*.pdf"))
    if not cands: return None
    rev = [c for c in cands if re.search(r"revis", os.path.basename(c), re.I)]
    if rev: return rev[0]
    cands = [c for c in cands if "(2)" not in c]
    cands.sort(key=lambda p: len(os.path.basename(p)))
    return cands[0]

pat = re.compile(r"\$\s*([\d,]+\.\d+)\s*\n(CONFIRM|LETTER|STATEMENT)\s+PAPER\s*-\s*PER\s*PAGE\s*\n\s*([\d.]+)\s*\n\s*([\d,]+)")

def extract(path):
    doc = fitz.open(path); txt = "".join(p.get_text() for p in doc); doc.close()
    rows = {}
    for m in pat.finditer(txt):
        amt, kind, rate, qty = m.groups()
        rows[kind] = (qty, rate, amt)
    return rows

print(f"{'Month':<10} {'CONFIRM':>9} {'LETTER':>9} {'STATEMENT':>10}   file")
print("-"*72)
prev = None
for year in [2023,2024,2025,2026]:
    for mnum in range(1,13):
        if str(mnum) not in folders[year]: continue
        p = find_pdf(year, mnum)
        label = f"{MONTHS[mnum-1]} {year}"
        if not p:
            print(f"{label:<10} {'NOT FOUND':>9}"); continue
        r = extract(p)
        c = r.get("CONFIRM",("","-",""))[1]
        l = r.get("LETTER",("","-",""))[1]
        s = r.get("STATEMENT",("","-",""))[1]
        flag = ""
        cur = (c,l,s)
        if prev is not None and cur != prev and all(x!="-" for x in cur):
            flag = "  <-- CHANGE"
        if all(x!="-" for x in cur): prev = cur
        print(f"{label:<10} {c:>9} {l:>9} {s:>10}   {os.path.basename(p)}{flag}")

print()
print("EXACT RATE (amount / qty), STATEMENT line, at each change point:")
print("-"*60)
checks = [(2023,1),(2023,5),(2024,1),(2025,1),(2026,1)]
for year,mnum in checks:
    p = find_pdf(year,mnum)
    r = extract(p)
    for kind in ["STATEMENT","CONFIRM","LETTER"]:
        if kind in r:
            qty,rate,amt = r[kind]
            q=float(qty.replace(",","")); a=float(amt.replace(",",""))
            print(f"{MONTHS[mnum-1]} {year} {kind:<10} shown={rate}  exact={a/q:.6f}  (amt {amt} / qty {qty})")
            break
