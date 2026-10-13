"""
Targeted checks on v2 pack for the 5 fix categories.
Rule 9: aggregates only.
"""
import sys, os, re
from collections import Counter
import openpyxl

DATA_ROOT = r"C:\dev\AR Migration"
V2_PATH   = os.path.join(DATA_ROOT, "hfp_signoff_pack_v2.xlsx")

def p(*a):
    print(*a, flush=True)

wb = openpyxl.load_workbook(V2_PATH, read_only=True, data_only=True)

# 1. Plans.basis values  (no-status rules)
p("=== 1. Plans.basis values ===")
ws = wb["Plans"]
hdrs = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
basis_idx = next((i for i, h in enumerate(hdrs) if h == "basis"), None)
if basis_idx is not None:
    basis_vals = []
    for r in range(2, ws.max_row + 1):
        v = ws.cell(r, basis_idx + 1).value
        if v:
            basis_vals.append(str(v))
    ctr = Counter(basis_vals)
    p(f"basis column (col {basis_idx+1}): {dict(ctr.most_common(15))}")
    ns = {k: v for k, v in ctr.items() if "status" in k.lower()}
    p(f"  'status' in basis: {ns}")
else:
    p("  basis column NOT FOUND in Plans")
p()

# 2. Charges sign-off sections
p("=== 2. Charges sign-off sections ===")
ws = wb["Charges sign-off"]
sec_vals = []
for r in range(1, ws.max_row + 1):
    v = ws.cell(r, 1).value
    if v:
        sec_vals.append(str(v))
ctr2 = Counter(sec_vals)
p(f"First column unique values: {dict(ctr2.most_common(20))}")
has_a = any("section a" in v.lower() or v.lower().startswith("a ") for v in sec_vals)
has_b = any("section b" in v.lower() or v.lower().startswith("b ") for v in sec_vals)
has_c = any("section c" in v.lower() or v.lower().startswith("c ") for v in sec_vals)
p(f"  Section A present: {has_a}")
p(f"  Section B present: {has_b}")
p(f"  Section C present: {has_c}")
total_charge_rows = ws.max_row - 1
p(f"  Total data rows in Charges sign-off: {total_charge_rows}")
p()

# 3. Summary tab - check for '42' or 'excluded' references
p("=== 3. Summary tab ===")
ws = wb["Summary"]
for r in range(1, ws.max_row + 1):
    v1 = ws.cell(r, 1).value
    v2 = ws.cell(r, 2).value
    if v1 or v2:
        p(f"  row {r}: [{v1}] | [{v2}]")
p()

# 4. Charges (all) - count by basis
p("=== 4. Charges (all) - basis + excluded info ===")
ws = wb["Charges (all)"]
hdrs = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
basis_idx = next((i for i, h in enumerate(hdrs) if h == "basis"), None)
if basis_idx is not None:
    basis_vals = []
    for r in range(2, ws.max_row + 1):
        v = ws.cell(r, basis_idx + 1).value
        if v:
            basis_vals.append(str(v))
    ctr3 = Counter(basis_vals)
    p(f"Charges.basis values: {dict(ctr3.most_common(10))}")
p(f"Charges (all) total rows: {ws.max_row - 1}")
p()

# 5. Exceptions tab - categories
p("=== 5. Exceptions - category breakdown ===")
ws = wb["Exceptions"]
hdrs = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
cat_idx = next((i for i, h in enumerate(hdrs) if h == "category"), None)
etype_idx = next((i for i, h in enumerate(hdrs) if h == "entity_type"), None)
if cat_idx is not None:
    cats = []
    for r in range(2, ws.max_row + 1):
        v = ws.cell(r, cat_idx + 1).value
        if v:
            cats.append(str(v))
    ctr4 = Counter(cats)
    p(f"Exception categories: {dict(ctr4.most_common(15))}")
    ns = {k: v for k, v in ctr4.items() if "status" in k.lower() or "no_status" in k.lower()}
    p(f"  status-related: {ns}")
p()

# 6. Change log - check for references to '42' exclusions
p("=== 6. Change log sample (first 10 rows) ===")
ws = wb["Change log"]
hdrs = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
p(f"  headers: {hdrs}")
for r in range(2, min(12, ws.max_row + 1)):
    row = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
    p(f"  row {r}: {row}")
p()

wb.close()
