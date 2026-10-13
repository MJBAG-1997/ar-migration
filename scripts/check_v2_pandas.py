"""
Check v2 fixes using pandas - much faster than openpyxl cell iteration.
Rule 9: aggregates only.
"""
import sys, os, re
from collections import Counter
import pandas as pd

DATA_ROOT = r"C:\dev\AR Migration"
V2_PATH   = os.path.join(DATA_ROOT, "hfp_signoff_pack_v2.xlsx")

def p(*a):
    print(*a, flush=True)

# Read specific sheets
p("Loading sheets...")

# 1. Plans - basis column only
p("=== 1. Plans.basis values (no-status rules) ===")
df_plans = pd.read_excel(V2_PATH, sheet_name="Plans", dtype=str)
p(f"Plans rows: {len(df_plans)}")
if "basis" in df_plans.columns:
    ctr = Counter(df_plans["basis"].dropna())
    p(f"basis values: {dict(sorted(ctr.items()))}")
    ns = {k: v for k, v in ctr.items() if "status" in k.lower()}
    p(f"  status-related: {ns}")
else:
    p("  'basis' column NOT FOUND")
p()

# 2. Charges sign-off sections
p("=== 2. Charges sign-off sections ===")
df_cs = pd.read_excel(V2_PATH, sheet_name="Charges sign-off", dtype=str)
p(f"Charges sign-off rows: {len(df_cs)}")
if "source_section" in df_cs.columns:
    ctr2 = Counter(df_cs["source_section"].dropna())
    p(f"source_section values: {dict(sorted(ctr2.items()))}")
    has_a = any("section a" in v.lower() or v.lower().startswith("a -") or v.lower().startswith("a-") for v in ctr2)
    has_b = any("section b" in v.lower() or v.lower().startswith("b -") or v.lower().startswith("b-") for v in ctr2)
    has_c = any("section c" in v.lower() or v.lower().startswith("c -") or v.lower().startswith("c-") for v in ctr2)
    p(f"  Section A: {has_a}, Section B: {has_b}, Section C: {has_c}")
else:
    p("  source_section column not found; checking first column:")
    fc = df_cs.iloc[:, 0].dropna()
    ctr2b = Counter(fc)
    p(f"  {dict(ctr2b.most_common(10))}")
p()

# 3. Summary tab
p("=== 3. Summary tab ===")
df_sum = pd.read_excel(V2_PATH, sheet_name="Summary", header=None, dtype=str)
for _, row in df_sum.iterrows():
    vals = [str(v) if pd.notna(v) else "" for v in row]
    if any(vals):
        p(f"  {vals}")
p()

# 4. Charges (all) - indicative_annual_income and basis
p("=== 4. Charges (all) ===")
df_ch = pd.read_excel(V2_PATH, sheet_name="Charges (all)", dtype=str)
p(f"Charges (all) rows: {len(df_ch)}")
p(f"Columns: {list(df_ch.columns)}")
if "indicative_annual_income" in df_ch.columns:
    non_null = df_ch["indicative_annual_income"].notna().sum()
    p(f"  indicative_annual_income filled: {non_null}")
if "basis" in df_ch.columns:
    ctr3 = Counter(df_ch["basis"].dropna())
    p(f"  basis: {dict(sorted(ctr3.items()))}")
p()

# 5. Exceptions - categories
p("=== 5. Exceptions - category breakdown ===")
df_exc = pd.read_excel(V2_PATH, sheet_name="Exceptions", dtype=str)
p(f"Exceptions rows: {len(df_exc)}")
if "category" in df_exc.columns:
    ctr4 = Counter(df_exc["category"].dropna())
    p(f"categories: {dict(sorted(ctr4.items()))}")
if "entity_type" in df_exc.columns:
    ctr5 = Counter(df_exc["entity_type"].dropna())
    p(f"entity_types: {dict(sorted(ctr5.items()))}")
p()

# 6. Adviser income
p("=== 6. Adviser income ===")
df_ai = pd.read_excel(V2_PATH, sheet_name="Adviser income", dtype=str)
p(f"Adviser income rows: {len(df_ai)}, cols: {list(df_ai.columns)}")
p()

p("Done.")
