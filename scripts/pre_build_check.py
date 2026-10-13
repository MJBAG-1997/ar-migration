"""
Pre-build checks before writing v3:
- Valuation scale (pence vs pounds)
- Charges sign-off workbook rates (to use for Fix 4)
- Snapshot income/employment/expenses availability
- Duplicate exceptions detail (Fix 5: which 39 to keep)
Rule 9: aggregates only.
"""
import sys, os, re, json
from collections import Counter, defaultdict
import pandas as pd

DATA_ROOT = r"C:\dev\AR Migration"
V2_PATH   = os.path.join(DATA_ROOT, "hfp_signoff_pack_v2.xlsx")
SNAP_DIR  = os.path.join(DATA_ROOT, "snapshot_20261012T090906Z")

def p(*a): print(*a, flush=True)

# ─── 1. Valuations from plans (pence check) ───────────────────────────────────
p("=== 1. Valuation scale check ===")
with open(os.path.join(SNAP_DIR, "plans_dedup.json")) as f:
    plans_raw = json.load(f)
plans = list(plans_raw.values()) if isinstance(plans_raw, dict) else plans_raw

# Check value field
vals_with_value = [pl for pl in plans if pl.get("value") is not None]
p(f"Plans with 'value': {len(vals_with_value)}")
sample_vals = sorted([float(pl["value"]) for pl in vals_with_value if pl["value"]], reverse=True)[:5]
p(f"Top 5 values: {sample_vals}")
total = sum(float(pl["value"]) for pl in vals_with_value if pl["value"])
p(f"Sum of all values: {total:,.0f}")
p(f"As pounds: {total:,.2f}  As pence->pounds: {total/100:,.2f}")
p()

# Check _valuations_latest structure
plans_with_valn = [pl for pl in plans if pl.get("_valuations_latest")]
p(f"Plans with _valuations_latest: {len(plans_with_valn)}")
if plans_with_valn:
    sample = plans_with_valn[0]["_valuations_latest"]
    p(f"_valuations_latest sample structure: {type(sample)} len={len(sample) if isinstance(sample, (list,dict)) else 'N/A'}")
    if isinstance(sample, list) and sample:
        p(f"  first entry keys: {list(sample[0].keys()) if isinstance(sample[0], dict) else sample[0]}")
    elif isinstance(sample, dict):
        p(f"  keys: {list(sample.keys())}")
p()

# ─── 2. Charges sign-off - read workbook rates for Section A ──────────────────
p("=== 2. Charges sign-off - Section A rates ===")
df_cs = pd.read_excel(V2_PATH, sheet_name="Charges sign-off", dtype=str)
p(f"Charges sign-off total rows: {len(df_cs)}")
p(f"Columns: {list(df_cs.columns)}")

sec_a = df_cs[df_cs["source_section"] == "A - Workbook confirmed"]
p(f"Section A rows: {len(sec_a)}")

# Check workbook_rate values
if "workbook_rate" in df_cs.columns:
    wr_vals = sec_a["workbook_rate"].dropna()
    p(f"workbook_rate non-null: {len(wr_vals)}")
    # Convert to float and check range
    try:
        wr_floats = [float(v) for v in wr_vals]
        if wr_floats:
            p(f"workbook_rate range: min={min(wr_floats):.4f} max={max(wr_floats):.4f} median={sorted(wr_floats)[len(wr_floats)//2]:.4f}")
    except:
        p(f"workbook_rate sample (str): {wr_vals.head(3).tolist()}")

if "workbook_amount" in df_cs.columns:
    wa_vals = sec_a["workbook_amount"].dropna()
    p(f"workbook_amount non-null: {len(wa_vals)}")

if "frequency" in df_cs.columns:
    freq_ctr = Counter(sec_a["frequency"].dropna())
    p(f"frequency: {dict(freq_ctr)}")

if "final_rate" in df_cs.columns:
    fr_vals = sec_a["final_rate"].dropna()
    p(f"final_rate non-null: {len(fr_vals)}")
    try:
        fr_floats = [float(v) for v in fr_vals]
        if fr_floats:
            p(f"final_rate range: min={min(fr_floats):.4f} max={max(fr_floats):.4f}")
    except:
        pass

# Section B and C
sec_b = df_cs[df_cs["source_section"].str.startswith("B", na=False)]
sec_c = df_cs[df_cs["source_section"].str.startswith("C", na=False)]
p(f"Section B rows: {len(sec_b)}, Section C rows: {len(sec_c)}")
p()

# ─── 3. Snapshot - accounts.json structure for income/employment/expenses ─────
p("=== 3. Snapshot accounts structure (income/employment/expenses) ===")
with open(os.path.join(SNAP_DIR, "accounts.json")) as f:
    accounts_raw = json.load(f)
# accounts.json might be a list or dict
if isinstance(accounts_raw, dict):
    accounts = list(accounts_raw.values())
else:
    accounts = accounts_raw
p(f"Accounts loaded: {len(accounts)}")

# Check first account structure
if accounts:
    acc = accounts[0]
    p(f"Account keys: {sorted(acc.keys())}")
    p()

    # Income
    income_key = next((k for k in acc if "income" in k.lower() or "employ" in k.lower()), None)
    p(f"Income/employment key: {income_key}")
    if income_key:
        sample = acc.get(income_key)
        p(f"  sample: {type(sample)} = {str(sample)[:200]}")

    # Expenses/expenditure
    exp_key = next((k for k in acc if "expend" in k.lower() or "expense" in k.lower()), None)
    p(f"Expenses key: {exp_key}")
    if exp_key:
        sample = acc.get(exp_key)
        p(f"  sample: {type(sample)} = {str(sample)[:200]}")
p()

# Count income/employment across all accounts
income_count = 0
expense_count = 0
for acc in accounts:
    for k, v in acc.items():
        if ("income" in k.lower() or "employ" in k.lower()) and isinstance(v, list):
            income_count += len(v)
        if ("expend" in k.lower() or "expense" in k.lower()) and isinstance(v, list):
            expense_count += len(v)
p(f"Total income/employment records: {income_count}")
p(f"Total expense records: {expense_count}")

# Check all keys across all accounts
all_keys = Counter()
for acc in accounts:
    all_keys.update(acc.keys())
p(f"All account keys: {dict(all_keys.most_common(40))}")
p()

# ─── 4. Duplicate exception detail (Fix 5) ────────────────────────────────────
p("=== 4. Duplicate policy number exceptions - Fix 5 ===")
df_exc = pd.read_excel(V2_PATH, sheet_name="Exceptions", dtype=str)
dup_rows = df_exc[df_exc["category"] == "Duplicate policy number"]
p(f"Total 'Duplicate policy number' rows: {len(dup_rows)}")
p(f"Columns: {list(dup_rows.columns)}")
# Show detail column patterns
if "detail" in dup_rows.columns:
    detail_ctr = Counter(dup_rows["detail"].dropna())
    p(f"detail values: {dict(detail_ctr.most_common(5))}")
# Show source_uuid column patterns (aggregated)
p(f"source_uuid fill: {dup_rows['source_uuid'].notna().sum()}")
# Group by policy number group - use 'current_value' which might hold the policy number
if "current_value" in dup_rows.columns:
    cv_ctr = Counter(dup_rows["current_value"].dropna())
    p(f"current_value (policy number) unique: {len(cv_ctr)}")
    p(f"groups with >1 row: {sum(1 for v in cv_ctr.values() if v > 1)}")
if "proposed_value" in dup_rows.columns:
    pv_sample = dup_rows["proposed_value"].dropna().head(3).tolist()
    p(f"proposed_value sample: {pv_sample}")
p()

# ─── 5. Check mappings_hfp for Expenditure tab ────────────────────────────────
p("=== 5. mappings_hfp Expenditure tab ===")
try:
    mappings_path = os.path.join(DATA_ROOT, "mappings_hfp.xlsx")
    xl = pd.ExcelFile(mappings_path)
    p(f"Sheets in mappings_hfp.xlsx: {xl.sheet_names}")
    if "Expenditure" in xl.sheet_names or "expenditure" in [s.lower() for s in xl.sheet_names]:
        exp_sheet = next(s for s in xl.sheet_names if s.lower() == "expenditure")
        df_exp = pd.read_excel(mappings_path, sheet_name=exp_sheet, dtype=str)
        p(f"Expenditure tab rows: {len(df_exp)}, cols: {list(df_exp.columns)}")
        p(df_exp.head(5).to_string())
    else:
        p("No Expenditure tab in mappings_hfp.xlsx")
except Exception as e:
    p(f"Error reading mappings: {e}")
p()

p("Done.")
