"""Count Plannr plans with any Rathbones provider."""
import json, os, re
from collections import Counter
import pandas as pd

DATA_ROOT = r"C:\dev\AR Migration"
SNAP_DIR  = os.path.join(DATA_ROOT, "snapshot_20261012T090906Z")

def p(*a): print(*a, flush=True)
def normalise_polno(s):
    if s is None: return None
    s = re.sub(r'[\s\-]', '', str(s).strip())
    s = s.lstrip('0')
    return s.upper() if s else None

with open(os.path.join(SNAP_DIR, "plans_dedup.json")) as f:
    plans_raw = json.load(f)
plans = list(plans_raw.values()) if isinstance(plans_raw, dict) else plans_raw

# Rathbones provider UUIDs (from reference.json)
RATH_UUIDS = {
    'd2c551bd-a2d1-4223-a3c6-d28a9d92ff91',  # Rathbone Greenbank Investments
    'de8a8f76-44d6-41d8-994b-f800ba05b1f5',  # Rathbone Investment Management Ltd
    '5c086784-b183-4964-bd38-bf8410275068',  # Rathbone Unit Trust Management Ltd
    '5d1737ed-195a-4323-8375-27bea2b0d6ec',  # Rathbones
}

def get_provider_uuid(pl):
    prov = pl.get("provider")
    if isinstance(prov, dict):
        return prov.get("uuid")
    return None

def get_provider_name(pl):
    prov = pl.get("provider")
    if isinstance(prov, dict):
        return prov.get("name")
    return None

# All Plannr plans with ANY Rathbones provider
rath_plans = [pl for pl in plans if get_provider_uuid(pl) in RATH_UUIDS]
p(f"Plannr plans with Rathbones provider: {len(rath_plans)}")
p(f"  By provider: {Counter(get_provider_name(pl) for pl in rath_plans)}")
p()

# Load Rathbones file
df_r = pd.read_excel(os.path.join(DATA_ROOT, "Jim list  - Rathbones.xlsx"), dtype=str)
rath_file_norms = {normalise_polno(v) for v in df_r["Fund Number"].dropna() if normalise_polno(v)}
p(f"Rathbones file unique Fund Numbers: {len(rath_file_norms)}")

# Match counts
plannr_rath_norms = {normalise_polno(pl.get("policy_number")): pl for pl in rath_plans
                     if normalise_polno(pl.get("policy_number"))}
no_polno = [pl for pl in rath_plans if not normalise_polno(pl.get("policy_number"))]

matched   = rath_file_norms & set(plannr_rath_norms.keys())
rath_only = rath_file_norms - set(plannr_rath_norms.keys())
plan_only = set(plannr_rath_norms.keys()) - rath_file_norms

p(f"\nFull match report:")
p(f"  Rathbones file rows:                   52")
p(f"  Rathbones file unique policy numbers:  {len(rath_file_norms)}")
p(f"  Matched (file AND Plannr):             {len(matched)}")
p(f"  Rathbones-only (file, not in Plannr):  {len(rath_only)}")
p(f"  Plannr-only (Plannr, not in file):     {len(plan_only)}")
p(f"  Plannr Rathbones plans no policy no:   {len(no_polno)}")
p(f"  Total Plannr Rathbones plans:          {len(rath_plans)}")

# Status breakdown of matched plans
status_ctr = Counter(pl.get("status") for pl in rath_plans
                     if normalise_polno(pl.get("policy_number")) in matched)
p(f"\n  Matched plans by status: {dict(status_ctr)}")
status_plan_only = Counter(plannr_rath_norms[n].get("status") for n in plan_only)
p(f"  Plannr-only plans by status: {dict(status_plan_only)}")

# Check SIPP rows in Rathbones file - what's the Col D pattern for those
sipp_rows = df_r[df_r["Fund Type"].str.lower().str.contains("sipp", na=False)]
p(f"\nSIPP rows in Rathbones file: {len(sipp_rows)}")
p(f"  Fund Type values: {Counter(sipp_rows['Fund Type'].tolist()).most_common()}")
p(f"  Fund Name col (col D) - masked patterns:")
pats = Counter()
for v in sipp_rows["Fund Name"].tolist():
    s = str(v)
    nd = sum(1 for c in s if c.isdigit())
    na = sum(1 for c in s if c.isalpha())
    pats[f"len={len(s)} d={nd} a={na}"] += 1
p(f"    {pats.most_common()}")

# Check for any column that looks like a fee/charge rate
p(f"\n=== Column check for charge/rate ===")
p(f"All columns: {list(df_r.columns)}")
for col in df_r.columns:
    vals = df_r[col].dropna().tolist()
    if not vals: continue
    # Check if looks like a percentage (small numbers between 0 and 5)
    try:
        nums = [float(v) for v in vals]
        if all(0 < n < 5 for n in nums):
            p(f"  Col '{col}' might be a rate: sample {nums[:3]}")
    except (ValueError, TypeError):
        pass

p("Done.")
