"""Check plan structure for provider and valuation fields."""
import json, os, re
from collections import Counter

DATA_ROOT = r"C:\dev\AR Migration"
SNAP_DIR  = os.path.join(DATA_ROOT, "snapshot_20261012T090906Z")

def p(*a): print(*a, flush=True)
def normalise_polno(s):
    if s is None:
        return None
    s = re.sub(r'[\s\-]', '', str(s).strip())
    s = s.lstrip('0')
    return s.upper() if s else None

with open(os.path.join(SNAP_DIR, "plans_dedup.json")) as f:
    plans_raw = json.load(f)

plans = list(plans_raw.values()) if isinstance(plans_raw, dict) else plans_raw

# Show ALL keys in first plan
p("=== All keys in first plan ===")
p(sorted(plans[0].keys()))
p()

# Check what 'provider' field looks like
for pl in plans[:5]:
    prov = pl.get("provider") or pl.get("provider_uuid") or pl.get("provider_name")
    p(f"  provider field: prov_uuid={pl.get('provider_uuid')}, prov_name={pl.get('provider_name')}")
    prov_full = {k: v for k, v in pl.items() if "prov" in k.lower()}
    p(f"  all 'prov' keys: {prov_full}")
    p()

# Count all keys across all plans
all_keys = Counter()
for pl in plans:
    all_keys.update(pl.keys())
p("Top 30 keys (by count):")
for k, v in all_keys.most_common(30):
    p(f"  {k}: {v}")
p()

# Match Rathbones file Fund Numbers to Plannr policy numbers
import pandas as pd
df_r = pd.read_excel(os.path.join(DATA_ROOT, "Jim list  - Rathbones.xlsx"), dtype=str)
rath_norms = {normalise_polno(v): str(v).strip()
              for v in df_r["Fund Number"].dropna()}

# Find matched plans
matched_plans = []
for pl in plans:
    n = normalise_polno(pl.get("policy_number"))
    if n and n in rath_norms:
        matched_plans.append(pl)

p(f"=== Plans matching Rathbones Fund Numbers: {len(matched_plans)} ===")
# Provider info on matched plans
for pl in matched_plans[:5]:
    prov_keys = {k: v for k, v in pl.items() if "prov" in k.lower()}
    p(f"  policy_number={pl.get('policy_number')} type={pl.get('type')} status={pl.get('status')}")
    p(f"  provider fields: {prov_keys}")
    p()

# Provider counts on matched plans
prov_counts = Counter()
for pl in matched_plans:
    prov_keys = {k: v for k, v in pl.items() if "prov" in k.lower() and v}
    prov_counts[str(prov_keys)[:80]] += 1
p(f"Provider key patterns in matched plans:")
for k, v in prov_counts.most_common(5):
    p(f"  {k}: {v}")
p()

# Check provider_uuid specifically on matched plans
prov_uuids = Counter(pl.get("provider_uuid") for pl in matched_plans)
p(f"provider_uuid values in matched plans: {dict(prov_uuids)}")

# For the 18 unmatched Rathbones plans, what are their Fund Types?
matched_fund_numbers = {normalise_polno(pl.get("policy_number")) for pl in matched_plans}
unmatched = df_r[~df_r["Fund Number"].apply(lambda v: normalise_polno(str(v)) if pd.notna(v) else None).isin(matched_fund_numbers)]
p(f"\n=== 18 unmatched Rathbones plans (in file, not in Plannr) ===")
p(f"Fund Type breakdown: {Counter(unmatched['Fund Type'].tolist()).most_common()}")
p(f"Fund Number patterns: {Counter(unmatched['Fund Number'].apply(lambda v: f'len={len(str(v).strip())}').tolist()).most_common()}")
p()

p("Done.")
