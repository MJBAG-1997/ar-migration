"""
Find Rathbones provider UUID in Plannr and count Plannr-only Rathbones plans.
Rule 9: aggregates, patterns only.
"""
import sys, os, re, json
from collections import Counter

DATA_ROOT = r"C:\dev\AR Migration"
SNAP_DIR  = os.path.join(DATA_ROOT, "snapshot_20261012T090906Z")

def p(*a): print(*a, flush=True)

def normalise_polno(s):
    if s is None:
        return None
    s = str(s).strip()
    s = re.sub(r'[\s\-]', '', s)
    s = s.lstrip('0')
    return s.upper() if s else None

# Load reference (providers)
ref_path = os.path.join(SNAP_DIR, "reference.json")
with open(ref_path) as f:
    ref = json.load(f)

providers = ref.get("providers", [])
p(f"Total providers in reference: {len(providers)}")

# Find Rathbones
rath_providers = [pr for pr in providers
                  if "rathbone" in str(pr.get("name", "")).lower()
                  or "rathbone" in str(pr.get("description", "")).lower()
                  or "rathbone" in str(pr.get("display_name", "")).lower()]
p(f"Rathbones providers found:")
for pr in rath_providers:
    p(f"  {pr}")
p()

# Also search by any key
if not rath_providers:
    p("No 'rathbone' match in name/description/display_name.")
    p("Searching all provider fields:")
    for pr in providers[:10]:
        p(f"  {pr}")

# Load plans
plans_path = os.path.join(SNAP_DIR, "plans_dedup.json")
with open(plans_path) as f:
    plans_raw = json.load(f)

if isinstance(plans_raw, dict):
    plans = list(plans_raw.values())
else:
    plans = plans_raw

p(f"Total plans: {len(plans)}")

# Check plan structure for provider fields
sample = plans[:3]
for pl in sample:
    p(f"  Plan keys sample: {list(pl.keys())[:20]}")
    p(f"  provider_uuid={pl.get('provider_uuid')}, provider_name={pl.get('provider_name')}, provider={pl.get('provider')}")
    break
p()

# Find plans with Rathbones provider UUIDs
if rath_providers:
    rath_uuids = set(pr.get("uuid") for pr in rath_providers if pr.get("uuid"))
    rath_names = set(pr.get("name") for pr in rath_providers)
    rath_plans = [pl for pl in plans
                  if pl.get("provider_uuid") in rath_uuids
                  or pl.get("provider_name") in rath_names]
    p(f"Plannr plans with Rathbones provider: {len(rath_plans)}")

    # Rathbones file policy numbers (Fund Number 6-digit)
    import pandas as pd
    df_rath = pd.read_excel(os.path.join(DATA_ROOT, "Jim list  - Rathbones.xlsx"), dtype=str)
    rath_norms = set(normalise_polno(v) for v in df_rath["Fund Number"].dropna() if normalise_polno(v))
    p(f"Rathbones file unique normalised policy numbers: {len(rath_norms)}")

    # Plannr Rathbones plans: which are in the file vs not
    plannr_rath_norms = {}
    for pl in rath_plans:
        n = normalise_polno(pl.get("policy_number"))
        if n:
            plannr_rath_norms[n] = pl.get("uuid")

    matched = rath_norms & set(plannr_rath_norms.keys())
    plannr_only = set(plannr_rath_norms.keys()) - rath_norms
    rath_file_only = rath_norms - set(plannr_rath_norms.keys())

    p(f"\nMatch report (Rathbones provider plans only):")
    p(f"  Rathbones file unique policy numbers:  {len(rath_norms)}")
    p(f"  Matched (file AND Plannr):             {len(matched)}")
    p(f"  Rathbones-only (file, not in Plannr):  {len(rath_file_only)}")
    p(f"  Plannr-only (Plannr, not in file):     {len(plannr_only)}")
    p(f"  Total Plannr Rathbones plans:          {len(rath_plans)}")

    # Plans with no policy number in Rathbones set
    no_polno = [pl for pl in rath_plans if not normalise_polno(pl.get("policy_number"))]
    p(f"  Plannr Rathbones plans with no policy number: {len(no_polno)}")
else:
    # Try searching plan provider names
    prov_names = Counter(pl.get("provider_name") or pl.get("provider") for pl in plans)
    rath_like = {k: v for k, v in prov_names.items()
                 if k and "rath" in str(k).lower()}
    p(f"Provider names containing 'rath': {rath_like}")
    all_prov = [(k, v) for k, v in prov_names.most_common(30)]
    p(f"Top 30 provider names in plans: {all_prov}")

p("Done.")
