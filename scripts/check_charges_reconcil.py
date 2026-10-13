"""
Check charges reconciliation state in v2 and snapshot.
Rule 9: aggregates only.
"""
import sys, os, re, json
from collections import Counter
import pandas as pd

DATA_ROOT = r"C:\dev\AR Migration"
V2_PATH   = os.path.join(DATA_ROOT, "hfp_signoff_pack_v2.xlsx")
SNAP_DIR  = os.path.join(DATA_ROOT, "snapshot_20261012T090906Z")

def p(*a): print(*a, flush=True)

# 1. Check original Plannr charges
p("=== Snapshot charges count ===")
charges_path = os.path.join(SNAP_DIR, "charges_dedup.json")
if os.path.exists(charges_path):
    with open(charges_path) as f:
        charges_raw = json.load(f)
    # May be a dict keyed by UUID or a list
    if isinstance(charges_raw, dict):
        charges = list(charges_raw.values())
    else:
        charges = charges_raw
    p(f"Total charges in snapshot (dedup): {len(charges)}")
    # Count by type/service_type/recurring
    by_recurring = Counter(c.get("recurring") for c in charges)
    by_service   = Counter(c.get("service_type") for c in charges)
    p(f"  by recurring: {dict(by_recurring)}")
    p(f"  by service_type: {dict(by_service)}")
    # Ongoing charges (recurring=True, service_type=Ongoing)
    ongoing = [c for c in charges if c.get("recurring") is True or str(c.get("recurring","")).lower() == "true"]
    p(f"  recurring=True: {len(ongoing)}")
    # Active (no end_date or end_date > 2026-11-02)
    GO_LIVE = "2026-11-02"
    active_ongoing = []
    for c in ongoing:
        ed = c.get("end_date")
        if ed is None or str(ed) >= GO_LIVE:
            active_ongoing.append(c)
    p(f"  active ongoing (no end or end >= go-live): {len(active_ongoing)}")
    excluded = len(ongoing) - len(active_ongoing)
    p(f"  excluded (ended before go-live): {excluded}")
else:
    p(f"  charges_dedup.json not found at {charges_path}")
    # Try plans_dedup.json
    plans_path = os.path.join(SNAP_DIR, "plans_dedup.json")
    if os.path.exists(plans_path):
        with open(plans_path) as f:
            plans = json.load(f)
        # Check if plans have embedded charges
        sample = plans[:5]
        for pl in sample:
            keys = list(pl.keys())
            p(f"  plan keys: {keys[:20]}")
            break
p()

# 2. Change log - charge-related rows
p("=== Change log charge-related entries ===")
df_cl = pd.read_excel(V2_PATH, sheet_name="Change log", dtype=str)
p(f"Change log total rows: {len(df_cl)}")
p(f"Columns: {list(df_cl.columns)}")
# Filter charge-related
charge_rows = df_cl[df_cl.apply(
    lambda r: any("charge" in str(v).lower() or "exclu" in str(v).lower()
                  for v in r if pd.notna(v)), axis=1)]
p(f"  Charge/exclusion rows: {len(charge_rows)}")
# Show entity_type breakdown
ctr = Counter(df_cl["entity_type"].dropna())
p(f"  entity_type breakdown: {dict(sorted(ctr.items()))}")
# Show rule_id breakdown (aggregated)
if "rule_id" in df_cl.columns:
    ctr2 = Counter(df_cl["rule_id"].dropna())
    p(f"  rule_id breakdown: {dict(sorted(ctr2.items()))}")
p()

# 3. Charges (all) basis + one-per-plan check
p("=== Charges (all) one-per-plan check ===")
df_ch = pd.read_excel(V2_PATH, sheet_name="Charges (all)", dtype=str)
p(f"  Total loadable charges: {len(df_ch)}")
# Count distinct plan_uuids
plan_uuids = df_ch["plan_uuid"].dropna().unique()
p(f"  Distinct plan_uuids: {len(plan_uuids)}")
# Check duplicates per plan
dupes = df_ch.groupby("plan_uuid").size()
multi = dupes[dupes > 1]
p(f"  Plans with >1 charge: {len(multi)}")
# Basis breakdown
if "basis" in df_ch.columns:
    ctr3 = Counter(df_ch["basis"].dropna())
    p(f"  basis: {dict(sorted(ctr3.items()))}")
# Service type
if "service_type" in df_ch.columns:
    ctr4 = Counter(df_ch["service_type"].dropna())
    p(f"  service_type: {dict(sorted(ctr4.items()))}")
p()

# 4. Does Charges (all) include workbook-sourced charges?
# Check if any row has basis='provider'
p("=== Provider-sourced charges in Charges (all) ===")
if "basis" in df_ch.columns:
    prov_charges = df_ch[df_ch["basis"].str.lower().str.contains("provider", na=False)]
    p(f"  basis contains 'provider': {len(prov_charges)}")
    co_charges = df_ch[df_ch["basis"] == "carried over"]
    p(f"  basis='carried over': {len(co_charges)}")
p()

p("Done.")
