"""
Four investigations before v4:
1. Rate units audit
2. Placeholder policy numbers
3. 69 zero-rate Section A plans
4. Plannr charge reconciliation (must sum to 330)
Rule 9: aggregates / UUIDs only in console.
"""
import json, os, re
from collections import Counter, defaultdict
import pandas as pd

DATA_ROOT = r"C:\dev\AR Migration"
SNAP_DIR  = os.path.join(DATA_ROOT, "snapshot_20261012T090906Z")
V3_PATH   = os.path.join(DATA_ROOT, "hfp_signoff_pack_v3.xlsx")
WB_PATH   = os.path.join(DATA_ROOT, "Hoban_Financial_Planning_Ltd_260922_Payment_Rates_113727875.xlsx")

def p(*a): print(*a, flush=True)

def normalise_polno(s):
    if s is None: return None
    s = re.sub(r"[\s\-]", "", str(s).strip())
    s = s.lstrip("0")
    return s.upper() if s else None

def pence_to_pounds(amt):
    try: return int(str(amt)) / 100.0
    except (TypeError, ValueError): return 0.0

def get_provider_uuid(pl):
    prov = pl.get("provider")
    if isinstance(prov, dict): return prov.get("uuid")
    return None

def get_client_uuids(pl):
    owners = pl.get("_owners", [])
    result = set()
    if isinstance(owners, list):
        for o in owners:
            if isinstance(o, dict):
                c = o.get("client_uuid")
                if c: result.add(c)
    return result

p("Loading data...")
with open(os.path.join(SNAP_DIR, "plans_dedup.json")) as f:
    plans_raw = json.load(f)
plans = list(plans_raw.values())
plans_by_uuid = {pl["uuid"]: pl for pl in plans}

with open(os.path.join(SNAP_DIR, "charges_dedup.json")) as f:
    charges_raw = json.load(f)
charges = list(charges_raw.values())

GO_LIVE = "2026-11-02"
RATH_PROVIDER_UUIDS = {
    "d2c551bd-a2d1-4223-a3c6-d28a9d92ff91",
    "de8a8f76-44d6-41d8-994b-f800ba05b1f5",
    "5c086784-b183-4964-bd38-bf8410275068",
    "5d1737ed-195a-4323-8375-27bea2b0d6ec",
}

# Read v3 Charges (all)
df_ca = pd.read_excel(V3_PATH, sheet_name="Charges (all)", dtype=str)
# Read v3 Charges sign-off
df_cs = pd.read_excel(V3_PATH, sheet_name="Charges sign-off", dtype=str)
# Read v3 Plans tab
df_plans = pd.read_excel(V3_PATH, sheet_name="Plans", dtype=str)

p(f"  Plans: {len(plans)}")
p(f"  Charges: {len(charges)}")
p(f"  v3 Charges (all) rows: {len(df_ca)}")
p()

# ─── 1. Rate units audit ───────────────────────────────────────────────────────

p("=" * 60)
p("INVESTIGATION 1: Rate units audit")
p("=" * 60)

# Plan value map from snapshot (pounds)
plan_value_pounds = {}
for pl in plans:
    val_dict = pl.get("value")
    if isinstance(val_dict, dict):
        plan_value_pounds[pl["uuid"]] = pence_to_pounds(val_dict.get("amount", 0))
    else:
        plan_value_pounds[pl["uuid"]] = 0.0

# Analyse each basis group
basis_groups = df_ca.groupby("basis")
for basis_name, grp in basis_groups:
    pcts = grp["percentage"].dropna()
    try:
        pct_vals = sorted([float(v) for v in pcts])
    except (ValueError, TypeError):
        pct_vals = []
    if not pct_vals:
        p(f"  Group '{basis_name}': no percentage data")
        continue
    n = len(pct_vals)
    pct_min = pct_vals[0]
    pct_med = pct_vals[n // 2]
    pct_max = pct_vals[-1]

    # Sum of plan valuations for this group
    plan_uuids_in_grp = grp["plan_uuid"].dropna().tolist()
    total_val = sum(plan_value_pounds.get(str(u), 0.0) for u in plan_uuids_in_grp)

    # Value-weighted average rate
    wt_sum = 0.0
    wt_count = 0.0
    for _, row in grp.iterrows():
        uuid = str(row.get("plan_uuid", "") or "")
        val = plan_value_pounds.get(uuid, 0.0)
        try:
            pct = float(row.get("percentage", 0) or 0)
        except (ValueError, TypeError):
            pct = 0.0
        wt_sum += pct * val
        wt_count += val
    vw_avg = wt_sum / wt_count if wt_count > 0 else 0.0

    # Check unit: Plannr stores 0.75 = 0.75%. Workbook stores 0.0075 = 0.75%.
    # After conversion in build_v3: all should be in percentage form (0.75 = 0.75%)
    unit_ok = 0.1 <= pct_max <= 5.0  # percentage form would be in 0.3–1.0 range
    unit_note = "OK (pct form 0.3-1.0)" if unit_ok else "WARNING: may be decimal form"

    p(f"\n  Group: '{basis_name}' ({n} rows)")
    p(f"    percentage range: min={pct_min:.5f}  median={pct_med:.5f}  max={pct_max:.5f}")
    p(f"    total plan valuation: GBP {total_val:,.2f}")
    p(f"    value-weighted avg rate: {vw_avg:.5f}%")
    p(f"    unit check: {unit_note}")

# Compute total IAI under current units
p()
p("  IAI computation using rate_decimal = percentage/100:")
total_iai = 0.0
for _, row in df_ca.iterrows():
    uuid = str(row.get("plan_uuid", "") or "")
    val = plan_value_pounds.get(uuid, 0.0)
    try:
        pct = float(row.get("percentage", 0) or 0)
    except (ValueError, TypeError):
        pct = 0.0
    total_iai += val * (pct / 100.0)
p(f"  Total IAI (v3 Charges all, rate_decimal=pct/100): GBP {total_iai:,.2f}")

# Compute total IAI if rates are ALREADY in decimal form (i.e., 0.0075 not 0.75)
total_iai_if_already_decimal = 0.0
for _, row in df_ca.iterrows():
    uuid = str(row.get("plan_uuid", "") or "")
    val = plan_value_pounds.get(uuid, 0.0)
    try:
        pct = float(row.get("percentage", 0) or 0)
    except (ValueError, TypeError):
        pct = 0.0
    total_iai_if_already_decimal += val * pct  # treat as decimal directly
p(f"  Total IAI (if percentage already decimal): GBP {total_iai_if_already_decimal:,.2f}")

# Expected ~193K
p(f"  Expected (from Mark): ~GBP 193,000")
p()

# Check workbook rates - what does the raw workbook store?
df_wb = pd.read_excel(WB_PATH, sheet_name="Payment Rates Summary", dtype=str)
wb_rates_sample = df_wb["Annual Payment - Investments"].dropna().tolist()[:10]
try:
    wb_rates_f = [float(v) for v in wb_rates_sample]
    p(f"  Raw workbook 'Annual Payment - Investments' sample: {wb_rates_f}")
    p(f"  (these are stored as e.g. 0.006=0.6% or 0.0075=0.75% — decimal form)")
except:
    p(f"  Raw workbook sample: {wb_rates_sample}")

p()
p("  Plannr charges.percentage sample (from charges_dedup):")
pct_sample = [float(ch.get("percentage", 0)) for ch in charges[:10] if ch.get("percentage")]
p(f"  {pct_sample}")
p("  (Plannr stores 0.75 = 0.75% — percentage form)")

# ─── 2. Placeholder policy numbers ────────────────────────────────────────────

p()
p("=" * 60)
p("INVESTIGATION 2: Placeholder policy numbers")
p("=" * 60)

PLACEHOLDER_UPPER = {
    "TBC", "NA", "N/A", "NOTKNOWN", "UNKNOWN", "NOTONFILE",
    "NOTSET", "NONE", "NOT SET", "NOTON", "NOTAPPLICABLE",
}

def is_placeholder(pn_raw):
    if not pn_raw or str(pn_raw).strip() in ("", "None", "nan"):
        return True
    pn_str = str(pn_raw).strip()
    # Blank or dash-ish
    if re.fullmatch(r"[\-\s\.]+", pn_str):
        return True
    # Normalised form
    norm = re.sub(r"[\s\-]", "", pn_str).upper()
    if norm in PLACEHOLDER_UPPER:
        return True
    # <= 5 chars that aren't clearly a real reference
    if len(norm) <= 5 and not re.match(r"^\d{4,5}$|^[A-Z]{2}\d{3,}$", norm):
        return True
    # Also check original (some have slashes etc.)
    if pn_str.upper().replace(" ", "").replace("-","") in PLACEHOLDER_UPPER:
        return True
    return False

PLAN_TYPES_NEED_POLNO = {"investment", "pension", "protection_policy", "group_protection", "annuity"}

placeholder_plans = []
for pl in plans:
    pn = pl.get("policy_number")
    if is_placeholder(pn):
        placeholder_plans.append(pl)

p(f"  Plans with placeholder policy numbers: {len(placeholder_plans)}")
ptype_ctr = Counter(pl.get("type", "unknown") for pl in placeholder_plans)
p(f"  By plan type: {dict(sorted(ptype_ctr.items()))}")

new_nopolno_exceptions = [pl for pl in placeholder_plans if pl.get("type") in PLAN_TYPES_NEED_POLNO]
p(f"  Plans needing 'No policy number' exception (right type): {len(new_nopolno_exceptions)}")

# Show placeholder value distribution (masked)
ph_vals = Counter()
for pl in placeholder_plans:
    pn = pl.get("policy_number") or ""
    norm = re.sub(r"[\s\-]", "", str(pn).strip()).upper()
    ph_vals[norm[:20]] += 1
p(f"  Placeholder value distribution: {dict(ph_vals.most_common(15))}")
p()

# Check how many Plans-tab rows had non-empty placeholder policy numbers (from v3 plans tab)
ph_in_plans_tab = df_plans[df_plans["policy_number"].apply(
    lambda x: is_placeholder(x) and bool(str(x).strip())
)]
p(f"  Plans tab rows with non-blank placeholder: {len(ph_in_plans_tab)}")

# ─── 3. 69 zero-rate Section A plans ─────────────────────────────────────────

p()
p("=" * 60)
p("INVESTIGATION 3: 69 zero-rate Section A plans")
p("=" * 60)

df_wb2 = pd.read_excel(WB_PATH, sheet_name="Payment Rates Summary", dtype=str)
wb_rate_all = {}  # wrapper_ref_norm -> rate (including zero)
for _, row in df_wb2.iterrows():
    wr = normalise_polno(row["Wrapper Ref"])
    if not wr: continue
    try:
        rate = float(row["Annual Payment - Investments"])
        wb_rate_all[wr] = rate
    except (ValueError, TypeError):
        wb_rate_all[wr] = None

# Get Section A plan UUIDs from v3 Charges sign-off
sec_a_rows = df_cs[df_cs["source_section"] == "A - Workbook confirmed"]
sec_a_uuids = set(sec_a_rows["plan_uuid"].dropna())
p(f"  Section A plans total: {len(sec_a_uuids)}")

# Find the 69 with no positive rate
zero_rate_uuids = []
for uuid in sec_a_uuids:
    pl = plans_by_uuid.get(uuid)
    if not pl: continue
    pn = normalise_polno(pl.get("policy_number"))
    rate = wb_rate_all.get(pn)
    if rate is None or rate == 0.0:
        zero_rate_uuids.append(uuid)

p(f"  Section A plans with zero/missing workbook rate: {len(zero_rate_uuids)}")

# Categorize reasons
reason_ctr = Counter()
for uuid in zero_rate_uuids:
    pl = plans_by_uuid.get(uuid, {})
    pn = normalise_polno(pl.get("policy_number"))
    rate = wb_rate_all.get(pn)
    status = pl.get("status", "unknown")
    if rate is None:
        reason = "Workbook rate missing (polno not found)"
    elif rate == 0.0:
        reason = "Workbook rate zero"
    else:
        reason = "Unexpected"
    if status != "active":
        reason += f" + plan {status}"
    reason_ctr[reason] += 1

p(f"  Reason breakdown:")
for reason, cnt in reason_ctr.most_common():
    p(f"    {cnt:4d}  {reason}")
p()

# Check: are these 69 currently in Section C in v3?
sec_c_rows = df_cs[df_cs["source_section"] == "C - Under advice, no charge any source"]
sec_c_uuids = set(sec_c_rows["plan_uuid"].dropna())
overlap = set(str(u) for u in zero_rate_uuids) & sec_c_uuids
p(f"  Zero-rate Section A plans also in Section C (v3): {len(overlap)}")
p(f"  -> These should NOT also be in Section C (workbook-confirmed, just zero rate)")

# ─── 4. Plannr charge reconciliation ─────────────────────────────────────────

p()
p("=" * 60)
p("INVESTIGATION 4: Plannr charge reconciliation (must sum to 330)")
p("=" * 60)

# Active ongoing Plannr charges
active_ongoing = [ch for ch in charges
                  if ch.get("recurring") and
                  (not ch.get("end_date") or str(ch.get("end_date","")) >= GO_LIVE)]
p(f"  Active ongoing Plannr charges: {len(active_ongoing)}")

# Which plan_uuids are in Section A?
sec_a_plan_uuids_full = set(sec_a_rows["plan_uuid"].dropna())

# Load Rathbones matched UUIDs
df_rath = pd.read_excel(os.path.join(DATA_ROOT, "Jim list  - Rathbones.xlsx"), dtype=str)
rath_file_norms = {normalise_polno(row["Fund Number"]) for _, row in df_rath.iterrows() if normalise_polno(row["Fund Number"])}
rath_plans_plannr = {pl["uuid"] for pl in plans if get_provider_uuid(pl) in RATH_PROVIDER_UUIDS}
rath_matched_uuids = {pl["uuid"] for pl in plans
                      if get_provider_uuid(pl) in RATH_PROVIDER_UUIDS
                      and normalise_polno(pl.get("policy_number")) in rath_file_norms}

# For each active ongoing charge, determine disposition
dispositions = Counter()
dup_plan_uuids_seen = set()

for ch in active_ongoing:
    plan_uuid = ch.get("_plan_uuid")
    pl = plans_by_uuid.get(plan_uuid, {}) if plan_uuid else {}
    status = pl.get("status", "unknown")

    if plan_uuid in sec_a_plan_uuids_full:
        dispositions["Replaced by workbook charge (Section A)"] += 1
    elif plan_uuid in rath_matched_uuids:
        dispositions["Replaced by Rathbones charge"] += 1
    elif status in ("inactive", "transferred"):
        dispositions["Dropped - plan inactive/transferred"] += 1
    else:
        # Check if there's another charge for the same plan (duplicate)
        # We'll count the disposition at the end
        dispositions["Kept as Section B / carried over"] += 1

p(f"\n  Reconciliation:")
total = 0
for disp, cnt in sorted(dispositions.items()):
    p(f"    {cnt:4d}  {disp}")
    total += cnt
p(f"    {'----'}")
p(f"    {total:4d}  TOTAL (expected 330)")
p()

# Investigate Section B count mismatch
# In v3: only 17 charges in Charges (all) with basis 'Plannr ongoing'
# But dispositions says "Kept as Section B" = ?
# Check how many distinct plan UUIDs are in the "Kept" bucket
kept_charges = [ch for ch in active_ongoing
                if ch.get("_plan_uuid") not in sec_a_plan_uuids_full
                and ch.get("_plan_uuid") not in rath_matched_uuids
                and plans_by_uuid.get(ch.get("_plan_uuid"), {}).get("status") not in ("inactive", "transferred")]

kept_plan_uuids = [ch.get("_plan_uuid") for ch in kept_charges]
kept_unique = set(kept_plan_uuids)
p(f"  'Kept' charges: {len(kept_charges)} total, {len(kept_unique)} unique plan UUIDs")
# Are there plans with multiple kept charges?
from collections import Counter as C
dup_ct = C(kept_plan_uuids)
multi_charge_plans = {k: v for k, v in dup_ct.items() if v > 1}
p(f"  Plans with multiple Plannr charges in 'kept': {len(multi_charge_plans)}")
p(f"    Multi-charge plan counts: {sorted(multi_charge_plans.values(), reverse=True)[:10]}")
p()

# Check the "1 charge per plan guard" — in v3 we took only the first charge per plan
# So Section B = unique plan UUIDs in kept bucket
p(f"  Expected Section B rows (unique plans): {len(kept_unique)}")
p(f"  Actual Section B rows in v3 Charges (all): {len(df_ca[df_ca['basis']=='Plannr ongoing'])}")
p()

# Investigate "Dropped" — status count
dropped = [ch for ch in active_ongoing
           if ch.get("_plan_uuid") not in sec_a_plan_uuids_full
           and ch.get("_plan_uuid") not in rath_matched_uuids
           and plans_by_uuid.get(ch.get("_plan_uuid"), {}).get("status") in ("inactive", "transferred")]
p(f"  Dropped (inactive/transferred) charges: {len(dropped)}")
p(f"  Status breakdown: {Counter(plans_by_uuid.get(ch.get('_plan_uuid'),{}).get('status','?') for ch in dropped)}")

# Check for charges on plans not in plans_dedup
no_plan = [ch for ch in active_ongoing if ch.get("_plan_uuid") not in plans_by_uuid]
p(f"  Active ongoing charges on plans NOT in plans_dedup: {len(no_plan)}")
if no_plan:
    for ch in no_plan[:3]:
        p(f"    charge_uuid=...{ch['uuid'][-8:]} plan_uuid=...{str(ch.get('_plan_uuid','?'))[-8:]}")
p()

# Summary reconciliation table
p("  Full reconciliation:")
repl_wb = sum(1 for ch in active_ongoing if ch.get("_plan_uuid") in sec_a_plan_uuids_full)
repl_rath = sum(1 for ch in active_ongoing
                if ch.get("_plan_uuid") not in sec_a_plan_uuids_full
                and ch.get("_plan_uuid") in rath_matched_uuids)
dropped_n = len(dropped)
no_plan_n = len(no_plan)
kept_n = len(kept_charges)
kept_unique_n = len(kept_unique)
multi_n = len(multi_charge_plans)
p(f"    {repl_wb:4d}  Replaced by workbook charge (Section A)")
p(f"    {repl_rath:4d}  Replaced by Rathbones charge")
p(f"    {kept_n:4d}  Kept charges (by plan, before dedup)")
p(f"    {kept_unique_n:4d}  Unique plans in Section B (after 1-per-plan guard)")
p(f"    {multi_n:4d}  Plans with >1 active Plannr charge (multi-dropped)")
p(f"    {dropped_n:4d}  Dropped - plan inactive/transferred")
p(f"    {no_plan_n:4d}  Active charges on plans not in plans_dedup")
p(f"    {'----'}")
p(f"    {repl_wb + repl_rath + kept_n + dropped_n + no_plan_n:4d}  TOTAL ({repl_wb}+{repl_rath}+{kept_n}+{dropped_n}+{no_plan_n})")
p()

# ─── Summary for v4 fixes ────────────────────────────────────────────────────

p("=" * 60)
p("SUMMARY: changes needed for v4")
p("=" * 60)
p()
p("1. Rate units: verify from sample above")
p("2. Placeholder policy numbers:")
p(f"   - {len(placeholder_plans)} plans to blank policy_number")
p(f"   - {len(new_nopolno_exceptions)} new 'No policy number' exceptions")
p(f"3. Zero-rate Section A plans: {len(zero_rate_uuids)} plans")
p(f"   - Keep in Section A (workbook confirmed, zero rate)")
p(f"   - Remove from Section C (they were incorrectly double-counted)")
p(f"4. Charge reconciliation:")
p(f"   - {repl_wb} workbook replacements")
p(f"   - {repl_rath} Rathbones replacements")
p(f"   - {kept_unique_n} Section B (unique plans)")
p(f"   - {len(dropped)} dropped (inactive/transferred)")
p(f"   - {no_plan_n} on plans not in dedup")
p(f"   - {kept_n - kept_unique_n} dropped by 1-per-plan guard (multi-charge plans)")
p()
