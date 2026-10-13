"""
Build hfp_signoff_pack_v5.xlsx from v2 base + snapshot + provider files.

All v4 fixes plus:
- Fix 1: Value-weighted rate report per charge group (console only)
- Fix 2: Remove 69 zero-rate Section A plans from Sections A and Charges (all);
          move them to Section C only with AR_comment
- Fix 3: "No policy number" exceptions - active plans only
- Fix 4: Plannr 330-charge reconciliation table (console only)

Rule 9: aggregates / UUIDs only in console. No PII.
"""
import shutil, json, os, re, sys
from collections import Counter, defaultdict
from datetime import date
import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter

DATA_ROOT = r"C:\dev\AR Migration"
V2_PATH   = os.path.join(DATA_ROOT, "hfp_signoff_pack_v2.xlsx")
V5_PATH   = os.path.join(DATA_ROOT, "hfp_signoff_pack_v5.xlsx")
SNAP_DIR  = os.path.join(DATA_ROOT, "snapshot_20261012T090906Z")
WB_PATH   = os.path.join(DATA_ROOT, "Hoban_Financial_Planning_Ltd_260922_Payment_Rates_113727875.xlsx")
RATH_PATH = os.path.join(DATA_ROOT, "Jim list  - Rathbones.xlsx")
MAP_PATH  = os.path.join(DATA_ROOT, "mappings_hfp.xlsx")

def p(*a): print(*a, flush=True)

# ── helpers ──────────────────────────────────────────────────────────────────

def normalise_polno(s):
    if s is None: return None
    s = re.sub(r"[\s\-]", "", str(s).strip())
    s = s.lstrip("0")
    return s.upper() if s else None

def pence_to_pounds(amt):
    try:
        return int(str(amt)) / 100.0
    except (TypeError, ValueError):
        return 0.0

PLACEHOLDER_UPPER = {
    "TBC", "NA", "N/A", "NOTKNOWN", "UNKNOWN", "NOTONFILE",
    "NOTSET", "NONE", "NOT SET", "NOTON", "NOTAPPLICABLE",
}

def is_placeholder(pn_raw):
    if not pn_raw or str(pn_raw).strip() in ("", "None", "nan"):
        return True
    pn_str = str(pn_raw).strip()
    if re.fullmatch(r"[\-\s\.]+", pn_str):
        return True
    norm = re.sub(r"[\s\-]", "", pn_str).upper()
    if norm in PLACEHOLDER_UPPER:
        return True
    if len(norm) <= 5 and not re.match(r"^\d{4,5}$|^[A-Z]{2}\d{3,}$", norm):
        return True
    if pn_str.upper().replace(" ", "").replace("-", "") in PLACEHOLDER_UPPER:
        return True
    return False

def get_provider_uuid(pl):
    prov = pl.get("provider")
    if isinstance(prov, dict): return prov.get("uuid")
    return None

def get_provider_name(pl):
    prov = pl.get("provider")
    if isinstance(prov, dict): return prov.get("name")
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

def plan_type_to_rmar(plan_type):
    if plan_type in ("protection_policy", "group_protection"):
        return {"type": "Commission", "rmar_type": "Commission"}
    return {"type": "Fee", "rmar_type": "Adviser Charge"}

def d15_defaults(plan_type):
    d = plan_type_to_rmar(plan_type)
    d.update({"tax_exempt": "Yes", "advice_type": "Independent", "exclude_from_rmar": "No"})
    return d

# ── load data ─────────────────────────────────────────────────────────────────

p("Loading plans_dedup.json...")
with open(os.path.join(SNAP_DIR, "plans_dedup.json")) as f:
    plans_raw = json.load(f)
plans = list(plans_raw.values())
plans_by_uuid = {pl["uuid"]: pl for pl in plans}
p(f"  Plans: {len(plans)}")

p("Loading charges_dedup.json...")
with open(os.path.join(SNAP_DIR, "charges_dedup.json")) as f:
    charges_raw = json.load(f)
charges = list(charges_raw.values())
p(f"  Charges: {len(charges)}")

p("Loading account_details.json...")
with open(os.path.join(SNAP_DIR, "account_details.json")) as f:
    account_details = json.load(f)
p(f"  Accounts: {len(account_details)}")

p("Loading v2 Charges sign-off tab...")
df_cs_v2 = pd.read_excel(V2_PATH, sheet_name="Charges sign-off", dtype=str)
p(f"  Charges sign-off rows: {len(df_cs_v2)}")

p("Loading v2 Plans tab...")
df_plans_v2 = pd.read_excel(V2_PATH, sheet_name="Plans", dtype=str)
p(f"  Plans tab rows: {len(df_plans_v2)}")

p("Loading v2 Exceptions tab...")
df_ex_v2 = pd.read_excel(V2_PATH, sheet_name="Exceptions", dtype=str)
p(f"  Exceptions rows: {len(df_ex_v2)}")

p("Loading v2 Provider-only plans tab...")
df_pop_v2 = pd.read_excel(V2_PATH, sheet_name="Provider-only plans", dtype=str)
p(f"  Provider-only plans rows: {len(df_pop_v2)}")

p("Loading workbook Payment Rates Summary...")
df_wb = pd.read_excel(WB_PATH, sheet_name="Payment Rates Summary", dtype=str)
p(f"  Workbook rows: {len(df_wb)}")

p("Loading Rathbones file...")
df_rath = pd.read_excel(RATH_PATH, dtype=str)
p(f"  Rathbones rows: {len(df_rath)}")

p("Loading expenditure mappings...")
df_exp_map = pd.read_excel(MAP_PATH, sheet_name="Expenditure", dtype=str)
p(f"  Expenditure mappings: {len(df_exp_map)}")

# ── workbook rate map ─────────────────────────────────────────────────────────

p("\nBuilding workbook rate map...")
wb_rate_map = {}   # normalised polno -> decimal rate (positive only)
wb_rate_all = {}   # includes zero rates
for _, row in df_wb.iterrows():
    wr = normalise_polno(row["Wrapper Ref"])
    if not wr: continue
    try:
        rate = float(row["Annual Payment - Investments"])
        wb_rate_all[wr] = rate
        if rate > 0:
            wb_rate_map[wr] = rate
    except (ValueError, TypeError):
        pass

p(f"  Positive-rate wrapper refs: {len(wb_rate_map)}")
p(f"  Zero-rate wrapper refs:     {sum(1 for v in wb_rate_all.values() if v == 0.0)}")

# ── Rathbones structures ──────────────────────────────────────────────────────

RATH_PROVIDER_UUIDS = {
    "d2c551bd-a2d1-4223-a3c6-d28a9d92ff91",
    "de8a8f76-44d6-41d8-994b-f800ba05b1f5",
    "5c086784-b183-4964-bd38-bf8410275068",
    "5d1737ed-195a-4323-8375-27bea2b0d6ec",
}
RATH_VALUATION_DATE = "2026-10-12"

p("\nBuilding Rathbones structures...")
rath_file_map = {}
for _, row in df_rath.iterrows():
    pn = normalise_polno(row["Fund Number"])
    if pn:
        rath_file_map[pn] = row

rath_plans_in_plannr = [pl for pl in plans if get_provider_uuid(pl) in RATH_PROVIDER_UUIDS]
rath_planno_to_pl = {}
for pl in rath_plans_in_plannr:
    pn = normalise_polno(pl.get("policy_number"))
    if pn:
        rath_planno_to_pl[pn] = pl

rath_matched_uuids = set()
for pl in rath_plans_in_plannr:
    pn = normalise_polno(pl.get("policy_number"))
    if pn and pn in rath_file_map:
        rath_matched_uuids.add(pl["uuid"])

rath_anomaly_uuids = set()
for pl in plans:
    if pl["uuid"] in rath_matched_uuids: continue
    pn = normalise_polno(pl.get("policy_number"))
    if pn and pn in rath_file_map:
        rath_anomaly_uuids.add(pl["uuid"])

rath_only_polnos = set(rath_file_map.keys()) - set(rath_planno_to_pl.keys())
rath_plannr_only_polnos = set(rath_planno_to_pl.keys()) - set(rath_file_map.keys())
rath_plannr_only_uuids = {pl["uuid"] for pl in rath_plans_in_plannr
                          if normalise_polno(pl.get("policy_number")) in rath_plannr_only_polnos}
rath_no_polno_uuids = {pl["uuid"] for pl in rath_plans_in_plannr
                       if not normalise_polno(pl.get("policy_number"))}

p(f"  Rathbones in Plannr: {len(rath_plans_in_plannr)}")
p(f"  Matched (file+Plannr): {len(rath_matched_uuids)}")
p(f"  Anomaly: {len(rath_anomaly_uuids)}")
p(f"  Rathbones-only (file, not Plannr): {len(rath_only_polnos)}")
p(f"  Plannr-only Rathbones: {len(rath_plannr_only_uuids)}")

# ── active ongoing charges map ────────────────────────────────────────────────

p("\nBuilding active ongoing charges map...")
GO_LIVE = "2026-11-02"
plan_to_active_charges = defaultdict(list)
for ch in charges:
    plan_uuid = ch.get("_plan_uuid")
    if not plan_uuid: continue
    end_date = ch.get("end_date")
    is_active = not end_date or str(end_date) >= GO_LIVE
    if ch.get("recurring") and is_active:
        plan_to_active_charges[plan_uuid].append(ch)

p(f"  Plans with active ongoing charges: {len(plan_to_active_charges)}")

# ── Section A plan UUIDs (from v2 Charges sign-off) ──────────────────────────

sec_a_rows = df_cs_v2[df_cs_v2["source_section"] == "A - Workbook confirmed"]
sec_a_plan_uuids = set(sec_a_rows["plan_uuid"].dropna())
p(f"\nSection A plans in v2 (all, inc. zero-rate): {len(sec_a_plan_uuids)}")

sec_a_plan_rate = {}  # plan_uuid -> decimal rate (positive only)
for uuid in sec_a_plan_uuids:
    pl = plans_by_uuid.get(uuid)
    if pl:
        pn = normalise_polno(pl.get("policy_number"))
        rate = wb_rate_map.get(pn)
        if rate:
            sec_a_plan_rate[uuid] = rate

sec_a_zero_rate_uuids = sec_a_plan_uuids - set(sec_a_plan_rate.keys())
p(f"  Section A positive-rate (v5 Section A): {len(sec_a_plan_rate)}")
p(f"  Section A zero-rate (v5 Section C only): {len(sec_a_zero_rate_uuids)}")

# ── Fix 4: Plannr charge reconciliation table ─────────────────────────────────

p()
p("=" * 60)
p("FIX 4: Plannr charge reconciliation")
p("=" * 60)

# All charges
total_charges = len(charges)
non_recurring = [ch for ch in charges if not ch.get("recurring")]
recurring = [ch for ch in charges if ch.get("recurring")]
ended = [ch for ch in recurring if ch.get("end_date") and str(ch.get("end_date","")) < GO_LIVE]
active_ongoing = [ch for ch in recurring if not ch.get("end_date") or str(ch.get("end_date","")) >= GO_LIVE]

p(f"  Total charges in snapshot:       {total_charges}")
p(f"  Non-recurring (one-off):         {len(non_recurring)}")
p(f"  Recurring:                       {len(recurring)}")
p(f"    Ended before go-live:          {len(ended)}")
p(f"    Active ongoing:                {len(active_ongoing)}")
p(f"  Expected active ongoing:         330")
if len(active_ongoing) != 330:
    p(f"  Discrepancy: {len(active_ongoing)} found vs 330 expected ({len(active_ongoing)-330:+d})")
p()

# Classify each active ongoing charge
recon = Counter()
for ch in active_ongoing:
    plan_uuid = ch.get("_plan_uuid")
    pl = plans_by_uuid.get(plan_uuid, {}) if plan_uuid else {}
    status = pl.get("status", "unknown")

    if plan_uuid in sec_a_plan_rate:
        recon["Replaced by workbook charge (Section A)"] += 1
    elif plan_uuid in rath_matched_uuids:
        recon["Replaced by Rathbones charge"] += 1
    elif plan_uuid in sec_a_zero_rate_uuids:
        recon["Dropped - zero-rate workbook plan (Section C, no charge built)"] += 1
    elif status == "active":
        recon["Kept in Section B (active, no workbook/Rathbones)"] += 1
    elif status in ("inactive", "transferred"):
        recon["Dropped - plan inactive or transferred"] += 1
    elif not plan_uuid or plan_uuid not in plans_by_uuid:
        recon["Dropped - plan not in plans_dedup"] += 1
    else:
        recon[f"Dropped - other (status={status})"] += 1

p("  Reconciliation of active ongoing charges:")
recon_total = 0
for disposition, cnt in sorted(recon.items(), key=lambda x: -x[1]):
    p(f"    {cnt:4d}  {disposition}")
    recon_total += cnt
p(f"    {'----'}")
p(f"    {recon_total:4d}  TOTAL (expected 330)")
if recon_total != 330:
    p(f"  NOTE: {recon_total} found vs 330 expected. Gap of {recon_total-330}.")
    p(f"  This is consistent with the snapshot containing {len(active_ongoing)} active")
    p(f"  ongoing charges vs the 330 in the v2 Charges (all) tab. The v2 build")
    p(f"  likely excluded {len(active_ongoing)-330} charge(s) with an additional filter")
    p(f"  not replicated here (possibly duplicate-polno plans or edge cases).")
p()

# ── per-plan charge summary ───────────────────────────────────────────────────

p("Building per-plan charge summary...")
plan_charge_info = {}

for uuid, rate in sec_a_plan_rate.items():
    plan_charge_info[uuid] = {"rate": rate, "source": "workbook"}

for uuid in rath_matched_uuids:
    if uuid in plan_charge_info: continue
    chgs = plan_to_active_charges.get(uuid, [])
    if chgs:
        ch = chgs[0]
        pct = ch.get("percentage")
        try:
            rate = float(pct) / 100.0
        except (TypeError, ValueError):
            continue
        plan_charge_info[uuid] = {"rate": rate, "source": "Rathbones", "charge_uuid": ch["uuid"]}

for pl in plans:
    uuid = pl["uuid"]
    if uuid in plan_charge_info: continue
    chgs = plan_to_active_charges.get(uuid, [])
    if chgs:
        ch = chgs[0]
        pct = ch.get("percentage")
        try:
            rate = float(pct) / 100.0
        except (TypeError, ValueError):
            continue
        plan_charge_info[uuid] = {"rate": rate, "source": "Plannr", "charge_uuid": ch["uuid"]}

p(f"  Plans with charge info: {len(plan_charge_info)}")
source_ctr = Counter(v["source"] for v in plan_charge_info.values())
p(f"  By source: {dict(source_ctr)}")

# ── IAI ───────────────────────────────────────────────────────────────────────

p("\nComputing indicative annual income...")
plan_value_pounds = {}
plan_iai = {}

for pl in plans:
    uuid = pl["uuid"]
    val_dict = pl.get("value")
    if isinstance(val_dict, dict):
        pounds = pence_to_pounds(val_dict.get("amount", 0))
    else:
        pounds = 0.0
    plan_value_pounds[uuid] = pounds
    ci = plan_charge_info.get(uuid)
    plan_iai[uuid] = pounds * ci["rate"] if ci else 0.0

total_iai_before = 19_310_879.66
total_iai_after = sum(plan_iai.values())
p(f"  IAI BEFORE fix (v2, pence scale): GBP {total_iai_before:,.2f}")
p(f"  IAI AFTER  fix (v5, pounds):      GBP {total_iai_after:,.2f}")

# ── build Charges (all) rows ──────────────────────────────────────────────────

p("\nBuilding Charges (all) rows...")
charges_all_rows = []

def get_plan_client_uuid(pl):
    owners = pl.get("_owners", [])
    if isinstance(owners, list) and owners:
        return owners[0].get("client_uuid")
    return None

plan_to_adviser = {}
plan_to_client = {}
for ch in charges:
    pid = ch.get("_plan_uuid")
    if not pid: continue
    acc = ch.get("account")
    if isinstance(acc, dict):
        plan_to_client[pid] = acc.get("uuid")
    seller = ch.get("seller")
    if isinstance(seller, dict):
        plan_to_adviser[pid] = seller.get("uuid")

def build_charge_row(plan_uuid, rate_decimal, freq, pct_freq, start_date,
                     payment_origin, basis, notes="", source_uuid=None, end_date=None):
    pl = plans_by_uuid.get(plan_uuid, {})
    plan_type = pl.get("type", "investment")
    d15 = d15_defaults(plan_type)
    val_pounds = plan_value_pounds.get(plan_uuid, 0.0)
    iai = val_pounds * rate_decimal
    client_uuid = plan_to_client.get(plan_uuid) or get_plan_client_uuid(pl)
    adviser_uuid = plan_to_adviser.get(plan_uuid)
    return {
        "uuid": source_uuid or "",
        "plan_uuid": plan_uuid,
        "client_uuid": client_uuid or "",
        "adviser_uuid": adviser_uuid or "",
        "payment_origin": payment_origin,
        "type": d15["type"],
        "category": "ongoing",
        "recurring": "True",
        "frequency": freq,
        "start_date": start_date,
        "end_date": end_date or "",
        "amount": "",
        "percentage": round(rate_decimal * 100, 6),
        "percentage_frequency": pct_freq,
        "tax": "",
        "tax_exempt": d15["tax_exempt"],
        "exclude_from_rmar": d15["exclude_from_rmar"],
        "advice_type": d15["advice_type"],
        "service_type": "ongoing_service",
        "rmar_type": d15["rmar_type"],
        "premium_uuid": "",
        "advice_area_deposit": "",
        "advice_area_non_regulated": "",
        "advice_area_non_rmar": "",
        "advice_area_non_investment_insurance": "",
        "advice_area_regulated_mortgage": "",
        "advice_area_retail_investment": "",
        "advice_area_other_fca_regulated": "",
        "advice_area_unknown": "",
        "notes": notes,
        "source_uuid": source_uuid or "",
        "basis": basis,
        "indicative_annual_income": round(iai, 2),
        "valuation_source": "Plannr",
    }

# 1. Workbook charges — positive-rate Section A plans only (Fix 2: zero-rate excluded)
for uuid, rate in sorted(sec_a_plan_rate.items()):
    row = build_charge_row(
        plan_uuid=uuid, rate_decimal=rate, freq="monthly", pct_freq="annually",
        start_date=GO_LIVE, payment_origin="Provider", basis="workbook",
    )
    charges_all_rows.append(row)

# 2. Rathbones charges
rath_charge_plan_uuids = set()
for uuid in sorted(rath_matched_uuids):
    if uuid in plan_charge_info and plan_charge_info[uuid]["source"] == "Rathbones":
        ci = plan_charge_info[uuid]
        chgs = plan_to_active_charges.get(uuid, [])
        if not chgs: continue
        ch = chgs[0]
        row = build_charge_row(
            plan_uuid=uuid, rate_decimal=ci["rate"], freq="quarterly", pct_freq="annually",
            start_date="2027-01-31", payment_origin="Provider", basis="Rathbones",
            notes="Rathbones: annual %, paid quarterly, first payment 31/01/2027",
            source_uuid=ch["uuid"],
        )
        charges_all_rows.append(row)
        rath_charge_plan_uuids.add(uuid)

# 3. Plannr ongoing (Section B) — active plans only, not workbook or Rathbones
plannr_remaining_uuids = set()
for pl in plans:
    uuid = pl["uuid"]
    if pl.get("status") != "active": continue
    if uuid in sec_a_plan_uuids: continue   # includes zero-rate plans (Fix 2)
    if uuid in rath_matched_uuids: continue
    ci = plan_charge_info.get(uuid)
    if ci and ci["source"] == "Plannr":
        chgs = plan_to_active_charges.get(uuid, [])
        if not chgs: continue
        ch = chgs[0]
        row = build_charge_row(
            plan_uuid=uuid, rate_decimal=ci["rate"],
            freq=ch.get("frequency", "monthly"),
            pct_freq=ch.get("percentage_frequency", "annually"),
            start_date=ch.get("start_date", GO_LIVE) or GO_LIVE,
            payment_origin=ch.get("payment_origin", "Provider"),
            basis="Plannr ongoing", source_uuid=ch["uuid"],
        )
        charges_all_rows.append(row)
        plannr_remaining_uuids.add(uuid)

workbook_count = sum(1 for r in charges_all_rows if r["basis"] == "workbook")
rathbones_count = sum(1 for r in charges_all_rows if r["basis"] == "Rathbones")
plannr_count = sum(1 for r in charges_all_rows if r["basis"] == "Plannr ongoing")
p(f"  Charges (all) rows: {len(charges_all_rows)}")
p(f"    workbook:        {workbook_count}")
p(f"    Rathbones:       {rathbones_count}")
p(f"    Plannr ongoing:  {plannr_count}")

# ── Fix 1: value-weighted rate report ────────────────────────────────────────

p()
p("=" * 60)
p("FIX 1: Value-weighted rate report (Charges all)")
p("=" * 60)
p(f"  {'Group':<22}  {'Count':>5}  {'Total_val_GBP':>15}  {'VW_avg_rate_%':>14}  {'Total_IAI_GBP':>14}")
p(f"  {'-'*22}  {'-'*5}  {'-'*15}  {'-'*14}  {'-'*14}")
for basis_label in ("workbook", "Rathbones", "Plannr ongoing"):
    grp = [r for r in charges_all_rows if r["basis"] == basis_label]
    n = len(grp)
    if n == 0:
        p(f"  {basis_label:<22}  {0:>5}  {'0.00':>15}  {'0.00000':>14}  {'0.00':>14}")
        continue
    total_val = sum(plan_value_pounds.get(str(r["plan_uuid"]), 0.0) for r in grp)
    wt_sum = sum(float(r["percentage"] or 0) * plan_value_pounds.get(str(r["plan_uuid"]), 0.0)
                 for r in grp)
    vw_avg = wt_sum / total_val if total_val > 0 else 0.0
    total_iai_grp = sum(float(r["indicative_annual_income"] or 0) for r in grp)
    p(f"  {basis_label:<22}  {n:>5}  {total_val:>15,.2f}  {vw_avg:>14.5f}  {total_iai_grp:>14,.2f}")
p()

# ── build Charges sign-off rows ───────────────────────────────────────────────

p("Building Charges sign-off rows...")

def get_plannr_rate_decimal(plan_uuid):
    chgs = plan_to_active_charges.get(plan_uuid, [])
    if not chgs: return None
    try:
        return float(chgs[0].get("percentage", 0)) / 100.0
    except (TypeError, ValueError):
        return None

def build_cs_row(source_section, plan_uuid, workbook_rate=None, plannr_rate=None,
                 final_rate=None, frequency=None, agency="Yes",
                 valuation_source="Plannr", basis="v5"):
    pl = plans_by_uuid.get(plan_uuid, {})
    val_pounds = plan_value_pounds.get(plan_uuid, 0.0)
    iai = val_pounds * final_rate if final_rate else 0.0
    wb_amount = val_pounds * workbook_rate if workbook_rate else None
    plannr_amount = val_pounds * plannr_rate if plannr_rate else None
    return {
        "source_section": source_section,
        "plan_uuid": plan_uuid,
        "plan_type": pl.get("type", ""),
        "policy_number": pl.get("policy_number", ""),
        "provider_name": get_provider_name(pl) or "",
        "workbook_rate": workbook_rate,
        "workbook_amount": round(wb_amount, 2) if wb_amount is not None else None,
        "plannr_rate_also": plannr_rate,
        "plannr_amount_also": round(plannr_amount, 2) if plannr_amount is not None else None,
        "final_rate": final_rate,
        "final_amount": round(iai, 2) if final_rate else None,
        "frequency": frequency,
        "agency_in_your_name": agency,
        "indicative_annual_income": round(iai, 2) if final_rate else None,
        "valuation_source": valuation_source,
        "basis": basis,
        "AR_decision": None,
        "AR_corrected_rate": None,
        "AR_corrected_amount": None,
        "AR_comment": None,
    }

cs_rows = []
SECTION_A_HEADER = "SECTION A - Workbook Confirmed Charges"
SECTION_B_HEADER = "SECTION B - Plannr Ongoing Charge, No Workbook Confirm (B1: agency=Yes, B2: agency=No)"
SECTION_C_HEADER = "SECTION C - Active Under-Advice, No Charge in Any Source"

cs_rows.append({"source_section": SECTION_A_HEADER})

# Section A: workbook — POSITIVE RATE ONLY (Fix 2: zero-rate plans excluded)
for uuid in sorted(sec_a_plan_rate.keys()):
    rate = sec_a_plan_rate[uuid]
    plannr_rate = get_plannr_rate_decimal(uuid)
    pl = plans_by_uuid.get(uuid, {})
    agency = "Yes" if pl.get("agency_in_your_name") else None
    row = build_cs_row(
        source_section="A - Workbook confirmed",
        plan_uuid=uuid,
        workbook_rate=rate,
        plannr_rate=plannr_rate,
        final_rate=rate,
        frequency="monthly",
        agency=agency,
        valuation_source="Plannr",
        basis="workbook",
    )
    cs_rows.append(row)

# Section A: Rathbones confirmed
rath_sec_a_uuids = set()
for uuid in sorted(rath_matched_uuids):
    chgs = plan_to_active_charges.get(uuid, [])
    if chgs:
        ch = chgs[0]
        try:
            rate_decimal = float(ch.get("percentage", 0)) / 100.0
        except (TypeError, ValueError):
            rate_decimal = None
        row = build_cs_row(
            source_section="A - Workbook confirmed",
            plan_uuid=uuid,
            workbook_rate=rate_decimal,
            plannr_rate=rate_decimal,
            final_rate=rate_decimal,
            frequency="quarterly",
            agency="Yes",
            valuation_source="Rathbones",
            basis="Rathbones",
        )
        row["provider_name"] = get_provider_name(plans_by_uuid.get(uuid, {})) or "Rathbones"
        cs_rows.append(row)
        rath_sec_a_uuids.add(uuid)

cs_rows.append({"source_section": SECTION_B_HEADER})

# Section B: Plannr ongoing (active plans only)
for uuid in sorted(plannr_remaining_uuids):
    pl = plans_by_uuid.get(uuid, {})
    plannr_rate = get_plannr_rate_decimal(uuid)
    agency = "Yes" if pl.get("agency_in_your_name") else "No"
    row = build_cs_row(
        source_section="B - Plannr ongoing, no workbook confirm",
        plan_uuid=uuid,
        workbook_rate=None,
        plannr_rate=plannr_rate,
        final_rate=plannr_rate,
        frequency="monthly",
        agency=agency,
        valuation_source="Plannr",
        basis="Plannr ongoing",
    )
    cs_rows.append(row)

cs_rows.append({"source_section": SECTION_C_HEADER})

# Section C: zero-rate Section A plans + Rathbones no-charge + other under-advice
sec_c_uuids = set()

# Fix 2: zero-rate Section A plans go here ONLY (not Section A)
for uuid in sec_a_zero_rate_uuids:
    sec_c_uuids.add(uuid)

# Rathbones matched with no charge
for uuid in rath_matched_uuids:
    if uuid not in rath_sec_a_uuids:
        sec_c_uuids.add(uuid)

# Other under-advice plans with no charge
for pl in plans:
    uuid = pl["uuid"]
    if uuid in plan_charge_info: continue
    if uuid in sec_c_uuids: continue
    if pl.get("under_advice") == "Yes" or pl.get("under_advice") is True:
        sec_c_uuids.add(uuid)

for uuid in sorted(sec_c_uuids):
    pl = plans_by_uuid.get(uuid, {})
    is_rath = uuid in rath_matched_uuids
    is_zero_rate = uuid in sec_a_zero_rate_uuids
    row = build_cs_row(
        source_section="C - Under advice, no charge any source",
        plan_uuid=uuid,
        valuation_source="Rathbones" if is_rath else "Plannr",
        basis="v5",
    )
    if is_zero_rate:
        # Fix 2: annotate with new message
        row["AR_comment"] = "Workbook rate zero or missing: confirm the rate, or no charge"
        row["workbook_rate"] = 0.0
    cs_rows.append(row)

sec_a_workbook_count = sum(1 for r in cs_rows if r.get("source_section") == "A - Workbook confirmed" and r.get("basis") == "workbook")
p(f"  Charges sign-off data rows: {sum(1 for r in cs_rows if r.get('plan_uuid'))}")
p(f"    Section A workbook:   {sec_a_workbook_count}")
p(f"    Section A Rathbones:  {len(rath_sec_a_uuids)}")
p(f"    Section B:            {len(plannr_remaining_uuids)}")
p(f"    Section C:            {len(sec_c_uuids)}")
p(f"      of which zero-rate annotated: {len(sec_a_zero_rate_uuids)}")

# ── Valuations ────────────────────────────────────────────────────────────────

p("\nBuilding Valuations rows...")
val_rows = []

rath_value_map = {}
for pl in plans:
    if pl["uuid"] in rath_matched_uuids:
        pn = normalise_polno(pl.get("policy_number"))
        if pn and pn in rath_file_map:
            rath_row = rath_file_map[pn]
            try:
                tv_val = None
                for col in rath_row.index:
                    if "total value" in str(col).lower():
                        tv_val = rath_row[col]
                        break
                val = float(tv_val) if tv_val is not None and str(tv_val).strip() not in ("", "nan") else 0.0
            except (ValueError, TypeError):
                val = 0.0
            rath_value_map[pl["uuid"]] = val

for pl in plans:
    uuid = pl["uuid"]
    policy_number = pl.get("policy_number", "")
    val_entries = pl.get("_valuations_latest", [])
    if not isinstance(val_entries, list): val_entries = []

    if uuid in rath_value_map:
        plannr_total_pounds = pence_to_pounds(
            pl.get("value", {}).get("amount", 0) if isinstance(pl.get("value"), dict) else 0)
        val_rows.append({
            "plan_uuid": uuid,
            "sub_account_uuid": "",
            "policy_number": policy_number,
            "valuation_date": RATH_VALUATION_DATE,
            "value": round(rath_value_map[uuid], 2),
            "source": "Rathbones",
            "plannr_value": round(plannr_total_pounds, 2),
        })
    elif val_entries:
        for entry in val_entries:
            sub_uuid = entry.get("_sub_account_uuid", "")
            valued_at = entry.get("valued_at", "")
            if valued_at and "T" in str(valued_at):
                valued_at = str(valued_at)[:10]
            val_pounds = pence_to_pounds(entry.get("value", {}).get("amount", 0))
            val_rows.append({
                "plan_uuid": uuid,
                "sub_account_uuid": sub_uuid,
                "policy_number": policy_number,
                "valuation_date": valued_at,
                "value": round(val_pounds, 2),
                "source": "Plannr",
                "plannr_value": "",
            })

p(f"  Valuation rows: {len(val_rows)}")

# ── Income & Employment ───────────────────────────────────────────────────────

p("Building Income & Employment rows...")
ie_rows = []

for acc_uuid, acc in account_details.items():
    income_list = acc.get("income", [])
    employment_list = acc.get("employment", [])
    if not income_list and not employment_list: continue
    emp = employment_list[0] if employment_list else {}
    for inc in income_list:
        amt = inc.get("amount", {})
        amt_pounds = pence_to_pounds(amt.get("amount", 0) if isinstance(amt, dict) else 0)
        emp_turnover = emp.get("turnover_history", [])
        turnover_year = emp_turnover[0].get("year", "") if emp_turnover else ""
        turnover_val = emp_turnover[0].get("value", {}) if emp_turnover else {}
        ie_rows.append({
            "client_uuid": acc_uuid,
            "name": inc.get("name", ""),
            "type": inc.get("type", ""),
            "type_other": inc.get("type_other", ""),
            "start_date": inc.get("start_date", ""),
            "end_date": inc.get("end_date", ""),
            "amount": round(amt_pounds, 2),
            "amount_type": inc.get("amount_type", ""),
            "frequency": inc.get("frequency", ""),
            "is_taxable": inc.get("is_taxable", ""),
            "employment_uuid": emp.get("uuid", ""),
            "employment_employer": emp.get("employer", ""),
            "employment_job_title": emp.get("job_title", ""),
            "employment_occupation_code": emp.get("occupation_code", ""),
            "employment_status": emp.get("status", ""),
            "employment_started_at": emp.get("started_at", ""),
            "employment_ended_at": emp.get("ended_at", ""),
            "employment_email": "",
            "employment_phone": "",
            "employment_address_line_1": "",
            "employment_address_line_2": "",
            "employment_address_line_3": "",
            "employment_address_line_4": "",
            "employment_address_town_city": "",
            "employment_address_county": "",
            "employment_address_postcode": "",
            "employment_address_country": "",
            "employment_is_primary": emp.get("primary", ""),
            "employment_is_business_owner": emp.get("is_business_owner", ""),
            "employment_percentage_of_business_owned": emp.get("percentage_of_business_owned", ""),
            "employment_directors": emp.get("directors", ""),
            "employment_business_year_end": emp.get("business_year_end", ""),
            "employment_shareholding_in_business": emp.get("shareholding_in_business", ""),
            "employment_profit": emp.get("profit", {}).get("amount", "") if isinstance(emp.get("profit"), dict) else emp.get("profit", ""),
            "employment_business_value": emp.get("business_value", {}).get("amount", "") if isinstance(emp.get("business_value"), dict) else emp.get("business_value", ""),
            "employment_accountant_company_name": emp.get("accountant_company_name", ""),
            "employment_accountant_email": emp.get("accountant_email", ""),
            "employment_accountant_telephone_number": emp.get("accountant_telephone_number", ""),
            "employment_accountant_address": emp.get("accountant_address") or "",
            "employment_turnover_history_year": turnover_year,
            "employment_turnover_history_value": pence_to_pounds(turnover_val.get("amount", 0)) if isinstance(turnover_val, dict) else "",
        })

p(f"  Income & Employment rows: {len(ie_rows)}")

# ── Expenses ──────────────────────────────────────────────────────────────────

p("Building Expenses rows...")
exp_map_lookup = {}
for _, mrow in df_exp_map.iterrows():
    src_sec = str(mrow.get("source_section", "") or "").strip()
    src_line = str(mrow.get("source_line", "") or "").strip()
    dst_sec = str(mrow.get("dest_section", "") or "").strip()
    dst_line = str(mrow.get("dest_line", "") or "").strip()
    is_custom = dst_line.upper() == "CUSTOM"
    if src_sec and src_line:
        exp_map_lookup[(src_sec, src_line)] = (dst_sec, dst_line, is_custom)

exp_rows = []
for acc_uuid, acc in account_details.items():
    for exp in acc.get("expenditure", []):
        src_sec = exp.get("section", "") or ""
        src_line = exp.get("name", "") or ""
        amt = exp.get("amount", {})
        amt_pounds = pence_to_pounds(amt.get("amount", 0) if isinstance(amt, dict) else 0)
        mapped = exp_map_lookup.get((src_sec, src_line))
        if mapped:
            dst_sec, dst_line, is_custom = mapped
        else:
            dst_sec, dst_line, is_custom = src_sec, src_line, False
        exp_rows.append({
            "client_uuid": acc_uuid,
            "life_stage_name": dst_sec,
            "life_stage_section_name": dst_sec,
            "expense_name": dst_line if not is_custom else src_line,
            "amount": round(amt_pounds, 2),
            "is_custom_expense": 1 if is_custom else 0,
            "external_reference_company": "",
            "external_reference_value": "",
        })

p(f"  Expenses rows: {len(exp_rows)}")

# ── Exceptions ────────────────────────────────────────────────────────────────

p("\nBuilding Exceptions rows...")

PLAN_TYPES_NEED_POLNO = {"investment", "pension", "protection_policy",
                          "group_protection", "annuity"}

ex_keep = df_ex_v2[~df_ex_v2["category"].isin(
    ["Duplicate policy number", "No policy number"]
)].copy()

# Rebuild duplicate exceptions (real groups only)
dup_ex_rows = []
polno_to_plans_dedup = defaultdict(list)
for pl in plans:
    pn = normalise_polno(pl.get("policy_number"))
    if pn and not is_placeholder(pl.get("policy_number")):
        polno_to_plans_dedup[pn].append(pl)

for pn, pls in polno_to_plans_dedup.items():
    if len(pls) < 2: continue
    all_clients = set()
    for pl in pls: all_clients.update(get_client_uuids(pl))
    if len(all_clients) < 2: continue
    if not any(pl.get("status") == "active" for pl in pls): continue
    for i, pl in enumerate(pls):
        dup_ex_rows.append({
            "category": "Duplicate policy number",
            "entity_type": "plan",
            "source_uuid": pl["uuid"],
            "detail": f"dup group {i+1} of {len(pls)}",
            "current_value": pn,
            "proposed_value": "AR to decide: merge, separate or correct",
            "AR_decision": None,
        })

# Fix 3: "No policy number" exceptions — active plans only
nopolno_before = 0
nopolno_rows = []
for pl in plans:
    pn = pl.get("policy_number")
    plan_type = pl.get("type", "")
    if plan_type not in PLAN_TYPES_NEED_POLNO: continue
    if not is_placeholder(pn): continue
    nopolno_before += 1
    if pl.get("status") != "active": continue  # Fix 3: active only
    nopolno_rows.append({
        "category": "No policy number",
        "entity_type": "plan",
        "source_uuid": pl["uuid"],
        "detail": f"type={plan_type}",
        "current_value": str(pn or ""),
        "proposed_value": "Add policy number",
        "AR_decision": None,
    })

p(f"  No-policy-number before active filter: {nopolno_before}")
p(f"  No-policy-number after  active filter: {len(nopolno_rows)}")

# Rathbones exceptions
rath_exc_rows = []
for uuid in rath_plannr_only_uuids:
    rath_exc_rows.append({
        "category": "Rathbones plan not on Rathbones list",
        "entity_type": "plan", "source_uuid": uuid,
        "detail": "In Plannr as Rathbones provider, not in Rathbones file",
        "current_value": "", "proposed_value": "Closed or transferred? AR to decide",
        "AR_decision": None,
    })
for uuid in rath_no_polno_uuids:
    rath_exc_rows.append({
        "category": "Rathbones plan not on Rathbones list",
        "entity_type": "plan", "source_uuid": uuid,
        "detail": "Rathbones provider plan with no policy number",
        "current_value": "", "proposed_value": "Add policy number or close",
        "AR_decision": None,
    })
for uuid in rath_anomaly_uuids:
    rath_exc_rows.append({
        "category": "Rathbones file: matched to non-Rathbones plan",
        "entity_type": "plan", "source_uuid": uuid,
        "detail": "Fund Number matches Rathbones file but plan has different provider in Plannr",
        "current_value": "", "proposed_value": "AR to investigate", "AR_decision": None,
    })

ex_rows_new = (
    ex_keep.to_dict("records") + dup_ex_rows + nopolno_rows + rath_exc_rows
)
ex_cat_ctr = Counter(r.get("category") for r in ex_rows_new)
p(f"  Total exceptions: {len(ex_rows_new)}")
for cat, cnt in sorted(ex_cat_ctr.items(), key=lambda x: -x[1]):
    p(f"    {cnt:4d}  {cat}")

# ── Provider-only plans ───────────────────────────────────────────────────────

p("\nBuilding Provider-only plans rows...")
pop_rows = df_pop_v2.to_dict("records")
for pn in sorted(rath_only_polnos):
    rath_row = rath_file_map.get(pn)
    if rath_row is None: continue
    try:
        tv_val = None
        for col in rath_row.index:
            if "total value" in str(col).lower():
                tv_val = rath_row[col]
                break
        val = float(tv_val) if tv_val is not None and str(tv_val).strip() not in ("", "nan") else 0.0
    except (ValueError, TypeError):
        val = 0.0
    fund_type = str(rath_row.get("Fund Type", "") or "")
    is_sipp = "sipp" in str(fund_type).lower()
    pop_rows.append({
        "policy_number": pn,
        "provider_name": "Rathbones (within SIPP wrap)" if is_sipp else "Rathbones",
        "wrapper_name": fund_type,
        "workbook_value": round(val, 2),
        "workbook_rate": None,
        "match_confidence": "name-only",
        "AR_decision": None,
    })
p(f"  Provider-only plans rows: {len(pop_rows)}")

# ── Plans tab update structures ───────────────────────────────────────────────

p("\nPreparing Plans tab update structures...")
rath_polno_norms = set()
for uuid in rath_matched_uuids:
    pl = plans_by_uuid.get(uuid, {})
    pn = normalise_polno(pl.get("policy_number"))
    if pn:
        rath_polno_norms.add(pn)

placeholder_blanked_count = sum(
    1 for pl in plans
    if pl.get("policy_number")
    and str(pl.get("policy_number", "")).strip() not in ("", "None", "nan")
    and is_placeholder(pl.get("policy_number"))
)
p(f"  Rathbones matched plans with real polno (under_advice/agency): {len(rath_polno_norms)}")
p(f"  Plans with non-blank placeholder polno to blank: {placeholder_blanked_count}")

# ── copy v2 -> v5 ─────────────────────────────────────────────────────────────

p(f"\nCopying v2 -> v5...")
shutil.copy2(V2_PATH, V5_PATH)
p(f"  v5 created: {V5_PATH}")

p("\nOpening v5 with openpyxl...")
wb = openpyxl.load_workbook(V5_PATH)

def clear_data_rows(ws, header_row=1):
    max_row = ws.max_row
    if max_row <= header_row: return
    ws.delete_rows(header_row + 1, max_row - header_row)

# ── write Charges (all) ───────────────────────────────────────────────────────

p("Writing Charges (all)...")
ws_ca = wb["Charges (all)"]
header_row_ca = [cell.value for cell in ws_ca[1]]
clear_data_rows(ws_ca, header_row=1)
for r_idx, row_data in enumerate(charges_all_rows, start=2):
    for c_idx, col_name in enumerate(header_row_ca, start=1):
        ws_ca.cell(row=r_idx, column=c_idx, value=row_data.get(col_name))
p(f"  Written {len(charges_all_rows)} charge rows")

# ── write Charges sign-off ────────────────────────────────────────────────────

p("Writing Charges sign-off...")
ws_cs = wb["Charges sign-off"]
header_row_cs = [cell.value for cell in ws_cs[1]]
clear_data_rows(ws_cs, header_row=1)
for r_idx, row_data in enumerate(cs_rows, start=2):
    for c_idx, col_name in enumerate(header_row_cs, start=1):
        ws_cs.cell(row=r_idx, column=c_idx, value=row_data.get(col_name))
p(f"  Written {len(cs_rows)} charge sign-off rows")

# ── write Valuations ──────────────────────────────────────────────────────────

p("Writing Valuations...")
ws_val = wb["Valuations"]
header_row_val = [cell.value for cell in ws_val[1]]
extra_cols = ["source", "plannr_value"]
for ec in extra_cols:
    if ec not in header_row_val:
        next_col = len(header_row_val) + 1
        ws_val.cell(row=1, column=next_col, value=ec)
        header_row_val.append(ec)
clear_data_rows(ws_val, header_row=1)
for r_idx, row_data in enumerate(val_rows, start=2):
    for c_idx, col_name in enumerate(header_row_val, start=1):
        ws_val.cell(row=r_idx, column=c_idx, value=row_data.get(col_name))
p(f"  Written {len(val_rows)} valuation rows")

# ── write Income & Employment ─────────────────────────────────────────────────

p("Writing Income & Employment...")
ws_ie = wb["Income & Employment"]
ie_headers_raw = [cell.value for cell in ws_ie[1]]
ie_col_start = 1
if ie_headers_raw and (ie_headers_raw[0] is None or "Unnamed" in str(ie_headers_raw[0])):
    ie_col_start = 2
ie_header_cols = ie_headers_raw[ie_col_start - 1:]
clear_data_rows(ws_ie, header_row=1)
ie_col_order = [c for c in ie_header_cols if c is not None and "Unnamed" not in str(c)]
for r_idx, row_data in enumerate(ie_rows, start=2):
    for local_idx, col_name in enumerate(ie_col_order):
        ws_ie.cell(row=r_idx, column=ie_col_start + local_idx, value=row_data.get(col_name))
p(f"  Written {len(ie_rows)} income & employment rows")

# ── write Expenses ────────────────────────────────────────────────────────────

p("Writing Expenses...")
ws_exp = wb["Expenses"]
exp_headers = [cell.value for cell in ws_exp[1]]
clear_data_rows(ws_exp, header_row=1)
exp_col_order = [c for c in exp_headers if c is not None]
for r_idx, row_data in enumerate(exp_rows, start=2):
    for c_idx, col_name in enumerate(exp_col_order, start=1):
        ws_exp.cell(row=r_idx, column=c_idx, value=row_data.get(col_name))
p(f"  Written {len(exp_rows)} expenses rows")

# ── write Exceptions ──────────────────────────────────────────────────────────

p("Writing Exceptions...")
ws_ex = wb["Exceptions"]
ex_headers = [cell.value for cell in ws_ex[1]]
clear_data_rows(ws_ex, header_row=1)
ex_col_order = [c for c in ex_headers if c is not None]
for r_idx, row_data in enumerate(ex_rows_new, start=2):
    for c_idx, col_name in enumerate(ex_col_order, start=1):
        ws_ex.cell(row=r_idx, column=c_idx, value=row_data.get(col_name))
p(f"  Written {len(ex_rows_new)} exception rows")

# ── write Provider-only plans ─────────────────────────────────────────────────

p("Writing Provider-only plans...")
ws_pop = wb["Provider-only plans"]
pop_headers = [cell.value for cell in ws_pop[1]]
clear_data_rows(ws_pop, header_row=1)
pop_col_order = [c for c in pop_headers if c is not None]
for r_idx, row_data in enumerate(pop_rows, start=2):
    for c_idx, col_name in enumerate(pop_col_order, start=1):
        ws_pop.cell(row=r_idx, column=c_idx, value=row_data.get(col_name))
p(f"  Written {len(pop_rows)} provider-only plan rows")

# ── update Plans tab ──────────────────────────────────────────────────────────

p("Updating Plans tab...")
ws_pl = wb["Plans"]
pl_headers = [cell.value for cell in ws_pl[1]]
pl_col_idx = {name: idx + 1 for idx, name in enumerate(pl_headers) if name}

polno_ci = pl_col_idx.get("policy_number")
ua_ci    = pl_col_idx.get("under_advice")
aia_ci   = pl_col_idx.get("agency_in_your_name")

plans_polno_blanked = 0
rath_ua_updated = 0

for row_idx in range(2, ws_pl.max_row + 1):
    pn_val = ws_pl.cell(row=row_idx, column=polno_ci).value if polno_ci else None

    if pn_val and is_placeholder(pn_val):
        ws_pl.cell(row=row_idx, column=polno_ci, value="")
        plans_polno_blanked += 1
        pn_val = None

    pn_norm = normalise_polno(pn_val) if pn_val else None
    if pn_norm and pn_norm in rath_polno_norms:
        if ua_ci:
            if ws_pl.cell(row=row_idx, column=ua_ci).value != "Yes":
                ws_pl.cell(row=row_idx, column=ua_ci, value="Yes")
                rath_ua_updated += 1
        if aia_ci:
            if ws_pl.cell(row=row_idx, column=aia_ci).value != "Yes":
                ws_pl.cell(row=row_idx, column=aia_ci, value="Yes")

p(f"  Placeholder policy numbers blanked: {plans_polno_blanked}")
p(f"  Rathbones plans under_advice/agency updated: {rath_ua_updated}")

# ── update Summary tab ────────────────────────────────────────────────────────

p("Updating Summary tab...")
ws_sum = wb["Summary"]

def set_summary_field(ws, field_name, value):
    for row in ws.iter_rows():
        for cell in row:
            if str(cell.value or "") == field_name:
                ws.cell(row=cell.row, column=cell.column + 1).value = value
                return True
    return False

set_summary_field(ws_sum, "Pack version", "v5")
set_summary_field(ws_sum, "Build date", str(date.today()))
set_summary_field(ws_sum, "Indicative annual income (GBP, estimate)", round(total_iai_after, 2))
set_summary_field(ws_sum, "Loadable ongoing charges", len(charges_all_rows))
set_summary_field(ws_sum, "Charges sign-off Section A (workbook)", len(sec_a_plan_rate) + len(rath_sec_a_uuids))
set_summary_field(ws_sum, "Charges sign-off Section B (Plannr ongoi", len(plannr_remaining_uuids))
set_summary_field(ws_sum, "Open exception items", len(ex_rows_new))
set_summary_field(ws_sum, "Provider-only plan rows", len(pop_rows))

# ── save v5 ───────────────────────────────────────────────────────────────────

p(f"\nSaving v5 to {V5_PATH}...")
wb.save(V5_PATH)
p("  Saved.")

# ── final report ─────────────────────────────────────────────────────────────

v5_size = os.path.getsize(V5_PATH)
p("\n" + "=" * 60)
p("FINAL REPORT (Rule 9: aggregates only)")
p("=" * 60)
p()

p("=== Fix 1: Value-weighted rate by charge group ===")
p(f"  {'Group':<22}  {'Count':>5}  {'Total_val_GBP':>15}  {'VW_avg_rate_%':>14}  {'Total_IAI_GBP':>14}")
p(f"  {'-'*22}  {'-'*5}  {'-'*15}  {'-'*14}  {'-'*14}")
for basis_label in ("workbook", "Rathbones", "Plannr ongoing"):
    grp = [r for r in charges_all_rows if r["basis"] == basis_label]
    n = len(grp)
    if n == 0:
        p(f"  {basis_label:<22}  {0:>5}  {'0.00':>15}  {'0.00000':>14}  {'0.00':>14}")
        continue
    total_val = sum(plan_value_pounds.get(str(r["plan_uuid"]), 0.0) for r in grp)
    wt_sum = sum(float(r["percentage"] or 0) * plan_value_pounds.get(str(r["plan_uuid"]), 0.0)
                 for r in grp)
    vw_avg = wt_sum / total_val if total_val > 0 else 0.0
    total_iai_grp = sum(float(r["indicative_annual_income"] or 0) for r in grp)
    p(f"  {basis_label:<22}  {n:>5}  {total_val:>15,.2f}  {vw_avg:>14.5f}  {total_iai_grp:>14,.2f}")
p()

p("=== Fix 3: No-policy-number exceptions ===")
p(f"  Before active filter: {nopolno_before}")
p(f"  After  active filter: {len(nopolno_rows)}")
p()

p("=== Fix 4: Plannr charge reconciliation ===")
p(f"  Active ongoing charges in snapshot: {len(active_ongoing)}")
p(f"  Expected: 330")
recon_total = 0
for disposition, cnt in sorted(recon.items(), key=lambda x: -x[1]):
    p(f"    {cnt:4d}  {disposition}")
    recon_total += cnt
p(f"    {'----'}")
p(f"    {recon_total:4d}  TOTAL")
p()

p("=== Exceptions by category (v5) ===")
for cat, cnt in sorted(ex_cat_ctr.items(), key=lambda x: -x[1]):
    p(f"  {cnt:4d}  {cat}")
p(f"  {'----'}")
p(f"  {sum(ex_cat_ctr.values()):4d}  TOTAL")
p()

p("=== Charges sign-off section counts ===")
p(f"  Section A workbook:   {sec_a_workbook_count}")
p(f"  Section A Rathbones:  {len(rath_sec_a_uuids)}")
p(f"  Section A total:      {sec_a_workbook_count + len(rath_sec_a_uuids)}")
p(f"  Section B:            {len(plannr_remaining_uuids)}")
p(f"  Section C:            {len(sec_c_uuids)}")
p(f"    of which zero-rate annotated: {len(sec_a_zero_rate_uuids)}")
p()

p(f"=== Total IAI ===")
p(f"  GBP {total_iai_after:,.2f}")
p()

p("=== v5 file ===")
p(f"  Size: {v5_size / 1024:.0f} KB")
p(f"  Path: {V5_PATH}")
p()
p("STOP - awaiting AR sign-off.")
