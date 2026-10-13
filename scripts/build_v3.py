"""
Build hfp_signoff_pack_v3.xlsx from v2 base + snapshot + provider files.

Applies:
- Fix 3: Pence->pounds conversion everywhere
- Fix 4: D18 workbook charges (Transact/Nucleus) for all 441 positive-rate plans
- Fix 5: Duplicate policy number exceptions - real groups only
- Step 2: Rathbones plans (matched, provider-only, Plannr-only exceptions)
- Step 3: Valuations, Income & Employment, Expenses tabs
- Step 4: No-policy-number exceptions filtered to relevant plan types

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
V3_PATH   = os.path.join(DATA_ROOT, "hfp_signoff_pack_v3.xlsx")
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
    """Convert Plannr pence amount (string or int) to float pounds."""
    try:
        return int(str(amt)) / 100.0
    except (TypeError, ValueError):
        return 0.0

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
    d.update({
        "tax_exempt": "Yes",
        "advice_type": "Independent",
        "exclude_from_rmar": "No",
    })
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

p("Loading v2 Plans tab (key columns only)...")
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

# ── build workbook rate map (Wrapper Ref -> rate) ──────────────────────────────

p("\nBuilding workbook rate map...")
wb_rate_map = {}    # wrapper_ref_norm -> rate (decimal)
wb_rate_all = {}    # includes zero rates
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

p(f"  Wrapper Refs with positive rate: {len(wb_rate_map)}")
p(f"  Wrapper Refs with zero rate: {sum(1 for v in wb_rate_all.values() if v == 0.0)}")

# ── build Rathbones match structures ─────────────────────────────────────────

RATH_PROVIDER_UUIDS = {
    "d2c551bd-a2d1-4223-a3c6-d28a9d92ff91",
    "de8a8f76-44d6-41d8-994b-f800ba05b1f5",
    "5c086784-b183-4964-bd38-bf8410275068",
    "5d1737ed-195a-4323-8375-27bea2b0d6ec",
}
RATH_VALUATION_DATE = "2026-10-12"

p("\nBuilding Rathbones match structures...")
rath_file_map = {}  # fund_number_norm -> row
for _, row in df_rath.iterrows():
    pn = normalise_polno(row["Fund Number"])
    if pn:
        rath_file_map[pn] = row

rath_plans_in_plannr = [pl for pl in plans
                        if get_provider_uuid(pl) in RATH_PROVIDER_UUIDS]
rath_planno_to_pl = {}
for pl in rath_plans_in_plannr:
    pn = normalise_polno(pl.get("policy_number"))
    if pn:
        rath_planno_to_pl[pn] = pl

# Matched = Rathbones-provider plan whose polno is in Rathbones file
rath_matched_uuids = set()
for pl in rath_plans_in_plannr:
    pn = normalise_polno(pl.get("policy_number"))
    if pn and pn in rath_file_map:
        rath_matched_uuids.add(pl["uuid"])

# Anomaly: file polno matches non-Rathbones plan
rath_anomaly_uuids = set()
for pl in plans:
    if pl["uuid"] in rath_matched_uuids:
        continue
    pn = normalise_polno(pl.get("policy_number"))
    if pn and pn in rath_file_map:
        rath_anomaly_uuids.add(pl["uuid"])

# Rathbones-only: in file, not in Plannr
rath_only_polnos = set(rath_file_map.keys()) - set(rath_planno_to_pl.keys())

# Plannr-only: Rathbones-provider plans not in file
rath_plannr_only_polnos = set(rath_planno_to_pl.keys()) - set(rath_file_map.keys())
rath_plannr_only_uuids = {pl["uuid"] for pl in rath_plans_in_plannr
                          if normalise_polno(pl.get("policy_number")) in rath_plannr_only_polnos}
# Also include Rathbones plans with no policy number
rath_no_polno_uuids = {pl["uuid"] for pl in rath_plans_in_plannr
                       if not normalise_polno(pl.get("policy_number"))}

p(f"  Rathbones in Plannr: {len(rath_plans_in_plannr)}")
p(f"  Matched (file+Plannr): {len(rath_matched_uuids)}")
p(f"  Anomaly (file but wrong provider in Plannr): {len(rath_anomaly_uuids)}")
p(f"  Rathbones-only (file, not Plannr): {len(rath_only_polnos)}")
p(f"  Plannr-only Rathbones: {len(rath_plannr_only_uuids)}")
p(f"  Rathbones plans no policy number: {len(rath_no_polno_uuids)}")

# ── build active ongoing charges map ─────────────────────────────────────────

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

# ── Section A plan UUIDs (from v2 Charges sign-off) ─────────────────────────

sec_a_rows = df_cs_v2[df_cs_v2["source_section"] == "A - Workbook confirmed"]
sec_a_plan_uuids = set(sec_a_rows["plan_uuid"].dropna())
p(f"\nSection A (workbook confirmed) plans: {len(sec_a_plan_uuids)}")

# For each Section A plan, get workbook rate
sec_a_plan_rate = {}  # plan_uuid -> rate (decimal) or None
for uuid in sec_a_plan_uuids:
    pl = plans_by_uuid.get(uuid)
    if pl:
        pn = normalise_polno(pl.get("policy_number"))
        rate = wb_rate_map.get(pn)  # positive rate only
        if rate:
            sec_a_plan_rate[uuid] = rate

p(f"  Section A plans with positive workbook rate: {len(sec_a_plan_rate)}")
p(f"  Section A plans with zero workbook rate: {len(sec_a_plan_uuids) - len(sec_a_plan_rate)}")

# ── per-plan charge summary for indicative income ────────────────────────────

p("\nBuilding per-plan charge summary...")
plan_charge_info = {}  # plan_uuid -> {rate_decimal, source, charge_uuid}

# 1. Workbook (Section A, positive rate) — highest priority
for uuid, rate in sec_a_plan_rate.items():
    plan_charge_info[uuid] = {"rate": rate, "source": "workbook"}

# 2. Rathbones matched (from Plannr charge)
for uuid in rath_matched_uuids:
    if uuid in plan_charge_info:
        continue  # already covered by workbook
    chgs = plan_to_active_charges.get(uuid, [])
    if chgs:
        ch = chgs[0]
        pct = ch.get("percentage")
        try:
            rate = float(pct) / 100.0  # percentage form -> decimal
        except (TypeError, ValueError):
            continue
        plan_charge_info[uuid] = {"rate": rate, "source": "Rathbones", "charge_uuid": ch["uuid"]}

# 3. Remaining Plannr ongoing (Section B)
for pl in plans:
    uuid = pl["uuid"]
    if uuid in plan_charge_info:
        continue
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

# ── indicative annual income per plan ────────────────────────────────────────

p("\nComputing indicative annual income (pence fix)...")
plan_value_pounds = {}  # plan_uuid -> float pounds
plan_iai = {}           # plan_uuid -> float pounds (indicative annual income)

for pl in plans:
    uuid = pl["uuid"]
    val_dict = pl.get("value")
    if isinstance(val_dict, dict):
        pounds = pence_to_pounds(val_dict.get("amount", 0))
    else:
        pounds = 0.0
    plan_value_pounds[uuid] = pounds
    ci = plan_charge_info.get(uuid)
    if ci:
        plan_iai[uuid] = pounds * ci["rate"]
    else:
        plan_iai[uuid] = 0.0

total_iai_before = 19_310_879.66  # from v2 summary (in pence × rate)
total_iai_after = sum(plan_iai.values())
p(f"  Indicative annual income BEFORE fix: £{total_iai_before:,.2f}")
p(f"  Indicative annual income AFTER fix:  £{total_iai_after:,.2f}")

# ── build Charges (all) rows ──────────────────────────────────────────────────

p("\nBuilding Charges (all) rows...")
charges_all_rows = []  # list of dicts matching Charges (all) columns

# Columns: uuid, plan_uuid, client_uuid, adviser_uuid, payment_origin, type, category,
#          recurring, frequency, start_date, end_date, amount, percentage,
#          percentage_frequency, tax, tax_exempt, exclude_from_rmar, advice_type,
#          service_type, rmar_type, premium_uuid, (advice_area_* cols), notes,
#          source_uuid, basis, indicative_annual_income, valuation_source

# Helper: get client_uuid for a plan
def get_plan_client_uuid(pl):
    owners = pl.get("_owners", [])
    if isinstance(owners, list) and owners:
        return owners[0].get("client_uuid")
    return None

def get_plan_adviser_uuid(pl):
    # from charges if available, else from plan seller_uuid
    return pl.get("seller_uuid") or pl.get("seller", {}).get("uuid") if isinstance(pl.get("seller"), dict) else None

# Build adviser UUID map from existing charges
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
                     payment_origin, basis, notes="", source_uuid=None,
                     end_date=None):
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
        "percentage": round(rate_decimal * 100, 6),  # back to percentage form
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

# 1. Workbook charges (Section A, positive rate)
for uuid, rate in sorted(sec_a_plan_rate.items()):
    row = build_charge_row(
        plan_uuid=uuid,
        rate_decimal=rate,
        freq="monthly",
        pct_freq="annually",
        start_date=GO_LIVE,
        payment_origin="Provider",
        basis="workbook",
    )
    charges_all_rows.append(row)

# 2. Rathbones charges (matched plans with existing Plannr charge)
rath_charge_plan_uuids = set()
for uuid in sorted(rath_matched_uuids):
    if uuid in plan_charge_info and plan_charge_info[uuid]["source"] == "Rathbones":
        ci = plan_charge_info[uuid]
        chgs = plan_to_active_charges.get(uuid, [])
        if not chgs: continue
        ch = chgs[0]
        rate = ci["rate"]  # already decimal
        row = build_charge_row(
            plan_uuid=uuid,
            rate_decimal=rate,
            freq="quarterly",
            pct_freq="annually",
            start_date="2027-01-31",
            payment_origin="Provider",
            basis="Rathbones",
            notes="Rathbones: annual %, paid quarterly, first payment 31/01/2027",
            source_uuid=ch["uuid"],
        )
        charges_all_rows.append(row)
        rath_charge_plan_uuids.add(uuid)

# 3. Remaining Plannr ongoing (Section B - not covered by workbook or Rathbones)
plannr_remaining_uuids = set()
for pl in plans:
    uuid = pl["uuid"]
    if uuid in sec_a_plan_uuids:
        continue  # covered by workbook
    if uuid in rath_matched_uuids:
        continue  # Rathbones matched (covered above or Section C)
    ci = plan_charge_info.get(uuid)
    if ci and ci["source"] == "Plannr":
        chgs = plan_to_active_charges.get(uuid, [])
        if not chgs: continue
        ch = chgs[0]
        rate = ci["rate"]
        row = build_charge_row(
            plan_uuid=uuid,
            rate_decimal=rate,
            freq=ch.get("frequency", "monthly"),
            pct_freq=ch.get("percentage_frequency", "annually"),
            start_date=ch.get("start_date", GO_LIVE) or GO_LIVE,
            payment_origin=ch.get("payment_origin", "Provider"),
            basis="Plannr ongoing",
            source_uuid=ch["uuid"],
        )
        charges_all_rows.append(row)
        plannr_remaining_uuids.add(uuid)

p(f"  Charges (all) rows: {len(charges_all_rows)}")
p(f"    workbook: {sum(1 for r in charges_all_rows if r['basis']=='workbook')}")
p(f"    Rathbones: {sum(1 for r in charges_all_rows if r['basis']=='Rathbones')}")
p(f"    Plannr ongoing: {sum(1 for r in charges_all_rows if r['basis']=='Plannr ongoing')}")

# ── build Charges sign-off rows ───────────────────────────────────────────────

p("\nBuilding Charges sign-off rows...")

# Helper: get Plannr rate for comparison
def get_plannr_rate_decimal(plan_uuid):
    chgs = plan_to_active_charges.get(plan_uuid, [])
    if not chgs: return None
    try:
        return float(chgs[0].get("percentage", 0)) / 100.0
    except (TypeError, ValueError):
        return None

def build_cs_row(source_section, plan_uuid, workbook_rate=None, plannr_rate=None,
                 final_rate=None, frequency=None, agency="Yes",
                 valuation_source="Plannr", basis="v3"):
    pl = plans_by_uuid.get(plan_uuid, {})
    plan_type = pl.get("type", "")
    policy_number = pl.get("policy_number", "")
    provider_name = get_provider_name(pl) or ""
    val_pounds = plan_value_pounds.get(plan_uuid, 0.0)
    iai = val_pounds * final_rate if final_rate else 0.0
    wb_amount = val_pounds * workbook_rate if workbook_rate else None
    plannr_amount = val_pounds * plannr_rate if plannr_rate else None

    return {
        "source_section": source_section,
        "plan_uuid": plan_uuid,
        "plan_type": plan_type,
        "policy_number": policy_number,
        "provider_name": provider_name,
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

# Section A header
cs_rows.append({"source_section": SECTION_A_HEADER})

# Section A: workbook (Transact/Nucleus) - positive rate
for uuid in sorted(sec_a_plan_uuids):
    rate = sec_a_plan_rate.get(uuid)  # None if zero rate
    plannr_rate = get_plannr_rate_decimal(uuid)
    pl = plans_by_uuid.get(uuid, {})
    agency = "Yes" if pl.get("agency_in_your_name") else None
    row = build_cs_row(
        source_section="A - Workbook confirmed",
        plan_uuid=uuid,
        workbook_rate=rate,
        plannr_rate=plannr_rate,
        final_rate=rate,
        frequency="monthly" if rate else None,
        agency=agency,
        valuation_source="Plannr",
        basis="workbook",
    )
    cs_rows.append(row)

# Section A: Rathbones confirmed (matched plans with Plannr charge)
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

# Section B header
cs_rows.append({"source_section": SECTION_B_HEADER})

# Section B: Plannr ongoing, not in workbook, not Rathbones
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

# Section C header
cs_rows.append({"source_section": SECTION_C_HEADER})

# Section C: under advice but no charge in any source
# = Section A zero-rate plans + Rathbones matched without charge + plans under_advice but no charge
sec_c_uuids = set()

# Section A zero-rate plans
for uuid in sec_a_plan_uuids:
    if uuid not in sec_a_plan_rate:
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
    row = build_cs_row(
        source_section="C - Under advice, no charge any source",
        plan_uuid=uuid,
        valuation_source="Rathbones" if is_rath else "Plannr",
        basis="v3",
    )
    cs_rows.append(row)

p(f"  Charges sign-off rows (data only): {sum(1 for r in cs_rows if r.get('plan_uuid'))}")
p(f"    Section A workbook: {sum(1 for r in cs_rows if r.get('source_section')=='A - Workbook confirmed' and r.get('basis')=='workbook')}")
p(f"    Section A Rathbones: {len(rath_sec_a_uuids)}")
p(f"    Section B: {len(plannr_remaining_uuids)}")
p(f"    Section C: {len(sec_c_uuids)}")

# ── build Valuations rows ─────────────────────────────────────────────────────

p("\nBuilding Valuations rows...")
val_rows = []

# Rathbones value map: plan_uuid -> value_pounds
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
    if not isinstance(val_entries, list):
        val_entries = []

    if uuid in rath_value_map:
        # Use Rathbones file value
        rath_val = rath_value_map[uuid]
        # Plannr value for reference
        plannr_total_pounds = pence_to_pounds(pl.get("value", {}).get("amount", 0) if isinstance(pl.get("value"), dict) else 0)
        val_rows.append({
            "plan_uuid": uuid,
            "sub_account_uuid": "",
            "policy_number": policy_number,
            "valuation_date": RATH_VALUATION_DATE,
            "value": round(rath_val, 2),
            "source": "Rathbones",
            "plannr_value": round(plannr_total_pounds, 2),
        })
    elif val_entries:
        for entry in val_entries:
            sub_uuid = entry.get("_sub_account_uuid", "")
            valued_at = entry.get("valued_at", "")
            if valued_at and "T" in str(valued_at):
                valued_at = str(valued_at)[:10]
            val_pence = entry.get("value", {}).get("amount", 0)
            val_pounds = pence_to_pounds(val_pence)
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
p(f"  Plans with Rathbones valuation: {sum(1 for r in val_rows if r['source']=='Rathbones')}")
p(f"  Plans with Plannr valuation: {len(set(r['plan_uuid'] for r in val_rows if r['source']=='Plannr'))}")

# ── build Income & Employment rows ────────────────────────────────────────────

p("\nBuilding Income & Employment rows...")
ie_rows = []

for acc_uuid, acc in account_details.items():
    income_list = acc.get("income", [])
    employment_list = acc.get("employment", [])

    if not income_list and not employment_list:
        continue

    # Match income records to employment by shared account UUID (heuristic)
    # We emit one row per income record, joined to first employment if any
    emp = employment_list[0] if employment_list else {}

    for inc in income_list:
        amt = inc.get("amount", {})
        amt_pounds = pence_to_pounds(amt.get("amount", 0) if isinstance(amt, dict) else 0)

        # Employment address (if present)
        emp_addr = emp.get("accountant_address") or ""
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
            "employment_accountant_address": emp_addr,
            "employment_turnover_history_year": turnover_year,
            "employment_turnover_history_value": pence_to_pounds(turnover_val.get("amount", 0)) if isinstance(turnover_val, dict) else "",
        })

p(f"  Income & Employment rows: {len(ie_rows)}")

# ── build Expenses rows ───────────────────────────────────────────────────────

p("\nBuilding Expenses rows...")

# Build mapping lookup: (source_section, source_line) -> (dest_section, dest_line, is_custom)
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
p(f"  Custom expenses: {sum(1 for r in exp_rows if r['is_custom_expense'])}")

# ── build Exceptions rows ─────────────────────────────────────────────────────

p("\nBuilding Exceptions rows...")

# Start with non-dup, non-no-polno exceptions from v2
PLAN_TYPES_NEED_POLNO = {"investment", "pension", "protection_policy",
                          "group_protection", "annuity"}
PLACEHOLDER_POLNOS = {
    "TBC", "NA", "NOTKNOWN", "UNKNOWN", "NOTONFILE",
    "NOTSET", "NONE", "N/A",
}

ex_keep = df_ex_v2[~df_ex_v2["category"].isin(
    ["Duplicate policy number", "No policy number"]
)].copy()

# Fix 5: rebuild Duplicate policy number exceptions (real groups only)
# Real duplicate = non-placeholder polno, different primary client UUIDs
dup_ex_rows = []
polno_to_plans_dedup = defaultdict(list)
for pl in plans:
    pn = normalise_polno(pl.get("policy_number"))
    if pn:
        polno_to_plans_dedup[pn].append(pl)

for pn, pls in polno_to_plans_dedup.items():
    if len(pls) < 2: continue
    if pn in PLACEHOLDER_POLNOS: continue
    # Get distinct primary clients
    all_clients = set()
    for pl in pls:
        all_clients.update(get_client_uuids(pl))
    if len(all_clients) < 2: continue
    # Keep if any active
    has_active = any(pl.get("status") == "active" for pl in pls)
    if not has_active: continue
    # These are real duplicates
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

p(f"  Real duplicate policy number groups: {len(set(r['current_value'] for r in dup_ex_rows))}")
p(f"  Real duplicate policy number plans: {len(dup_ex_rows)}")

# Step 4: No policy number exceptions — only keep for relevant plan types
nopolno_rows = []
for pl in plans:
    pn = pl.get("policy_number")
    pn_norm = normalise_polno(pn)
    plan_type = pl.get("type", "")
    if plan_type not in PLAN_TYPES_NEED_POLNO: continue
    if pn_norm:
        continue  # has a policy number
    # No policy number for a type that needs one -> exception
    nopolno_rows.append({
        "category": "No policy number",
        "entity_type": "plan",
        "source_uuid": pl["uuid"],
        "detail": f"type={plan_type}",
        "current_value": pn or "",
        "proposed_value": "Add policy number",
        "AR_decision": None,
    })

p(f"  No-policy-number exceptions (after Step 4 filter): {len(nopolno_rows)}")
p(f"  No-policy-number exceptions BEFORE filter: {len(df_ex_v2[df_ex_v2['category']=='No policy number'])}")

# Add Rathbones exceptions
rath_exc_rows = []
# Plannr-only Rathbones plans
for uuid in rath_plannr_only_uuids:
    rath_exc_rows.append({
        "category": "Rathbones plan not on Rathbones list",
        "entity_type": "plan",
        "source_uuid": uuid,
        "detail": "In Plannr as Rathbones provider, not in Rathbones file",
        "current_value": "",
        "proposed_value": "Closed or transferred? AR to decide",
        "AR_decision": None,
    })
# No-polno Rathbones plans
for uuid in rath_no_polno_uuids:
    rath_exc_rows.append({
        "category": "Rathbones plan not on Rathbones list",
        "entity_type": "plan",
        "source_uuid": uuid,
        "detail": "Rathbones provider plan with no policy number",
        "current_value": "",
        "proposed_value": "Add policy number or close",
        "AR_decision": None,
    })
# Anomaly plan
for uuid in rath_anomaly_uuids:
    rath_exc_rows.append({
        "category": "Rathbones file: matched to non-Rathbones plan",
        "entity_type": "plan",
        "source_uuid": uuid,
        "detail": "Fund Number matches Rathbones file but plan has different provider in Plannr",
        "current_value": "",
        "proposed_value": "AR to investigate",
        "AR_decision": None,
    })

p(f"  Rathbones exceptions: {len(rath_exc_rows)}")

# Combine all exceptions
ex_rows_new = (
    ex_keep.to_dict("records")
    + dup_ex_rows
    + nopolno_rows
    + rath_exc_rows
)
p(f"  Total exceptions: {len(ex_rows_new)}")
ex_cat_ctr = Counter(r.get("category") for r in ex_rows_new)
for cat, cnt in sorted(ex_cat_ctr.items(), key=lambda x: -x[1]):
    p(f"    {cat}: {cnt}")

# ── build Provider-only plans rows ───────────────────────────────────────────

p("\nBuilding Provider-only plans rows...")
# Keep existing rows from v2 (non-Rathbones)
pop_rows = df_pop_v2.to_dict("records")

# Add Rathbones-only rows
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
    # Check if SIPP row
    is_sipp = "sipp" in str(fund_type).lower()
    provider = "Rathbones (within SIPP wrap)" if is_sipp else "Rathbones"
    pop_rows.append({
        "policy_number": pn,
        "provider_name": provider,
        "wrapper_name": fund_type,
        "workbook_value": round(val, 2),
        "workbook_rate": None,
        "match_confidence": "name-only",
        "AR_decision": None,
    })

p(f"  Provider-only plans rows (v2 + Rathbones-only): {len(pop_rows)}")

# ── update Plans tab ──────────────────────────────────────────────────────────

p("\nPreparing Plans tab updates...")
# Track which plan_uuids need under_advice/agency updates
plans_to_update = {}  # uuid -> {col: value}

# Rathbones matched: set under_advice=Yes, agency_in_your_name=Yes
for uuid in rath_matched_uuids:
    plans_to_update[uuid] = plans_to_update.get(uuid, {})
    plans_to_update[uuid]["under_advice"] = "Yes"
    plans_to_update[uuid]["agency_in_your_name"] = "Yes"

# All plans: set indicative_annual_income
for pl in plans:
    uuid = pl["uuid"]
    iai = plan_iai.get(uuid, 0.0)
    if uuid not in plans_to_update:
        plans_to_update[uuid] = {}
    plans_to_update[uuid]["indicative_annual_income"] = round(iai, 2) if iai else ""

p(f"  Plans with column updates: {len(plans_to_update)}")
p(f"  Plans with Rathbones under_advice/agency update: {len(rath_matched_uuids)}")

# ── copy v2 to v3 ─────────────────────────────────────────────────────────────

p(f"\nCopying v2 -> v3...")
shutil.copy2(V2_PATH, V3_PATH)
p(f"  v3 created: {V3_PATH}")

# ── open v3 and write changes ─────────────────────────────────────────────────

p("\nOpening v3 with openpyxl...")
wb = openpyxl.load_workbook(V3_PATH)

def clear_data_rows(ws, header_row=1):
    """Delete all rows below header_row."""
    max_row = ws.max_row
    if max_row <= header_row: return
    ws.delete_rows(header_row + 1, max_row - header_row)

def get_col_index(ws, col_name, header_row=1):
    """Return 1-based column index for col_name in header row."""
    for cell in ws[header_row]:
        if str(cell.value or "") == col_name:
            return cell.column
    return None

def write_rows_to_ws(ws, rows, col_order, start_row=2):
    """Write list-of-dicts to worksheet starting at start_row."""
    for r_idx, row_data in enumerate(rows, start=start_row):
        for c_idx, col_name in enumerate(col_order, start=1):
            ws.cell(row=r_idx, column=c_idx, value=row_data.get(col_name))

# ── write Charges (all) ───────────────────────────────────────────────────────

p("Writing Charges (all)...")
ws_ca = wb["Charges (all)"]
header_row_ca = [cell.value for cell in ws_ca[1]]
p(f"  Charges (all) header cols: {len(header_row_ca)}")
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
# Add extra columns if not present
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
# The template has 'Unnamed: 0' as first col (example label) - skip it when writing real data
ie_headers_raw = [cell.value for cell in ws_ie[1]]
# Find actual data columns (skip 'Unnamed: 0')
ie_col_start = 1
if ie_headers_raw and ie_headers_raw[0] is None or (ie_headers_raw and "Unnamed" in str(ie_headers_raw[0])):
    ie_col_start = 2
ie_header_cols = ie_headers_raw[ie_col_start - 1:]  # 0-indexed
# Clear data rows
clear_data_rows(ws_ie, header_row=1)
# Write new rows
ie_col_order = [c for c in ie_header_cols if c is not None and "Unnamed" not in str(c)]
for r_idx, row_data in enumerate(ie_rows, start=2):
    for local_idx, col_name in enumerate(ie_col_order):
        c_idx = ie_col_start + local_idx
        ws_ie.cell(row=r_idx, column=c_idx, value=row_data.get(col_name))
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

UPDATE_COLS = ["under_advice", "agency_in_your_name", "indicative_annual_income"]
uuid_col_idx = pl_col_idx.get("uuid")

if uuid_col_idx:
    for row_idx in range(2, ws_pl.max_row + 1):
        plan_uuid = ws_pl.cell(row=row_idx, column=uuid_col_idx).value
        if not plan_uuid: continue
        updates = plans_to_update.get(str(plan_uuid))
        if not updates: continue
        for col_name, new_val in updates.items():
            ci = pl_col_idx.get(col_name)
            if ci:
                # Only update under_advice/agency if new value is stronger
                if col_name in ("under_advice", "agency_in_your_name"):
                    existing = ws_pl.cell(row=row_idx, column=ci).value
                    if existing == "Yes": continue  # already Yes, don't downgrade
                ws_pl.cell(row=row_idx, column=ci, value=new_val)

p("  Plans tab updated")

# ── update Summary tab ────────────────────────────────────────────────────────

p("Updating Summary tab...")
ws_sum = wb["Summary"]

def set_summary_field(ws, field_name, value):
    for row in ws.iter_rows():
        for cell in row:
            if str(cell.value or "") == field_name:
                # Value is in the next column
                val_cell = ws.cell(row=cell.row, column=cell.column + 1)
                val_cell.value = value
                return True
    return False

set_summary_field(ws_sum, "Pack version", "v3")
set_summary_field(ws_sum, "Build date", str(date.today()))
set_summary_field(ws_sum, "Indicative annual income (GBP, estimate)", round(total_iai_after, 2))
set_summary_field(ws_sum, "Loadable ongoing charges", len(charges_all_rows))
set_summary_field(ws_sum, "Charges sign-off Section A (workbook)", len(sec_a_plan_uuids) + len(rath_sec_a_uuids))
set_summary_field(ws_sum, "Charges sign-off Section B (Plannr ongoi", len(plannr_remaining_uuids))
set_summary_field(ws_sum, "Open exception items", len(ex_rows_new))
set_summary_field(ws_sum, "Provider-only plan rows", len(pop_rows))

p("  Summary tab updated")

# ── save v3 ───────────────────────────────────────────────────────────────────

p(f"\nSaving v3 to {V3_PATH}...")
wb.save(V3_PATH)
p("  Saved.")

# ── final report ─────────────────────────────────────────────────────────────

v3_size = os.path.getsize(V3_PATH)
p("\n" + "="*60)
p("FINAL REPORT (Rule 9: aggregates only)")
p("="*60)
p()
p("=== Rathbones counts ===")
p(f"  Rathbones plans in Plannr:           {len(rath_plans_in_plannr)}")
p(f"  Matched (file + Plannr):             {len(rath_matched_uuids)}")
p(f"  With existing Plannr charge (Sec A): {len(rath_sec_a_uuids)}")
p(f"  Without Plannr charge (Sec C):       {len(rath_matched_uuids) - len(rath_sec_a_uuids)}")
p(f"  Rathbones-only -> Provider-only tab:  {len(rath_only_polnos)}")
p(f"  Plannr-only -> Exceptions:            {len(rath_plannr_only_uuids)}")
p(f"  Anomaly (wrong provider):            {len(rath_anomaly_uuids)}")
p()
p("=== Charges by basis ===")
p(f"  Workbook (Transact):  {sum(1 for r in charges_all_rows if r['basis']=='workbook')}")
p(f"  Rathbones:            {sum(1 for r in charges_all_rows if r['basis']=='Rathbones')}")
p(f"  Plannr ongoing (B):   {sum(1 for r in charges_all_rows if r['basis']=='Plannr ongoing')}")
p(f"  Total Charges (all):  {len(charges_all_rows)}")
p()
p("=== Indicative annual income ===")
p(f"  BEFORE fix (v2, pence scale): £{total_iai_before:>15,.2f}")
p(f"  AFTER fix  (v3, pounds):      £{total_iai_after:>15,.2f}")
p()
p("=== Exceptions by category (v3) ===")
for cat, cnt in sorted(ex_cat_ctr.items(), key=lambda x: -x[1]):
    p(f"  {cnt:4d}  {cat}")
p(f"  {'----'}")
p(f"  {sum(ex_cat_ctr.values()):4d}  TOTAL")
p()
p("=== Exceptions: No policy number before/after Step 4 ===")
no_polno_before = len(df_ex_v2[df_ex_v2["category"] == "No policy number"])
p(f"  Before (v2): {no_polno_before}")
p(f"  After  (v3): {len(nopolno_rows)}")
p()
p("=== v3 tab row counts ===")
for tab_name in wb.sheetnames:
    ws = wb[tab_name]
    data_rows = max(ws.max_row - 1, 0)
    if data_rows > 0:
        p(f"  {tab_name:<40s} {data_rows:>5d} data rows")
p()
p(f"v3 file size: {v3_size / 1024:.0f} KB")
p(f"v3 path: {V3_PATH}")
p()
p("STOP — awaiting AR sign-off.")
