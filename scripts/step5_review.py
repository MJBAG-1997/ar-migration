"""
HFP AR Migration - Step 5: Review columns + AR tables.
Rule 9: no client names; adviser names OK.
"""
import os, json, re, csv
from collections import defaultdict, Counter
import openpyxl
from openpyxl.styles import Font, PatternFill

DATA_ROOT   = r"C:\dev\AR Migration"
SNAP_DIR    = os.path.join(DATA_ROOT, "snapshot_20261012T090906Z")
EXPORTS_DIR = os.path.join(DATA_ROOT, "exports_20261014")
MANIFEST    = os.path.join(EXPORTS_DIR, "_manifest.json")
RATH_PATH   = os.path.join(DATA_ROOT, "Jim list  - Rathbones.xlsx")
WB_PATH     = os.path.join(DATA_ROOT,
    "Hoban_Financial_Planning_Ltd_260922_Payment_Rates_113727875.xlsx")
MAPPINGS    = os.path.join(DATA_ROOT, "mappings_hfp.xlsx")
DEST_JSON   = os.path.join(DATA_ROOT, "block6_destination.json")
V6_PATH     = os.path.join(DATA_ROOT, "hfp_signoff_pack_v6.xlsx")
SUMMARY_V2  = os.path.join(DATA_ROOT, "step3_summary_v2.json")

PLANS_XLSX   = os.path.join(DATA_ROOT, "Plannr_Plans.xlsx")
CLIENTS_XLSX = os.path.join(DATA_ROOT, "Plannr_Clients.xlsx")
CHARGES_XLSX = os.path.join(DATA_ROOT, "Plannr_Charges.xlsx")
DEPS_XLSX    = os.path.join(DATA_ROOT, "Plannr_Dependants.xlsx")
RELS_XLSX    = os.path.join(DATA_ROOT, "Plannr_Relationships.xlsx")
VALS_XLSX    = os.path.join(DATA_ROOT, "Plannr_Valuations.xlsx")
EXP_XLSX     = os.path.join(DATA_ROOT, "Plannr_Expenditure.xlsx")
IAE_XLSX     = os.path.join(DATA_ROOT, "Plannr_Income_Employment.xlsx")

# Class C = pension/retirement, Class D = investment/savings
CLASS_CD = {t.lower() for t in [
    "sipp","personal pension","group sipp","stakeholder pension","ssas",
    "pension","group personal pension","annuity","drawdown","pension drawdown",
    "retirement account","defined benefit","defined contribution","final salary",
    "qrops","isa","stocks and shares isa","junior isa","lifetime isa","lisa",
    "gia","general investment account","investment bond","offshore bond",
    "onshore bond","unit trust","oeic","investment account","dealing account",
    "with-profits bond","endowment","flexible investment account","bond",
]}

UK_PC = re.compile(r'^[A-Z]{1,2}[0-9][0-9A-Z]?\s*[0-9][A-Z]{2}$', re.I)
REVIEW_FILL = PatternFill(fill_type="solid", fgColor="FFF2CC")
REVIEW_FONT = Font(bold=True, color="7F6000")

def p(*a): print(*a, flush=True)
def blank(v): return v is None or str(v).strip() in ("", "None", "nan")

PLACEHOLDER_NORM = {"TBC","NA","N/A","NOTKNOWN","UNKNOWN","NOTONFILE",
                    "NOTSET","NONE","NOTAPPLICABLE"}
def is_placeholder_pn(v):
    if blank(v): return True
    s = str(v).strip()
    n = re.sub(r"[\s\-]","",s).upper()
    if n in PLACEHOLDER_NORM: return True
    if len(n) <= 5 and not re.match(r"^\d{4,5}$|^[A-Z]{2}\d{3,}$", n): return True
    return False

# ── Load ──────────────────────────────────────────────────────────────────────

with open(MANIFEST, encoding="utf-8") as f:
    manifest = json.load(f)
results = manifest.get("results", {})
def ep(label): return results.get(label, {}).get("path", "")

with open(os.path.join(SNAP_DIR, "plans_dedup.json"), encoding="utf-8") as f:
    snap_plans = json.load(f)
with open(os.path.join(SNAP_DIR, "circles.json"), encoding="utf-8") as f:
    circles_snap = json.load(f)
with open(SUMMARY_V2, encoding="utf-8") as f:
    summary = json.load(f)

with open(ep("plan_export"),   encoding="utf-8-sig", newline="") as f:
    plan_rows_raw = list(csv.DictReader(f))
with open(ep("client_export"), encoding="utf-8-sig", newline="") as f:
    client_rows_raw = list(csv.DictReader(f))
with open(ep("charges_export"), encoding="utf-8-sig", newline="") as f:
    charges_rows_raw = list(csv.DictReader(f))

staff_rows = []
sp = ep("staff_member_export")
if sp and os.path.exists(sp):
    with open(sp, encoding="utf-8-sig", newline="") as f:
        staff_rows = list(csv.DictReader(f))
staff_by_uuid = {r.get("uuid",""): r for r in staff_rows}

with open(DEST_JSON, encoding="utf-8") as f:
    dest = json.load(f)

# Rathbones fund numbers
_wb_r = openpyxl.load_workbook(RATH_PATH, read_only=True, data_only=True)
rath_rows = list(_wb_r.active.iter_rows(values_only=True))
_wb_r.close()
rath_fund_nos = {str(r[0] or "").strip().upper() for r in rath_rows[1:] if r[0]}

# Transact/Nucleus workbook: headers + advice col
_wb_tn = openpyxl.load_workbook(WB_PATH, read_only=True, data_only=True)
tn_rows_all = list(_wb_tn["Payment Rates Summary"].iter_rows(values_only=True))
_wb_tn.close()
tn_hdrs = [str(h or "").strip() for h in tn_rows_all[0]]
p(f"Workbook columns: {tn_hdrs}")
# Identify advice/service/agency column
advice_idx = next(
    (i for i, h in enumerate(tn_hdrs)
     if any(kw in h.lower() for kw in ("advice","agency","service type","service_type"))),
    None
)
p(f"Advice/service column: col {advice_idx} = '{tn_hdrs[advice_idx] if advice_idx is not None else 'none found'}'")
# Build portfolio/wrapper -> workbook advice value
tn_advice_by_pn = {}  # normalised_pn -> workbook value
for row_t in tn_rows_all[1:]:
    pn1 = str(row_t[3] or "").strip().upper()   # col D Portfolio Number
    pn2 = str(row_t[4] or "").strip().upper()   # col E Wrapper Ref
    adv = str(row_t[advice_idx] or "").strip() if advice_idx is not None else ""
    for pn in (pn1, pn2):
        if pn:
            tn_advice_by_pn[pn] = adv

# Plannr original charge advice_type per plan_uuid
orig_advice_by_plan = {r.get("plan_uuid",""): (r.get("advice_type","") or "")
                       for r in charges_rows_raw}

# Provider name from snap_plans
def provider_name_from_snap(plan_uuid):
    pl = snap_plans.get(plan_uuid, {})
    prov = pl.get("provider") or pl.get("provider_name") or ""
    if isinstance(prov, dict):
        prov = prov.get("name","")
    return str(prov or "").strip()

# Also check plan_export for provider_name column
plan_export_cols = list(plan_rows_raw[0].keys()) if plan_rows_raw else []
client_export_cols_all = list(client_rows_raw[0].keys()) if client_rows_raw else []
p(f"Plan export cols (first 30): {plan_export_cols[:30]}")
p(f"Client export cols (first 30): {client_export_cols_all[:30]}")
# Advisers: plan export uses active_adviser_uuid, not adviser_uuid
PLAN_ADVISER_COL   = "active_adviser_uuid" if "active_adviser_uuid" in plan_export_cols else "adviser_uuid"
# Prefer UUID column for client adviser (so staff lookup works by UUID)
CLIENT_ADVISER_COL = next((c for c in client_export_cols_all
                           if c in ("assigned_adviser_uuid","adviser_uuid")), None)
if not CLIENT_ADVISER_COL:
    CLIENT_ADVISER_COL = next((c for c in client_export_cols_all
                               if "adviser" in c.lower() and "uuid" in c.lower()), "adviser_uuid")
p(f"Plan adviser col: {PLAN_ADVISER_COL}")
p(f"Client adviser col: {CLIENT_ADVISER_COL}")
# Client status column: prefer status_name over status_uuid
CLIENT_STATUS_COL = next((c for c in client_export_cols_all if c == "status_name"), None)
if not CLIENT_STATUS_COL:
    CLIENT_STATUS_COL = next((c for c in client_export_cols_all if c == "status"), None)
if not CLIENT_STATUS_COL:
    CLIENT_STATUS_COL = next((c for c in client_export_cols_all
                              if "status" in c.lower() and "name" in c.lower()), None)
if not CLIENT_STATUS_COL:
    CLIENT_STATUS_COL = next((c for c in client_export_cols_all if "status" in c.lower()), "status")
p(f"Client status col: {CLIENT_STATUS_COL}")
# Adviser name lookup: from staff export + plan export active_adviser_name
adviser_name_lookup = {}
for sr in staff_rows:
    uid = sr.get("uuid","")
    fn = (sr.get("first_name") or sr.get("name") or "").strip()
    ln = (sr.get("last_name") or "").strip()
    if uid and (fn or ln):
        adviser_name_lookup[uid] = f"{fn} {ln}".strip()
# Also pick up names from plan export active_adviser_name
for r in plan_rows_raw:
    uid  = (r.get("active_adviser_uuid") or "").strip()
    name = (r.get("active_adviser_name") or "").strip()
    if uid and name and uid not in adviser_name_lookup:
        adviser_name_lookup[uid] = name
# And from client export assigned_adviser (name) keyed by assigned_adviser_uuid
for r in client_rows_raw:
    uid  = (r.get("assigned_adviser_uuid") or "").strip()
    name = (r.get("assigned_adviser") or "").strip()
    if uid and name and uid not in adviser_name_lookup:
        adviser_name_lookup[uid] = name
p(f"Adviser names resolved: {len(adviser_name_lookup)}")
has_provider_col = any("provider" in c.lower() for c in plan_export_cols)

def plan_provider(row):
    """Best-effort provider name from plan row + snapshot."""
    puuid = row.get("uuid","")
    for col in plan_export_cols:
        if "provider" in col.lower():
            v = row.get(col,"")
            if v and "rathbone" in str(v).lower():
                return str(v)
    snap_prov = provider_name_from_snap(puuid)
    if snap_prov:
        return snap_prov
    # fallback: sub_accounts provider names in snapshot
    pl = snap_plans.get(puuid, {})
    for sa in pl.get("_sub_accounts", []):
        prov = sa.get("provider") or sa.get("provider_name") or ""
        if isinstance(prov, dict):
            prov = prov.get("name","")
        if "rathbone" in str(prov or "").lower():
            return str(prov)
    return ""

# ── Plan data maps ────────────────────────────────────────────────────────────

plan_by_uuid  = {r["uuid"]: r for r in plan_rows_raw}
client_uuids_in_plans = set()
for r in plan_rows_raw:
    for c in (r.get("client_uuids","") or "").split(","):
        c = c.strip()
        if c:
            client_uuids_in_plans.add(c)

plans_per_adviser   = Counter(r.get(PLAN_ADVISER_COL,"") or "" for r in plan_rows_raw)
clients_per_adviser = Counter(r.get(CLIENT_ADVISER_COL,"") or "" for r in client_rows_raw)

# Circles: client -> number of circles (households)
client_circle_count = Counter()
for circle in circles_snap:
    for m in circle.get("_members", []):
        muuid = m.get("uuid","") if isinstance(m,dict) else str(m)
        if muuid:
            client_circle_count[muuid] += 1

# ── Compute plan review flags ─────────────────────────────────────────────────

# Duplicate policy numbers in classes C/D
pn_to_plan_uuids = defaultdict(list)
for r in plan_rows_raw:
    pn = (r.get("policy_number","") or "").strip().upper()
    if pn and not is_placeholder_pn(pn):
        ptype = (r.get("type","") or "").strip().lower()
        if ptype in CLASS_CD:
            pn_to_plan_uuids[pn].append(r["uuid"])
dup_plan_uuids = {uid for uids in pn_to_plan_uuids.values() if len(uids)>1 for uid in uids}

# No policy number (active only)
no_pn_active = {r["uuid"] for r in plan_rows_raw
                if (r.get("status","") or "").strip().lower() == "active"
                and is_placeholder_pn(r.get("policy_number",""))}

# No adviser
no_adviser = {r["uuid"] for r in plan_rows_raw if blank(r.get(PLAN_ADVISER_COL))}

# Rathbones plans in Plannr (by provider name) not on Jim list
rath_plans_in_plannr = set()
for r in plan_rows_raw:
    prov = plan_provider(r)
    if "rathbone" in prov.lower():
        rath_plans_in_plannr.add(r["uuid"])
p(f"Rathbones plans in Plannr (by provider): {len(rath_plans_in_plannr)}")

not_on_rath_list = set()
for puuid in rath_plans_in_plannr:
    row = plan_by_uuid.get(puuid,{})
    pn = (row.get("policy_number","") or "").strip().upper()
    if pn not in rath_fund_nos:
        not_on_rath_list.add(puuid)

# Agency/under-advice changed from workbook
changed_advice = set()
for r in plan_rows_raw:
    pn = (r.get("policy_number","") or "").strip().upper()
    if pn and pn in tn_advice_by_pn:
        wb_adv  = tn_advice_by_pn[pn].strip().lower()
        pl_adv  = orig_advice_by_plan.get(r["uuid"],"").strip().lower()
        if wb_adv and pl_adv and wb_adv != pl_adv:
            changed_advice.add(r["uuid"])

# Plan types found (for reference)
plan_types_found = sorted({(r.get("type","") or "").strip() for r in plan_rows_raw} - {""})
p(f"Plan types in export: {plan_types_found}")

# Build plan review dict
plan_review = {}
for r in plan_rows_raw:
    uid = r["uuid"]
    labels = []
    if uid in no_pn_active:
        labels.append("No policy number")
    if uid in dup_plan_uuids:
        labels.append("Duplicate policy number")
    if uid in no_adviser:
        labels.append("No adviser")
    if uid in not_on_rath_list:
        labels.append("Not on Rathbones list")
    plan_review[uid] = "; ".join(labels)

# ── Compute client review flags ───────────────────────────────────────────────

client_export_cols = list(client_rows_raw[0].keys()) if client_rows_raw else []
postcode_col = next((c for c in client_export_cols
                     if "postcode" in c.lower() or "post_code" in c.lower()), None)
p(f"Postcode col: {postcode_col}")

client_uuids_all = {r["uuid"] for r in client_rows_raw}
no_status_no_plans = {r["uuid"] for r in client_rows_raw
                      if blank(r.get(CLIENT_STATUS_COL)) and r["uuid"] not in client_uuids_in_plans}
multi_household = {uid for uid, cnt in client_circle_count.items()
                   if cnt > 1 and uid in client_uuids_all}
bad_postcode = set()
if postcode_col:
    for r in client_rows_raw:
        pc = (r.get(postcode_col,"") or "").strip()
        if pc and not UK_PC.match(pc):
            bad_postcode.add(r["uuid"])

client_review = {}
for r in client_rows_raw:
    uid = r["uuid"]
    labels = []
    if uid in no_status_no_plans:
        labels.append("No status and no plans")
    if uid in multi_household:
        labels.append("In more than one household")
    if uid in bad_postcode:
        labels.append("Check postcode")
    client_review[uid] = "; ".join(labels)

# ── Insert Review column in Plannr_Plans.xlsx ─────────────────────────────────

p("\nModifying Plannr_Plans.xlsx...")
wb_p = openpyxl.load_workbook(PLANS_XLSX)
ws_p = wb_p["Plans"]
ws_p.insert_cols(1)
ws_p.cell(1, 1, "Review").font   = REVIEW_FONT
ws_p.cell(1, 1).fill = REVIEW_FILL
# Find uuid column (shifted right by 1)
uuid_col_p = next((ci for ci in range(1, ws_p.max_column+1)
                   if ws_p.cell(1, ci).value == "uuid"), None)
if uuid_col_p:
    for ri in range(2, ws_p.max_row + 1):
        uid = ws_p.cell(ri, uuid_col_p).value
        if uid:
            ws_p.cell(ri, 1, plan_review.get(str(uid), ""))
wb_p.save(PLANS_XLSX)
p("  Saved Plannr_Plans.xlsx")

# ── Insert Review column in Plannr_Clients.xlsx ───────────────────────────────

p("Modifying Plannr_Clients.xlsx...")
wb_c = openpyxl.load_workbook(CLIENTS_XLSX)
ws_c = wb_c["Clients"]
ws_c.insert_cols(1)
ws_c.cell(1, 1, "Review").font   = REVIEW_FONT
ws_c.cell(1, 1).fill = REVIEW_FILL
uuid_col_c = next((ci for ci in range(1, ws_c.max_column+1)
                   if ws_c.cell(1, ci).value == "uuid"), None)
if uuid_col_c:
    for ri in range(2, ws_c.max_row + 1):
        uid = ws_c.cell(ri, uuid_col_c).value
        if uid:
            ws_c.cell(ri, 1, client_review.get(str(uid), ""))
wb_c.save(CLIENTS_XLSX)
p("  Saved Plannr_Clients.xlsx")

# ── Read Plannr files for counts ─────────────────────────────────────────────

def read_sheet(path, sheet_index=0):
    _wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    _ws = list(_wb.worksheets)[sheet_index]
    data = list(_ws.iter_rows(values_only=True))
    _wb.close()
    return data

# Charges
ch_data = read_sheet(CHARGES_XLSX)
ch_hdrs = list(ch_data[0])
basis_col_i  = ch_hdrs.index("basis") if "basis" in ch_hdrs else None
orig_cols_i  = [i for i,h in enumerate(ch_hdrs) if str(h or "").startswith("ORIGINAL_")]
basis_counts = Counter()
sec_b_with_orig = sec_c_with_orig = 0
for row_t in ch_data[1:]:
    basis = str(row_t[basis_col_i] or "") if basis_col_i is not None else ""
    basis_counts[basis] += 1
    has_orig = any(not blank(row_t[i]) for i in orig_cols_i)
    if basis == "Your Plannr" and has_orig:   sec_b_with_orig += 1
    if basis == "Rate needed"  and has_orig:  sec_c_with_orig += 1

# Provider tab
prov_data = read_sheet(PLANS_XLSX, sheet_index=1)
prov_hdrs  = list(prov_data[0])
action_col_i = prov_hdrs.index("action") if "action" in prov_hdrs else None
action_counts = Counter(
    str(row_t[action_col_i] or "").strip()
    for row_t in prov_data[1:]
    if action_col_i is not None
)

# Dependants: no DOB
dep_data = read_sheet(DEPS_XLSX)
dep_hdrs = list(dep_data[0])
dob_i = dep_hdrs.index("date_of_birth") if "date_of_birth" in dep_hdrs else None
no_dob = sum(1 for r in dep_data[1:] if dob_i is not None and blank(r[dob_i]))

# Relationships: no type
rel_data = read_sheet(RELS_XLSX)
rel_hdrs = list(rel_data[0])
rel_i = rel_hdrs.index("relationship") if "relationship" in rel_hdrs else None
no_rel = sum(1 for r in rel_data[1:] if rel_i is not None and blank(r[rel_i]))

# Valuations: date range
val_data = read_sheet(VALS_XLSX)
val_hdrs = list(val_data[0])
vd_i = val_hdrs.index("valuation_date") if "valuation_date" in val_hdrs else None
vdates = sorted(str(r[vd_i]) for r in val_data[1:]
                if vd_i is not None and not blank(r[vd_i]))
earliest_val = vdates[0]  if vdates else "n/a"
latest_val   = vdates[-1] if vdates else "n/a"

# I&E: employment-only rows
iae_data = read_sheet(IAE_XLSX)
iae_hdrs = list(iae_data[0])
inc_name_i   = iae_hdrs.index("name")   if "name"   in iae_hdrs else None
inc_type_i   = iae_hdrs.index("type")   if "type"   in iae_hdrs else None
emp_uuid_i   = next((i for i,h in enumerate(iae_hdrs) if str(h or "")=="employment_uuid"), None)
emp_only_rows = 0
for r in iae_data[1:]:
    has_emp = emp_uuid_i is not None and not blank(r[emp_uuid_i])
    has_inc = (inc_name_i is not None and not blank(r[inc_name_i])) or \
              (inc_type_i is not None and not blank(r[inc_type_i]))
    if has_emp and not has_inc:
        emp_only_rows += 1

# Expenditure: custom vs standard
exp_data = read_sheet(EXP_XLSX)
exp_hdrs = list(exp_data[0])
cust_i = exp_hdrs.index("is_custom_expense") if "is_custom_expense" in exp_hdrs else None
custom_exp = 0
mapped_exp  = 0
for r in exp_data[1:]:
    v = r[cust_i] if cust_i is not None else None
    if v is True or str(v or "").upper() in ("TRUE","YES","1"):
        custom_exp += 1
    else:
        mapped_exp += 1

# ── Status mapping ────────────────────────────────────────────────────────────

status_map = {}      # Plannr status -> BA status
ba_status_list = []
try:
    _wb_map = openpyxl.load_workbook(MAPPINGS, read_only=True, data_only=True)
    if "client_statuses" in _wb_map.sheetnames:
        sm_data = list(_wb_map["client_statuses"].iter_rows(values_only=True))
        sm_hdrs = [str(h or "") for h in sm_data[0]]
        p(f"Status mapping cols: {sm_hdrs}")
        # Columns: source_uuid, source_name, dest_uuid, dest_name, notes
        src_name_i = next((i for i,h in enumerate(sm_hdrs) if h.lower() == "source_name"), None)
        dst_name_i = next((i for i,h in enumerate(sm_hdrs) if h.lower() == "dest_name"), None)
        if src_name_i is None: src_name_i = 1
        if dst_name_i is None: dst_name_i = 3
        for row_t in sm_data[1:]:
            src = str(row_t[src_name_i] or "").strip()
            dst = str(row_t[dst_name_i] or "").strip()
            if src:
                status_map[src] = dst
            if dst and dst not in ba_status_list:
                ba_status_list.append(dst)
    _wb_map.close()
except Exception as e:
    p(f"  Status mapping error: {e}")

client_status_counts = Counter((r.get(CLIENT_STATUS_COL,"") or "").strip() for r in client_rows_raw)

# ── TABLE 1 ───────────────────────────────────────────────────────────────────

p("\n" + "=" * 115)
p("TABLE 1: REVIEW ITEMS")
p("=" * 115)

COL = 52
HDR = f"{'Item':<{COL}} {'Count':>6}  {'File / tab':<44}  How to find (col: filter value)"
p(HDR)
p("-" * 115)

def row(item, cnt, file_, col_, filt):
    p(f"  {item:<{COL}} {str(cnt):>6}  {file_:<44}  {col_}: {filt}")

# Plan Review labels
row("Plan: No policy number (active plans only)",
    sum(1 for v in plan_review.values() if "No policy number" in v),
    "Plannr_Plans.xlsx > Plans", "Review", "contains 'No policy number'")
row("Plan: Duplicate policy number (class C/D only)",
    sum(1 for v in plan_review.values() if "Duplicate policy number" in v),
    "Plannr_Plans.xlsx > Plans", "Review", "contains 'Duplicate policy number'")
row("Plan: No adviser",
    sum(1 for v in plan_review.values() if "No adviser" in v),
    "Plannr_Plans.xlsx > Plans", "Review", "contains 'No adviser'")
row("Plan: Not on Rathbones list",
    sum(1 for v in plan_review.values() if "Not on Rathbones list" in v),
    "Plannr_Plans.xlsx > Plans", "Review", "contains 'Not on Rathbones list'")

# Client Review labels
row("Client: No status and no plans",
    sum(1 for v in client_review.values() if "No status and no plans" in v),
    "Plannr_Clients.xlsx", "Review", "contains 'No status and no plans'")
row("Client: In more than one household",
    sum(1 for v in client_review.values() if "In more than one household" in v),
    "Plannr_Clients.xlsx", "Review", "contains 'In more than one household'")
row("Client: Check postcode",
    sum(1 for v in client_review.values() if "Check postcode" in v),
    "Plannr_Clients.xlsx", "Review", "contains 'Check postcode'")

p("")
# Placeholder policy numbers blanked
row("Plans: placeholder policy number blanked",
    sum(1 for r in plan_rows_raw
        if not blank(r.get("policy_number",""))
        and is_placeholder_pn(r.get("policy_number",""))),
    "Plannr_Plans.xlsx > Plans", "ORIGINAL_policy_number", "not blank")

# Agency/under-advice changed
if advice_idx is not None:
    row(f"Plans: {tn_hdrs[advice_idx]} differs from workbook",
        len(changed_advice),
        "Plannr_Plans.xlsx > Plans",
        f"Review or advice_type",
        f"workbook col '{tn_hdrs[advice_idx]}' != Plannr charge advice_type")
else:
    p(f"  {'Plans: advice/service col not found in workbook':<{COL}} {'n/a':>6}  {'--':<44}  workbook has no advice/agency col")

p("")
# Provider plans not in Plannr
total_prov = len(prov_data) - 1
row("Provider plans not in Plannr (total)", total_prov,
    "Plannr_Plans.xlsx > Provider_plans_not_in_Plannr",
    "action", "any value (drop-down)")
for act, cnt in sorted(action_counts.items(), key=lambda x: (x[0] != "", x[0])):
    label = act if act else "(blank - not yet decided)"
    row(f"  action = {label}", cnt,
        "Plannr_Plans.xlsx > Provider_plans_not_in_Plannr",
        "action", f"= '{act}'" if act else "is blank")

p("")
# Charges by basis
for basis_label in ["Provider statement", "Rathbones", "Your Plannr", "Rate needed"]:
    row(f"Charges: {basis_label}", basis_counts.get(basis_label, 0),
        "Plannr_Charges.xlsx", "basis", f"= '{basis_label}'")
row("Charges: Section B with previous Plannr rate", sec_b_with_orig,
    "Plannr_Charges.xlsx", "basis",
    "= 'Your Plannr' AND any ORIGINAL_ column not blank")
row("Charges: Section C with previous Plannr rate", sec_c_with_orig,
    "Plannr_Charges.xlsx", "basis",
    "= 'Rate needed' AND any ORIGINAL_ column not blank")

p("")
row("Relationships: no relationship type", no_rel,
    "Plannr_Relationships.xlsx", "relationship", "is blank")
row("Dependants: no date of birth", no_dob,
    "Plannr_Dependants.xlsx", "date_of_birth", "is blank")
row("Employment rows with no income (name + type both blank)", emp_only_rows,
    "Plannr_Income_Employment.xlsx", "employment_uuid",
    "not blank AND name is blank")
row("Expenditure: mapped to BA headings", mapped_exp,
    "Plannr_Expenditure.xlsx", "is_custom_expense", "is blank")
row("Expenditure: kept as custom", custom_exp,
    "Plannr_Expenditure.xlsx", "is_custom_expense", "= TRUE")
row(f"Valuations: earliest date  {earliest_val}", "",
    "Plannr_Valuations.xlsx", "valuation_date",
    f"sort A-Z, first value = {earliest_val}")
row(f"Valuations: latest date    {latest_val}", "",
    "Plannr_Valuations.xlsx", "valuation_date",
    f"sort Z-A, first value = {latest_val}")

# ── TABLE 2 ───────────────────────────────────────────────────────────────────

p("\n" + "=" * 68)
p("TABLE 2: CLIENTS BY STATUS")
p("=" * 68)
p(f"  {'Plannr status':<28} {'Count':>6}  BA status")
p("  " + "-" * 55)
for status, cnt in sorted(client_status_counts.items(), key=lambda x: -x[1]):
    ba_st = status_map.get(status, "(check mapping)")
    p(f"  {status or '(blank)':<28} {cnt:>6}  {ba_st}")
p("")
if ba_status_list:
    p(f"  BA destination status options available: {', '.join(sorted(set(ba_status_list)))}")
else:
    p("  BA destination status options: check mappings_hfp.xlsx > client_statuses")

# ── TABLE 3 ───────────────────────────────────────────────────────────────────

p("\n" + "=" * 68)
p("TABLE 3: PER ADVISER")
p("=" * 68)
p(f"  {'Adviser':<35} {'Clients':>8}  {'Plans':>7}")
p("  " + "-" * 55)

all_adviser_uuids = set(plans_per_adviser) | set(clients_per_adviser)
adviser_rows = []
for auuid in all_adviser_uuids:
    if not auuid:
        name = "(no adviser)"
    else:
        name = adviser_name_lookup.get(auuid) or auuid[:8]
    adviser_rows.append((name, clients_per_adviser.get(auuid,0), plans_per_adviser.get(auuid,0)))

adviser_rows.sort(key=lambda x: (-x[2], x[0]))
for name, clients, plans in adviser_rows:
    p(f"  {name:<35} {clients:>8}  {plans:>7}")

# ── TABLE 4 ───────────────────────────────────────────────────────────────────

p("\n" + "=" * 68)
p("TABLE 4: OTHER AR DECISIONS")
p("=" * 68)

notes_s = summary.get("notes", {})
by_model = notes_s.get("by_model", {})
task_n = by_model.get("task", 0)
case_n = by_model.get("case", 0)
board_n = by_model.get("board_task", 0)
other_n = sum(v for k, v in by_model.items() if k not in ("task","case","board_task"))

decisions = [
    (f"Charges - Rate needed ({basis_counts.get('Rate needed',0)} plans): "
     "open Plannr_Charges.xlsx, filter basis = 'Rate needed', fill in rate or delete row."),

    (f"Provider plans not in Plannr ({total_prov} rows): "
     "open Provider_plans_not_in_Plannr tab, mark each row as: "
     "Add to Plannr / Already in Plannr / Not ours / Closed."),

    (f"Notes - which types to import: "
     f"tasks and case notes only ({task_n + case_n:,}), "
     f"or all transferable types ({task_n + case_n + other_n:,} total). "
     f"Board tasks ({board_n:,}) stay on old Plannr regardless."),

    (f"Plans - No adviser ({len(no_adviser)} plans): "
     "assign an adviser or confirm the plan is correctly unassigned."),

    (f"Plans - Duplicate policy number, class C/D ({len(dup_plan_uuids)} plans): "
     "confirm each pair is correct (e.g. two sub-accounts of the same wrapper) or an error."),

    (f"Employment rows with no income ({emp_only_rows} rows): "
     "fill in income name/type/amount in Plannr_Income_Employment.xlsx or delete the row."),

    (f"Plans - Not on Rathbones list ({len(not_on_rath_list)} plans): "
     "Plannr shows provider = Rathbones but the fund number is not in the Rathbones file. "
     "Confirm whether the policy number is correct or the provider is wrong."),

    ("Valuations are from the 12 Oct snapshot. "
     "Any new valuations added after that date are not included."),

    ("Tags: not in scope for this run. Confirm whether tags should be migrated."),

    ("Vision IFP: not in scope. Confirm separately if required."),
]

for i, d in enumerate(decisions, 1):
    p(f"  {i}. {d}")

p("\nSTOP")
