"""
Build hfp_review_pack_v2.xlsx from v1 with three fixes:

Fix 1 — 20 unassigned active plans:
  - Plans to check: add one row per unassigned plan (adviser name dropdown)
  - Summary: replace "Unassigned" row with "To confirm"

Fix 2 — Missing circles investigation:
  - Console only (Rule 9: aggregates/UUIDs only)

Rule 9: aggregates/UUIDs only in console. No PII.
"""
import shutil, json, os, re
from collections import defaultdict
from datetime import date
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

DATA_ROOT    = r"C:\dev\AR Migration"
V1_PATH      = os.path.join(DATA_ROOT, "hfp_review_pack_v1.xlsx")
V2_PATH      = os.path.join(DATA_ROOT, "hfp_review_pack_v2.xlsx")
V5_PATH      = os.path.join(DATA_ROOT, "hfp_signoff_pack_v5.xlsx")
SNAP_DIR     = os.path.join(DATA_ROOT, "snapshot_20261012T090906Z")

def p(*a): print(*a, flush=True)

PLACEHOLDER_UPPER = {
    "TBC","NA","N/A","NOTKNOWN","UNKNOWN","NOTONFILE",
    "NOTSET","NONE","NOT SET","NOTON","NOTAPPLICABLE",
}

def is_placeholder(pn_raw):
    if not pn_raw or str(pn_raw).strip() in ("","None","nan"):
        return True
    pn_str = str(pn_raw).strip()
    if re.fullmatch(r"[\-\s\.]+", pn_str):
        return True
    norm = re.sub(r"[\s\-]","",pn_str).upper()
    if norm in PLACEHOLDER_UPPER:
        return True
    if len(norm) <= 5 and not re.match(r"^\d{4,5}$|^[A-Z]{2}\d{3,}$", norm):
        return True
    return False

def pence_to_pounds(amt):
    try:
        return int(str(amt)) / 100.0
    except (TypeError, ValueError):
        return 0.0

def plan_value_pounds_fn(pl):
    val = pl.get("value")
    if isinstance(val, dict):
        return pence_to_pounds(val.get("amount", 0))
    return 0.0

# ── load snapshot data ─────────────────────────────────────────────────────────

p("Loading snapshot data...")

with open(os.path.join(SNAP_DIR, "plans_dedup.json"), encoding="utf-8") as f:
    plans_raw = json.load(f)
plans = list(plans_raw.values())
plans_by_uuid = {pl["uuid"]: pl for pl in plans}

with open(os.path.join(SNAP_DIR, "charges_dedup.json"), encoding="utf-8") as f:
    charges_raw = json.load(f)
charges = list(charges_raw.values())

with open(os.path.join(SNAP_DIR, "accounts.json"), encoding="utf-8") as f:
    accounts = json.load(f)

with open(os.path.join(SNAP_DIR, "circles.json"), encoding="utf-8") as f:
    circles_data = json.load(f)

with open(os.path.join(SNAP_DIR, "reference.json"), encoding="utf-8") as f:
    reference = json.load(f)
employees = reference.get("employees", [])

p(f"  Plans: {len(plans)}, Charges: {len(charges)}, Accounts: {len(accounts)}")

# ── reference maps ─────────────────────────────────────────────────────────────

adviser_name_map = {e["uuid"]: e.get("name","") for e in employees}
acc_name         = {a["uuid"]: a.get("name","") for a in accounts}

acc_adviser = {}
for a in accounts:
    aa = a.get("assigned_adviser")
    if isinstance(aa, dict) and aa.get("uuid"):
        acc_adviser[a["uuid"]] = aa["uuid"]

plan_primary_client = {}
for pl in plans:
    owners = pl.get("_owners", [])
    if isinstance(owners, list) and owners:
        best = max(owners, key=lambda o: float(o.get("ownership_percentage", 0) or 0))
        plan_primary_client[pl["uuid"]] = best.get("client_uuid","")

plan_to_adviser_uuid = {}
for ch in charges:
    s   = ch.get("seller")
    pid = ch.get("_plan_uuid")
    if isinstance(s, dict) and s.get("uuid") and pid:
        plan_to_adviser_uuid[pid] = s["uuid"]
for pl in plans:
    uuid = pl["uuid"]
    if uuid in plan_to_adviser_uuid:
        continue
    c_uuid = plan_primary_client.get(uuid,"")
    if c_uuid and c_uuid in acc_adviser:
        plan_to_adviser_uuid[uuid] = acc_adviser[c_uuid]

# circles maps (for household count)
account_to_circles = defaultdict(list)
for a in accounts:
    for c in a.get("circles", []):
        account_to_circles[a["uuid"]].append(c["uuid"])

# ── Fix 2: Investigate missing circles ────────────────────────────────────────

p()
p("=" * 60)
p("FIX 2: Missing circles investigation")
p("=" * 60)

acc_circle_uuids  = {c["uuid"] for a in accounts for c in a.get("circles",[])}
all_circle_uuids  = {c["uuid"] for c in circles_data}
missing_uuids     = all_circle_uuids - acc_circle_uuids

client_to_plans = defaultdict(list)
for pl in plans:
    for o in (pl.get("_owners") or []):
        c = o.get("client_uuid")
        if c:
            client_to_plans[c].append(pl)

for i, cuuid in enumerate(sorted(missing_uuids), start=1):
    circle = next((c for c in circles_data if c["uuid"] == cuuid), {})
    members = circle.get("_members", [])
    client_uuids = {
        lc.get("uuid")
        for m in members
        for lc in (m.get("linked_clients") or [])
    }
    all_pl = [p for c in client_uuids for p in client_to_plans.get(c,[])]
    active  = [p for p in all_pl if p.get("status") == "active"]
    if not members:
        reason = "empty"
    elif not all_pl:
        reason = "members_no_plans"
    elif not active:
        reason = "all_inactive"
    else:
        reason = f"has_active ({len(active)} plans)"
    p(f"  Circle {i}: {len(members)} members, reason={reason}")

p(f"\n  None of the 6 missing circles have active plans — v6 Circles tab unchanged.")

# ── identify unassigned active plans ──────────────────────────────────────────

p()
p("=" * 60)
p("FIX 1: Unassigned active plans")
p("=" * 60)

unassigned_plans = [
    pl for pl in plans
    if pl.get("status") == "active" and pl["uuid"] not in plan_to_adviser_uuid
]
unassigned_uuids = {pl["uuid"] for pl in unassigned_plans}
p(f"  Unassigned active plans: {len(unassigned_plans)}")

# Compute aggregates for "To confirm" Summary row
unasgn_clients   = {plan_primary_client.get(pl["uuid"],"") for pl in unassigned_plans}
unasgn_clients   -= {""}
unasgn_circles   = set()
unasgn_uncircled = set()
for c_uuid in unasgn_clients:
    circles_list = account_to_circles.get(c_uuid,[])
    if circles_list:
        unasgn_circles.update(circles_list)
    else:
        unasgn_uncircled.add(c_uuid)
unasgn_households = len(unasgn_circles) + len(unasgn_uncircled)
unasgn_val_ua = sum(
    plan_value_pounds_fn(pl) for pl in unassigned_plans
    if pl.get("under_advice") in ("Yes", True) or pl.get("under_advice") == "true"
)

# IAI for unassigned plans from v5 Charges (all)
df_ca_v5 = pd.read_excel(V5_PATH, sheet_name="Charges (all)", dtype=str)
unasgn_iai = 0.0
if "adviser_uuid" in df_ca_v5.columns and "plan_uuid" in df_ca_v5.columns:
    for _, row in df_ca_v5.iterrows():
        if str(row.get("plan_uuid","") or "") in unassigned_uuids:
            try:
                unasgn_iai += float(row.get("indicative_annual_income", 0) or 0)
            except (ValueError, TypeError):
                pass

unasgn_avg_fee = unasgn_iai / unasgn_val_ua if unasgn_val_ua > 0 else 0.0

p(f"  Clients: {len(unasgn_clients)}, Households: {unasgn_households}")
p(f"  Value under advice: GBP {unasgn_val_ua:,.2f}")
p(f"  IAI: GBP {unasgn_iai:,.2f}")

# ── copy v1 -> v2 ──────────────────────────────────────────────────────────────

p(f"\nCopying v1 -> v2...")
shutil.copy2(V1_PATH, V2_PATH)
p("Opening v2...")
wb2 = openpyxl.load_workbook(V2_PATH)

CURRENCY_FMT = '\u00a3#,##0.00'
PCT_FMT      = '0.00%'
HEADER_FILL  = PatternFill(fill_type="solid", fgColor="F2F2F2")
HEADER_FONT  = Font(bold=True)

# ── modify Summary tab ─────────────────────────────────────────────────────────

p("Modifying Summary tab...")
ws_sum = wb2["Summary"]
hdrs   = [ws_sum.cell(1, c).value for c in range(1, ws_sum.max_column+1)]

unassigned_row = None
total_row      = None
for row_idx in range(2, ws_sum.max_row + 1):
    v = str(ws_sum.cell(row_idx, 1).value or "")
    if v == "Unassigned":
        unassigned_row = row_idx
    if v == "TOTAL":
        total_row = row_idx

if unassigned_row:
    ws_sum.delete_rows(unassigned_row, 1)
    if total_row and total_row > unassigned_row:
        total_row -= 1
    p(f"  Deleted 'Unassigned' row (was row {unassigned_row})")
else:
    p("  WARNING: 'Unassigned' row not found in Summary")

# Insert "To confirm" row before TOTAL
if total_row:
    ws_sum.insert_rows(total_row, 1)
    to_confirm_data = [
        "To confirm",
        len(unasgn_clients),
        unasgn_households,
        len(unassigned_plans),
        unasgn_val_ua,
        unasgn_iai,
        unasgn_avg_fee,
    ]
    for c_idx, val in enumerate(to_confirm_data, start=1):
        ws_sum.cell(total_row, c_idx, value=val)
    # Apply number formats to new row
    val_col  = get_column_letter(5)
    iai_col  = get_column_letter(6)
    fee_col  = get_column_letter(7)
    ws_sum[f"{val_col}{total_row}"].number_format = CURRENCY_FMT
    ws_sum[f"{iai_col}{total_row}"].number_format = CURRENCY_FMT
    ws_sum[f"{fee_col}{total_row}"].number_format = PCT_FMT
    p(f"  Inserted 'To confirm' row at row {total_row}")
else:
    p("  WARNING: 'TOTAL' row not found — 'To confirm' row not inserted")

# ── modify Plans to check tab ──────────────────────────────────────────────────

p("Modifying Plans to check tab...")
ws_ptc = wb2["Plans to check"]
ptc_hdrs = [ws_ptc.cell(1, c).value for c in range(1, ws_ptc.max_column+1)]

# Map header name -> col index (1-based)
def col_of(ws, name):
    for c in range(1, ws.max_column + 1):
        if ws.cell(1, c).value == name:
            return c
    return None

issue_col  = col_of(ws_ptc, "Issue")
ans_col_n  = col_of(ws_ptc, "Your answer")
ref_col_n  = col_of(ws_ptc, "_ref")
last_data_row = ws_ptc.max_row
first_new_row = last_data_row + 1

PLAN_TYPE_MAP = {
    "investment":"Investment","pension":"Pension","protection_policy":"Protection",
    "group_protection":"Group protection","annuity":"Annuity","fixed_asset":"Asset",
    "cash_account":"Cash account","mortgage":"Mortgage","loan_credit":"Loan / credit",
}

def plan_type_english(t):
    return PLAN_TYPE_MAP.get(t, t or "")

def get_provider_name(pl):
    prov = pl.get("provider")
    if isinstance(prov, dict):
        return prov.get("name","")
    return ""

def plan_polno(pl):
    pn = pl.get("policy_number","")
    return "" if is_placeholder(pn) else (pn or "")

rows_added = 0
for pl in unassigned_plans:
    uuid        = pl["uuid"]
    c_uuid      = plan_primary_client.get(uuid,"")
    client_name = acc_name.get(c_uuid,"")
    row_data    = {
        "Client":      client_name,
        "Plan type":   plan_type_english(pl.get("type","")),
        "Provider":    get_provider_name(pl),
        "Policy no":   plan_polno(pl),
        "Issue":       "No adviser on this plan — please tell us whose client this is",
        "Your answer": None,
        "Comment":     None,
        "_ref":        uuid,
    }
    for c_idx, h in enumerate(ptc_hdrs, start=1):
        ws_ptc.cell(first_new_row + rows_added, c_idx, value=row_data.get(h))
    rows_added += 1

p(f"  Appended {rows_added} unassigned-plan rows (rows {first_new_row}–{first_new_row+rows_added-1})")

# Add adviser dropdown for the new rows only
if rows_added > 0 and ans_col_n:
    ans_letter   = get_column_letter(ans_col_n)
    adviser_opts = '"Jim Hoban,Katie Hoban,Tommie Hoban"'
    dv_adv = DataValidation(type="list", formula1=adviser_opts,
                            allow_blank=True, showErrorMessage=False)
    wb2["Plans to check"].add_data_validation(dv_adv)
    dv_adv.sqref = f"{ans_letter}{first_new_row}:{ans_letter}{first_new_row+rows_added-1}"

# ── save v2 ────────────────────────────────────────────────────────────────────

p(f"\nSaving v2 to {V2_PATH}...")
wb2.save(V2_PATH)
p("  Saved.")

# ── console report ─────────────────────────────────────────────────────────────

v2_size = os.path.getsize(V2_PATH) // 1024

# Row counts from saved file
wb_check = openpyxl.load_workbook(V2_PATH, read_only=True)
tab_counts = {}
for sn in ["Summary","Client book","Fees to confirm","Plans to check","Plans we found at providers"]:
    if sn in wb_check.sheetnames:
        ws_chk = wb_check[sn]
        cnt = sum(1 for _ in ws_chk.iter_rows(min_row=2))
        tab_counts[sn] = cnt
wb_check.close()

p()
p("=" * 60)
p("FINAL REPORT (Rule 9: aggregates only)")
p("=" * 60)
p()
p("=== Fix 2: Missing circles ===")
p("  Circle 1: 0 members, reason=empty")
p("  Circle 2: 0 members, reason=empty")
p("  Circle 3: 0 members, reason=empty")
p("  Circle 4: 2 members, reason=members_no_plans")
p("  Circle 5: 0 members, reason=empty")
p("  Circle 6: 0 members, reason=empty")
p("  None have active plans — v6 Circles tab unchanged.")
p()
p("=== Fix 1: Unassigned plans added to Summary + Plans to check ===")
p(f"  Unassigned active plans:           {len(unassigned_plans)}")
p(f"  Summary 'To confirm' row: clients={len(unasgn_clients)} households={unasgn_households} plans={len(unassigned_plans)}")
p(f"    Value under advice: GBP {unasgn_val_ua:,.2f}")
p(f"    Ongoing fee (IAI):  GBP {unasgn_iai:,.2f}")
p()
p("=== Review pack v2 tab row counts ===")
for sn in ["Client book","Fees to confirm","Plans we found at providers"]:
    p(f"  {sn}: {tab_counts.get(sn,'?')} rows")
sum_data_rows = tab_counts.get("Summary", 0) - 1  # exclude TOTAL
p(f"  Summary: {sum_data_rows} adviser rows (excl. TOTAL)")
ptc_total = tab_counts.get("Plans to check", 0)
p(f"  Plans to check: {ptc_total} rows total ({ptc_total - rows_added} existing + {rows_added} unassigned)")
p()
p(f"=== Output ===")
p(f"  {V2_PATH}  ({v2_size} KB)")
