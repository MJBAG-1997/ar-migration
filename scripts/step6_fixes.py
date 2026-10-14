"""
HFP AR Migration - Step 6: Apply 5 corrections to Plans and Clients sheets.
Rule 9: no client names; adviser names OK.
"""
import os, json, re, csv
from collections import Counter
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

DATA_ROOT   = r"C:\dev\AR Migration"
SNAP_DIR    = os.path.join(DATA_ROOT, "snapshot_20261012T090906Z")
EXPORTS_DIR = os.path.join(DATA_ROOT, "exports_20261014")
MANIFEST    = os.path.join(EXPORTS_DIR, "_manifest.json")
MAPPINGS    = os.path.join(DATA_ROOT, "mappings_hfp.xlsx")
RATH_PATH   = os.path.join(DATA_ROOT, "Jim list  - Rathbones.xlsx")

PLANS_XLSX   = os.path.join(DATA_ROOT, "Plannr_Plans.xlsx")
CLIENTS_XLSX = os.path.join(DATA_ROOT, "Plannr_Clients.xlsx")
CHARGES_XLSX = os.path.join(DATA_ROOT, "Plannr_Charges.xlsx")

# Plan types excluded from "No policy number" check (v6 rule: fixed assets + others)
NO_PN_EXCLUDED = {
    "fixed_asset", "cash_account", "mortgage", "loan_credit",
    "general_insurance", "group_protection",
}

PLACEHOLDER_NORM = {"TBC","NA","N/A","NOTKNOWN","UNKNOWN","NOTONFILE",
                    "NOTSET","NONE","NOTAPPLICABLE"}

def blank(v):
    return v is None or str(v).strip() in ("", "None", "nan")

def is_placeholder_value(v):
    if blank(v): return False
    s = str(v).strip()
    n = re.sub(r"[\s\-]", "", s).upper()
    if n in PLACEHOLDER_NORM: return True
    if len(n) <= 5 and not re.match(r"^\d{4,5}$|^[A-Z]{2}\d{3,}$", n): return True
    return False

def to_bool(v):
    if blank(v): return None
    return str(v).strip().lower() in ("true", "yes", "1", "t")

def p(*a): print(*a, flush=True)

HDR_FONT  = Font(bold=True)
HDR_FILL  = PatternFill(fill_type="solid", fgColor="F2F2F2")
ORIG_FONT = Font(bold=True, color="7F6000")
ORIG_FILL = PatternFill(fill_type="solid", fgColor="FFF2CC")

def style_header(cell):
    h = str(cell.value or "")
    if h == "Review" or h.startswith("ORIGINAL_"):
        cell.font = ORIG_FONT
        cell.fill = ORIG_FILL
    else:
        cell.font = HDR_FONT
        cell.fill = HDR_FILL

def read_sheet(path, name):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[name]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    if not rows: return [], []
    hdrs = list(rows[0])
    data = [{hdrs[i]: r[i] for i in range(len(hdrs))}
            for r in rows[1:] if any(v is not None for v in r)]
    return hdrs, data

def write_sheet(path, name, hdrs, rows, preserve_sheets=None):
    """Delete sheet `name`, recreate it, write data. Preserve other sheets."""
    wb = openpyxl.load_workbook(path)
    idx = wb.sheetnames.index(name) if name in wb.sheetnames else 0
    del wb[name]
    ws = wb.create_sheet(name, idx)
    for i, h in enumerate(hdrs, 1):
        c = ws.cell(1, i, h)
        style_header(c)
    for ri, row in enumerate(rows, 2):
        for ci, h in enumerate(hdrs, 1):
            ws.cell(ri, ci, row.get(h))
    ws.freeze_panes = "A2"
    if len(rows) > 0:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(hdrs))}{len(rows)+1}"
    wb.save(path)

def keep_col(h, rows):
    if h in ("Review", "uuid"): return True
    if str(h or "").startswith("ORIGINAL_"):
        return any(not blank(r.get(h)) for r in rows)
    return any(not blank(r.get(h)) for r in rows)

# ── Load ──────────────────────────────────────────────────────────────────────

with open(MANIFEST, encoding="utf-8") as f:
    results = json.load(f).get("results", {})
def ep(l): return results.get(l, {}).get("path", "")

with open(ep("plan_export"),   encoding="utf-8-sig", newline="") as f:
    plan_rows_raw = list(csv.DictReader(f))
with open(ep("client_export"), encoding="utf-8-sig", newline="") as f:
    client_rows_raw = list(csv.DictReader(f))
circle_export_path = ep("circle_export")
if circle_export_path:
    with open(circle_export_path, encoding="utf-8-sig", newline="") as f:
        circle_export_rows = list(csv.DictReader(f))
else:
    circle_export_rows = []

plan_by_uuid   = {r["uuid"]: r for r in plan_rows_raw}
client_by_uuid = {r["uuid"]: r for r in client_rows_raw}

# Client -> assigned adviser UUID
client_adv = {r["uuid"]: (r.get("assigned_adviser_uuid","") or "").strip()
              for r in client_rows_raw}

# Adviser name lookup
adv_names = {}
for r in plan_rows_raw:
    u = (r.get("active_adviser_uuid","") or "").strip()
    n = (r.get("active_adviser_name","") or "").strip()
    if u and n: adv_names[u] = n
for r in client_rows_raw:
    u = (r.get("assigned_adviser_uuid","") or "").strip()
    n = (r.get("assigned_adviser","") or "").strip()
    if u and n: adv_names[u] = n

# Rathbones provider check
_wb_r = openpyxl.load_workbook(RATH_PATH, read_only=True, data_only=True)
rath_fund_nos = {str(r[0] or "").strip().upper()
                 for r in list(_wb_r.active.iter_rows(values_only=True))[1:] if r[0]}
_wb_r.close()
rath_plans = {r["uuid"] for r in plan_rows_raw
              if "rathbone" in str(r.get("provider","") or "").lower()}
not_on_rath = {p for p in rath_plans
               if (plan_by_uuid[p].get("policy_number","") or "").strip().upper()
               not in rath_fund_nos}

# Provider-matched plans for under_advice / agency fix
hdrs_ch, data_ch = read_sheet(CHARGES_XLSX, "Charges")
provider_matched = {str(r.get("plan_uuid",""))
                    for r in data_ch
                    if str(r.get("basis","")) in ("Provider statement","Rathbones")}
p(f"Provider-matched plans (Provider statement + Rathbones): {len(provider_matched)}")

# Data for client Review label rebuild
# Circles per client (account_uuid → count of circles)
from collections import defaultdict as _dd
client_circle_count = _dd(int)
for r in circle_export_rows:
    cuuid = (r.get("account_uuid","") or "").strip()
    if cuuid: client_circle_count[cuuid] += 1

# Active plans per client (client_uuid → count of active plans)
client_active_plan_count = _dd(int)
for r in plan_rows_raw:
    if str(r.get("status","")).strip().lower() == "active":
        for cuuid in (r.get("client_uuids","") or "").split(","):
            cuuid = cuuid.strip()
            if cuuid: client_active_plan_count[cuuid] += 1

UK_POSTCODE_RE = re.compile(
    r"^[A-Z]{1,2}\d[A-Z\d]?\s*\d[A-Z]{2}$", re.I
)

# Status mapping
status_map = {}
ba_statuses = []
_wb_m = openpyxl.load_workbook(MAPPINGS, read_only=True, data_only=True)
if "client_statuses" in _wb_m.sheetnames:
    sm = list(_wb_m["client_statuses"].iter_rows(values_only=True))
    sh = [str(h or "") for h in sm[0]]
    src_i = next((i for i,h in enumerate(sh) if h.lower()=="source_name"), 1)
    dst_i = next((i for i,h in enumerate(sh) if h.lower()=="dest_name"), 3)
    for r in sm[1:]:
        s = str(r[src_i] or "").strip()
        d = str(r[dst_i] or "").strip()
        if s: status_map[s] = d
        if d and d not in ba_statuses: ba_statuses.append(d)
_wb_m.close()
p(f"Status map: {status_map}")

# ── Process Plans sheet ────────────────────────────────────────────────────────

p("\nProcessing Plans...")
hdrs_pl, data_pl = read_sheet(PLANS_XLSX, "Plans")
hdrs_pl = list(hdrs_pl)

# Ensure adviser_uuid column exists
if "adviser_uuid" not in hdrs_pl:
    hdrs_pl.append("adviser_uuid")
# Ensure under_advice and agency columns
for col in ("under_advice", "agency_in_your_name",
            "ORIGINAL_under_advice", "ORIGINAL_agency_in_your_name"):
    if col not in hdrs_pl:
        hdrs_pl.append(col)
# Remove ORIGINAL_policy_number from headers (it will be cleared and dropped)
# We'll decide at end via keep_col

# Counters
ph_active = ph_total = 0
no_pn_count = 0
changed_ua = changed_aia = 0
review_pl = Counter()
no_pn_by_type = Counter()

def resolve_plan_adviser(puuid):
    pe = plan_by_uuid.get(puuid, {})
    own = (pe.get("active_adviser_uuid","") or "").strip()
    if own: return own
    for cuuid in (pe.get("client_uuids","") or "").split(","):
        cuuid = cuuid.strip()
        if cuuid:
            a = client_adv.get(cuuid,"")
            if a: return a
    return ""

updated_pl = []
for orig_row in data_pl:
    row = dict(orig_row)
    puuid   = str(row.get("uuid","") or "")
    pstatus = str(row.get("status","") or "").strip().lower()
    is_act  = pstatus == "active"
    pe      = plan_by_uuid.get(puuid, {})
    ptype   = str(pe.get("type","") or "").strip().lower()

    # Fix 1: restore placeholder policy number
    orig_pn = row.get("ORIGINAL_policy_number")
    if not blank(orig_pn):
        row["policy_number"] = orig_pn
        row["ORIGINAL_policy_number"] = None
        ph_total += 1
        if is_act: ph_active += 1

    pn = (row.get("policy_number","") or "").strip()

    # Build Review labels (fresh)
    labels = []

    if is_act and is_placeholder_value(pn):
        labels.append("Placeholder policy number")
    elif is_act and blank(pn) and ptype not in NO_PN_EXCLUDED:
        labels.append("No policy number")
        no_pn_count += 1
        no_pn_by_type[ptype or "(blank type)"] += 1

    adv_uuid = resolve_plan_adviser(puuid)
    if not adv_uuid:
        labels.append("No adviser")
    else:
        row["adviser_uuid"] = adv_uuid

    if puuid in not_on_rath:
        labels.append("Not on Rathbones list")

    row["Review"] = "; ".join(labels) if labels else ""
    for lbl in labels:
        review_pl[lbl] += 1

    # Fix 5: under_advice and agency_in_your_name
    if puuid in provider_matched:
        curr_ua  = str(pe.get("under_advice","")  or "").strip()
        curr_aia = str(pe.get("agency_in_your_name","") or "").strip()
        ua_bool  = to_bool(curr_ua)
        aia_bool = to_bool(curr_aia)

        if ua_bool is not True:
            row["ORIGINAL_under_advice"] = curr_ua if curr_ua else None
            row["under_advice"] = True
            changed_ua += 1
        else:
            row["under_advice"] = True
            row["ORIGINAL_under_advice"] = None

        if aia_bool is not True:
            row["ORIGINAL_agency_in_your_name"] = curr_aia if curr_aia else None
            row["agency_in_your_name"] = True
            changed_aia += 1
        else:
            row["agency_in_your_name"] = True
            row["ORIGINAL_agency_in_your_name"] = None
    else:
        # Copy current values from export (populate if columns were missing)
        curr_ua  = pe.get("under_advice")
        curr_aia = pe.get("agency_in_your_name")
        if blank(row.get("under_advice")):
            row["under_advice"] = curr_ua if not blank(curr_ua) else None
        if blank(row.get("agency_in_your_name")):
            row["agency_in_your_name"] = curr_aia if not blank(curr_aia) else None
        row.setdefault("ORIGINAL_under_advice", None)
        row.setdefault("ORIGINAL_agency_in_your_name", None)

    updated_pl.append(row)

# Drop empty cols (ORIGINAL_policy_number now all blank → dropped)
final_hdrs_pl = [h for h in hdrs_pl if keep_col(h, updated_pl)]
write_sheet(PLANS_XLSX, "Plans", final_hdrs_pl, updated_pl,
            preserve_sheets=["Provider_plans_not_in_Plannr"])
p(f"  Saved Plannr_Plans.xlsx ({len(updated_pl)} rows, {len(final_hdrs_pl)} cols)")
p(f"  No-policy-number by type: {dict(no_pn_by_type)}")

# ── Process Clients sheet ──────────────────────────────────────────────────────

p("\nProcessing Clients...")
hdrs_cl, data_cl = read_sheet(CLIENTS_XLSX, "Clients")
hdrs_cl = list(hdrs_cl)

# Ensure status column exists
STATUS_COL = "status"
if STATUS_COL not in hdrs_cl:
    # Insert after Review (position 1)
    review_pos = hdrs_cl.index("Review") if "Review" in hdrs_cl else 0
    hdrs_cl.insert(review_pos + 1, STATUS_COL)
    p(f"  Inserting '{STATUS_COL}' column at position {review_pos+1}")

p(f"  Status column header: '{STATUS_COL}'")

review_cl = Counter()
status_after = Counter()
src_status_breakdown = Counter()

updated_cl = []
for orig_row in data_cl:
    row = dict(orig_row)
    cuuid = str(row.get("uuid","") or "")
    ce = client_by_uuid.get(cuuid, {})

    # Fix 4: populate status
    src = (ce.get("status_name","") or "").strip()
    src_status_breakdown[src] += 1
    if blank(src):
        dest_status = "Level 2 / Navigator"
    else:
        dest_status = status_map.get(src)
        if not dest_status:
            # fallback heuristic
            sl = src.lower()
            if "inactive" in sl or "archived" in sl or "deceased" in sl:
                dest_status = "ARCHIVED"
            elif "prospect" in sl or "onboarding" in sl:
                dest_status = "Prospect"
            else:
                dest_status = "Level 2 / Navigator"
    row[STATUS_COL] = dest_status
    status_after[dest_status] += 1

    # Rebuild ALL client Review labels from scratch
    labels = []
    src_status = (ce.get("status_uuid","") or "").strip()
    if blank(src_status) and client_active_plan_count.get(cuuid, 0) == 0:
        labels.append("No status and no plans")
    if client_circle_count.get(cuuid, 0) > 1:
        labels.append("In more than one household")
    pc = str(row.get("postcode","") or "").strip()
    if pc and not UK_POSTCODE_RE.match(pc):
        labels.append("Check postcode")
    # Fix 3: No adviser
    adv = (ce.get("assigned_adviser_uuid","") or "").strip()
    if blank(adv):
        labels.append("No adviser")
    row["Review"] = "; ".join(labels) if labels else ""
    for lbl in labels:
        review_cl[lbl] += 1

    updated_cl.append(row)

final_hdrs_cl = [h for h in hdrs_cl if keep_col(h, updated_cl)]
# status must always be present
if STATUS_COL not in final_hdrs_cl:
    final_hdrs_cl.insert(1, STATUS_COL)

write_sheet(CLIENTS_XLSX, "Clients", final_hdrs_cl, updated_cl)
p(f"  Saved Plannr_Clients.xlsx ({len(updated_cl)} rows, {len(final_hdrs_cl)} cols)")

# ── Print results ─────────────────────────────────────────────────────────────

p("\n" + "=" * 68)
p("COUNTS")
p("=" * 68)

p(f"\nReview labels - Plannr_Plans.xlsx > Plans:")
for lbl, cnt in sorted(review_pl.items(), key=lambda x: -x[1]):
    p(f"  {lbl:<50} {cnt:>6}")
p(f"  (no Review label)                                  "
  f"{len(updated_pl) - sum(review_pl.values()):>6}")

p(f"\nReview labels - Plannr_Clients.xlsx:")
for lbl, cnt in sorted(review_cl.items(), key=lambda x: -x[1]):
    p(f"  {lbl:<50} {cnt:>6}")
p(f"  (no Review label)                                  "
  f"{len(updated_cl) - sum(review_cl.values()):>6}")

p(f"\nPlaceholder policy numbers:")
p(f"  Active plans with placeholder (now shown in policy_number): {ph_active}")
p(f"  All plans with placeholder restored:                        {ph_total}")

p(f"\nNo policy number (active, v6 rule):")
p(f"  Count: {no_pn_count}")
p(f"  By type: {dict(no_pn_by_type)}")

p(f"\nUnder_advice / agency_in_your_name changes:")
p(f"  Plans where under_advice changed to True:           {changed_ua}")
p(f"  Plans where agency_in_your_name changed to True:    {changed_aia}")
p(f"  (Provider statement: {sum(1 for r in data_ch if str(r.get('basis',''))=='Provider statement')},  "
  f"Rathbones: {sum(1 for r in data_ch if str(r.get('basis',''))=='Rathbones')})")

# Per-adviser table (fix-3 logic)
p(f"\n{'Adviser':<35} {'Clients':>8}  {'Plans':>7}")
p("-" * 55)
plan_adv_ctr  = Counter(resolve_plan_adviser(r["uuid"]) for r in plan_rows_raw)
client_adv_ctr = Counter((r.get("assigned_adviser_uuid","") or "").strip()
                          for r in client_rows_raw)
all_uuids = set(plan_adv_ctr) | set(client_adv_ctr)
adv_table = []
for uid in all_uuids:
    name = adv_names.get(uid) or (uid[:8] if uid else "(no adviser)")
    adv_table.append((name, client_adv_ctr.get(uid,0), plan_adv_ctr.get(uid,0)))
adv_table.sort(key=lambda x: (-x[2], x[0]))
for name, clients, plans in adv_table:
    p(f"  {name:<35} {clients:>8}  {plans:>7}")

# Status table after fix
p(f"\n{'Plannr source status':<28} {'Count':>6}  BA status (set in sheet)")
p("-" * 68)
for src, cnt in sorted(src_status_breakdown.items(), key=lambda x: -x[1]):
    if blank(src):
        dest = "Level 2 / Navigator"
    else:
        dest = status_map.get(src)
        if not dest:
            sl = src.lower()
            dest = ("ARCHIVED" if any(k in sl for k in ("inactive","archived","deceased"))
                    else "Prospect" if any(k in sl for k in ("prospect","onboarding"))
                    else "Level 2 / Navigator")
    p(f"  {src or '(blank)':<28} {cnt:>6}  {dest}")

p(f"\nStatus column header in Plannr_Clients.xlsx: '{STATUS_COL}'")
p(f"BA status options the AR can choose from: {', '.join(sorted(set(ba_statuses)))}")

p("\nSTOP")
