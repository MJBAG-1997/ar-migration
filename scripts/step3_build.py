"""
HFP AR Migration – Step 3: Build hfp_importer_v1.xlsx

Sheets: Clients, Plans, Charges, Circles, Expenditure, Plannr_Premiums,
        Relationships, Dependants, Valuations, Income_and_Employment,
        Provider_plans_not_in_Plannr

Also: Notes investigation (console counts, Rule 9, no PII).

Rules applied:
- G3: source-site GET-only (no API calls here)
- G4: no destination writes
- Rule 9: no PII in console; aggregates and UUIDs only
- external_reference_company = data_migration; external_reference_value = source UUID
- ORIGINAL_ columns for every changed value
- Drop importer columns empty on every row (ORIGINAL_ columns always kept)
- Source UUIDs placed in external_reference_value (rightmost importer cols)
"""
import os, sys, json, re, csv, glob
from collections import defaultdict
from datetime import datetime, timezone, date
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

# ── Constants ─────────────────────────────────────────────────────────────────

DATA_ROOT   = r"C:\dev\AR Migration"
SNAP_DIR    = os.path.join(DATA_ROOT, "snapshot_20261012T090906Z")
EXPORTS_DIR = os.path.join(DATA_ROOT, "exports_20261014")
DEST_JSON   = os.path.join(DATA_ROOT, "block6_destination.json")
MAPPINGS    = os.path.join(DATA_ROOT, "mappings_hfp.xlsx")
RATH_PATH   = os.path.join(DATA_ROOT, "Jim list  - Rathbones.xlsx")
WB_PATH     = os.path.join(DATA_ROOT,
    "Hoban_Financial_Planning_Ltd_260922_Payment_Rates_113727875.xlsx")
MANIFEST    = os.path.join(EXPORTS_DIR, "_manifest.json")
OUT_PATH    = os.path.join(DATA_ROOT, "hfp_importer_v1.xlsx")
EXT_REF_CO  = "data_migration"
BA_GROUP_UUID = "b5749882-0a12-4be3-a022-7cf98d4b56b9"   # BA "default" group

def p(*a): print(*a, flush=True)

# ── Helpers ───────────────────────────────────────────────────────────────────

PLACEHOLDER_UPPER = {
    "TBC","NA","N/A","NOTKNOWN","UNKNOWN","NOTONFILE",
    "NOTSET","NONE","NOT SET","NOTON","NOTAPPLICABLE",
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
    return False

def pence_to_pounds(v):
    try:
        return round(int(str(v)) / 100.0, 2)
    except (TypeError, ValueError):
        try:
            return round(float(str(v)) / 100.0, 2)
        except Exception:
            return None

def first_of(csv_val):
    """Return first UUID from a comma-separated multi-value CSV field."""
    if not csv_val:
        return ""
    return str(csv_val).split(",")[0].strip()

def pipe_join(items):
    return "|".join(str(x) for x in items if x)

def blank(v):
    return v is None or str(v).strip() in ("", "None", "nan")

def clean_val(v):
    """Flatten nested dicts (e.g. monetary amounts) to a plain Python scalar."""
    if v is None:
        return None
    if isinstance(v, (int, float, bool)):
        return v
    if isinstance(v, str):
        return v or None
    if isinstance(v, dict):
        # Monetary amount dict
        if "amount" in v:
            try:
                return round(int(str(v["amount"])) / 100.0, 2)
            except (TypeError, ValueError):
                pass
        # Fall back to string representation (safe for Excel)
        return str(v)
    if isinstance(v, list):
        return "|".join(str(x) for x in v if x is not None) or None
    return str(v)

# ── Style helpers ─────────────────────────────────────────────────────────────

NAVY = "020849"
HDR_FILL = PatternFill(fill_type="solid", fgColor="F2F2F2")
HDR_FONT = Font(bold=True)
ORIG_FILL = PatternFill(fill_type="solid", fgColor="FFF2CC")
ORIG_FONT = Font(color="7F6000")

def style_ws(ws):
    for cell in ws[1]:
        cell.font = HDR_FONT
        cell.fill = HDR_FILL
        if str(cell.value or "").startswith("ORIGINAL_"):
            cell.font = Font(bold=True, color="7F6000")
            cell.fill = ORIG_FILL
    ws.freeze_panes = "A2"
    if ws.max_row > 1 and ws.max_column >= 1:
        ws.auto_filter.ref = (
            f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
        )
    for col_cells in ws.columns:
        max_len = 12
        col_letter = get_column_letter(col_cells[0].column)
        for cell in col_cells:
            if cell.value is not None:
                try:
                    max_len = max(max_len, min(50, len(str(cell.value)) + 2))
                except Exception:
                    pass
        ws.column_dimensions[col_letter].width = max_len


def write_sheet(ws, headers, rows):
    """Write headers + rows to ws. No column filtering here."""
    for i, h in enumerate(headers, 1):
        ws.cell(row=1, column=i, value=h)
    for r_idx, row in enumerate(rows, 2):
        for c_idx, h in enumerate(headers, 1):
            ws.cell(row=r_idx, column=c_idx, value=row.get(h))
    style_ws(ws)


def drop_empty_cols(headers, rows):
    """
    Return (new_headers, new_rows) dropping columns that are entirely blank,
    except ORIGINAL_ columns are always kept.
    """
    keep = []
    for h in headers:
        if h.startswith("ORIGINAL_"):
            keep.append(h)
            continue
        if any(not blank(r.get(h)) for r in rows):
            keep.append(h)
    new_rows = [{h: r.get(h) for h in keep} for r in rows]
    return keep, new_rows


# ── Load data sources ─────────────────────────────────────────────────────────

p("=" * 68)
p("HFP AR Migration — Step 3: Build importer workbook")
p("=" * 68)

p("\n[0] Loading manifest and importer keys...")
with open(MANIFEST, encoding="utf-8") as f:
    manifest = json.load(f)
results = manifest.get("results", {})

with open(DEST_JSON, encoding="utf-8") as f:
    dest = json.load(f)
imp_keys = {k: v["store_keys"] for k, v in dest["importer_keys"].items()}

def export_path(label):
    return results.get(label, {}).get("path", "")

p("\n[1] Loading snapshot data...")
with open(os.path.join(SNAP_DIR, "plans_dedup.json"), encoding="utf-8") as f:
    snap_plans = json.load(f)          # dict uuid→plan
p(f"  Plans: {len(snap_plans)}")

with open(os.path.join(SNAP_DIR, "accounts.json"), encoding="utf-8") as f:
    snap_accounts = json.load(f)       # list
snap_client_uuids = {a["uuid"] for a in snap_accounts if a.get("type") == "client"}
p(f"  Accounts: {len(snap_accounts)}  clients: {len(snap_client_uuids)}")

with open(os.path.join(SNAP_DIR, "account_details.json"), encoding="utf-8") as f:
    acc_details = json.load(f)         # dict uuid→detail
p(f"  Account details: {len(acc_details)}")

with open(os.path.join(SNAP_DIR, "circles.json"), encoding="utf-8") as f:
    snap_circles = json.load(f)        # list
p(f"  Circles: {len(snap_circles)}")

with open(os.path.join(SNAP_DIR, "notes_dedup.json"), encoding="utf-8") as f:
    snap_notes = json.load(f)          # dict uuid→note
p(f"  Notes (dedup): {len(snap_notes)}")

p("\n[2] Loading mappings...")
import openpyxl as _ox
_wb_map = _ox.load_workbook(MAPPINGS, read_only=True, data_only=True)
# Expenditure mapping: source_section+source_line → (dest_section, dest_line, confidence)
_ws_exp = _wb_map["Expenditure"]
exp_map = {}   # (norm_section, norm_line) → (dest_section, dest_line, is_custom)
_exp_rows = list(_ws_exp.iter_rows(min_row=2, values_only=True))
for row in _exp_rows:
    if row[0] is None:
        continue
    src_sec = str(row[0] or "").strip()
    src_ln  = str(row[1] or "").strip()
    dst_sec = str(row[3] or "").strip()
    dst_ln  = str(row[4] or "").strip()
    conf    = str(row[5] or "").strip().lower()
    is_cust = (conf == "custom" or dst_ln.upper() == "CUSTOM")
    exp_map[(src_sec.lower(), src_ln.lower())] = (dst_sec, dst_ln, is_cust)
_wb_map.close()
p(f"  Expenditure mapping rows: {len(exp_map)}")

# ── Notes investigation ───────────────────────────────────────────────────────

p("\n" + "=" * 68)
p("NOTES INVESTIGATION (counts only, Rule 9)")
p("=" * 68)

note_export_path = export_path("note_export")
export_uuids_note = set()
export_by_model = {}
with open(note_export_path, encoding="utf-8-sig", newline="") as f:
    reader = csv.DictReader(f)
    for row in reader:
        uuid = row.get("uuid", "")
        if uuid not in export_uuids_note:
            export_uuids_note.add(uuid)
            mt = row.get("model_type", "")
            export_by_model[mt] = export_by_model.get(mt, 0) + 1

snap_uuids_note = set(snap_notes.keys())
snap_by_type = {}
for v in snap_notes.values():
    t = v.get("notable_type", "") or v.get("type", "")
    snap_by_type[t] = snap_by_type.get(t, 0) + 1

in_both     = export_uuids_note & snap_uuids_note
export_only = export_uuids_note - snap_uuids_note
snap_only   = snap_uuids_note - export_uuids_note

# Export-only breakdown by model_type
eo_by_model = {}
with open(note_export_path, encoding="utf-8-sig", newline="") as f:
    reader = csv.DictReader(f)
    for row in reader:
        if row.get("uuid", "") in export_only:
            mt = row.get("model_type", "")
            eo_by_model[mt] = eo_by_model.get(mt, 0) + 1

p(f"\nExport (unique UUIDs): {len(export_uuids_note)}")
p(f"Snapshot notes_dedup:  {len(snap_uuids_note)}")
p(f"  In both:             {len(in_both)}")
p(f"  Export only:         {len(export_only)}")
p(f"  Snapshot only:       {len(snap_only)}")

p(f"\nExport by model_type:")
for mt, n in sorted(export_by_model.items()):
    p(f"  {mt:<20} {n:>5}")

p(f"\nSnapshot by notable_type:")
for nt, n in sorted(snap_by_type.items()):
    p(f"  {nt:<20} {n:>5}")

p(f"\nExport-only by model_type (not in snapshot):")
for mt, n in sorted(eo_by_model.items()):
    p(f"  {mt:<20} {n:>5}")

# Counts for Step 4
note_task_count  = export_by_model.get("task", 0)
note_case_count  = export_by_model.get("case", 0)
note_board_task  = export_by_model.get("board_task", 0)
note_total_exportable = sum(
    v for k, v in export_by_model.items() if k != "board_task"
)

p(f"\nFor Step 4 import: task={note_task_count}, case={note_case_count}")
p(f"board_task={note_board_task} (not importable via notes importer)")

# ── Build importer workbook ───────────────────────────────────────────────────

p("\n" + "=" * 68)
p("Building hfp_importer_v1.xlsx")
p("=" * 68)

wb = openpyxl.Workbook()
wb.remove(wb.active)

# ─── Sheet 1: Clients ────────────────────────────────────────────────────────

p("\n[Sheet 1] Clients...")
client_path = export_path("client_export")
client_importer_keys = imp_keys["clients"]

with open(client_path, encoding="utf-8-sig", newline="") as f:
    reader = csv.DictReader(f)
    export_hdrs = reader.fieldnames
    export_rows = list(reader)

# Build rows in importer key order, take direct matches only
direct_match = set(export_hdrs) & set(client_importer_keys)
rows_clients = []
for row in export_rows:
    out = {}
    for k in client_importer_keys:
        if k in row:
            out[k] = row[k] or None
        else:
            out[k] = None
    out["external_reference_company"] = EXT_REF_CO
    out["external_reference_value"]   = row.get("uuid", "")
    rows_clients.append(out)

# Build ordered headers: importer keys in order
hdrs_clients = list(client_importer_keys)
if "external_reference_company" not in hdrs_clients:
    hdrs_clients += ["external_reference_company", "external_reference_value"]

hdrs_clients, rows_clients = drop_empty_cols(hdrs_clients, rows_clients)
ws = wb.create_sheet("Clients")
write_sheet(ws, hdrs_clients, rows_clients)
p(f"  Rows: {len(rows_clients)}  Cols: {len(hdrs_clients)}")

# ─── Sheet 2: Plans ──────────────────────────────────────────────────────────

p("\n[Sheet 2] Plans...")
plan_path = export_path("plan_export")
plan_importer_keys = imp_keys["plans"]
snap_time = datetime(2026, 10, 12, 9, 9, 6, tzinfo=timezone.utc)

# Build sub_account policy_number lookup: plan_uuid → {sub_acc_uuid: policy_number}
sub_acc_polno = {}
for uuid, pl in snap_plans.items():
    for sa in (pl.get("_sub_accounts") or []):
        sa_uuid = sa.get("uuid")
        pn = sa.get("policy_number") or ""
        sub_acc_polno.setdefault(uuid, {})[sa_uuid] = pn

# Build plan owners lookup from snapshot
snap_plan_owners = {}   # plan_uuid → [client_uuid, ...]
for uuid, pl in snap_plans.items():
    owners = pl.get("_owners") or []
    snap_plan_owners[uuid] = [o.get("client_uuid", "") for o in owners if o.get("client_uuid")]

with open(plan_path, encoding="utf-8-sig", newline="") as f:
    reader = csv.DictReader(f)
    plan_export_rows = list(reader)

p(f"  Export rows: {len(plan_export_rows)}")
new_plan_uuids = [r["uuid"] for r in plan_export_rows if r.get("uuid") not in snap_plans]
p(f"  New (post-snapshot): {len(new_plan_uuids)}")

# Identify joint plans for ownership_split report
joint_blank_count = 0
rows_plans = []
for row in plan_export_rows:
    out = {}
    for k in plan_importer_keys:
        if k == "external_reference_company":
            out[k] = EXT_REF_CO
        elif k == "external_reference_value":
            out[k] = row.get("uuid", "")
        elif k == "ownership_split":
            # Always None in source; leave blank; count joints
            plan_uuid = row.get("uuid", "")
            owners_from_snap = snap_plan_owners.get(plan_uuid)
            if owners_from_snap and len(owners_from_snap) > 1:
                joint_blank_count += 1
            out[k] = None
        elif k in row:
            out[k] = row[k] or None
        else:
            out[k] = None

    # Apply v6 rule: blank placeholder policy_number, keep ORIGINAL_
    raw_pn = row.get("policy_number", "")
    if is_placeholder(raw_pn):
        out["ORIGINAL_policy_number"] = raw_pn if raw_pn else None
        out["policy_number"] = None
    else:
        out["ORIGINAL_policy_number"] = None

    rows_plans.append(out)

# Insert ORIGINAL_policy_number after policy_number in headers
hdrs_plans = list(plan_importer_keys)
if "policy_number" in hdrs_plans:
    pn_idx = hdrs_plans.index("policy_number")
    if "ORIGINAL_policy_number" not in hdrs_plans:
        hdrs_plans.insert(pn_idx + 1, "ORIGINAL_policy_number")
else:
    hdrs_plans.append("ORIGINAL_policy_number")

if "external_reference_company" not in hdrs_plans:
    hdrs_plans += ["external_reference_company", "external_reference_value"]

hdrs_plans, rows_plans = drop_empty_cols(hdrs_plans, rows_plans)
ws = wb.create_sheet("Plans")
write_sheet(ws, hdrs_plans, rows_plans)
placeholder_plan_count = sum(1 for r in rows_plans if r.get("ORIGINAL_policy_number"))
p(f"  Rows: {len(rows_plans)}  Cols: {len(hdrs_plans)}")
p(f"  Placeholder policy numbers blanked: {placeholder_plan_count}")
p(f"  Joint plans (ownership_split left blank): {joint_blank_count}")

# ─── Sheet 3: Charges ────────────────────────────────────────────────────────

p("\n[Sheet 3] Charges...")
charges_path = export_path("charges_export")
charges_importer_keys = imp_keys["charges"]

# Build plan→first_owner_uuid from plan export
plan_first_owner = {}
for row in plan_export_rows:
    plan_first_owner[row.get("uuid", "")] = first_of(row.get("client_uuids", ""))

with open(charges_path, encoding="utf-8-sig", newline="") as f:
    reader = csv.DictReader(f)
    charges_export_rows = list(reader)

rows_charges = []
for row in charges_export_rows:
    out = {}
    for k in charges_importer_keys:
        if k == "client_uuid":
            out[k] = plan_first_owner.get(row.get("plan_uuid", ""), "") or first_of(row.get("client_uuids", ""))
        elif k == "premium_uuid":
            out[k] = None    # always blank per instruction
        elif k == "external_reference_company":
            out[k] = EXT_REF_CO
        elif k == "external_reference_value":
            out[k] = row.get("uuid", "")
        elif k in row:
            out[k] = row[k] or None
        else:
            out[k] = None
    rows_charges.append(out)

hdrs_charges = list(charges_importer_keys)
if "external_reference_company" not in hdrs_charges:
    hdrs_charges += ["external_reference_company", "external_reference_value"]

hdrs_charges, rows_charges = drop_empty_cols(hdrs_charges, rows_charges)
ws = wb.create_sheet("Charges")
write_sheet(ws, hdrs_charges, rows_charges)
p(f"  Rows: {len(rows_charges)}  Cols: {len(hdrs_charges)}")

# ─── Sheet 4: Circles ────────────────────────────────────────────────────────

p("\n[Sheet 4] Circles...")
circle_path = export_path("circle_export")
circles_importer_keys = imp_keys["circles"]

# Load circle export and group by circle_uuid
circle_members = defaultdict(list)   # circle_uuid → [account_uuid, ...]
circle_names   = {}
with open(circle_path, encoding="utf-8-sig", newline="") as f:
    reader = csv.DictReader(f)
    for row in reader:
        cuuid = row.get("circle_uuid", "")
        auuid = row.get("account_uuid", "")
        cname = row.get("circle_name", "")
        if cuuid:
            if auuid and auuid not in circle_members[cuuid]:
                circle_members[cuuid].append(auuid)
            circle_names[cuuid] = cname

p(f"  Unique circles in export: {len(circle_members)}")

# Build circle rows (one per circle, include only those with members)
rows_circles = []
for cuuid, members in circle_members.items():
    if not members:
        continue     # skip empty
    out = {
        "uuid":                     cuuid,
        "name":                     circle_names.get(cuuid, ""),
        "account_uuids":            pipe_join(members),
        "group_uuids":              BA_GROUP_UUID,
        "external_reference_company": EXT_REF_CO,
        "external_reference_value": cuuid,
        "external_reference_company_2": None,   # placeholder removed below
    }
    rows_circles.append(out)

hdrs_circles = list(circles_importer_keys)
if "external_reference_company" not in hdrs_circles:
    hdrs_circles += ["external_reference_company", "external_reference_value"]

# Clean helper key
rows_circles = [{k: r.get(k) for k in hdrs_circles} for r in rows_circles]
hdrs_circles, rows_circles = drop_empty_cols(hdrs_circles, rows_circles)
ws = wb.create_sheet("Circles")
write_sheet(ws, hdrs_circles, rows_circles)
p(f"  Rows: {len(rows_circles)}  Cols: {len(hdrs_circles)}")

# ─── Sheet 5: Expenditure ────────────────────────────────────────────────────

p("\n[Sheet 5] Expenditure...")
exp_exp_path = export_path("client_expenditure_export")
expenses_importer_keys = imp_keys["expenses"]

with open(exp_exp_path, encoding="utf-8-sig", newline="") as f:
    reader = csv.DictReader(f)
    exp_export_rows = list(reader)

rows_exp = []
unmatched_exp = 0
for row in exp_export_rows:
    src_sec = (row.get("section") or "").strip()
    src_ln  = (row.get("type")    or "").strip()
    mapping = exp_map.get((src_sec.lower(), src_ln.lower()))
    if mapping:
        dst_sec, dst_ln, is_cust = mapping
    else:
        dst_sec  = src_sec
        dst_ln   = src_ln
        is_cust  = False
        unmatched_exp += 1

    out = {
        "client_uuid":             row.get("client_uuid") or None,
        "life_stage_name":         row.get("life_stage") or None,
        "life_stage_section_name": dst_sec or None,
        "expense_name":            dst_ln or None,
        "is_custom_expense":       True if is_cust else None,
        "amount":                  row.get("amount") or None,
        "external_reference_company": EXT_REF_CO,
        "external_reference_value":   row.get("client_uuid") or None,
        # ORIGINAL_ cols for mapped section/line
        "ORIGINAL_section":        src_sec if (mapping and dst_sec != src_sec) else None,
        "ORIGINAL_type":           src_ln  if (mapping and dst_ln  != src_ln)  else None,
    }
    rows_exp.append(out)

# Build headers: importer keys + ORIGINAL_ cols
hdrs_exp = list(expenses_importer_keys)
if "external_reference_company" not in hdrs_exp:
    hdrs_exp += ["external_reference_company", "external_reference_value"]
# Insert ORIGINAL_ after the relevant cols
for orig in ["ORIGINAL_section", "ORIGINAL_type"]:
    if orig not in hdrs_exp:
        hdrs_exp.append(orig)

hdrs_exp, rows_exp = drop_empty_cols(hdrs_exp, rows_exp)
ws = wb.create_sheet("Expenditure")
write_sheet(ws, hdrs_exp, rows_exp)
p(f"  Rows: {len(rows_exp)}  Cols: {len(hdrs_exp)}  Unmatched in mapping: {unmatched_exp}")

# ─── Sheet 6: Plannr_Premiums ────────────────────────────────────────────────

p("\n[Sheet 6] Plannr_Premiums...")
prem_path = export_path("premium_export")
premiums_importer_keys = imp_keys["premiums"]

with open(prem_path, encoding="utf-8-sig", newline="") as f:
    reader = csv.DictReader(f)
    prem_export_rows = list(reader)

rows_prem = []
for row in prem_export_rows:
    out = {}
    for k in premiums_importer_keys:
        if k in row:
            out[k] = row[k] or None
        else:
            out[k] = None
    rows_prem.append(out)

hdrs_prem = list(premiums_importer_keys)
hdrs_prem, rows_prem = drop_empty_cols(hdrs_prem, rows_prem)
ws = wb.create_sheet("Plannr_Premiums")
write_sheet(ws, hdrs_prem, rows_prem)
p(f"  Rows: {len(rows_prem)}  Cols: {len(hdrs_prem)}")

# ─── Sheet 7: Relationships ──────────────────────────────────────────────────

p("\n[Sheet 7] Relationships...")
rels_importer_keys = imp_keys["relationships"]

# Build linked-client map: acc_uuid → {linked_uuid: relationship}
lc_map = {}
for acc_uuid, detail in acc_details.items():
    for lc in detail.get("linked_clients", []):
        lu = lc.get("uuid", "")
        lr = lc.get("relationship", "")
        lc_map.setdefault(acc_uuid, {})[lu] = lr

# Deduplicate pairs: keep one row per (sorted A, B) pair
seen_pairs = set()
rows_rels = []
for acc_uuid, linked in lc_map.items():
    for linked_uuid, rel in linked.items():
        pair_key = tuple(sorted([acc_uuid, linked_uuid]))
        if pair_key in seen_pairs:
            continue
        seen_pairs.add(pair_key)
        # inverse_relationship = what linked_uuid thinks of acc_uuid
        inv_rel = lc_map.get(linked_uuid, {}).get(acc_uuid, "")
        rows_rels.append({
            "account_uuid":         acc_uuid,
            "related_account_uuid": linked_uuid,
            "relationship":         rel or None,
            "inverse_relationship": inv_rel or None,
        })

hdrs_rels = list(rels_importer_keys)
hdrs_rels, rows_rels = drop_empty_cols(hdrs_rels, rows_rels)
ws = wb.create_sheet("Relationships")
write_sheet(ws, hdrs_rels, rows_rels)
p(f"  Rows: {len(rows_rels)}  Cols: {len(hdrs_rels)}")

# ─── Sheet 8: Dependants ─────────────────────────────────────────────────────

p("\n[Sheet 8] Dependants...")
deps_importer_keys = imp_keys["dependants"]

# Collect all unique dependants from snapshot, accumulate account_uuids
all_deps = {}   # dep_uuid → {dep: {...}, accounts: [uuid, ...]}
for acc_uuid, detail in acc_details.items():
    for dep in detail.get("dependants", []):
        dep_uuid = dep.get("uuid", "")
        if not dep_uuid:
            continue
        if dep_uuid not in all_deps:
            all_deps[dep_uuid] = {"dep": dep, "accounts": []}
        if acc_uuid not in all_deps[dep_uuid]["accounts"]:
            all_deps[dep_uuid]["accounts"].append(acc_uuid)

rows_deps = []
for dep_uuid, entry in all_deps.items():
    dep = entry["dep"]
    accs = entry["accounts"]
    out = {
        "uuid":                     dep_uuid,
        "first_name":               dep.get("first_name") or None,
        "last_name":                dep.get("last_name") or None,
        "known_as":                 dep.get("known_as") or None,
        "relationship":             dep.get("relationship") or None,
        "date_of_birth":            dep.get("date_of_birth") or None,
        "is_financially_dependant": dep.get("is_financially_dependant"),
        "notes":                    dep.get("notes") or None,
        "account_uuids":            pipe_join(accs),
        "external_reference_company": EXT_REF_CO,
        "external_reference_value": dep_uuid,
    }
    rows_deps.append(out)

hdrs_deps = list(deps_importer_keys)
if "external_reference_company" not in hdrs_deps:
    hdrs_deps += ["external_reference_company", "external_reference_value"]

hdrs_deps, rows_deps = drop_empty_cols(hdrs_deps, rows_deps)
ws = wb.create_sheet("Dependants")
write_sheet(ws, hdrs_deps, rows_deps)
p(f"  Rows: {len(rows_deps)}  Cols: {len(hdrs_deps)}")

# ─── Sheet 9: Valuations ─────────────────────────────────────────────────────

p("\n[Sheet 9] Valuations...")
vals_importer_keys = imp_keys["valuations"]

rows_vals = []
for plan_uuid, pl in snap_plans.items():
    vl_list = pl.get("_valuations_latest") or []
    sa_list  = pl.get("_sub_accounts") or []
    sa_polno = {sa["uuid"]: (sa.get("policy_number") or "") for sa in sa_list}

    for vl in vl_list:
        sa_uuid = vl.get("_sub_account_uuid", "")
        val_raw = vl.get("value", {})
        if isinstance(val_raw, dict):
            pence_str = val_raw.get("amount", "0")
        else:
            pence_str = str(val_raw or 0)
        pounds = pence_to_pounds(pence_str)

        valued_at_raw = vl.get("valued_at", "")
        # Normalise to date string
        try:
            valued_at = datetime.fromisoformat(
                valued_at_raw.replace("Z", "+00:00")
            ).strftime("%Y-%m-%d")
        except Exception:
            valued_at = valued_at_raw

        rows_vals.append({
            "plan_uuid":       plan_uuid,
            "sub_account_uuid":sa_uuid or None,
            "policy_number":   sa_polno.get(sa_uuid) or None,
            "valuation_date":  valued_at or None,
            "value":           pounds,
            "ORIGINAL_value":  pence_str,
        })

hdrs_vals = list(vals_importer_keys)
if "ORIGINAL_value" not in hdrs_vals:
    # Insert after value
    if "value" in hdrs_vals:
        hdrs_vals.insert(hdrs_vals.index("value") + 1, "ORIGINAL_value")
    else:
        hdrs_vals.append("ORIGINAL_value")

hdrs_vals, rows_vals = drop_empty_cols(hdrs_vals, rows_vals)
ws = wb.create_sheet("Valuations")
write_sheet(ws, hdrs_vals, rows_vals)
p(f"  Rows: {len(rows_vals)}  Cols: {len(hdrs_vals)}")

# ─── Sheet 10: Income_and_Employment ────────────────────────────────────────

p("\n[Sheet 10] Income_and_Employment...")
iae_importer_keys = imp_keys["income_and_employment"]

rows_iae = []

for acc_uuid, detail in acc_details.items():
    # Income rows
    for inc in detail.get("income", []):
        raw_amt = None
        amt_obj = inc.get("amount")
        if isinstance(amt_obj, dict):
            raw_amt = amt_obj.get("amount", "0")
        elif amt_obj is not None:
            raw_amt = str(amt_obj)
        pounds_inc = pence_to_pounds(raw_amt) if raw_amt is not None else None

        out = {
            "client_uuid":              acc_uuid,
            "name":                     clean_val(inc.get("name")),
            "type":                     clean_val(inc.get("type")),
            "type_other":               clean_val(inc.get("type_other")),
            "start_date":               clean_val(inc.get("start_date")),
            "end_date":                 clean_val(inc.get("end_date")),
            "amount":                   pounds_inc,
            "amount_type":              clean_val(inc.get("amount_type")),
            "frequency":                clean_val(inc.get("frequency")),
            "is_taxable":               inc.get("is_taxable"),
            "external_reference_company": EXT_REF_CO,
            "external_reference_value": inc.get("uuid") or acc_uuid,
            "ORIGINAL_amount":          raw_amt,
        }
        rows_iae.append(out)

    # Employment rows
    for emp in detail.get("employment", []):
        emp_uuid = emp.get("uuid", "")
        out = {
            "client_uuid":                    acc_uuid,
            "employment_uuid":                emp_uuid or None,
            "employment_employer":            clean_val(emp.get("employer")),
            "employment_job_title":           clean_val(emp.get("job_title")),
            "employment_occupation_code":     clean_val(emp.get("occupation_code")),
            "employment_status":              clean_val(emp.get("status")),
            "employment_started_at":          clean_val(emp.get("started_at")),
            "employment_ended_at":            clean_val(emp.get("ended_at")),
            "employment_is_primary":          emp.get("primary"),
            "employment_is_business_owner":   emp.get("is_business_owner"),
            "employment_percentage_of_business_owned":
                clean_val(emp.get("percentage_of_business_owned")),
            "employment_business_year_end":   clean_val(emp.get("business_year_end")),
            "employment_profit":              clean_val(emp.get("profit")),
            "employment_business_value":      clean_val(emp.get("business_value")),
            "employment_accountant_company_name":
                clean_val(emp.get("accountant_company_name")),
            "employment_external_reference_company": EXT_REF_CO,
            "employment_external_reference_value":   emp_uuid or None,
            "external_reference_company": EXT_REF_CO,
            "external_reference_value":  emp_uuid or acc_uuid,
        }
        rows_iae.append(out)

# Build headers from importer keys
hdrs_iae = list(iae_importer_keys)
# Add employment_external_reference keys if in spec
for k in ["employment_external_reference_company",
          "employment_external_reference_value"]:
    if k not in hdrs_iae:
        hdrs_iae.append(k)
if "external_reference_company" not in hdrs_iae:
    hdrs_iae += ["external_reference_company", "external_reference_value"]
if "ORIGINAL_amount" not in hdrs_iae:
    if "amount" in hdrs_iae:
        hdrs_iae.insert(hdrs_iae.index("amount") + 1, "ORIGINAL_amount")
    else:
        hdrs_iae.append("ORIGINAL_amount")

hdrs_iae, rows_iae = drop_empty_cols(hdrs_iae, rows_iae)
ws = wb.create_sheet("Income_and_Employment")
write_sheet(ws, hdrs_iae, rows_iae)
income_count = sum(1 for d in acc_details.values() for _ in d.get("income", []))
employ_count = sum(1 for d in acc_details.values() for _ in d.get("employment", []))
p(f"  Income rows: {income_count}  Employment rows: {employ_count}  Total: {len(rows_iae)}")

# ─── Sheet 11: Provider_plans_not_in_Plannr ─────────────────────────────────

p("\n[Sheet 11] Provider_plans_not_in_Plannr...")

# Build set of plan policy_numbers from the 1,871 export
plan_policy_numbers = set()
for row in plan_export_rows:
    pn = (row.get("policy_number") or "").strip()
    if pn and not is_placeholder(pn):
        plan_policy_numbers.add(pn.upper())

rows_provider = []

# Rathbones file
_wb_rath = _ox.load_workbook(RATH_PATH, read_only=True, data_only=True)
_ws_rath = _wb_rath.active
rath_hdrs = [_ws_rath.cell(1, c).value for c in range(1, (_ws_rath.max_column or 12) + 1)]
for row in _ws_rath.iter_rows(min_row=2, values_only=True):
    fund_no = str(row[0] or "").strip()
    if not fund_no:
        continue
    matched = fund_no.upper() in plan_policy_numbers
    rows_provider.append({
        "source":         "Rathbones",
        "ref_1":          fund_no,
        "ref_2":          "",
        "extra_1":        str(row[3] or ""),   # Fund Name
        "extra_2":        str(row[6] or ""),   # Total Value
        "matched_to_plan": "Y" if matched else "N",
    })
_wb_rath.close()

# Transact/Nucleus workbook
_wb_tn = _ox.load_workbook(WB_PATH, read_only=True, data_only=True)
_ws_tn = _wb_tn["Payment Rates Summary"]
for row in _ws_tn.iter_rows(min_row=2, values_only=True):
    port_no = str(row[3] or "").strip()   # Portfolio Number col D
    wrap_ref = str(row[4] or "").strip()  # Wrapper Ref col E
    if not port_no and not wrap_ref:
        continue
    matched = (port_no.upper() in plan_policy_numbers or
               wrap_ref.upper() in plan_policy_numbers)
    rows_provider.append({
        "source":          "Transact/Nucleus",
        "ref_1":           port_no,
        "ref_2":           wrap_ref,
        "extra_1":         str(row[6] or ""),   # Wrapper Name
        "extra_2":         str(row[8] or ""),   # Annual Payment - Cash
        "matched_to_plan": "Y" if matched else "N",
    })
_wb_tn.close()

hdrs_provider = ["source", "ref_1", "ref_2", "extra_1", "extra_2", "matched_to_plan"]
ws = wb.create_sheet("Provider_plans_not_in_Plannr")
write_sheet(ws, hdrs_provider, rows_provider)
not_matched = sum(1 for r in rows_provider if r.get("matched_to_plan") == "N")
matched_ct  = sum(1 for r in rows_provider if r.get("matched_to_plan") == "Y")
p(f"  Total provider rows: {len(rows_provider)}  Matched: {matched_ct}  Not matched: {not_matched}")

# ─── Save workbook ────────────────────────────────────────────────────────────

p(f"\nSaving {OUT_PATH}...")
wb.save(OUT_PATH)
p("  Saved.")

# ─── Final summary ────────────────────────────────────────────────────────────

p("\n" + "=" * 68)
p("STEP 3 SUMMARY")
p("=" * 68)
sheet_counts = {
    "Clients":                  len(rows_clients),
    "Plans":                    len(rows_plans),
    "Charges":                  len(rows_charges),
    "Circles":                  len(rows_circles),
    "Expenditure":              len(rows_exp),
    "Plannr_Premiums":          len(rows_prem),
    "Relationships":            len(rows_rels),
    "Dependants":               len(rows_deps),
    "Valuations":               len(rows_vals),
    "Income_and_Employment":    len(rows_iae),
    "Provider_plans (ref)":     len(rows_provider),
}
for sheet, n in sheet_counts.items():
    p(f"  {sheet:<35} {n:>6} rows")

p(f"\n  Joint plans with ownership_split blank: {joint_blank_count}")
p(f"  Plans with placeholder policy_number blanked: {placeholder_plan_count}")
p(f"  Expenditure rows without mapping match: {unmatched_exp}")
p(f"  Provider rows not matched to any plan: {not_matched}")
p(f"\n  Output: {OUT_PATH}")

# Write summary dict for step4
_summary = {
    "sheet_counts": sheet_counts,
    "joint_blank":  joint_blank_count,
    "placeholder_plans": placeholder_plan_count,
    "unmatched_exp": unmatched_exp,
    "provider_not_matched": not_matched,
    "provider_matched": matched_ct,
    "note_task": note_task_count,
    "note_case": note_case_count,
    "note_board_task": note_board_task,
    "note_total_exportable": note_total_exportable,
    "note_export_by_model": export_by_model,
    "note_snap_by_type": snap_by_type,
    "note_in_both": len(in_both),
    "note_export_only": len(export_only),
    "note_snap_only": len(snap_only),
    "note_eo_by_model": eo_by_model,
}
summary_path = os.path.join(DATA_ROOT, "step3_summary.json")
with open(summary_path, "w", encoding="utf-8") as f:
    json.dump(_summary, f, indent=2)
p(f"\nStep 3 summary JSON -> {summary_path}")
p("\nReady for Step 4.")
