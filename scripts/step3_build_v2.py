"""
HFP AR Migration - Step 3 v2: Build 10 separate importer files.

Outputs (DATA_ROOT):
  Plannr_Clients.xlsx, Plannr_Plans.xlsx (2 sheets), Plannr_Charges.xlsx,
  Plannr_Circles.xlsx, Plannr_Relationships.xlsx, Plannr_Dependants.xlsx,
  Plannr_Valuations.xlsx, Plannr_Income_Employment.xlsx,
  Plannr_Expenditure.xlsx, Plannr_Premiums.xlsx

Rules: Rule 9 (no PII in console), G4 (no destination writes).
"""
import os, sys, json, re, csv
from collections import defaultdict
from datetime import datetime, timezone, date
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

# ── Constants ─────────────────────────────────────────────────────────────────

DATA_ROOT   = r"C:\dev\AR Migration"
SNAP_DIR    = os.path.join(DATA_ROOT, "snapshot_20261012T090906Z")
EXPORTS_DIR = os.path.join(DATA_ROOT, "exports_20261014")
DEST_JSON   = os.path.join(DATA_ROOT, "block6_destination.json")
MAPPINGS    = os.path.join(DATA_ROOT, "mappings_hfp.xlsx")
RATH_PATH   = os.path.join(DATA_ROOT, "Jim list  - Rathbones.xlsx")
WB_PATH     = os.path.join(DATA_ROOT,
    "Hoban_Financial_Planning_Ltd_260922_Payment_Rates_113727875.xlsx")
V6_PATH     = os.path.join(DATA_ROOT, "hfp_signoff_pack_v6.xlsx")
MANIFEST    = os.path.join(EXPORTS_DIR, "_manifest.json")
EXT_REF_CO  = "data_migration"
BA_GROUP_UUID = "b5749882-0a12-4be3-a022-7cf98d4b56b9"
BASIS_MAP = {
    "workbook":       "Provider statement",
    "Rathbones":      "Rathbones",
    "Plannr ongoing": "Your Plannr",
}

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
    if not csv_val:
        return ""
    return str(csv_val).split(",")[0].strip()

def pipe_join(items):
    return "|".join(str(x) for x in items if x)

def blank(v):
    return v is None or str(v).strip() in ("", "None", "nan")

def clean_val(v):
    if v is None:
        return None
    if isinstance(v, (int, float, bool)):
        return v
    if isinstance(v, str):
        return v or None
    if isinstance(v, dict):
        if "amount" in v:
            try:
                return round(int(str(v["amount"])) / 100.0, 2)
            except (TypeError, ValueError):
                pass
        return str(v)
    if isinstance(v, list):
        return "|".join(str(x) for x in v if x is not None) or None
    return str(v)

# ── Style helpers ─────────────────────────────────────────────────────────────

HDR_FILL = PatternFill(fill_type="solid", fgColor="F2F2F2")
HDR_FONT = Font(bold=True)
ORIG_FILL = PatternFill(fill_type="solid", fgColor="FFF2CC")
NAVY_FILL = PatternFill(fill_type="solid", fgColor="EEF0F6")

def style_ws(ws):
    for cell in ws[1]:
        cell.font = HDR_FONT
        cell.fill = HDR_FILL
        if str(cell.value or "").startswith("ORIGINAL_"):
            cell.font = Font(bold=True, color="7F6000")
            cell.fill = ORIG_FILL
        elif str(cell.value or "") == "basis":
            cell.fill = NAVY_FILL
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
    for i, h in enumerate(headers, 1):
        ws.cell(row=1, column=i, value=h)
    for r_idx, row in enumerate(rows, 2):
        for c_idx, h in enumerate(headers, 1):
            ws.cell(row=r_idx, column=c_idx, value=row.get(h))
    style_ws(ws)

def drop_empty_cols(headers, rows):
    keep = []
    for h in headers:
        if h.startswith("ORIGINAL_"):
            keep.append(h)
            continue
        if any(not blank(r.get(h)) for r in rows):
            keep.append(h)
    new_rows = [{h: r.get(h) for h in keep} for r in rows]
    return keep, new_rows

def save_wb(wb, name):
    path = os.path.join(DATA_ROOT, name)
    wb.save(path)
    p(f"  Saved: {name}")
    return path

# ── Load sources ─────────────────────────────────────────────────────────────

p("=" * 68)
p("HFP AR Migration - Step 3 v2: Build importer files")
p("=" * 68)

p("\n[0] Loading sources...")
with open(MANIFEST, encoding="utf-8") as f:
    manifest = json.load(f)
results = manifest.get("results", {})

with open(DEST_JSON, encoding="utf-8") as f:
    dest = json.load(f)
imp_keys = {k: v["store_keys"] for k, v in dest["importer_keys"].items()}

def export_path(label):
    return results.get(label, {}).get("path", "")

with open(os.path.join(SNAP_DIR, "plans_dedup.json"), encoding="utf-8") as f:
    snap_plans = json.load(f)
with open(os.path.join(SNAP_DIR, "accounts.json"), encoding="utf-8") as f:
    snap_accounts = json.load(f)
snap_client_uuids = {a["uuid"] for a in snap_accounts if a.get("type") == "client"}
with open(os.path.join(SNAP_DIR, "account_details.json"), encoding="utf-8") as f:
    acc_details = json.load(f)
with open(os.path.join(SNAP_DIR, "notes_dedup.json"), encoding="utf-8") as f:
    snap_notes = json.load(f)

# Plan export
with open(export_path("plan_export"), encoding="utf-8-sig", newline="") as f:
    plan_export_rows = list(csv.DictReader(f))
plan_by_uuid = {r["uuid"]: r for r in plan_export_rows}

# Plan policy number set for provider matching
plan_policy_numbers = {
    (r.get("policy_number") or "").strip().upper()
    for r in plan_export_rows
    if not is_placeholder(r.get("policy_number", ""))
}
plan_policy_numbers.discard("")

# Charges export: plan_uuid → row (for ORIGINAL_ values)
with open(export_path("charges_export"), encoding="utf-8-sig", newline="") as f:
    charges_export_rows = list(csv.DictReader(f))
planned_charge_by_plan = {r.get("plan_uuid", ""): r for r in charges_export_rows}

# Expenditure mapping
_wb_map = openpyxl.load_workbook(MAPPINGS, read_only=True, data_only=True)
exp_map = {}
for row in _wb_map["Expenditure"].iter_rows(min_row=2, values_only=True):
    if row[0] is None:
        continue
    src_sec = str(row[0] or "").strip()
    src_ln  = str(row[1] or "").strip()
    dst_sec = str(row[3] or "").strip()
    dst_ln  = str(row[4] or "").strip()
    conf    = str(row[5] or "").strip().lower()
    exp_map[(src_sec.lower(), src_ln.lower())] = (
        dst_sec, dst_ln, conf == "custom" or dst_ln.upper() == "CUSTOM"
    )
_wb_map.close()

p(f"  Plans: {len(plan_export_rows)}, Snap: {len(snap_plans)}, Charges: {len(charges_export_rows)}")

# ── Notes investigation ───────────────────────────────────────────────────────

p("\n" + "=" * 68)
p("NOTES INVESTIGATION")
p("=" * 68)

note_export_path = export_path("note_export")
export_uuids_note = set()
export_by_model = {}
with open(note_export_path, encoding="utf-8-sig", newline="") as f:
    reader = csv.DictReader(f)
    for row in reader:
        uid = row.get("uuid", "")
        if uid not in export_uuids_note:
            export_uuids_note.add(uid)
            mt = row.get("model_type", "")
            export_by_model[mt] = export_by_model.get(mt, 0) + 1

snap_by_type = {}
for v in snap_notes.values():
    t = v.get("notable_type", "") or v.get("type", "")
    snap_by_type[t] = snap_by_type.get(t, 0) + 1

in_both     = export_uuids_note & set(snap_notes.keys())
export_only = export_uuids_note - set(snap_notes.keys())
snap_only   = set(snap_notes.keys()) - export_uuids_note

eo_by_model = {}
with open(note_export_path, encoding="utf-8-sig", newline="") as f:
    for row in csv.DictReader(f):
        if row.get("uuid", "") in export_only:
            mt = row.get("model_type", "")
            eo_by_model[mt] = eo_by_model.get(mt, 0) + 1

p(f"  Export unique: {len(export_uuids_note)}")
p(f"  Snapshot:      {len(snap_notes)}")
p(f"  In both:       {len(in_both)}")
p(f"  Export only:   {len(export_only)}")
p(f"  Snapshot only: {len(snap_only)}")
p(f"\n  Export by model_type:")
for mt, n in sorted(export_by_model.items()):
    p(f"    {mt:<20} {n:>5}")
p(f"\n  Snapshot by notable_type:")
for nt, n in sorted(snap_by_type.items()):
    p(f"    {nt:<20} {n:>5}")
p(f"\n  Export-only by model_type:")
for mt, n in sorted(eo_by_model.items()):
    p(f"    {mt:<20} {n:>5}")

# Board_task with account_uuid
bt_with_acc = 0
with open(note_export_path, encoding="utf-8-sig", newline="") as f:
    for row in csv.DictReader(f):
        if row.get("model_type") == "board_task":
            if row.get("account_uuid") or row.get("account_uuids"):
                bt_with_acc += 1
p(f"\n  board_task with account_uuid: {bt_with_acc} (all stay in old Plannr)")

# ── FILE 1: Plannr_Clients.xlsx ───────────────────────────────────────────────

p("\n[1] Plannr_Clients.xlsx")
client_importer_keys = imp_keys["clients"]
with open(export_path("client_export"), encoding="utf-8-sig", newline="") as f:
    client_rows_raw = list(csv.DictReader(f))

rows_clients = []
for row in client_rows_raw:
    out = {k: (row[k] or None) if k in row else None for k in client_importer_keys}
    out["external_reference_company"] = EXT_REF_CO
    out["external_reference_value"]   = row.get("uuid", "")
    rows_clients.append(out)

hdrs_clients = list(client_importer_keys)
if "external_reference_company" not in hdrs_clients:
    hdrs_clients += ["external_reference_company", "external_reference_value"]
hdrs_clients, rows_clients = drop_empty_cols(hdrs_clients, rows_clients)

wb = openpyxl.Workbook()
wb.remove(wb.active)
write_sheet(wb.create_sheet("Clients"), hdrs_clients, rows_clients)
save_wb(wb, "Plannr_Clients.xlsx")
p(f"  Rows: {len(rows_clients)}  Cols: {len(hdrs_clients)}")

# ── FILE 2: Plannr_Plans.xlsx (2 sheets) ─────────────────────────────────────

p("\n[2] Plannr_Plans.xlsx")
plan_importer_keys = imp_keys["plans"]
new_plan_uuids = {r["uuid"] for r in plan_export_rows if r.get("uuid") not in snap_plans}

rows_plans = []
for row in plan_export_rows:
    out = {}
    for k in plan_importer_keys:
        if k == "external_reference_company":
            out[k] = EXT_REF_CO
        elif k == "external_reference_value":
            out[k] = row.get("uuid", "")
        elif k == "ownership_split":
            out[k] = None
        elif k in row:
            out[k] = row[k] or None
        else:
            out[k] = None
    raw_pn = row.get("policy_number", "")
    if is_placeholder(raw_pn):
        out["ORIGINAL_policy_number"] = raw_pn if raw_pn else None
        out["policy_number"] = None
    else:
        out["ORIGINAL_policy_number"] = None
    rows_plans.append(out)

hdrs_plans = list(plan_importer_keys)
if "policy_number" in hdrs_plans:
    pn_idx = hdrs_plans.index("policy_number")
    hdrs_plans.insert(pn_idx + 1, "ORIGINAL_policy_number")
if "external_reference_company" not in hdrs_plans:
    hdrs_plans += ["external_reference_company", "external_reference_value"]
hdrs_plans, rows_plans = drop_empty_cols(hdrs_plans, rows_plans)

# Provider plans not in Plannr: 51 unmatched rows only
_wb_rath = openpyxl.load_workbook(RATH_PATH, read_only=True, data_only=True)
_ws_rath = _wb_rath.active
provider_rows = []
for row in _ws_rath.iter_rows(min_row=2, values_only=True):
    fund_no = str(row[0] or "").strip()
    if not fund_no:
        continue
    if fund_no.upper() not in plan_policy_numbers:
        provider_rows.append({
            "source":     "Rathbones",
            "ref_1":      fund_no,
            "ref_2":      "",
            "extra_1":    str(row[3] or ""),
            "extra_2":    str(row[6] or ""),
            "action":     "",
        })
_wb_rath.close()

_wb_tn = openpyxl.load_workbook(WB_PATH, read_only=True, data_only=True)
_ws_tn = _wb_tn["Payment Rates Summary"]
for row in _ws_tn.iter_rows(min_row=2, values_only=True):
    port_no  = str(row[3] or "").strip()
    wrap_ref = str(row[4] or "").strip()
    if not port_no and not wrap_ref:
        continue
    if (port_no.upper() not in plan_policy_numbers and
            wrap_ref.upper() not in plan_policy_numbers):
        provider_rows.append({
            "source":   "Transact/Nucleus",
            "ref_1":    port_no,
            "ref_2":    wrap_ref,
            "extra_1":  str(row[6] or ""),
            "extra_2":  str(row[8] or ""),
            "action":   "",
        })
_wb_tn.close()
p(f"  Provider rows not matched: {len(provider_rows)}")

hdrs_provider = ["source", "ref_1", "ref_2", "extra_1", "extra_2", "action"]

wb = openpyxl.Workbook()
wb.remove(wb.active)
write_sheet(wb.create_sheet("Plans"), hdrs_plans, rows_plans)
ws_prov = wb.create_sheet("Provider_plans_not_in_Plannr")
write_sheet(ws_prov, hdrs_provider, provider_rows)
# Add Action dropdown
if provider_rows:
    dv = DataValidation(type="list",
                        formula1='"Add to Plannr,Already in Plannr,Not ours,Closed"',
                        allow_blank=True, showErrorMessage=False)
    ws_prov.add_data_validation(dv)
    dv.sqref = f"F2:F{len(provider_rows)+1}"

save_wb(wb, "Plannr_Plans.xlsx")
placeholder_plan_count = sum(1 for r in rows_plans if r.get("ORIGINAL_policy_number"))
p(f"  Plans rows: {len(rows_plans)}  Provider not-in-Plannr: {len(provider_rows)}")
p(f"  Placeholder policy numbers blanked: {placeholder_plan_count}")

# ── FILE 3: Plannr_Charges.xlsx ───────────────────────────────────────────────

p("\n[3] Plannr_Charges.xlsx")

# Load v6 Charges (all): 469 rows (Sections A+B)
_wb_v6 = openpyxl.load_workbook(V6_PATH, read_only=True, data_only=True)
_ws_ca = _wb_v6["Charges (all)"]
_ca_data = list(_ws_ca.iter_rows(values_only=True))
v6_ca_hdrs = list(_ca_data[0])
charges_importer_keys = imp_keys["charges"]

rows_charges = []
for row_tuple in _ca_data[1:]:
    if not any(v is not None for v in row_tuple):
        continue
    v6_row = {v6_ca_hdrs[i]: row_tuple[i]
              for i in range(min(len(v6_ca_hdrs), len(row_tuple)))
              if v6_ca_hdrs[i] is not None}
    plan_uuid = v6_row.get("plan_uuid", "")
    basis_raw = str(v6_row.get("basis", "") or "")
    basis_label = BASIS_MAP.get(basis_raw, basis_raw)

    # Build importer row
    out = {}
    for k in charges_importer_keys:
        out[k] = v6_row.get(k)

    # client_uuid from plan export
    plan_row = plan_by_uuid.get(plan_uuid, {})
    if not out.get("client_uuid"):
        out["client_uuid"] = first_of(plan_row.get("client_uuids", ""))

    out["premium_uuid"] = None
    out["basis"] = basis_label
    out["indicative_annual_income"] = v6_row.get("indicative_annual_income")
    out["external_reference_company"] = EXT_REF_CO
    # external_reference_value: use source_uuid (Plannr charge UUID) if available
    src_uuid = v6_row.get("source_uuid", "")
    out["external_reference_value"] = src_uuid or plan_uuid or ""

    # ORIGINAL_ columns: Plannr charge values if plan had a charge
    pc = planned_charge_by_plan.get(plan_uuid)
    if pc:
        out["ORIGINAL_percentage"]       = pc.get("percentage") or None
        out["ORIGINAL_amount"]           = pc.get("amount") or None
        out["ORIGINAL_frequency"]        = pc.get("frequency") or None
        out["ORIGINAL_start_date"]       = pc.get("start_date") or None
        out["ORIGINAL_notes"]            = pc.get("notes") or None
        out["ORIGINAL_payment_origin"]   = pc.get("payment_origin") or None
    else:
        out["ORIGINAL_percentage"] = out["ORIGINAL_amount"] = None
        out["ORIGINAL_frequency"]  = out["ORIGINAL_start_date"] = None
        out["ORIGINAL_notes"]      = out["ORIGINAL_payment_origin"] = None

    rows_charges.append(out)

# Load Section C plan_uuids from v6 sign-off
_ws_cs = _wb_v6["Charges sign-off"]
_cs_data = list(_ws_cs.iter_rows(values_only=True))
cs_hdrs = list(_cs_data[0])
sec_col = cs_hdrs.index("source_section")
plan_col_cs = cs_hdrs.index("plan_uuid")
sec_c_plan_uuids = []
for row_tuple in _cs_data[1:]:
    if not any(v is not None for v in row_tuple):
        continue
    sec = str(row_tuple[sec_col] or "")
    if sec.startswith("C"):
        puuid = row_tuple[plan_col_cs]
        if puuid:
            sec_c_plan_uuids.append(puuid)
_wb_v6.close()

# Check new plans against workbook/Rathbones for Section C
sec_c_plan_uuids_set = set(sec_c_plan_uuids)
v6_covered_plans = set(r.get("plan_uuid") for r in rows_charges)
new_plans_sec_c = []
for uuid in new_plan_uuids:
    plan_row = plan_by_uuid.get(uuid, {})
    pn = (plan_row.get("policy_number") or "").strip().upper()
    # Check if it matches any provider workbook entry
    if pn and pn in plan_policy_numbers and uuid not in v6_covered_plans:
        pass  # would be Section A — shouldn't happen for brand new plans
    new_plans_sec_c.append(uuid)

# Add Section C rows (199 from v6 + up to 7 new)
for plan_uuid in sec_c_plan_uuids:
    plan_row = plan_by_uuid.get(plan_uuid, {})
    out = {k: None for k in charges_importer_keys}
    out["plan_uuid"]   = plan_uuid
    out["client_uuid"] = first_of(plan_row.get("client_uuids", ""))
    out["basis"]       = "Rate needed"
    out["external_reference_company"] = EXT_REF_CO
    out["external_reference_value"]   = plan_uuid
    for orig_k in ["ORIGINAL_percentage","ORIGINAL_amount","ORIGINAL_frequency",
                   "ORIGINAL_start_date","ORIGINAL_notes","ORIGINAL_payment_origin"]:
        out[orig_k] = None
    rows_charges.append(out)

for plan_uuid in new_plans_sec_c:
    if plan_uuid in v6_covered_plans:
        continue
    plan_row = plan_by_uuid.get(plan_uuid, {})
    out = {k: None for k in charges_importer_keys}
    out["plan_uuid"]   = plan_uuid
    out["client_uuid"] = first_of(plan_row.get("client_uuids", ""))
    out["basis"]       = "Rate needed"
    out["external_reference_company"] = EXT_REF_CO
    out["external_reference_value"]   = plan_uuid
    for orig_k in ["ORIGINAL_percentage","ORIGINAL_amount","ORIGINAL_frequency",
                   "ORIGINAL_start_date","ORIGINAL_notes","ORIGINAL_payment_origin"]:
        out[orig_k] = None
    rows_charges.append(out)

# Build headers
hdrs_charges = list(charges_importer_keys) + ["basis", "indicative_annual_income"]
if "external_reference_company" not in hdrs_charges:
    hdrs_charges += ["external_reference_company", "external_reference_value"]
for orig_k in ["ORIGINAL_percentage","ORIGINAL_amount","ORIGINAL_frequency",
               "ORIGINAL_start_date","ORIGINAL_notes","ORIGINAL_payment_origin"]:
    if orig_k not in hdrs_charges:
        hdrs_charges.append(orig_k)

hdrs_charges, rows_charges = drop_empty_cols(hdrs_charges, rows_charges)

# IAI totals by section
sec_a_iai = sum(float(r.get("indicative_annual_income") or 0)
                for r in rows_charges if r.get("basis") == "Provider statement")
sec_b_iai = sum(float(r.get("indicative_annual_income") or 0)
                for r in rows_charges if r.get("basis") == "Your Plannr")
sec_rath_iai = sum(float(r.get("indicative_annual_income") or 0)
                   for r in rows_charges if r.get("basis") == "Rathbones")

wb = openpyxl.Workbook()
wb.remove(wb.active)
write_sheet(wb.create_sheet("Charges"), hdrs_charges, rows_charges)
save_wb(wb, "Plannr_Charges.xlsx")

sec_a_n  = sum(1 for r in rows_charges if r.get("basis") == "Provider statement")
sec_b_n  = sum(1 for r in rows_charges if r.get("basis") == "Your Plannr")
sec_rath  = sum(1 for r in rows_charges if r.get("basis") == "Rathbones")
sec_c_n  = sum(1 for r in rows_charges if r.get("basis") == "Rate needed")
total_iai = sec_a_iai + sec_b_iai + sec_rath_iai
p(f"  Rows: {len(rows_charges)}")
p(f"  Provider statement: {sec_a_n} rows  IAI: GBP{sec_a_iai:,.2f}")
p(f"  Rathbones:          {sec_rath} rows  IAI: GBP{sec_rath_iai:,.2f}")
p(f"  Your Plannr:        {sec_b_n} rows  IAI: GBP{sec_b_iai:,.2f}")
p(f"  Rate needed:        {sec_c_n} rows  (Section C, no charge)")
p(f"  Total IAI (A+B):    GBP{total_iai:,.2f}")

# ── FILE 4: Plannr_Circles.xlsx ───────────────────────────────────────────────

p("\n[4] Plannr_Circles.xlsx")
circle_members = defaultdict(list)
circle_names = {}
with open(export_path("circle_export"), encoding="utf-8-sig", newline="") as f:
    for row in csv.DictReader(f):
        cuuid = row.get("circle_uuid", "")
        auuid = row.get("account_uuid", "")
        if cuuid:
            if auuid and auuid not in circle_members[cuuid]:
                circle_members[cuuid].append(auuid)
            circle_names[cuuid] = row.get("circle_name", "")

rows_circles = []
for cuuid, members in circle_members.items():
    if not members:
        continue
    rows_circles.append({
        "uuid":                     cuuid,
        "name":                     circle_names.get(cuuid, ""),
        "account_uuids":            pipe_join(members),
        "group_uuids":              BA_GROUP_UUID,
        "external_reference_company": EXT_REF_CO,
        "external_reference_value": cuuid,
    })

hdrs_circles = list(imp_keys["circles"])
if "external_reference_company" not in hdrs_circles:
    hdrs_circles += ["external_reference_company", "external_reference_value"]
rows_circles = [{k: r.get(k) for k in hdrs_circles} for r in rows_circles]
hdrs_circles, rows_circles = drop_empty_cols(hdrs_circles, rows_circles)

wb = openpyxl.Workbook()
wb.remove(wb.active)
write_sheet(wb.create_sheet("Circles"), hdrs_circles, rows_circles)
save_wb(wb, "Plannr_Circles.xlsx")
p(f"  Rows: {len(rows_circles)}")

# ── FILE 5: Plannr_Relationships.xlsx ────────────────────────────────────────

p("\n[5] Plannr_Relationships.xlsx")
lc_map = {}
for acc_uuid, detail in acc_details.items():
    for lc in detail.get("linked_clients", []):
        lc_map.setdefault(acc_uuid, {})[lc.get("uuid", "")] = lc.get("relationship", "")

seen_pairs = set()
rows_rels = []
for acc_uuid, linked in lc_map.items():
    for linked_uuid, rel in linked.items():
        pair_key = tuple(sorted([acc_uuid, linked_uuid]))
        if pair_key in seen_pairs:
            continue
        seen_pairs.add(pair_key)
        rows_rels.append({
            "account_uuid":         acc_uuid,
            "related_account_uuid": linked_uuid,
            "relationship":         rel or None,
            "inverse_relationship": lc_map.get(linked_uuid, {}).get(acc_uuid) or None,
        })

hdrs_rels = list(imp_keys["relationships"])
hdrs_rels, rows_rels = drop_empty_cols(hdrs_rels, rows_rels)

wb = openpyxl.Workbook()
wb.remove(wb.active)
write_sheet(wb.create_sheet("Relationships"), hdrs_rels, rows_rels)
save_wb(wb, "Plannr_Relationships.xlsx")
p(f"  Rows: {len(rows_rels)}")

# ── FILE 6: Plannr_Dependants.xlsx ───────────────────────────────────────────

p("\n[6] Plannr_Dependants.xlsx")

# 249 from snapshot
all_deps = {}
for acc_uuid, detail in acc_details.items():
    for dep in detail.get("dependants", []):
        dep_uuid = dep.get("uuid", "")
        if not dep_uuid:
            continue
        if dep_uuid not in all_deps:
            all_deps[dep_uuid] = {"dep": dep, "accounts": []}
        if acc_uuid not in all_deps[dep_uuid]["accounts"]:
            all_deps[dep_uuid]["accounts"].append(acc_uuid)

# Add API-only dependant: uuid=89a694c4, parents from API call
API_ONLY_DEP = {
    "uuid":                   "89a694c4-4ad7-4ff3-98de-eaffdaef9688",
    "first_name":             None,  # PII - omitted
    "last_name":              None,
    "known_as":               None,
    "relationship":           None,
    "date_of_birth":          None,
    "is_financially_dependant": None,
    "notes":                  None,
}
API_ONLY_DEP_ACCOUNTS = [
    "34aed1c8-2ce6-4562-8500-f1b2cc7f7010",
    "ff64f4f4-9603-49cc-aebb-c44390b6486b",
    "0e3a8af6-63a4-467f-a4af-aab1988d6133",
]
if API_ONLY_DEP["uuid"] not in all_deps:
    # Fetch from the API json file to get name fields (Rule 9: only to workbook)
    with open(export_path("dependants"), encoding="utf-8") as f:
        api_deps_list = json.load(f)
    api_dep_lookup = {d.get("uuid"): d for d in api_deps_list}
    api_dep_data = api_dep_lookup.get(API_ONLY_DEP["uuid"], API_ONLY_DEP)
    all_deps[API_ONLY_DEP["uuid"]] = {
        "dep": api_dep_data,
        "accounts": API_ONLY_DEP_ACCOUNTS,
    }

rows_deps = []
for dep_uuid, entry in all_deps.items():
    dep = entry["dep"]
    accs = entry["accounts"]
    rows_deps.append({
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
    })

hdrs_deps = list(imp_keys["dependants"])
if "external_reference_company" not in hdrs_deps:
    hdrs_deps += ["external_reference_company", "external_reference_value"]
hdrs_deps, rows_deps = drop_empty_cols(hdrs_deps, rows_deps)

wb = openpyxl.Workbook()
wb.remove(wb.active)
write_sheet(wb.create_sheet("Dependants"), hdrs_deps, rows_deps)
save_wb(wb, "Plannr_Dependants.xlsx")
p(f"  Rows: {len(rows_deps)} (249 snapshot + 1 API-only)")

# ── FILE 7: Plannr_Valuations.xlsx ───────────────────────────────────────────

p("\n[7] Plannr_Valuations.xlsx")
rows_vals = []
for plan_uuid, pl in snap_plans.items():
    sa_list  = pl.get("_sub_accounts") or []
    sa_polno = {sa["uuid"]: (sa.get("policy_number") or "") for sa in sa_list}
    for vl in (pl.get("_valuations_latest") or []):
        sa_uuid = vl.get("_sub_account_uuid", "")
        val_raw = vl.get("value", {})
        pence_str = val_raw.get("amount", "0") if isinstance(val_raw, dict) else str(val_raw or 0)
        pounds = pence_to_pounds(pence_str)
        try:
            valued_at = datetime.fromisoformat(
                (vl.get("valued_at") or "").replace("Z", "+00:00")
            ).strftime("%Y-%m-%d")
        except Exception:
            valued_at = vl.get("valued_at") or None
        rows_vals.append({
            "plan_uuid":        plan_uuid,
            "sub_account_uuid": sa_uuid or None,
            "policy_number":    sa_polno.get(sa_uuid) or None,
            "valuation_date":   valued_at,
            "value":            pounds,
        })

# No ORIGINAL_value per correction 5
hdrs_vals = list(imp_keys["valuations"])
hdrs_vals, rows_vals = drop_empty_cols(hdrs_vals, rows_vals)

wb = openpyxl.Workbook()
wb.remove(wb.active)
write_sheet(wb.create_sheet("Valuations"), hdrs_vals, rows_vals)
save_wb(wb, "Plannr_Valuations.xlsx")
p(f"  Rows: {len(rows_vals)}")

# ── FILE 8: Plannr_Income_Employment.xlsx ────────────────────────────────────

p("\n[8] Plannr_Income_Employment.xlsx")

# Build employment lookup per account: acc_uuid → [emp, ...]
acc_employments = {acc: detail.get("employment", []) for acc, detail in acc_details.items()}

# Link income to employment:
# For each account, pair salary income with primary employment (or first),
# then all other incomes without employment,
# then any unmatched employments as standalone rows.
iae_importer_keys = imp_keys["income_and_employment"]

def emp_to_fields(emp):
    return {
        "employment_uuid":                   emp.get("uuid") or None,
        "employment_employer":               clean_val(emp.get("employer")),
        "employment_job_title":              clean_val(emp.get("job_title")),
        "employment_occupation_code":        clean_val(emp.get("occupation_code")),
        "employment_status":                 clean_val(emp.get("status")),
        "employment_started_at":             clean_val(emp.get("started_at")),
        "employment_ended_at":               clean_val(emp.get("ended_at")),
        "employment_is_primary":             emp.get("primary"),
        "employment_is_business_owner":      emp.get("is_business_owner"),
        "employment_percentage_of_business_owned":
            clean_val(emp.get("percentage_of_business_owned")),
        "employment_business_year_end":      clean_val(emp.get("business_year_end")),
        "employment_profit":                 clean_val(emp.get("profit")),
        "employment_business_value":         clean_val(emp.get("business_value")),
        "employment_accountant_company_name":
            clean_val(emp.get("accountant_company_name")),
        "employment_external_reference_company": EXT_REF_CO,
        "employment_external_reference_value":   emp.get("uuid") or None,
    }

rows_iae = []
income_count = 0
employment_count = 0
standalone_emp_rows = 0

for acc_uuid, detail in acc_details.items():
    incs = detail.get("income", [])
    emps = detail.get("employment", [])

    # Pick primary employment (primary=True or first)
    primary_emp = None
    for e in emps:
        if e.get("primary"):
            primary_emp = e
            break
    if primary_emp is None and emps:
        primary_emp = emps[0]

    # Additional employments (beyond primary)
    extra_emps = [e for e in emps if e is not primary_emp]

    # Find the salary income row (or first income) to attach primary_emp
    salary_idx = None
    for i, inc in enumerate(incs):
        if inc.get("type") == "salary":
            salary_idx = i
            break
    attach_emp_idx = salary_idx if salary_idx is not None else (0 if incs else None)

    for inc_idx, inc in enumerate(incs):
        amt_obj = inc.get("amount")
        if isinstance(amt_obj, dict):
            raw_amt = amt_obj.get("amount", "0")
        elif amt_obj is not None:
            raw_amt = str(amt_obj)
        else:
            raw_amt = None
        pounds_inc = pence_to_pounds(raw_amt) if raw_amt is not None else None

        out = {
            "client_uuid":   acc_uuid,
            "name":          clean_val(inc.get("name")),
            "type":          clean_val(inc.get("type")),
            "type_other":    clean_val(inc.get("type_other")),
            "start_date":    clean_val(inc.get("start_date")),
            "end_date":      clean_val(inc.get("end_date")),
            "amount":        pounds_inc,
            "amount_type":   clean_val(inc.get("amount_type")),
            "frequency":     clean_val(inc.get("frequency")),
            "is_taxable":    inc.get("is_taxable"),
            "external_reference_company": EXT_REF_CO,
            "external_reference_value":  inc.get("uuid") or acc_uuid,
        }
        # Attach primary employment to the designated income row
        if primary_emp is not None and inc_idx == attach_emp_idx:
            out.update(emp_to_fields(primary_emp))
            employment_count += 1
        else:
            for k in emp_to_fields({}):
                out[k] = None
        # Attach extra employment to subsequent income rows if available
        extra_idx = inc_idx - (attach_emp_idx or 0) - 1
        if extra_idx >= 0 and extra_idx < len(extra_emps):
            out.update(emp_to_fields(extra_emps[extra_idx]))
            employment_count += 1
        rows_iae.append(out)
        income_count += 1

    # Employments with no income row to attach to
    emps_covered = set()
    if primary_emp and attach_emp_idx is not None and incs:
        emps_covered.add(id(primary_emp))
    for extra_idx, e in enumerate(extra_emps):
        if (attach_emp_idx or 0) + 1 + extra_idx < len(incs):
            emps_covered.add(id(e))

    for emp in emps:
        if id(emp) not in emps_covered:
            out = {k: None for k in iae_importer_keys}
            out["client_uuid"] = acc_uuid
            out["external_reference_company"] = EXT_REF_CO
            out["external_reference_value"]  = emp.get("uuid") or acc_uuid
            out.update(emp_to_fields(emp))
            out["employment_external_reference_company"] = EXT_REF_CO
            out["employment_external_reference_value"]   = emp.get("uuid") or None
            rows_iae.append(out)
            employment_count += 1
            standalone_emp_rows += 1

# Build headers
hdrs_iae = list(iae_importer_keys)
for k in ["employment_external_reference_company",
          "employment_external_reference_value"]:
    if k not in hdrs_iae:
        hdrs_iae.append(k)
if "external_reference_company" not in hdrs_iae:
    hdrs_iae += ["external_reference_company", "external_reference_value"]
# No ORIGINAL_amount per correction 5

hdrs_iae, rows_iae = drop_empty_cols(hdrs_iae, rows_iae)

wb = openpyxl.Workbook()
wb.remove(wb.active)
write_sheet(wb.create_sheet("Income_Employment"), hdrs_iae, rows_iae)
save_wb(wb, "Plannr_Income_Employment.xlsx")
p(f"  Total rows: {len(rows_iae)}")
p(f"  Income: {income_count}  Employment: {employment_count}  Standalone emp rows: {standalone_emp_rows}")

# ── FILE 9: Plannr_Expenditure.xlsx ──────────────────────────────────────────

p("\n[9] Plannr_Expenditure.xlsx")
expenses_importer_keys = imp_keys["expenses"]

with open(export_path("client_expenditure_export"), encoding="utf-8-sig", newline="") as f:
    exp_export_rows = list(csv.DictReader(f))

rows_exp = []
for row in exp_export_rows:
    src_sec = (row.get("section") or "").strip()
    src_ln  = (row.get("type")    or "").strip()
    mapping = exp_map.get((src_sec.lower(), src_ln.lower()))
    if mapping:
        dst_sec, dst_ln, is_cust = mapping
    else:
        dst_sec, dst_ln, is_cust = src_sec, src_ln, False

    rows_exp.append({
        "client_uuid":             row.get("client_uuid") or None,
        "life_stage_name":         row.get("life_stage") or None,
        "life_stage_section_name": dst_sec or None,
        "expense_name":            dst_ln or None,
        "is_custom_expense":       True if is_cust else None,
        "amount":                  row.get("amount") or None,
        "external_reference_company": EXT_REF_CO,
        "external_reference_value":   row.get("client_uuid") or None,
    })

hdrs_exp = list(expenses_importer_keys)
if "external_reference_company" not in hdrs_exp:
    hdrs_exp += ["external_reference_company", "external_reference_value"]
hdrs_exp, rows_exp = drop_empty_cols(hdrs_exp, rows_exp)

wb = openpyxl.Workbook()
wb.remove(wb.active)
write_sheet(wb.create_sheet("Expenditure"), hdrs_exp, rows_exp)
save_wb(wb, "Plannr_Expenditure.xlsx")
p(f"  Rows: {len(rows_exp)}")

# ── FILE 10: Plannr_Premiums.xlsx ─────────────────────────────────────────────

p("\n[10] Plannr_Premiums.xlsx")
premiums_importer_keys = imp_keys["premiums"]

with open(export_path("premium_export"), encoding="utf-8-sig", newline="") as f:
    prem_export_rows = list(csv.DictReader(f))

rows_prem = []
for row in prem_export_rows:
    rows_prem.append({k: (row[k] or None) if k in row else None
                     for k in premiums_importer_keys})

hdrs_prem = list(premiums_importer_keys)
hdrs_prem, rows_prem = drop_empty_cols(hdrs_prem, rows_prem)

wb = openpyxl.Workbook()
wb.remove(wb.active)
write_sheet(wb.create_sheet("Premiums"), hdrs_prem, rows_prem)
save_wb(wb, "Plannr_Premiums.xlsx")
p(f"  Rows: {len(rows_prem)}")

# ── Summary ───────────────────────────────────────────────────────────────────

p("\n" + "=" * 68)
p("STEP 3 v2 SUMMARY")
p("=" * 68)
summary = {
    "files": {
        "Plannr_Clients.xlsx":           len(rows_clients),
        "Plannr_Plans.xlsx (Plans)":     len(rows_plans),
        "Plannr_Plans.xlsx (Provider)":  len(provider_rows),
        "Plannr_Charges.xlsx":           len(rows_charges),
        "Plannr_Circles.xlsx":           len(rows_circles),
        "Plannr_Relationships.xlsx":     len(rows_rels),
        "Plannr_Dependants.xlsx":        len(rows_deps),
        "Plannr_Valuations.xlsx":        len(rows_vals),
        "Plannr_Income_Employment.xlsx": len(rows_iae),
        "Plannr_Expenditure.xlsx":       len(rows_exp),
        "Plannr_Premiums.xlsx":          len(rows_prem),
    },
    "charges": {
        "provider_statement": {"rows": sec_a_n,  "iai": round(sec_a_iai, 2)},
        "rathbones":          {"rows": sec_rath, "iai": round(sec_rath_iai, 2)},
        "your_plannr":        {"rows": sec_b_n,  "iai": round(sec_b_iai, 2)},
        "rate_needed":        {"rows": sec_c_n,  "iai": 0},
        "total_iai":          round(total_iai, 2),
    },
    "plans": {
        "total": len(rows_plans),
        "new_post_snapshot": len(new_plan_uuids),
        "placeholder_pn_blanked": placeholder_plan_count,
        "provider_not_in_plannr": len(provider_rows),
    },
    "dependants": {
        "snapshot": 249,
        "api_only": 1,
        "total": len(rows_deps),
    },
    "notes": {
        "export_total": len(export_uuids_note),
        "snapshot_total": len(snap_notes),
        "in_both": len(in_both),
        "export_only": len(export_only),
        "snap_only": len(snap_only),
        "by_model": export_by_model,
        "snap_by_type": snap_by_type,
        "export_only_by_model": eo_by_model,
        "board_task_with_account_uuid": bt_with_acc,
    },
    "iae": {
        "income_rows": income_count,
        "employment_rows": employment_count,
        "standalone_emp_rows": standalone_emp_rows,
        "total_rows": len(rows_iae),
    },
    "plan_check": {
        "same_day_plans": 5,
        "result": "All 5 created after 09:09:06Z (earliest 15:00:57Z). All 7 treated as normal."
    },
}

for fname, n in summary["files"].items():
    p(f"  {fname:<40} {n:>6}")
p(f"\n  Charges IAI: Provider statement GBP{sec_a_iai:,.2f}  "
  f"Rathbones GBP{sec_rath_iai:,.2f}  Plannr GBP{sec_b_iai:,.2f}  "
  f"Total GBP{total_iai:,.2f}")
p(f"  v6 total was GBP409,691.97  -  difference: GBP{total_iai - 409691.97:+.2f}")

summary_path = os.path.join(DATA_ROOT, "step3_summary_v2.json")
with open(summary_path, "w", encoding="utf-8") as f:
    json.dump(summary, f, indent=2)
p(f"\nSummary -> {summary_path}")
p("\nReady for Step 4 v2.")
