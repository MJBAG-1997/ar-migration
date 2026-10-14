"""
step7_readable.py — Add INFO columns, rearrange, sort all importer sheets for AR.
Rule 9: names go into Excel only; nothing identifying in print/logs.
"""
import os, json, csv, re
from collections import defaultdict
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

DATA_ROOT   = r"C:\dev\AR Migration"
EXPORTS_DIR = os.path.join(DATA_ROOT, "exports_20261014")
MANIFEST    = os.path.join(EXPORTS_DIR, "_manifest.json")
SNAP_DIR    = os.path.join(DATA_ROOT, "snapshot_20261012T090906Z")

PLANS_XLSX   = os.path.join(DATA_ROOT, "Plannr_Plans.xlsx")
CLIENTS_XLSX = os.path.join(DATA_ROOT, "Plannr_Clients.xlsx")
CIRCLES_XLSX = os.path.join(DATA_ROOT, "Plannr_Circles.xlsx")
RELS_XLSX    = os.path.join(DATA_ROOT, "Plannr_Relationships.xlsx")
DEPS_XLSX    = os.path.join(DATA_ROOT, "Plannr_Dependants.xlsx")
IAE_XLSX     = os.path.join(DATA_ROOT, "Plannr_Income_Employment.xlsx")
EXP_XLSX     = os.path.join(DATA_ROOT, "Plannr_Expenditure.xlsx")
VAL_XLSX     = os.path.join(DATA_ROOT, "Plannr_Valuations.xlsx")
PREM_XLSX    = os.path.join(DATA_ROOT, "Plannr_Premiums.xlsx")
CHRG_XLSX    = os.path.join(DATA_ROOT, "Plannr_Charges.xlsx")

# ── Styles ────────────────────────────────────────────────────────────────────
HDR_FONT  = Font(bold=True)
HDR_FILL  = PatternFill(fill_type="solid", fgColor="F2F2F2")
INFO_FONT = Font(bold=True, color="1F4E79")
INFO_FILL = PatternFill(fill_type="solid", fgColor="DEEAF1")
ORIG_FONT = Font(bold=True, color="7F6000")
ORIG_FILL = PatternFill(fill_type="solid", fgColor="FFF2CC")
REV_FONT  = Font(bold=True, color="7F6000")
REV_FILL  = PatternFill(fill_type="solid", fgColor="FFF2CC")

def style_hdr(c):
    h = str(c.value or "")
    if h.startswith("INFO:"):
        c.font = INFO_FONT; c.fill = INFO_FILL
    elif h == "Review":
        c.font = REV_FONT;  c.fill = REV_FILL
    elif h.startswith("ORIGINAL_"):
        c.font = ORIG_FONT; c.fill = ORIG_FILL
    else:
        c.font = HDR_FONT;  c.fill = HDR_FILL

def blank(v): return v is None or str(v).strip() in ("", "None", "nan")

UUID_RE = re.compile(r'^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$', re.I)
def looks_uuid(v): return bool(UUID_RE.match(str(v or "").strip()))

def to_plain(v):
    """protection_policy → Protection policy"""
    if not v: return ""
    return str(v).replace("_", " ").capitalize()

def p(*a): print(*a, flush=True)

SPECIAL_UUID_COLS = frozenset([
    "external_reference_company", "external_reference_value",
    "owners", "groups", "account_uuids", "group_uuids",
    "visible_to_circles_uuids",
])
def is_uuid_col(h):
    h = str(h or "").lower()
    return (h == "uuid" or h.endswith("_uuid") or h.endswith("_uuids")
            or h in SPECIAL_UUID_COLS)

# ── Load exports ──────────────────────────────────────────────────────────────
with open(MANIFEST, encoding="utf-8") as f:
    results = json.load(f).get("results", {})
def ep(l): return results.get(l, {}).get("path", "")

with open(ep("plan_export"),   encoding="utf-8-sig", newline="") as f:
    plan_rows_raw = list(csv.DictReader(f))
with open(ep("client_export"), encoding="utf-8-sig", newline="") as f:
    client_rows_raw = list(csv.DictReader(f))
with open(ep("circle_export"), encoding="utf-8-sig", newline="") as f:
    circle_export_rows = list(csv.DictReader(f))
with open(ep("premium_export"), encoding="utf-8-sig", newline="") as f:
    prem_export_rows = list(csv.DictReader(f))

plan_by_uuid   = {r["uuid"]: r for r in plan_rows_raw}
client_by_uuid = {r["uuid"]: r for r in client_rows_raw}

# ── Name helpers (names stay in Excel, not in logs) ───────────────────────────

def client_display_name(uuid):
    """Full name for display; blank if UUID not found. Never logs."""
    r = client_by_uuid.get(str(uuid or "").strip())
    if not r: return ""
    fn = (r.get("first_name","") or "").strip()
    ln = (r.get("last_name","") or "").strip()
    if not fn and not ln:
        return (r.get("name","") or "").strip()
    t = (r.get("title","") or "").strip()
    return " ".join(p for p in [t, fn, ln] if p)

def client_sort_key(uuid):
    """Sort key (last_name lowercase). Not displayed."""
    r = client_by_uuid.get(str(uuid or "").strip(), {})
    ln = (r.get("last_name","") or "").strip().lower()
    return ln if ln else (r.get("name","") or "zzz").lower()

def plan_owners(plan_uuid):
    """(owner1_name, owner2_name_plus) tuple."""
    pe = plan_by_uuid.get(str(plan_uuid or "").strip(), {})
    uuids = [u.strip() for u in (pe.get("client_uuids","") or "").split(",") if u.strip()]
    names = [client_display_name(u) for u in uuids]
    o1 = names[0] if names else ""
    o2 = "; ".join(n for n in names[1:] if n)
    return o1, o2

def plan_adviser_name(plan_uuid):
    """Adviser name from plan export active_adviser_name."""
    pe = plan_by_uuid.get(str(plan_uuid or "").strip(), {})
    return (pe.get("active_adviser_name","") or "").strip()

def plan_sort_key(plan_uuid):
    """(owner1_lastname, provider) for sorting."""
    pe = plan_by_uuid.get(str(plan_uuid or "").strip(), {})
    uuids = [u.strip() for u in (pe.get("client_uuids","") or "").split(",") if u.strip()]
    sk = client_sort_key(uuids[0]) if uuids else "zzz"
    prov = (pe.get("provider","") or "").lower()
    return (sk, prov)

# Adviser name lookup (UUID → name)
adv_names = {}
for r in plan_rows_raw:
    u = (r.get("active_adviser_uuid","") or "").strip()
    n = (r.get("active_adviser_name","") or "").strip()
    if u and n: adv_names[u] = n
for r in client_rows_raw:
    u = (r.get("assigned_adviser_uuid","") or "").strip()
    n = (r.get("assigned_adviser","") or "").strip()
    if u and n: adv_names[u] = n
# From snapshot accounts.json
with open(os.path.join(SNAP_DIR, "accounts.json")) as f:
    _accounts = json.load(f)
for a in _accounts:
    u = (a.get("uuid","") or "").strip()
    n = (a.get("name","") or "").strip()
    if u and n: adv_names[u] = n

# Circle membership: client_uuid → [circle_name, ...]
client_circles = defaultdict(list)
circle_name_by_uuid = {}
for r in circle_export_rows:
    cuuid = (r.get("account_uuid","") or "").strip()
    cname = (r.get("circle_name","") or "").strip()
    circu = (r.get("circle_uuid","") or "").strip()
    if cuuid and cname and cname not in client_circles[cuuid]:
        client_circles[cuuid].append(cname)
    if circu and cname:
        circle_name_by_uuid[circu] = cname

# Latest valuation per plan (for Charges INFO: Latest value)
latest_val = {}
wb_v = openpyxl.load_workbook(VAL_XLSX, read_only=True, data_only=True)
ws_v = wb_v.worksheets[0]
vrows = list(ws_v.iter_rows(values_only=True))
vhdrs = list(vrows[0])
pu_i = vhdrs.index("plan_uuid")
vd_i = vhdrs.index("valuation_date")
vl_i = vhdrs.index("value")
for r in vrows[1:]:
    pu = str(r[pu_i] or "").strip()
    vd = r[vd_i]
    vl = r[vl_i]
    if pu and vd is not None and vl is not None:
        if pu not in latest_val or vd > latest_val[pu][0]:
            latest_val[pu] = (vd, vl)
wb_v.close()

def fmt_val(v):
    """Format value as £X,XXX (already in pounds from step3)."""
    if v is None: return ""
    try:
        return f"£{float(v):,.0f}"
    except (ValueError, TypeError):
        return str(v)

# Premium UUID → export row (for transferred_from lookup)
prem_export_by_uuid = {r.get("uuid","").strip(): r for r in prem_export_rows}

# Plan policy_number from Plannr_Plans.xlsx (after step6 restores placeholders)
plan_pn = {}
wb_pp = openpyxl.load_workbook(PLANS_XLSX, read_only=True, data_only=True)
ws_pp = wb_pp["Plans"]
pprows = list(ws_pp.iter_rows(values_only=True))
pphdrs = list(pprows[0])
_ppu_i = pphdrs.index("uuid") if "uuid" in pphdrs else -1
_ppn_i = pphdrs.index("policy_number") if "policy_number" in pphdrs else -1
for r in pprows[1:]:
    if _ppu_i >= 0 and _ppn_i >= 0:
        pu = str(r[_ppu_i] or "").strip()
        pn = r[_ppn_i]
        if pu: plan_pn[pu] = pn
wb_pp.close()

# ── Sheet I/O ─────────────────────────────────────────────────────────────────

def read_sheet(path, name=None, idx=0):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[name] if name else wb.worksheets[idx]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    if not rows: return [], []
    hdrs = list(rows[0])
    data = [{hdrs[i]: r[i] for i in range(len(hdrs))}
            for r in rows[1:] if any(v is not None for v in r)]
    return hdrs, data

def dedup_headers(hdrs, data):
    """Keep only first occurrence of each column header; others are dropped."""
    seen = set()
    keep = []
    for h in hdrs:
        if h not in seen:
            seen.add(h)
            keep.append(h)
    # Row dicts already have last-write-wins for duplicate keys (which is fine
    # since step6 wrote the same value to all duplicate Review columns).
    data = [{h: r.get(h) for h in keep} for r in data]
    return keep, data

def write_sheet_to_wb(wb, name, idx, hdrs, rows, freeze_col, extra_dv=None):
    if name in wb.sheetnames:
        del wb[name]
    ws = wb.create_sheet(name, idx)
    for i, h in enumerate(hdrs, 1):
        c = ws.cell(1, i, h)
        style_hdr(c)
    for ri, row in enumerate(rows, 2):
        for ci, h in enumerate(hdrs, 1):
            ws.cell(ri, ci, row.get(h))
    fp_col = get_column_letter(freeze_col + 1)
    ws.freeze_panes = f"{fp_col}2"
    if rows:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(hdrs))}{len(rows)+1}"
    if extra_dv:
        for dv in extra_dv:
            ws.add_data_validation(dv)
    return ws

def build_col_order(orig_hdrs, first_cols, info_col_names, move_cols=None):
    """
    Returns ordered list of column headers:
    first_cols → info_col_names → move_cols → body → uuid_cols
    """
    fixed = set(first_cols) | set(info_col_names) | set(move_cols or [])
    body  = [h for h in orig_hdrs if h not in fixed and not is_uuid_col(h)]
    uuids = [h for h in orig_hdrs if h not in fixed and is_uuid_col(h)]
    return list(first_cols) + list(info_col_names) + list(move_cols or []) + body + uuids

def freeze_pos(first_cols, info_cols, move_cols):
    return len(first_cols) + len(info_cols) + len(move_cols or [])

# ── Verify INFO cells contain no UUIDs ───────────────────────────────────────
def check_no_uuid_in_info(rows, info_cols, sheet_name):
    bad = 0
    for r in rows:
        for col in info_cols:
            v = r.get(col, "")
            if v and looks_uuid(str(v).strip()):
                bad += 1
    if bad:
        p(f"  WARNING: {sheet_name} has {bad} INFO cell(s) containing UUIDs")

# ── 1. Clients ─────────────────────────────────────────────────────────────────
p("Processing Clients...")
hdrs_cl, data_cl = read_sheet(CLIENTS_XLSX, "Clients")
hdrs_cl, data_cl = dedup_headers(hdrs_cl, data_cl)

INFO_COLS_CL = ["INFO: Adviser", "INFO: Administrator", "INFO: Introduced by", "INFO: Circle"]
FIRST_COLS_CL = ["Review", "title", "first_name", "last_name", "entity_name"]
MOVE_COLS_CL = []

new_rows_cl = []
for r in data_cl:
    row = dict(r)
    cuuid = str(row.get("uuid","") or "")
    ce = client_by_uuid.get(cuuid, {})
    row["INFO: Adviser"]       = (ce.get("assigned_adviser","") or "").strip() or None
    row["INFO: Administrator"] = (ce.get("assigned_administrator","") or "").strip() or None
    row["INFO: Introduced by"] = (ce.get("introduced_by","") or "").strip() or None
    circs = client_circles.get(cuuid, [])
    row["INFO: Circle"]        = "; ".join(circs) if circs else None
    new_rows_cl.append(row)

# Drop INFO cols that are all empty
INFO_COLS_CL = [c for c in INFO_COLS_CL if any(not blank(r.get(c)) for r in new_rows_cl)]

new_hdrs_cl = build_col_order(hdrs_cl, FIRST_COLS_CL, INFO_COLS_CL, MOVE_COLS_CL)
for info_col in INFO_COLS_CL:
    if info_col not in new_hdrs_cl:
        new_hdrs_cl.insert(len(FIRST_COLS_CL), info_col)

# Sort by last_name
new_rows_cl.sort(key=lambda r: (str(r.get("last_name","") or r.get("entity_name","") or "zzz").lower(),
                                str(r.get("first_name","") or "").lower()))

check_no_uuid_in_info(new_rows_cl, INFO_COLS_CL, "Clients")
fp_cl = freeze_pos(FIRST_COLS_CL, INFO_COLS_CL, MOVE_COLS_CL)

wb_cl = openpyxl.load_workbook(CLIENTS_XLSX)
write_sheet_to_wb(wb_cl, "Clients", 0, new_hdrs_cl, new_rows_cl, fp_cl)
wb_cl.save(CLIENTS_XLSX)
p(f"  Clients: {len(new_rows_cl)} rows | added {INFO_COLS_CL} | moved {FIRST_COLS_CL[1:]}")

# ── 2. Circles ─────────────────────────────────────────────────────────────────
p("Processing Circles...")
hdrs_cir, data_cir = read_sheet(CIRCLES_XLSX, "Circles")
hdrs_cir, data_cir = dedup_headers(hdrs_cir, data_cir)

INFO_COLS_CIR = ["INFO: Members"]
FIRST_COLS_CIR = []
MOVE_COLS_CIR = ["name"]

new_rows_cir = []
for r in data_cir:
    row = dict(r)
    uuid_list_raw = str(row.get("account_uuids","") or "")
    uuids = [u.strip() for u in uuid_list_raw.split(",") if u.strip()]
    names = [client_display_name(u) for u in uuids if client_display_name(u)]
    row["INFO: Members"] = "; ".join(names) if names else None
    new_rows_cir.append(row)

INFO_COLS_CIR = [c for c in INFO_COLS_CIR if any(not blank(r.get(c)) for r in new_rows_cir)]
new_hdrs_cir = build_col_order(hdrs_cir, FIRST_COLS_CIR, INFO_COLS_CIR, MOVE_COLS_CIR)

# Sort by name
new_rows_cir.sort(key=lambda r: str(r.get("name","") or "zzz").lower())

check_no_uuid_in_info(new_rows_cir, INFO_COLS_CIR, "Circles")
fp_cir = freeze_pos(FIRST_COLS_CIR, INFO_COLS_CIR, MOVE_COLS_CIR)

wb_cir = openpyxl.load_workbook(CIRCLES_XLSX)
write_sheet_to_wb(wb_cir, "Circles", 0, new_hdrs_cir, new_rows_cir, fp_cir)
wb_cir.save(CIRCLES_XLSX)
p(f"  Circles: {len(new_rows_cir)} rows | added {INFO_COLS_CIR}")

# ── 3. Relationships ──────────────────────────────────────────────────────────
p("Processing Relationships...")
hdrs_rel, data_rel = read_sheet(RELS_XLSX, "Relationships")
hdrs_rel, data_rel = dedup_headers(hdrs_rel, data_rel)

new_rows_rel = []
for r in data_rel:
    row = dict(r)
    row["INFO: Client"]     = client_display_name(row.get("account_uuid"))
    row["INFO: Related to"] = client_display_name(row.get("related_account_uuid"))
    new_rows_rel.append(row)

# Custom column order: INFO: Client, relationship, inverse_relationship, INFO: Related to, [UUIDs]
REL_INFO = ["INFO: Client", "INFO: Related to"]
REL_INFO = [c for c in REL_INFO if any(not blank(r.get(c)) for r in new_rows_rel)]
mid_cols = [h for h in hdrs_rel if not is_uuid_col(h)]
uuid_cols = [h for h in hdrs_rel if is_uuid_col(h)]
new_hdrs_rel = (["INFO: Client", "relationship", "inverse_relationship", "INFO: Related to"]
                + [h for h in mid_cols if h not in ("relationship","inverse_relationship")]
                + uuid_cols)
fp_rel = 4  # through INFO: Related to

# Sort by INFO: Client surname
new_rows_rel.sort(key=lambda r: (
    client_sort_key(r.get("account_uuid")),
    str(r.get("relationship","") or "").lower()
))

check_no_uuid_in_info(new_rows_rel, REL_INFO, "Relationships")
wb_rel = openpyxl.load_workbook(RELS_XLSX)
write_sheet_to_wb(wb_rel, "Relationships", 0, new_hdrs_rel, new_rows_rel, fp_rel)
wb_rel.save(RELS_XLSX)
p(f"  Relationships: {len(new_rows_rel)} rows | added INFO: Client, INFO: Related to")

# ── 4. Dependants ──────────────────────────────────────────────────────────────
p("Processing Dependants...")
hdrs_dep, data_dep = read_sheet(DEPS_XLSX, "Dependants")
hdrs_dep, data_dep = dedup_headers(hdrs_dep, data_dep)

INFO_COLS_DEP = ["INFO: Client 1", "INFO: Client 2"]
FIRST_COLS_DEP = []
MOVE_COLS_DEP = []

new_rows_dep = []
for r in data_dep:
    row = dict(r)
    uuids = [u.strip() for u in str(row.get("account_uuids","") or "").split(",") if u.strip()]
    row["INFO: Client 1"] = client_display_name(uuids[0]) if len(uuids) > 0 else None
    row["INFO: Client 2"] = client_display_name(uuids[1]) if len(uuids) > 1 else None
    new_rows_dep.append(row)

INFO_COLS_DEP = [c for c in INFO_COLS_DEP if any(not blank(r.get(c)) for r in new_rows_dep)]
new_hdrs_dep = build_col_order(hdrs_dep, FIRST_COLS_DEP, INFO_COLS_DEP, MOVE_COLS_DEP)
fp_dep = freeze_pos(FIRST_COLS_DEP, INFO_COLS_DEP, MOVE_COLS_DEP)

# Sort by Client 1 surname
new_rows_dep.sort(key=lambda r: client_sort_key(
    [u.strip() for u in str(r.get("account_uuids","") or "").split(",") if u.strip()][0]
    if str(r.get("account_uuids","") or "").strip() else ""))

check_no_uuid_in_info(new_rows_dep, INFO_COLS_DEP, "Dependants")
wb_dep = openpyxl.load_workbook(DEPS_XLSX)
write_sheet_to_wb(wb_dep, "Dependants", 0, new_hdrs_dep, new_rows_dep, fp_dep)
wb_dep.save(DEPS_XLSX)
p(f"  Dependants: {len(new_rows_dep)} rows | added {INFO_COLS_DEP}")

# ── 5. Income & Employment ─────────────────────────────────────────────────────
p("Processing Income & Employment...")
hdrs_iae, data_iae = read_sheet(IAE_XLSX)
hdrs_iae, data_iae = dedup_headers(hdrs_iae, data_iae)

INFO_COLS_IAE = ["INFO: Client"]

new_rows_iae = []
for r in data_iae:
    row = dict(r)
    row["INFO: Client"] = client_display_name(row.get("client_uuid"))
    new_rows_iae.append(row)

INFO_COLS_IAE = [c for c in INFO_COLS_IAE if any(not blank(r.get(c)) for r in new_rows_iae)]
new_hdrs_iae = build_col_order(hdrs_iae, [], INFO_COLS_IAE, [])
fp_iae = freeze_pos([], INFO_COLS_IAE, [])

new_rows_iae.sort(key=lambda r: client_sort_key(r.get("client_uuid")))

check_no_uuid_in_info(new_rows_iae, INFO_COLS_IAE, "Income & Employment")
wb_iae = openpyxl.load_workbook(IAE_XLSX)
write_sheet_to_wb(wb_iae, "Income_Employment", 0, new_hdrs_iae, new_rows_iae, fp_iae)
wb_iae.save(IAE_XLSX)
p(f"  Income & Employment: {len(new_rows_iae)} rows | added {INFO_COLS_IAE}")

# ── 6. Expenditure ──────────────────────────────────────────────────────────────
p("Processing Expenditure...")
hdrs_exp, data_exp = read_sheet(EXP_XLSX)
hdrs_exp, data_exp = dedup_headers(hdrs_exp, data_exp)

INFO_COLS_EXP = ["INFO: Client"]

new_rows_exp = []
for r in data_exp:
    row = dict(r)
    row["INFO: Client"] = client_display_name(row.get("client_uuid"))
    new_rows_exp.append(row)

INFO_COLS_EXP = [c for c in INFO_COLS_EXP if any(not blank(r.get(c)) for r in new_rows_exp)]
new_hdrs_exp = build_col_order(hdrs_exp, [], INFO_COLS_EXP, [])
fp_exp = freeze_pos([], INFO_COLS_EXP, [])

new_rows_exp.sort(key=lambda r: client_sort_key(r.get("client_uuid")))

check_no_uuid_in_info(new_rows_exp, INFO_COLS_EXP, "Expenditure")
wb_exp = openpyxl.load_workbook(EXP_XLSX)
write_sheet_to_wb(wb_exp, "Expenditure", 0, new_hdrs_exp, new_rows_exp, fp_exp)
wb_exp.save(EXP_XLSX)
p(f"  Expenditure: {len(new_rows_exp)} rows | added {INFO_COLS_EXP}")

# ── 7a. Plans (main sheet) ──────────────────────────────────────────────────────
p("Processing Plans (main)...")
hdrs_pl, data_pl = read_sheet(PLANS_XLSX, "Plans")
hdrs_pl, data_pl = dedup_headers(hdrs_pl, data_pl)

INFO_COLS_PL = ["INFO: Owner 1", "INFO: Owner 2", "INFO: Adviser"]
FIRST_COLS_PL = ["Review"]
MOVE_COLS_PL = ["name", "provider", "type", "sub_type", "policy_number", "status"]

new_rows_pl = []
for r in data_pl:
    row = dict(r)
    puuid = str(row.get("uuid","") or "")
    o1, o2 = plan_owners(puuid)
    adv    = plan_adviser_name(puuid)
    # Fallback: if plan export has no adviser name, try adviser_uuid in this row
    if not adv:
        advu = str(row.get("adviser_uuid","") or "").strip()
        adv  = adv_names.get(advu, "")
    row["INFO: Owner 1"] = o1 or None
    row["INFO: Owner 2"] = o2 or None
    row["INFO: Adviser"] = adv or None
    # type/sub_type: plain words already in moved columns (raw value kept)
    new_rows_pl.append(row)

INFO_COLS_PL = [c for c in INFO_COLS_PL if any(not blank(r.get(c)) for r in new_rows_pl)]
new_hdrs_pl = build_col_order(hdrs_pl, FIRST_COLS_PL, INFO_COLS_PL, MOVE_COLS_PL)
fp_pl = freeze_pos(FIRST_COLS_PL, INFO_COLS_PL, MOVE_COLS_PL)

# Sort by Owner 1 surname, then provider
new_rows_pl.sort(key=lambda r: (
    client_sort_key([u.strip() for u in str(plan_by_uuid.get(str(r.get("uuid",""))
        or "").get("client_uuids","") or "").split(",") if u.strip()][0]
        if [u.strip() for u in str(plan_by_uuid.get(str(r.get("uuid",""))
        or "").get("client_uuids","") or "").split(",") if u.strip()] else ""),
    str(r.get("provider","") or "").lower()
))

check_no_uuid_in_info(new_rows_pl, INFO_COLS_PL, "Plans")

# ── 7b. Provider_plans_not_in_Plannr ──────────────────────────────────────────
p("Processing Plans (provider tab)...")
hdrs_prov, data_prov = read_sheet(PLANS_XLSX, "Provider_plans_not_in_Plannr")
hdrs_prov, data_prov = dedup_headers(hdrs_prov, data_prov)

# Drop 'source' column (= Firm Name / provider source — constant across each source group)
# Build "Client in your Plannr" lookup
# Investor name is in extra_1 for Rathbones; Transact/Nucleus have no investor name

PLAN_STRIP_SUFFIXES = re.compile(
    r'\b(STPS|ISA|GIA|SIPP|EIS|JISA|Junior ISA|Bond|Drawdown|Annuity|BOND|'
    r'Pension|Investment|Protection|Endowment|AVC|FSAVC|'
    r'Junior|Jr\.?|Trust|Ltd|plc|Co\.?|Inc\.?)\b',
    re.I
)
PLAN_STRIP_TITLES = re.compile(
    r'\b(Mr|Mrs|Ms|Miss|Dr|Prof|Emeritus|Rev|Sir|Lord|Lady|Mx|'
    r'The|Hon|Capt|Major|Col|Brig)\b\.?',
    re.I
)
PLAN_STRIP_PARENS = re.compile(r'\(.*?\)')

def normalize_investor_name(v):
    s = str(v or "")
    s = PLAN_STRIP_PARENS.sub(" ", s)
    s = PLAN_STRIP_SUFFIXES.sub(" ", s)
    s = PLAN_STRIP_TITLES.sub(" ", s)
    return s.strip().lower()

# Build client last_name → uuid map (for matching)
client_ln_map = defaultdict(list)
for r in client_rows_raw:
    ln = (r.get("last_name","") or "").strip().lower()
    fn = (r.get("first_name","") or "").strip().lower()
    en = (r.get("name","") or "").strip().lower()
    if ln:
        client_ln_map[ln].append(r["uuid"])
    # Also add entity name words
    for w in en.split():
        if len(w) > 3 and w not in ("the","and","ltd","plc","llp"):
            client_ln_map[w].append(r["uuid"])

def match_investor_to_client(investor_name_raw):
    norm = normalize_investor_name(investor_name_raw)
    words = [w.strip("'.,-") for w in norm.split() if len(w.strip("'.,-")) > 1]
    matches = set()
    for w in words:
        for cuuid in client_ln_map.get(w, []):
            matches.add(cuuid)
    if len(matches) == 1:
        return client_display_name(matches.pop())
    return None

provider_match_count = 0
new_rows_prov = []
for r in data_prov:
    row = dict(r)
    inv = row.get("extra_1")
    matched_name = None
    if inv:
        matched_name = match_investor_to_client(str(inv))
    row["Client in your Plannr"] = matched_name
    if matched_name:
        provider_match_count += 1
    new_rows_prov.append(row)

# Column order for provider tab: drop 'source' (firm name); add 'Client in your Plannr'
# Keep: ref_1, ref_2, extra_1, extra_2, action, Client in your Plannr
prov_hdrs_new = ["Client in your Plannr", "ref_1", "ref_2", "extra_1", "extra_2", "action", "source"]
# source goes last (it's constant "Rathbones"/"Transact/Nucleus" — still useful ref)

# No freeze for provider tab (small, no Review)
# Re-apply the Action dropdown DV
action_dv = DataValidation(
    type="list",
    formula1='"Add to Plannr,Already in Plannr,Not ours,Closed"',
    allow_blank=True,
    showDropDown=False
)
action_dv.sqref = f"F2:F{len(new_rows_prov)+1}"

# Write both Plans sheets together
wb_pl = openpyxl.load_workbook(PLANS_XLSX)
write_sheet_to_wb(wb_pl, "Plans", 0, new_hdrs_pl, new_rows_pl, fp_pl)
write_sheet_to_wb(wb_pl, "Provider_plans_not_in_Plannr", 1, prov_hdrs_new, new_rows_prov, 0,
                  extra_dv=[action_dv])
wb_pl.save(PLANS_XLSX)
p(f"  Plans (main): {len(new_rows_pl)} rows | added {INFO_COLS_PL} | moved {MOVE_COLS_PL}")
p(f"  Plans (provider tab): {len(new_rows_prov)} rows | added Client in your Plannr | {provider_match_count} matched")

# ── 8. Valuations ──────────────────────────────────────────────────────────────
p("Processing Valuations...")
hdrs_val, data_val = read_sheet(VAL_XLSX)
hdrs_val, data_val = dedup_headers(hdrs_val, data_val)

INFO_COLS_VAL = ["INFO: Owner 1", "INFO: Owner 2", "INFO: Provider", "INFO: Type", "INFO: Sub-type"]
# INFO: Sub-account omitted — sub-account names not available without API
MOVE_COLS_VAL = ["policy_number"]

new_rows_val = []
for r in data_val:
    row = dict(r)
    puuid = str(row.get("plan_uuid","") or "")
    o1, o2 = plan_owners(puuid)
    pe = plan_by_uuid.get(puuid, {})
    row["INFO: Owner 1"]  = o1 or None
    row["INFO: Owner 2"]  = o2 or None
    row["INFO: Provider"] = (pe.get("provider","") or "").strip() or None
    row["INFO: Type"]     = to_plain(pe.get("type","")) or None
    row["INFO: Sub-type"] = to_plain(pe.get("sub_type","")) or None
    new_rows_val.append(row)

INFO_COLS_VAL = [c for c in INFO_COLS_VAL if any(not blank(r.get(c)) for r in new_rows_val)]
new_hdrs_val = build_col_order(hdrs_val, [], INFO_COLS_VAL, MOVE_COLS_VAL)
fp_val = freeze_pos([], INFO_COLS_VAL, MOVE_COLS_VAL)

new_rows_val.sort(key=lambda r: plan_sort_key(r.get("plan_uuid")))

check_no_uuid_in_info(new_rows_val, INFO_COLS_VAL, "Valuations")
wb_val = openpyxl.load_workbook(VAL_XLSX)
write_sheet_to_wb(wb_val, "Valuations", 0, new_hdrs_val, new_rows_val, fp_val)
wb_val.save(VAL_XLSX)
p(f"  Valuations: {len(new_rows_val)} rows | added {INFO_COLS_VAL} | moved {MOVE_COLS_VAL}")

# ── 9. Premiums ────────────────────────────────────────────────────────────────
p("Processing Premiums...")
hdrs_prm, data_prm = read_sheet(PREM_XLSX)
hdrs_prm, data_prm = dedup_headers(hdrs_prm, data_prm)

INFO_COLS_PRM = ["INFO: Owner 1", "INFO: Owner 2", "INFO: Provider", "INFO: Type",
                 "INFO: Sub-type", "INFO: Policy number", "INFO: Transferred from"]

new_rows_prm = []
for r in data_prm:
    row = dict(r)
    puuid = str(row.get("plan_uuid","") or "")
    o1, o2 = plan_owners(puuid)
    pe = plan_by_uuid.get(puuid, {})
    row["INFO: Owner 1"]      = o1 or None
    row["INFO: Owner 2"]      = o2 or None
    row["INFO: Provider"]     = (pe.get("provider","") or "").strip() or None
    row["INFO: Type"]         = to_plain(pe.get("type","")) or None
    row["INFO: Sub-type"]     = to_plain(pe.get("sub_type","")) or None
    pn = plan_pn.get(puuid)
    row["INFO: Policy number"] = pn if not blank(pn) else None

    tfuuid = str(row.get("transferred_from_plan_uuid","") or "").strip()
    tf_info = None
    if tfuuid and tfuuid != "None":
        tf_pe = plan_by_uuid.get(tfuuid, {})
        tf_prov = (tf_pe.get("provider","") or "").strip()
        tf_pn   = plan_pn.get(tfuuid)
        if tf_prov and not blank(tf_pn):
            tf_info = f"{tf_prov} ({tf_pn})"
        elif tf_prov:
            tf_info = tf_prov
        else:
            # fallback to columns already in the row
            tf_type = to_plain(row.get("transferred_from_plan_type","") or "")
            if tf_type: tf_info = tf_type
    row["INFO: Transferred from"] = tf_info

    new_rows_prm.append(row)

INFO_COLS_PRM = [c for c in INFO_COLS_PRM if any(not blank(r.get(c)) for r in new_rows_prm)]
new_hdrs_prm = build_col_order(hdrs_prm, [], INFO_COLS_PRM, [])
fp_prm = freeze_pos([], INFO_COLS_PRM, [])

new_rows_prm.sort(key=lambda r: plan_sort_key(r.get("plan_uuid")))

check_no_uuid_in_info(new_rows_prm, INFO_COLS_PRM, "Premiums")
wb_prm = openpyxl.load_workbook(PREM_XLSX)
write_sheet_to_wb(wb_prm, "Premiums", 0, new_hdrs_prm, new_rows_prm, fp_prm)
wb_prm.save(PREM_XLSX)
p(f"  Premiums: {len(new_rows_prm)} rows | added {INFO_COLS_PRM}")

# ── 10. Charges ────────────────────────────────────────────────────────────────
p("Processing Charges...")
hdrs_chr, data_chr = read_sheet(CHRG_XLSX, "Charges")
hdrs_chr, data_chr = dedup_headers(hdrs_chr, data_chr)

INFO_COLS_CHR = ["INFO: Owner 1", "INFO: Owner 2", "INFO: Provider", "INFO: Type",
                 "INFO: Sub-type", "INFO: Policy number", "INFO: Adviser", "INFO: Latest value"]

new_rows_chr = []
for r in data_chr:
    row = dict(r)
    puuid = str(row.get("plan_uuid","") or "")
    o1, o2 = plan_owners(puuid)
    pe = plan_by_uuid.get(puuid, {})
    row["INFO: Owner 1"]      = o1 or None
    row["INFO: Owner 2"]      = o2 or None
    row["INFO: Provider"]     = (pe.get("provider","") or "").strip() or None
    row["INFO: Type"]         = to_plain(pe.get("type","")) or None
    row["INFO: Sub-type"]     = to_plain(pe.get("sub_type","")) or None
    pn = plan_pn.get(puuid)
    row["INFO: Policy number"] = pn if not blank(pn) else None
    advu = str(row.get("adviser_uuid","") or "").strip()
    row["INFO: Adviser"]       = adv_names.get(advu, "") or None
    lv = latest_val.get(puuid)
    row["INFO: Latest value"]  = fmt_val(lv[1]) if lv else None
    new_rows_chr.append(row)

INFO_COLS_CHR = [c for c in INFO_COLS_CHR if any(not blank(r.get(c)) for r in new_rows_chr)]
new_hdrs_chr = build_col_order(hdrs_chr, [], INFO_COLS_CHR, [])
fp_chr = freeze_pos([], INFO_COLS_CHR, [])

new_rows_chr.sort(key=lambda r: plan_sort_key(r.get("plan_uuid")))

check_no_uuid_in_info(new_rows_chr, INFO_COLS_CHR, "Charges")
wb_chr = openpyxl.load_workbook(CHRG_XLSX)
write_sheet_to_wb(wb_chr, "Charges", 0, new_hdrs_chr, new_rows_chr, fp_chr)
wb_chr.save(CHRG_XLSX)
p(f"  Charges: {len(new_rows_chr)} rows | added {INFO_COLS_CHR}")

# ── Final row-count check ──────────────────────────────────────────────────────
p("\nRow count check:")
EXPECTED = {
    CLIENTS_XLSX: 375, CIRCLES_XLSX: 92, RELS_XLSX: 190, DEPS_XLSX: 250,
    IAE_XLSX: 350, EXP_XLSX: 2381, VAL_XLSX: 1528, PREM_XLSX: 173, CHRG_XLSX: 675,
}
for path, exp in EXPECTED.items():
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    actual = len([r for r in rows[1:] if any(v is not None for v in r)])
    wb.close()
    status = "OK" if actual == exp else f"MISMATCH (expected {exp})"
    p(f"  {os.path.basename(path)}: {actual} {status}")
# Plans
wb_pl2 = openpyxl.load_workbook(PLANS_XLSX, read_only=True, data_only=True)
for sname, exp in [("Plans", 1871), ("Provider_plans_not_in_Plannr", 51)]:
    ws2 = wb_pl2[sname]
    rows2 = list(ws2.iter_rows(values_only=True))
    actual2 = len([r for r in rows2[1:] if any(v is not None for v in r)])
    status2 = "OK" if actual2 == exp else f"MISMATCH (expected {exp})"
    p(f"  Plannr_Plans.xlsx / {sname}: {actual2} {status2}")
wb_pl2.close()

p(f"\nProvider tab: {provider_match_count} of {len(new_rows_prov)} rows matched to a client")
p("STOP")
