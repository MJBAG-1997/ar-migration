"""
step8_two_fixes.py
Fix 1: Dependants — populate account_uuids for shared dependants; add INFO: Client 2.
Fix 2: Provider tab — rename generic headers; add Investor name; redo client match.
Rule 9: names in Excel only; counts in terminal.
"""
import os, json, csv, re
from collections import defaultdict
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

DATA_ROOT   = r"C:\dev\AR Migration"
SNAP_DIR    = os.path.join(DATA_ROOT, "snapshot_20261012T090906Z")
EXPORTS_DIR = os.path.join(DATA_ROOT, "exports_20261014")
MANIFEST    = os.path.join(EXPORTS_DIR, "_manifest.json")
RATH_PATH   = os.path.join(DATA_ROOT, "Jim list  - Rathbones.xlsx")
WB_PATH     = os.path.join(DATA_ROOT,
              "Hoban_Financial_Planning_Ltd_260922_Payment_Rates_113727875.xlsx")

DEPS_PATH   = os.path.join(DATA_ROOT, "Plannr_Dependants.xlsx")
PLANS_PATH  = os.path.join(DATA_ROOT, "Plannr_Plans.xlsx")

def blank(v): return v is None or str(v).strip() in ("","None","nan")
def s(v):     return str(v).strip() if not blank(v) else ""
def p(*a):    print(*a, flush=True)

HDR_FONT  = Font(bold=True)
HDR_FILL  = PatternFill(fill_type="solid", fgColor="F2F2F2")
INFO_FONT = Font(bold=True, color="1F497D")
INFO_FILL = PatternFill(fill_type="solid", fgColor="DCE6F1")
ORIG_FONT = Font(bold=True, color="7F6000")
ORIG_FILL = PatternFill(fill_type="solid", fgColor="FFF2CC")

def style_hdr(cell):
    h = str(cell.value or "")
    if h.startswith("INFO:"):
        cell.font = INFO_FONT; cell.fill = INFO_FILL
    elif h == "Review" or h.startswith("ORIGINAL_"):
        cell.font = ORIG_FONT; cell.fill = ORIG_FILL
    else:
        cell.font = HDR_FONT; cell.fill = HDR_FILL

def read_sheet(path, name):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[name] if name in wb.sheetnames else wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    if not rows: return [], []
    hdrs = list(rows[0])
    data = [{hdrs[i]: r[i] for i in range(len(hdrs))}
            for r in rows[1:] if any(v is not None for v in r)]
    return hdrs, data

def rewrite_sheet(path, sheet_name, hdrs, rows, freeze_col,
                  other_sheet_fn=None, dvs=None):
    """Replace sheet in workbook, preserve others via other_sheet_fn callback."""
    wb = openpyxl.load_workbook(path)
    # Collect other sheets' data
    preserved = {}
    for sname in wb.sheetnames:
        if sname != sheet_name:
            ws_o = wb[sname]
            preserved[sname] = (
                list(ws_o.iter_rows(values_only=True)),
                list(ws_o.data_validations.dataValidation),
            )
    idx = wb.sheetnames.index(sheet_name) if sheet_name in wb.sheetnames else 0
    del wb[sheet_name]
    ws = wb.create_sheet(sheet_name, idx)
    for i, h in enumerate(hdrs, 1):
        c = ws.cell(1, i, h); style_hdr(c)
    for ri, row in enumerate(rows, 2):
        for ci, h in enumerate(hdrs, 1):
            ws.cell(ri, ci, row.get(h))
    ws.freeze_panes = f"{get_column_letter(freeze_col + 1)}2"
    if rows:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(hdrs))}{len(rows)+1}"
    if dvs:
        for dv in dvs: ws.add_data_validation(dv)
    # Restore other sheets
    for sname, (sdata, sdvs) in preserved.items():
        if sname not in wb.sheetnames:
            ws_x = wb.create_sheet(sname)
        else:
            ws_x = wb[sname]
        for ri, r in enumerate(sdata, 1):
            for ci, val in enumerate(r, 1):
                ws_x.cell(ri, ci, val)
        for dv in sdvs:
            ws_x.add_data_validation(dv)
    wb.save(path)

# ── Load reference data ─────────────────────────────────────────────────────

with open(MANIFEST, encoding="utf-8") as f:
    man = json.load(f).get("results", {})
def ep(l): return man.get(l, {}).get("path","")

with open(ep("client_export"), encoding="utf-8-sig", newline="") as f:
    client_rows = list(csv.DictReader(f))
client_by_uuid = {r["uuid"]: r for r in client_rows}

def client_name(cuuid):
    """Rule 9: for Excel only."""
    ce = client_by_uuid.get(s(cuuid), {})
    ent = s(ce.get("entity_name",""))
    if ent: return ent
    return " ".join(x for x in [s(ce.get("first_name","")), s(ce.get("last_name",""))] if x)

def client_last_name_sort(cuuid):
    ce = client_by_uuid.get(s(cuuid), {})
    return s(ce.get("last_name","")).lower() or s(ce.get("entity_name","")).lower()

# ── FIX 1: Dependants ─────────────────────────────────────────────────────────

p("=== Fix 1: Dependants ===")

# Build dep_uuid → [account_uuids] from account_details snapshot
dep_to_accounts = defaultdict(list)
ad_path = os.path.join(SNAP_DIR, "account_details.json")
with open(ad_path, encoding="utf-8") as f:
    account_details = json.load(f)   # dict: account_uuid → {dependants: [...]}

for acct_uuid, details in account_details.items():
    for dep in (details.get("dependants") or []):
        duuid = s(dep.get("uuid",""))
        if duuid and acct_uuid not in dep_to_accounts[duuid]:
            dep_to_accounts[duuid].append(acct_uuid)

multi_dep = {d: accts for d, accts in dep_to_accounts.items() if len(accts) > 1}
p(f"Dependants with >1 account in snapshot: {len(multi_dep)}")

# Read current Dependants sheet
hdrs_dep, data_dep = read_sheet(DEPS_PATH, "Dependants")
hdrs_dep = list(hdrs_dep)

# Find account_uuid column; rename to account_uuids
if "account_uuid" in hdrs_dep and "account_uuids" not in hdrs_dep:
    idx_au = hdrs_dep.index("account_uuid")
    hdrs_dep[idx_au] = "account_uuids"
    # Rename key in every row
    for row in data_dep:
        if "account_uuid" in row:
            row["account_uuids"] = row.pop("account_uuid")

# Ensure INFO: Client 2 column exists
if "INFO: Client 2" not in hdrs_dep:
    # Insert after INFO: Client 1
    if "INFO: Client 1" in hdrs_dep:
        ins = hdrs_dep.index("INFO: Client 1") + 1
        hdrs_dep.insert(ins, "INFO: Client 2")
    else:
        hdrs_dep.insert(0, "INFO: Client 2")

# Also ensure INFO: Client 1 exists
if "INFO: Client 1" not in hdrs_dep:
    hdrs_dep.insert(0, "INFO: Client 1")

two_client_count = 0
for row in data_dep:
    # Get all account UUIDs for this dependant
    duuid = s(row.get("external_reference_value","") or row.get("uuid",""))
    if not duuid:
        # try to find from account_uuids
        existing = s(row.get("account_uuids",""))
        if existing:
            first_acct = existing.split("|")[0].strip()
            # find dep uuid from snapshot lookup (reverse)
            for d, accts in dep_to_accounts.items():
                if first_acct in accts:
                    # Only valid if this dep is in the row's context
                    pass
    # Better: use account_uuids current value to identify account, then look up
    # Actually, use the row's first account to find dep_uuid mapping
    # We rely on external_reference_value = dep_uuid (Plannr UUID)
    existing_acct = s(row.get("account_uuids",""))
    current_acct  = existing_acct.split("|")[0].strip()

    # Find dep UUID: external_reference_value is the source Plannr UUID
    dep_uuid = s(row.get("external_reference_value",""))
    if dep_uuid and dep_uuid in dep_to_accounts:
        all_accts = dep_to_accounts[dep_uuid]
    elif current_acct:
        # Fallback: use only the current account
        all_accts = [current_acct] if current_acct else []
    else:
        all_accts = []

    if len(all_accts) > 1:
        two_client_count += 1

    # Update account_uuids
    if all_accts:
        row["account_uuids"] = "|".join(all_accts)
    # Also keep INFO: Client 1 (first account name) if not set
    if all_accts:
        row["INFO: Client 1"] = client_name(all_accts[0])
    # Set INFO: Client 2 (2nd+ account names)
    if len(all_accts) >= 2:
        row["INFO: Client 2"] = "; ".join(client_name(u) for u in all_accts[1:] if s(u))
    else:
        row["INFO: Client 2"] = None

p(f"Dependants now with two clients: {two_client_count}")

# Drop INFO: Client 2 if all blank
if all(blank(r.get("INFO: Client 2")) for r in data_dep):
    hdrs_dep = [h for h in hdrs_dep if h != "INFO: Client 2"]
    p("INFO: Client 2 dropped (all blank after fix)")
else:
    p(f"INFO: Client 2 populated on {sum(1 for r in data_dep if not blank(r.get('INFO: Client 2')))} rows")

# Sort by INFO: Client 1 surname
def dep_sort_key(row):
    name = s(row.get("INFO: Client 1",""))
    parts = name.split()
    surname = parts[-1].lower() if parts else "zzz"
    return surname

data_dep.sort(key=dep_sort_key)

# Freeze col = last INFO col position
freeze_dep = max((i+1 for i,h in enumerate(hdrs_dep) if str(h or "").startswith("INFO:")),
                 default=1)
rewrite_sheet(DEPS_PATH, "Dependants", hdrs_dep, data_dep, freeze_dep)
p(f"Saved Plannr_Dependants.xlsx ({len(data_dep)} rows)")

# ── FIX 2: Provider_plans_not_in_Plannr ─────────────────────────────────────

p("\n=== Fix 2: Provider tab ===")

# ── Build investor name lookups from source workbooks ──

# Rathbones: fund_no → (fund_name, [used as investor name proxy])
wb_r = openpyxl.load_workbook(RATH_PATH, read_only=True, data_only=True)
rath_inv = {}   # fund_no → fund_name (which embeds investor surname)
for row in wb_r.active.iter_rows(min_row=2, values_only=True):
    fn = s(row[0])
    fund_name = s(row[3])   # Fund Name (extra_1 source)
    if fn: rath_inv[fn] = fund_name
wb_r.close()

# Transact/Nucleus: portfolio_no → investor_name
wb_tn = openpyxl.load_workbook(WB_PATH, read_only=True, data_only=True)
tn_inv = {}     # portfolio_no → investor_name
for row in wb_tn["Payment Rates Summary"].iter_rows(min_row=2, values_only=True):
    port_no = s(row[3])
    investor = s(row[2])   # Investor Name column
    if port_no: tn_inv[port_no] = investor
wb_tn.close()

# ── Client matching helpers ──

# Noise tokens to ignore in investor names
STOP = {
    'mr','mrs','ms','miss','dr','prof','sir','rev','lord','lady','emeritus',
    'hon','lt','col','capt','the','bsc','mba','cfa','cfp','qc',
    # Rathbones plan type abbreviations
    'stps','sipp','isa','gia','eis','bond','pension','trust','ltd','plc',
    'plan','portfolio','fund','co','and','or','jr','sr','esq',
}

def name_tokens(v):
    v = re.sub(r"[^a-z0-9\s]", " ", s(v).lower())
    return [t for t in v.split() if t and t not in STOP and len(t) > 1]

def match_investor(investor_name):
    """
    Returns list of client UUIDs where surname matches AND
    first name or initial matches, case/title/punctuation-insensitive.
    Rule 9: never print names.
    """
    tokens = name_tokens(investor_name)
    if not tokens: return []
    matches = []
    for ce in client_rows:
        ln = re.sub(r"[^a-z]", "", s(ce.get("last_name","")).lower())
        fn = re.sub(r"[^a-z]", "", s(ce.get("first_name","")).lower())
        if not ln: continue
        # Surname match: must appear as a token (or near-match for hyphenated names)
        ln_parts = re.sub(r"[^a-z\s]", " ", s(ce.get("last_name","")).lower()).split()
        ln_parts = [re.sub(r"[^a-z]","",p) for p in ln_parts if p]
        surname_match = any(lp in tokens for lp in ln_parts if lp)
        if not surname_match: continue
        # First name / initial match
        if fn and fn in tokens:
            matches.append(ce["uuid"]); continue
        if fn and fn[:1] in [t[:1] for t in tokens if t not in (set(ln_parts))]:
            matches.append(ce["uuid"]); continue
    return matches

# ── Read and rebuild Provider tab ──

prov_sheet = "Provider_plans_not_in_Plannr"
hdrs_prov, data_prov = read_sheet(PLANS_PATH, prov_sheet)
p(f"Provider tab: {len(data_prov)} rows, current headers: {hdrs_prov}")

# Determine new header names
# Current order after step7: Client in your Plannr, ref_1, ref_2, extra_1, extra_2, action, source
# Rathbones columns: ref_1=Fund number, ref_2='', extra_1=Fund name, extra_2=Total value
# Transact cols:      ref_1=Portfolio number, ref_2=Wrapper ref, extra_1=Wrapper name, extra_2=Annual payment

RENAME = {
    "ref_1":   "Portfolio number",
    "ref_2":   "Wrapper ref",
    "extra_1": "Portfolio name",
    "extra_2": "Annual payment",
}

# Rename keys in headers and rows
new_hdrs = [RENAME.get(h, h) for h in hdrs_prov]
for row in data_prov:
    for old, new in RENAME.items():
        if old in row:
            row[new] = row.pop(old)

# Add Investor name column (after Client in your Plannr)
inv_col = "Investor name"
if inv_col not in new_hdrs:
    # Insert after Client in your Plannr
    insert_pos = new_hdrs.index("Client in your Plannr") + 1 \
                 if "Client in your Plannr" in new_hdrs else 0
    new_hdrs.insert(insert_pos, inv_col)

# Populate Investor name
for row in data_prov:
    src      = s(row.get("source","")).lower()
    port_no  = s(row.get("Portfolio number",""))
    wrap_ref = s(row.get("Wrapper ref",""))
    if "rathbones" in src:
        # Fund Name (extra_1) used as investor name proxy
        row[inv_col] = s(row.get("Portfolio name",""))
    else:
        # Transact/Nucleus: look up from workbook
        inv = tn_inv.get(port_no) or tn_inv.get(wrap_ref) or ""
        row[inv_col] = inv

# Drop any column whose every row is blank (catches Wrapper ref if all empty for Rathbones)
firm_col_names = {"firm name","firm","company"}
final_hdrs_prov = []
for h in new_hdrs:
    h_norm = s(h).lower()
    # Drop if it's a firm name column
    if h_norm in firm_col_names:
        p(f"  Dropping firm-name column: {h}")
        continue
    final_hdrs_prov.append(h)

# Redo Client in your Plannr matching
new_match_count = 0
for row in data_prov:
    inv_name = s(row.get(inv_col,""))
    if not inv_name:
        row["Client in your Plannr"] = None
        continue
    candidates = match_investor(inv_name)
    if len(candidates) == 1:
        row["Client in your Plannr"] = client_name(candidates[0])  # Rule 9: Excel only
        new_match_count += 1
    else:
        row["Client in your Plannr"] = None

p(f"Client in your Plannr matches: {new_match_count}")

# Get existing Action dropdown DV
wb_tmp = openpyxl.load_workbook(PLANS_PATH, data_only=True)
prov_dvs = list(wb_tmp[prov_sheet].data_validations.dataValidation)
wb_tmp.close()

# Rebuild DV for new Action column position
action_col_idx = final_hdrs_prov.index("action") + 1 \
                 if "action" in final_hdrs_prov else None
new_dvs = []
if action_col_idx and prov_dvs:
    orig_dv = prov_dvs[0]
    dv = DataValidation(
        type="list",
        formula1=orig_dv.formula1,
        allow_blank=True,
        showErrorMessage=False,
    )
    dv.sqref = f"{get_column_letter(action_col_idx)}2:{get_column_letter(action_col_idx)}{len(data_prov)+1}"
    new_dvs = [dv]

# Freeze after Investor name column
freeze_prov = max(
    (i+1 for i,h in enumerate(final_hdrs_prov)
     if h in ("Client in your Plannr", inv_col)),
    default=1,
)
rewrite_sheet(PLANS_PATH, prov_sheet, final_hdrs_prov, data_prov,
              freeze_prov, dvs=new_dvs)
p(f"Saved Provider tab ({len(data_prov)} rows)")
p(f"New headers: {final_hdrs_prov}")

# ── Final check ──────────────────────────────────────────────────────────────
p("\n=== Verification ===")
_, dep_final = read_sheet(DEPS_PATH, "Dependants")
p(f"Dependants: {len(dep_final)} rows")
c2_filled = sum(1 for r in dep_final if not blank(r.get("INFO: Client 2","")))
c1_filled = sum(1 for r in dep_final if not blank(r.get("INFO: Client 1","")))
p(f"  INFO: Client 1 filled: {c1_filled}")
p(f"  INFO: Client 2 filled: {c2_filled}  <- dependants with two clients")

_, prov_final = read_sheet(PLANS_PATH, prov_sheet)
p(f"Provider tab: {len(prov_final)} rows")
m = sum(1 for r in prov_final if not blank(r.get("Client in your Plannr","")))
p(f"  Client in your Plannr matched: {m}")
p("STOP")
