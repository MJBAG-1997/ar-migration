"""
Step 0: inspect hfp_signoff_pack_v2.xlsx for which v2 fixes are present.
Step 1 (profile only): profile Jim list - Rathbones.xlsx and propose column mapping.
Rule 9: aggregates only to console. No PII printed.
"""
import sys
import os
import json
import re
from collections import Counter, defaultdict

import openpyxl

DATA_ROOT = r"C:\dev\AR Migration"
V2_PATH   = os.path.join(DATA_ROOT, "hfp_signoff_pack_v2.xlsx")
RATH_PATH = os.path.join(DATA_ROOT, "Jim list  - Rathbones.xlsx")
SNAP_DIR  = os.path.join(DATA_ROOT, "snapshot_20261012T090906Z")

# ─── helpers ──────────────────────────────────────────────────────────────────

def sheet_names(wb):
    return wb.sheetnames

def col_headers(ws):
    """Return list of cell values in row 1."""
    return [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]

def data_rows(ws):
    """Return all non-empty rows below row 1 as lists."""
    rows = []
    for r in range(2, ws.max_row + 1):
        row = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        if any(v is not None for v in row):
            rows.append(row)
    return rows

def mask_pattern(val):
    if val is None:
        return "null"
    s = str(val)
    p = re.sub(r'[A-Za-z]', 'a', re.sub(r'[0-9]', 'd', s))
    return f"len={len(s)} pat={p[:30]}"

# ─── Step 0: v2 inspection ────────────────────────────────────────────────────

print("=" * 70)
print("STEP 0: v2 pack inspection")
print("=" * 70)

wb2 = openpyxl.load_workbook(V2_PATH, read_only=True, data_only=True)
tabs = sheet_names(wb2)
print(f"Tabs ({len(tabs)}): {tabs}")
print()

# Check 1: Exceptions tab – does it have detail rows?
exc_check = "MISSING"
if "Exceptions" in tabs:
    ws = wb2["Exceptions"]
    hdrs = col_headers(ws)
    rows = data_rows(ws)
    print(f"Exceptions tab: headers={hdrs}")
    print(f"  data rows: {len(rows)}")
    if len(rows) > 0:
        exc_check = f"PRESENT ({len(rows)} rows)"
    else:
        exc_check = "PRESENT but empty"
else:
    print("Exceptions tab: NOT FOUND")
print()

# Check 2: no-status rules – look for a 'no_status_rules' or similar tab, or a
#          column named 'no_status' in Exceptions or Plans
no_status_check = "MISSING"
for tab in tabs:
    ws = wb2[tab]
    hdrs = [str(h).lower() if h else "" for h in col_headers(ws)]
    if any("no_status" in h or "nostatus" in h or "status_rule" in h for h in hdrs):
        no_status_check = f"PRESENT in tab '{tab}'"
        break
# Also check if Plans tab has a 'basis' column referencing no-status
if "Plans" in tabs:
    ws = wb2["Plans"]
    hdrs = col_headers(ws)
    rows = data_rows(ws)
    basis_col = next((i for i,h in enumerate(hdrs) if h and "basis" in str(h).lower()), None)
    if basis_col is not None:
        basis_vals = [r[basis_col] for r in rows if r[basis_col]]
        basis_ctr = Counter(basis_vals)
        ns_count = sum(v for k,v in basis_ctr.items() if k and "no" in str(k).lower() and "status" in str(k).lower())
        if ns_count > 0:
            no_status_check = f"PRESENT in Plans.basis ({ns_count} rows)"
        print(f"Plans.basis values: {dict(basis_ctr.most_common(10))}")
print(f"no-status rules: {no_status_check}")
print()

# Check 3: indicative annual income – look in Summary or Charges sign-off
iai_check = "MISSING"
for tab in tabs:
    ws = wb2[tab]
    hdrs = col_headers(ws)
    if any(h and "indicative" in str(h).lower() for h in hdrs):
        iai_check = f"PRESENT in tab '{tab}'"
        break
    rows = data_rows(ws)
    for row in rows:
        if any(v and "indicative" in str(v).lower() for v in row):
            iai_check = f"PRESENT (cell value) in tab '{tab}'"
            break
    if iai_check != "MISSING":
        break
print(f"Indicative annual income: {iai_check}")
print()

# Check 4: Charges reconciliation items
#   a) the 42 excluded
#   b) one ongoing charge per plan
#   c) Charges sign-off sections A, B, C
charges_signoff_tab = None
for tab in tabs:
    if "charge" in tab.lower() and "sign" in tab.lower():
        charges_signoff_tab = tab
        break
    if "charge" in tab.lower():
        charges_signoff_tab = tab  # fallback

excl_42_check = "MISSING"
one_ongoing_check = "MISSING"
sections_abc_check = "MISSING"

if charges_signoff_tab:
    ws = wb2[charges_signoff_tab]
    hdrs = col_headers(ws)
    rows = data_rows(ws)
    print(f"Charges tab '{charges_signoff_tab}': headers={hdrs}")
    print(f"  data rows: {len(rows)}")
    # look for section markers
    all_text = []
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            if v:
                all_text.append(str(v).lower())
    text_blob = " ".join(all_text)
    if "42" in text_blob or "exclud" in text_blob:
        excl_42_check = "PRESENT"
    if "section a" in text_blob:
        sections_abc_check = "section A found"
    if "section b" in text_blob:
        sections_abc_check += " + B"
    if "section c" in text_blob:
        sections_abc_check += " + C"
    if sections_abc_check == "MISSING":
        # Look for any section labelling in first column
        col1 = [ws.cell(r, 1).value for r in range(1, ws.max_row + 1)]
        for v in col1:
            if v and re.search(r'section\s+[abc]', str(v), re.I):
                sections_abc_check = f"PRESENT: {v}"
                break
    # one ongoing per plan: check if there's a dedup note
    if "one ongoing" in text_blob or "ongoing charge per plan" in text_blob:
        one_ongoing_check = "PRESENT"

    print(f"42 excluded: {excl_42_check}")
    print(f"One ongoing charge per plan: {one_ongoing_check}")
    print(f"Charges sign-off sections A/B/C: {sections_abc_check}")
else:
    print("No Charges sign-off tab found")
print()

# Show all tab names and row counts
print("All tabs in v2:")
for tab in tabs:
    ws = wb2[tab]
    print(f"  {tab}: max_row={ws.max_row}, max_col={ws.max_column}")

wb2.close()
print()

# ─── Step 1: Rathbones profile ────────────────────────────────────────────────

print("=" * 70)
print("STEP 1: Rathbones file profile")
print("=" * 70)

wbr = openpyxl.load_workbook(RATH_PATH, read_only=True, data_only=True)
r_tabs = sheet_names(wbr)
print(f"Sheets: {r_tabs}")
print()

for sheet_name in r_tabs:
    ws = wbr[sheet_name]
    print(f"--- Sheet: '{sheet_name}' ---")

    # Find header row (first row with >= 3 non-null cells)
    header_row_idx = None
    for r in range(1, min(10, ws.max_row + 1)):
        row_vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        non_null = sum(1 for v in row_vals if v is not None)
        if non_null >= 3:
            header_row_idx = r
            break
    if header_row_idx is None:
        print("  Could not find header row")
        continue

    headers = [ws.cell(header_row_idx, c).value for c in range(1, ws.max_column + 1)]
    print(f"  Header row: {header_row_idx}")
    print(f"  Columns ({len([h for h in headers if h is not None])}): {[h for h in headers if h is not None]}")

    # Count data rows
    data_row_count = 0
    col_data = defaultdict(list)
    for r in range(header_row_idx + 1, ws.max_row + 1):
        row = [ws.cell(r, c).value for c in range(1, len(headers) + 1)]
        if any(v is not None for v in row):
            data_row_count += 1
            for ci, v in enumerate(row):
                if v is not None:
                    col_data[ci].append(v)

    print(f"  Data rows: {data_row_count}")
    print()

    # Per-column profile
    for ci, hdr in enumerate(headers):
        if hdr is None:
            continue
        vals = col_data.get(ci, [])
        fill_rate = len(vals) / data_row_count if data_row_count else 0
        stored_as_number = any(isinstance(v, (int, float)) for v in vals)
        # Infer type
        types = set()
        for v in vals:
            if isinstance(v, (int, float)):
                types.add("numeric")
            elif isinstance(v, str):
                types.add("text")
            else:
                types.add(str(type(v).__name__))
        inferred = "numeric" if types == {"numeric"} else ("text" if "text" in types else str(types))

        # Pattern analysis (masked - no PII)
        patterns = Counter()
        for v in vals[:500]:
            s = str(v)
            num_d = sum(1 for c in s if c.isdigit())
            num_a = sum(1 for c in s if c.isalpha())
            patterns[f"len={len(s)} d={num_d} a={num_a}"] += 1
        top3 = patterns.most_common(3)

        print(f"  Col {ci+1}: '{hdr}'")
        print(f"    fill={fill_rate:.0%}  type={inferred}  stored_as_number={stored_as_number}")
        print(f"    top patterns: {top3}")
        if stored_as_number and inferred == "text":
            print(f"    WARNING: Excel stored as number - leading zeros likely lost")
        print()

wbr.close()

# ─── Step 1b: Trial match Rathbones vs Plannr plans ──────────────────────────

print("=" * 70)
print("STEP 1b: Trial match Rathbones vs Plannr plans (normalised policy number)")
print("=" * 70)

import json

with open(os.path.join(SNAP_DIR, "plans_dedup.json")) as f:
    plans = json.load(f)

def normalise_polno(s):
    if s is None:
        return None
    s = str(s).strip()
    # strip spaces, hyphens, leading zeros
    s = re.sub(r'[\s\-]', '', s)
    s = s.lstrip('0')
    return s.upper() if s else None

# Build Plannr normalised policy numbers
plannr_norm = {}
for p in plans:
    pn = p.get("policy_number")
    n = normalise_polno(pn)
    if n:
        if n not in plannr_norm:
            plannr_norm[n] = []
        plannr_norm[n].append(p.get("uuid"))

print(f"Plannr plans total: {len(plans)}")
print(f"Plannr plans with policy_number: {sum(1 for p in plans if p.get('policy_number'))}")
print(f"Plannr unique normalised policy numbers: {len(plannr_norm)}")
print()

# Now load Rathbones and attempt match
# Need to identify which column holds the policy number
# We'll profile each sheet and try to match
wbr2 = openpyxl.load_workbook(RATH_PATH, read_only=True, data_only=True)
for sheet_name in sheet_names(wbr2):
    ws = wbr2[sheet_name]
    print(f"--- Matching sheet: '{sheet_name}' ---")

    # Find header row
    header_row_idx = None
    for r in range(1, min(10, ws.max_row + 1)):
        row_vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        non_null = sum(1 for v in row_vals if v is not None)
        if non_null >= 3:
            header_row_idx = r
            break
    if not header_row_idx:
        continue

    headers = [ws.cell(header_row_idx, c).value for c in range(1, ws.max_column + 1)]

    # Load all data rows
    rathbones_rows = []
    for r in range(header_row_idx + 1, ws.max_row + 1):
        row = {headers[ci]: ws.cell(r, ci+1).value
               for ci in range(len(headers)) if headers[ci] is not None}
        if any(v is not None for v in row.values()):
            rathbones_rows.append(row)

    print(f"  Rathbones data rows: {len(rathbones_rows)}")

    # Try each column as policy number candidate
    for col in headers:
        if col is None:
            continue
        vals = [str(r.get(col, "")).strip() for r in rathbones_rows if r.get(col)]
        if not vals:
            continue
        norms = [normalise_polno(v) for v in vals]
        norms_set = set(n for n in norms if n)
        matched = norms_set & set(plannr_norm.keys())
        if matched:
            print(f"  Column '{col}': {len(vals)} values, {len(norms_set)} unique, {len(matched)} match Plannr")

    # Also check column E (index 4) for 'sipp' keyword
    col_e_vals = []
    col_d_vals = []
    for row in rathbones_rows:
        col_keys = list(row.keys())
        if len(col_keys) > 4:
            col_e_vals.append(row.get(col_keys[4]))
        if len(col_keys) > 3:
            col_d_vals.append(row.get(col_keys[3]))
    sipp_count = sum(1 for v in col_e_vals if v and "sipp" in str(v).lower())
    print(f"  Col E (index 4) values: {Counter(str(v).lower() for v in col_e_vals if v).most_common(5)}")
    print(f"  Col E 'sipp' rows: {sipp_count}")
    print()

wbr2.close()
