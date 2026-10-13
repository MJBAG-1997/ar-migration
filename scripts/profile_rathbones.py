"""
Step 1: Profile Jim list - Rathbones.xlsx and trial match vs Plannr plans.
Rule 9: aggregates, patterns only. No PII.
"""
import sys, os, re, json
from collections import Counter, defaultdict
import openpyxl
import pandas as pd

DATA_ROOT = r"C:\dev\AR Migration"
RATH_PATH = os.path.join(DATA_ROOT, "Jim list  - Rathbones.xlsx")
SNAP_DIR  = os.path.join(DATA_ROOT, "snapshot_20261012T090906Z")

def p(*a): print(*a, flush=True)

def mask_pattern(val):
    if val is None:
        return "null"
    s = str(val)
    num_d = sum(1 for c in s if c.isdigit())
    num_a = sum(1 for c in s if c.isalpha())
    return f"len={len(s)} d={num_d} a={num_a}"

def normalise_polno(s):
    if s is None:
        return None
    s = str(s).strip()
    s = re.sub(r'[\s\-]', '', s)
    s = s.lstrip('0')
    return s.upper() if s else None

# ─── profile ──────────────────────────────────────────────────────────────────
p("=== Rathbones file profile ===")

# Use openpyxl to detect header row and see raw values
wb = openpyxl.load_workbook(RATH_PATH, read_only=True, data_only=True)
sheets = wb.sheetnames
p(f"Sheets: {sheets}")
p()

for sname in sheets:
    ws = wb[sname]
    p(f"--- Sheet: '{sname}' ---")
    p(f"  max_row={ws.max_row}, max_col={ws.max_column}")

    # Find header row: scan first 10 rows for row with most non-null cells
    best_row = 1
    best_count = 0
    all_rows_sample = []
    for r in range(1, min(11, ws.max_row + 1)):
        row_vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        nonnull = sum(1 for v in row_vals if v is not None)
        all_rows_sample.append((r, row_vals, nonnull))
        if nonnull > best_count:
            best_count = nonnull
            best_row = r

    # Show first few rows
    p("  First 5 rows (raw):")
    for (r, rv, nn) in all_rows_sample[:5]:
        p(f"    row {r} ({nn} non-null): {rv}")
    p()

    # Use first row with >= 3 non-null as header
    header_row_idx = next((r for r, rv, nn in all_rows_sample if nn >= 3), 1)
    headers = all_rows_sample[header_row_idx - 1][1]
    p(f"  Header row: {header_row_idx}")
    p(f"  Column names: {[h for h in headers if h is not None]}")

    # Load all data using pandas (faster)
    df = pd.read_excel(RATH_PATH, sheet_name=sname, header=header_row_idx - 1, dtype=str)
    p(f"  Data rows (pandas): {len(df)}")
    p(f"  Columns: {list(df.columns)}")
    p()

    # Per-column analysis
    for col in df.columns:
        vals = df[col].dropna().tolist()
        fill_pct = len(vals) / len(df) if len(df) else 0

        # Detect if likely stored as number (check original wb)
        # We read as str so check if all values look numeric
        all_numeric = all(re.match(r'^-?\d+(\.\d+)?$', str(v).strip()) for v in vals) if vals else False
        stored_as_num_warn = all_numeric and any(len(str(v)) >= 10 for v in vals)

        # Type inference
        if all_numeric:
            inferred = "numeric"
        else:
            inferred = "text"

        # Pattern analysis (masked)
        pats = Counter(mask_pattern(v) for v in vals[:200])
        p(f"  Col '{col}':")
        p(f"    fill={fill_pct:.0%}  inferred={inferred}  stored_as_number_risk={stored_as_num_warn}")
        p(f"    top patterns: {pats.most_common(3)}")
        if stored_as_num_warn:
            p(f"    WARNING: large numeric values - may have lost leading zeros/precision")
        p()

    # Special check: column E (index 4) for SIPP
    if len(df.columns) > 4:
        col_e = df.columns[4]
        col_d = df.columns[3] if len(df.columns) > 3 else None
        e_vals = df[col_e].dropna().tolist()
        sipp_rows = [v for v in e_vals if "sipp" in str(v).lower()]
        p(f"  Column E ('{col_e}'): {len(e_vals)} non-null")
        p(f"    unique values: {Counter(e_vals).most_common(10)}")
        p(f"    'SIPP' rows: {len(sipp_rows)}")
        if col_d and sipp_rows:
            # Show Col D values for SIPP rows (masked - will show structure not values)
            sipp_mask = df[col_e].str.lower().str.contains("sipp", na=False)
            d_vals = df[col_d][sipp_mask].dropna().tolist()
            d_pats = Counter(mask_pattern(v) for v in d_vals)
            p(f"    Col D ('{col_d}') pattern for SIPP rows: {d_pats.most_common(3)}")
    p()

wb.close()

# ─── Trial match vs Plannr ────────────────────────────────────────────────────
p("=== Trial match Rathbones vs Plannr plans ===")

with open(os.path.join(SNAP_DIR, "plans_dedup.json")) as f:
    plans_raw = json.load(f)

if isinstance(plans_raw, dict):
    plans = list(plans_raw.values())
else:
    plans = plans_raw

p(f"Plannr plans total: {len(plans)}")

# Build normalised policy number index
plannr_norm = defaultdict(list)
for pl in plans:
    pn = pl.get("policy_number")
    n = normalise_polno(pn)
    if n:
        plannr_norm[n].append(pl.get("uuid"))

p(f"Plannr plans with policy_number: {sum(1 for pl in plans if pl.get('policy_number'))}")
p(f"Plannr unique normalised policy numbers: {len(plannr_norm)}")
p()

# Load Rathbones using pandas
for sname in pd.ExcelFile(RATH_PATH).sheet_names:
    p(f"--- Matching sheet: '{sname}' ---")
    df = pd.read_excel(RATH_PATH, sheet_name=sname, dtype=str)
    p(f"  Rathbones rows: {len(df)}")
    p(f"  Columns: {list(df.columns)}")

    # Try each column as policy number candidate
    best_col = None
    best_match = 0
    results = {}
    for col in df.columns:
        vals = df[col].dropna().tolist()
        if not vals:
            continue
        norms = [normalise_polno(v) for v in vals]
        norms_set = set(n for n in norms if n)
        if not norms_set:
            continue
        matched = norms_set & set(plannr_norm.keys())
        if matched:
            p(f"  Column '{col}': {len(vals)} values, {len(norms_set)} unique normalised, {len(matched)} match Plannr")
            results[col] = {"total": len(vals), "unique": len(norms_set), "matched": len(matched)}
            if len(matched) > best_match:
                best_match = len(matched)
                best_col = col

    if best_col:
        p(f"  Best policy number column: '{best_col}'")
        # Detailed match report
        rath_norms = set(normalise_polno(v) for v in df[best_col].dropna() if normalise_polno(v))
        matched   = rath_norms & set(plannr_norm.keys())
        rath_only = rath_norms - set(plannr_norm.keys())
        # Plannr-only: Plannr plans whose policy_number normalised is NOT in Rathbones
        # but we need to know which Plannr plans are Rathbones-provider plans
        # For now just show counts
        p(f"\n  MATCH REPORT using '{best_col}':")
        p(f"    Rathbones unique policy numbers: {len(rath_norms)}")
        p(f"    Matched (in both):               {len(matched)}")
        p(f"    Rathbones-only:                  {len(rath_only)}")
        p(f"    (Plannr-only needs provider filter - shown separately)")
    else:
        p("  No column produced matches - check if policy numbers are stored as numbers")
    p()

p("Done.")
