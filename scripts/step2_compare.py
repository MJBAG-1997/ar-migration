"""
HFP AR Migration - Step 2: Compare export columns to importer keys.

Reads exports from DATA_ROOT/exports_YYYYMMDD/ (uses most recent if no date given).
For each export:
  - lists columns
  - shows which match importer keys exactly
  - shows which differ by name only (case / space / underscore)
  - shows which importer keys are missing from export

Reconciles row counts to the deduplicated snapshot targets.

Rule 9: no PII in output; column names and counts only.
"""
import os, sys, json, re, csv, glob
from collections import defaultdict

DATA_ROOT = r"C:\dev\AR Migration"
SNAP_DIR  = os.path.join(DATA_ROOT, "snapshot_20261012T090906Z")
DEST_JSON = os.path.join(DATA_ROOT, "block6_destination.json")

# Expected counts from the deduplicated snapshot
EXPECTED = {
    "client_export":                  375,
    "plan_export":                    1864,
    "charges_export":                 372,
    "note_export":                    6792,
    "circle_export":                  98,    # (92 have members)
    "valuation_report":               1528,
    "premium_export":                 None,  # not in v6 counts
    "income_and_employment_export":   316,
    "client_expenditure_export":      2381,
    "staff_member_export":            None,
    "group_scheme_members_export":    None,
    "dependants":                     367,
}

# Map report_type_id → importer key name in block6_destination.json
IMPORTER_MAP = {
    "client_export":                 "clients",
    "plan_export":                   "plans",
    "charges_export":                "charges",
    "circle_export":                 "circles",
    "valuation_report":              "valuations",
    "premium_export":                "premiums",
    "income_and_employment_export":  "income_and_employment",
    "client_expenditure_export":     "expenses",
    "note_export":                   "notes",
    "staff_member_export":           "business_staff_members",
    "group_scheme_members_export":   "group_scheme_members",
    "dependants":                    "dependants",
}

def p(*a): print(*a, flush=True)


def normalise(name: str) -> str:
    """Collapse spaces, hyphens, and underscores; lowercase."""
    return re.sub(r"[\s\-_]+", "_", str(name).strip()).lower()


def load_importer_keys(dest_json_path: str) -> dict:
    with open(dest_json_path, encoding="utf-8") as fh:
        data = json.load(fh)
    keys_by_importer = {}
    for name, spec in data.get("importer_keys", {}).items():
        keys_by_importer[name] = spec.get("store_keys", [])
    return keys_by_importer


def find_exports_dir(data_root: str) -> str:
    """Return most recent exports_YYYYMMDD directory."""
    candidates = sorted(
        d for d in glob.glob(os.path.join(data_root, "exports_*"))
        if os.path.isdir(d)
    )
    if not candidates:
        raise FileNotFoundError(
            f"No exports_* directory found under {data_root}. Run step1_exports.py first."
        )
    return candidates[-1]


def read_csv_columns_and_rows(path: str):
    """Return (columns_list, row_count) for a CSV file."""
    with open(path, encoding="utf-8-sig", newline="") as fh:
        reader = csv.reader(fh)
        try:
            header = next(reader)
        except StopIteration:
            return [], 0
        # Count remaining rows (data rows)
        count = sum(1 for _ in reader)
    return header, count


def read_excel_columns_and_rows(path: str):
    """Return (columns_list, row_count) for an Excel file (first sheet)."""
    try:
        import openpyxl
    except ImportError:
        return None, None
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    # First non-empty row as header
    header = []
    row_count = 0
    first = True
    for row in ws.iter_rows(values_only=True):
        if first:
            header = [str(c) if c is not None else "" for c in row]
            first = False
        else:
            if any(c is not None for c in row):
                row_count += 1
    wb.close()
    return header, row_count


def read_json_columns_and_rows(path: str):
    """Return (columns_list, row_count) for a JSON array file."""
    with open(path, encoding="utf-8") as fh:
        data = json.load(fh)
    if not isinstance(data, list) or not data:
        return [], len(data) if isinstance(data, list) else 0
    cols = list(data[0].keys()) if isinstance(data[0], dict) else []
    return cols, len(data)


def load_export(path: str):
    """Auto-detect format and return (columns, row_count)."""
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xls"):
        return read_excel_columns_and_rows(path)
    elif ext == ".json":
        return read_json_columns_and_rows(path)
    else:
        # default: try CSV
        return read_csv_columns_and_rows(path)


def compare_columns(export_cols: list, importer_keys: list):
    """
    Returns:
      exact   - cols present in export that exactly match an importer key
      renamed - cols present in export that match after normalisation
      export_extra - cols in export not matching any importer key
      missing - importer keys with no match in export (exact or normalised)
    """
    imp_set     = set(importer_keys)
    imp_norm    = {normalise(k): k for k in importer_keys}

    exact         = []
    renamed       = []
    export_extra  = []

    matched_importer = set()

    for col in export_cols:
        if col in imp_set:
            exact.append(col)
            matched_importer.add(col)
        elif normalise(col) in imp_norm:
            renamed.append((col, imp_norm[normalise(col)]))
            matched_importer.add(imp_norm[normalise(col)])
        else:
            export_extra.append(col)

    missing = [k for k in importer_keys if k not in matched_importer]

    return exact, renamed, export_extra, missing


def main():
    if len(sys.argv) > 1:
        exports_dir = sys.argv[1]
    else:
        exports_dir = find_exports_dir(DATA_ROOT)

    p("=" * 68)
    p("HFP AR Migration — Step 2: Column comparison and row reconciliation")
    p("=" * 68)
    p(f"Exports dir: {exports_dir}")

    # Load manifest
    manifest_path = os.path.join(exports_dir, "_manifest.json")
    if os.path.exists(manifest_path):
        with open(manifest_path, encoding="utf-8") as fh:
            manifest = json.load(fh)
        p(f"Pulled at:   {manifest.get('pulled_at', 'unknown')}")
        results = manifest.get("results", {})
    else:
        p("No _manifest.json found; scanning directory.")
        results = {}

    # Load importer keys
    importer_keys = load_importer_keys(DEST_JSON)

    # Build label → path from manifest results (which have actual paths)
    export_files = {}
    for label, v in results.items():
        if v.get("status") == "ok" and v.get("path"):
            export_files[label] = v["path"]

    # Also scan dir for any files not in manifest (fallback)
    for f in os.listdir(exports_dir):
        if f.startswith("_"):
            continue
        fpath = os.path.join(exports_dir, f)
        # Check if any label matches
        found = False
        for lbl in IMPORTER_MAP:
            key_part = lbl.replace("_", " ")
            if key_part.lower() in f.lower() and lbl not in export_files:
                export_files[lbl] = fpath
                found = True
                break

    all_labels = set(IMPORTER_MAP.keys())

    p("\n" + "-" * 68)
    p(f"{'Export':<42}  {'Rows':>6}  {'Expect':>6}  {'Match?'}")
    p("-" * 68)

    summaries = {}

    for label in sorted(all_labels):
        imp_name = IMPORTER_MAP.get(label)
        keys     = importer_keys.get(imp_name, []) if imp_name else []
        expected = EXPECTED.get(label)

        # Find file
        path = export_files.get(label)
        if path is None:
            # Try with date suffix or similar name in manifest
            for base, fp in export_files.items():
                if label in base:
                    path = fp
                    break

        if path is None or not os.path.exists(path):
            p(f"  {label:<42}  {'--':>6}  {str(expected or '--'):>6}  NO FILE")
            summaries[label] = {"status": "missing"}
            continue

        cols, rows = load_export(path)
        if cols is None:
            p(f"  {label:<42}  {'ERR':>6}  {str(expected or '--'):>6}  READ ERROR")
            summaries[label] = {"status": "error"}
            continue

        match_sym = ""
        if expected is not None:
            match_sym = "=" if rows == expected else f"delta={rows - expected:+d}"

        p(f"  {label:<42}  {rows:>6,}  {str(expected or '--'):>6}  {match_sym}")

        if keys:
            exact, renamed, extra, missing = compare_columns(cols, keys)
        else:
            exact, renamed, extra, missing = cols, [], [], []

        summaries[label] = {
            "status":   "ok",
            "rows":     rows,
            "expected": expected,
            "cols_total":    len(cols),
            "exact_match":   len(exact),
            "renamed_match": len(renamed),
            "export_extra":  len(extra),
            "missing_from_export": len(missing),
            "exact":   exact,
            "renamed": renamed,
            "extra":   extra,
            "missing": missing,
        }

    # Detailed breakdown per export
    p("\n" + "=" * 68)
    p("DETAILED COLUMN ANALYSIS")
    p("=" * 68)

    for label in sorted(all_labels):
        s = summaries.get(label, {})
        if s.get("status") != "ok":
            continue

        imp_name = IMPORTER_MAP.get(label, "(no mapping)")
        p(f"\n{'-' * 68}")
        p(f"  {label}  ->  importer: {imp_name}")
        p(f"  Columns in export: {s['cols_total']}  "
          f"Exact: {s['exact_match']}  "
          f"Renamed: {s['renamed_match']}  "
          f"Export-only: {s['export_extra']}  "
          f"Missing from export: {s['missing_from_export']}")

        if s["renamed"]:
            p(f"\n  Differ by name only (export col → importer key):")
            for exp_col, imp_key in s["renamed"]:
                p(f"    {exp_col!r:40s} → {imp_key!r}")

        if s["extra"]:
            p(f"\n  Export columns not in importer spec:")
            for col in s["extra"]:
                p(f"    {col!r}")

        if s["missing"]:
            p(f"\n  Importer keys missing from export:")
            for key in s["missing"]:
                p(f"    {key!r}")

    # Row count reconciliation summary
    p("\n" + "=" * 68)
    p("ROW COUNT RECONCILIATION  (target: deduplicated 12 Oct snapshot)")
    p("=" * 68)
    p(f"{'Export':<42}  {'Export':>7}  {'Target':>7}  {'Delta':>7}")
    p("-" * 68)
    for label in sorted(all_labels):
        s = summaries.get(label, {})
        rows = s.get("rows")
        exp  = s.get("expected")
        if rows is None:
            p(f"  {label:<42}  {'--':>7}  {str(exp or '--'):>7}  {'--':>7}")
        elif exp is None:
            p(f"  {label:<42}  {rows:>7,}  {'--':>7}  {'--':>7}")
        else:
            delta = rows - exp
            marker = "" if delta == 0 else " <<<"
            p(f"  {label:<42}  {rows:>7,}  {exp:>7,}  {delta:>+7,}{marker}")

    p("\nStep 2 complete. Review and confirm before Step 3.")
    p("STOP")


if __name__ == "__main__":
    main()
