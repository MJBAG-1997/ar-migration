"""
higgs_remap.py  —  Higgs AR migration: reference-swap and round-assembly

Each upload round replaces source-system UUIDs with real Plannr UUIDs taken
from Plannr exports, then writes clean CSV + xlsx files ready to upload.

Usage:
    python higgs_remap.py r1
    python higgs_remap.py r2
    python higgs_remap.py r3
    python higgs_remap.py r4

Rounds:
    r1  01 Clients (individuals)
            Sets assigned_adviser_uuid, send_welcome_email=No,
            client_regulatory_classification=Retail Client.
            No reference swap needed (clients have no cross-refs yet).
    r2  02 Clients BT (businesses and trusts)
            Reads client export dropped after R1.
            Swaps 'owners' source person UUID -> Plannr UUID.
            Trusts with blank owners are held back and listed.
    r3  03 Plans + 06 I&E + 07 Expenses + 08 Notes + 09 Circles
            Reads combined client export (after R1+R2 upload).
            Swaps client_uuids / client_uuid / account_uuid.
    r4  04 Valuations + 05 Charges
            Reads client export + plan export dropped after R3.
            Valuations: policy_number already set, no UUID swap needed.
            Charges: populates plan_uuid and client_uuid from review-pack
            helper columns (x_plan_ref, x_client_refs).

NOT included: 04 Premiums — these need plan_uuid which is not recoverable
from the current importer CSV; link premiums manually in Plannr after R3,
or request a follow-up mapping step.

Reads:
    DATA_ROOT/stage2_build/importer/       clean importer CSVs from build step
    DATA_ROOT/stage2_build/higgs_Merge_Review.xlsx      helper columns for charges
    DATA_ROOT/stage2_build/higgs_Merge_Review_BT.xlsx   BT client data for R2
    DATA_ROOT/plannr_exports/              Plannr exports (client, then plan)

Writes:
    DATA_ROOT/upload/R<n>/  xlsx + csv per record type, ready to upload

Security rule: no client names in terminal output or files — counts and refs only.
"""

import sys
import os
import re
import glob

try:
    import pandas as pd
    from openpyxl import Workbook, load_workbook
except ImportError:
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "--quiet", "pandas", "openpyxl"])
    import pandas as pd
    from openpyxl import Workbook, load_workbook

# ---------------------------------------------------------------------------
DATA_ROOT    = os.path.dirname(os.path.abspath(__file__))
IMPORTER_DIR = os.path.join(DATA_ROOT, "stage2_build", "importer")
REVIEW_PACK  = os.path.join(DATA_ROOT, "stage2_build", "higgs_Merge_Review.xlsx")
REVIEW_BT    = os.path.join(DATA_ROOT, "stage2_build", "higgs_Merge_Review_BT.xlsx")
EXPORTS_DIR  = os.path.join(DATA_ROOT, "plannr_exports")
UPLOAD_DIR   = os.path.join(DATA_ROOT, "upload")
HELD_PATH    = os.path.join(UPLOAD_DIR, "held_back_trusts.txt")

DEST_ADVISER_UUID = "8dff5edc-2830-42d2-8a23-38ea8d6006bb"


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _stop(msg):
    print(f"\nSTOP: {msg}", file=sys.stderr)
    sys.exit(1)


def _read_csv(path):
    try:
        return pd.read_csv(path, dtype=str, keep_default_na=False, encoding="utf-8-sig")
    except UnicodeDecodeError:
        return pd.read_csv(path, dtype=str, keep_default_na=False, encoding="latin-1")


def _read_xlsx_sheet(path, sheet):
    try:
        return pd.read_excel(path, sheet_name=sheet, dtype=str, engine="openpyxl")
    except Exception as e:
        _stop(f"Cannot read sheet '{sheet}' from {os.path.relpath(path, DATA_ROOT)}: {e}")


def _write_round(rdir, name, df):
    """Write one file as csv + xlsx; return row count."""
    os.makedirs(rdir, exist_ok=True)
    df = df.fillna("").astype(str).replace("nan", "")
    csv_path  = os.path.join(rdir, f"{name}.csv")
    xlsx_path = os.path.join(rdir, f"{name}.xlsx")
    df.to_csv(csv_path, index=False, encoding="utf-8")
    wb = Workbook()
    ws = wb.active
    ws.title = name[:31]
    ws.append(list(df.columns))
    for tup in df.itertuples(index=False):
        ws.append(list(tup))
    wb.save(xlsx_path)
    print(f"  {name}: {len(df)} rows")
    return len(df)


def _get_col(df, col, default=""):
    if col in df.columns:
        return list(df[col].fillna(default).astype(str))
    return [default] * len(df)


_ALWAYS_EXCLUDE = ("supplement", "failed_import", "failed_imports")

def _latest_export(pattern_list, exclude_keyword=None):
    """Return the latest file matching any of the glob patterns."""
    files = []
    for pat in pattern_list:
        files.extend(glob.glob(pat))
    # Always exclude internal helper / error files
    files = [f for f in files
             if not any(kw in os.path.basename(f).lower() for kw in _ALWAYS_EXCLUDE)]
    if exclude_keyword:
        files = [f for f in files if exclude_keyword.lower() not in os.path.basename(f).lower()]
    files = list({os.path.abspath(f) for f in files})  # dedup
    if not files:
        return None
    return max(files, key=os.path.getmtime)


# ---------------------------------------------------------------------------
# Export loaders
# ---------------------------------------------------------------------------

UUID_RE = re.compile(
    r"^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$", re.I
)


def _detect_uuid_col(df):
    """Heuristic: find the column whose values look most like UUIDs."""
    best, best_n = None, 0
    for c in df.columns:
        n = sum(1 for v in df[c].dropna().head(20) if UUID_RE.match(str(v).strip()))
        if n > best_n:
            best, best_n = c, n
    return best if best_n >= 2 else None


def _load_export(kind):
    """Load the latest Plannr export of `kind` ('client' or 'plan').
    Returns dict: source_externalreference -> plannr_uuid.
    """
    if kind == "client":
        patterns = [os.path.join(EXPORTS_DIR, p) for p in
                    ("*[Cc]lient*.csv", "*[Pp]erson*.csv", "*[Mm]igration*.csv")]
        exclude = "plan"
    else:
        patterns = [os.path.join(EXPORTS_DIR, p) for p in
                    ("*[Pp]lan*.csv", "*[Pp]olic*.csv")]
        exclude = "client"

    path = _latest_export(patterns, exclude_keyword=exclude)
    if path is None:
        _stop(
            f"No {kind} export found in {EXPORTS_DIR}.\n"
            f"  Drop a Plannr {kind} export CSV there, then re-run."
        )

    rel = os.path.relpath(path, DATA_ROOT)
    print(f"  Loading {kind} export: {rel}")
    df = _read_csv(path)
    cols_lower = {c.strip().lower(): c for c in df.columns}

    if "externalreference" not in cols_lower:
        _stop(
            f"{kind.title()} export '{os.path.basename(path)}' has no 'externalreference' column.\n"
            f"  Columns found: {list(df.columns)}\n"
            "  Re-export from Plannr with the 'externalreference' custom field included."
        )

    ext_col = cols_lower["externalreference"]

    uuid_col = None
    for candidate in ("uuid", "id", f"plannr_uuid", f"{kind}_uuid", "plan_uuid", "client_uuid"):
        if candidate in cols_lower:
            uuid_col = cols_lower[candidate]
            break
    if uuid_col is None:
        uuid_col = _detect_uuid_col(df)
    if uuid_col is None:
        _stop(
            f"Cannot identify the Plannr UUID column in {kind} export.\n"
            f"  Columns: {list(df.columns)}"
        )

    mapping = {}
    for _, row in df.iterrows():
        ext = str(row[ext_col]).strip()
        uid = str(row[uuid_col]).strip()
        if ext and uid and ext.lower() not in ("nan", "", "none"):
            mapping[ext] = uid

    print(f"  {len(mapping)} {kind} UUID mappings loaded (column '{uuid_col}')")

    # Merge supplement file for clients that were refused as duplicates
    if kind == "client":
        supp_path = os.path.join(EXPORTS_DIR, "client_map_supplement.csv")
        if os.path.isfile(supp_path):
            df_s = _read_csv(supp_path)
            added = 0
            for _, row in df_s.iterrows():
                ext = str(row.get("externalreference", "")).strip()
                uid = str(row.get("uuid", "")).strip()
                if ext and uid and ext not in mapping:
                    mapping[ext] = uid
                    added += 1
            if added:
                print(f"  +{added} supplement mappings loaded (duplicate-refused clients)")

    return mapping


# ---------------------------------------------------------------------------
# Reference-swap helpers
# ---------------------------------------------------------------------------

def _swap(val, mapping):
    """Swap a single source ref. Returns (mapped, was_unresolved)."""
    v = str(val).strip()
    if not v or v.lower() in ("nan", "", "none"):
        return "", False
    mapped = mapping.get(v, "")
    return mapped, (mapped == "")


def _swap_col(df, col, mapping):
    """Swap a column of single source refs. Returns modified df + unresolved list."""
    new_vals, unresolved = [], []
    for v in df[col]:
        mapped, bad = _swap(v, mapping)
        new_vals.append(mapped)
        if bad and str(v).strip() not in ("", "nan"):
            unresolved.append(str(v).strip())
    df = df.copy()
    df[col] = new_vals
    return df, unresolved


def _swap_multi_col(df, col, mapping):
    """Swap a comma-separated column of source refs. Returns df + unresolved list."""
    new_vals, unresolved = [], []
    for v in df[col]:
        v = str(v).strip()
        if not v or v.lower() in ("nan", ""):
            new_vals.append("")
            continue
        parts = [p.strip() for p in v.split(",") if p.strip()]
        out = []
        for p in parts:
            mapped, bad = _swap(p, mapping)
            out.append(mapped)
            if bad:
                unresolved.append(p)
        new_vals.append(",".join(out))
    df = df.copy()
    df[col] = new_vals
    return df, unresolved


def _require_resolved(unresolved, label):
    if unresolved:
        unique = sorted(set(unresolved))
        _stop(
            f"{len(unresolved)} unresolved reference(s) in '{label}' "
            f"({len(unique)} unique).\n"
            f"  First 10: {unique[:10]}\n"
            "  Ensure the preceding upload rounds are complete and the export is current."
        )


# ---------------------------------------------------------------------------
# Held-back trust tracking
# ---------------------------------------------------------------------------

def _save_held_refs(refs):
    os.makedirs(UPLOAD_DIR, exist_ok=True)
    with open(HELD_PATH, "w") as f:
        f.write("\n".join(r for r in refs if r))
    print(f"  Held-back trust refs saved: {os.path.relpath(HELD_PATH, DATA_ROOT)}")


def _load_held_refs():
    if not os.path.isfile(HELD_PATH):
        return set()
    with open(HELD_PATH) as f:
        return {line.strip() for line in f if line.strip()}


# ---------------------------------------------------------------------------
# Round 1 — Individuals clients
# ---------------------------------------------------------------------------

def run_r1():
    print("\n=== R1: Individuals Clients ===")

    src = os.path.join(IMPORTER_DIR, "01 Clients.csv")
    if not os.path.isfile(src):
        _stop(f"Source file not found: {src}\n  Run the build step first.")
    df = _read_csv(src)
    print(f"  Source: {len(df)} rows")

    # Set firm-side fields
    df["assigned_adviser_uuid"] = DEST_ADVISER_UUID
    df["send_welcome_email"] = "No"
    df["client_regulatory_classification"] = "Retail Client"
    if "is_prospect" in df.columns:
        df["is_prospect"] = df["is_prospect"].replace({"": "no"})

    rdir = os.path.join(UPLOAD_DIR, "R1")
    _write_round(rdir, "01 Clients", df)

    print(
        f"\nR1 done. Upload 01 Clients.csv / xlsx to Plannr.\n"
        f"Then drop a Plannr client export into:\n"
        f"  {EXPORTS_DIR}\n"
        f"and run: python higgs_remap.py r2"
    )


# ---------------------------------------------------------------------------
# Round 2 — Business and trust clients
# ---------------------------------------------------------------------------

def run_r2():
    print("\n=== R2: Business / Trust Clients ===")

    # Load BT source — importer CSV preferred, review pack as fallback
    bt_csv = os.path.join(IMPORTER_DIR, "01 Clients BT.csv")
    if os.path.isfile(bt_csv):
        df = _read_csv(bt_csv)
        print(f"  Source (CSV): {len(df)} rows")
    elif os.path.isfile(REVIEW_BT):
        print(f"  BT importer CSV not found; reading Clients sheet from review pack.")
        df_raw = _read_xlsx_sheet(REVIEW_BT, "Clients")
        # Keep template columns + externalreference, drop x_/helper cols
        keep = [c for c in df_raw.columns
                if not str(c).startswith("x_")
                and not str(c).startswith("Unnamed")
                and str(c) not in ("_miss",)]
        df = df_raw[keep].copy()
        if "externalreference" in df_raw.columns and "externalreference" not in df.columns:
            df["externalreference"] = df_raw["externalreference"]
        # Drop descriptor/blank rows
        if "type" in df.columns:
            df = df[df["type"].isin(["business", "trust"])].reset_index(drop=True)
        print(f"  Source (review pack): {len(df)} rows")
    else:
        _stop(
            f"No BT client source found.\n"
            f"  Looked for: {bt_csv}\n"
            f"           or: {REVIEW_BT}"
        )

    # Load client export (needed to swap owners)
    client_map = _load_export("client")

    # Set firm-side fields for all BT clients
    df["assigned_adviser_uuid"] = DEST_ADVISER_UUID
    df["send_welcome_email"] = "No"
    df["client_regulatory_classification"] = "Retail Client"

    # Identify trusts with no owner
    is_trust = df.get("type", pd.Series(dtype=str)).fillna("") == "trust"
    owners_col = df.get("owners", pd.Series([""] * len(df)))
    no_owner = is_trust & (owners_col.fillna("").str.strip() == "")

    held_df = df[no_owner].copy()
    upload_df = df[~no_owner].copy()

    if len(held_df) > 0:
        held_ext = list(held_df.get("externalreference", pd.Series(dtype=str)).fillna(""))
        print(f"  HELD BACK: {len(held_df)} trust(s) with no owner")
        print(f"  Their externalrefs: {held_ext}")
        _save_held_refs(held_ext)
        rdir = os.path.join(UPLOAD_DIR, "R2")
        _write_round(rdir, "02 Clients_HELD_BACK_TRUSTS", held_df)

    # Swap owners source UUID -> Plannr UUID
    # Owners outside the migration scope can't be resolved; Plannr requires an owner
    # for every BT client, so any row with an unresolvable owner is held back.
    if "owners" in upload_df.columns and len(upload_df) > 0:
        upload_df, unres = _swap_col(upload_df, "owners", client_map)
        if unres:
            unique_unres = sorted(set(unres))
            print(f"  WARNING: {len(unres)} owner ref(s) unresolvable (outside migration scope) — rows held back")
            print(f"  Blanked owner refs: {unique_unres}")

    rdir = os.path.join(UPLOAD_DIR, "R2")

    # Hold back any row whose owners column is blank after swapping
    # (Plannr requires an owner; empty string fails validation)
    has_owner_after_swap = upload_df.get("owners", pd.Series([""] * len(upload_df))).str.strip() != ""
    extra_held = upload_df[~has_owner_after_swap].copy()
    upload_df  = upload_df[has_owner_after_swap].copy()

    if len(extra_held) > 0:
        extra_held_exts = list(extra_held.get("externalreference", pd.Series(dtype=str)).fillna(""))
        print(f"  HELD BACK (owner outside migration scope): {len(extra_held)} business(es)")
        print(f"  Their externalrefs: {extra_held_exts}")
        # Append to held-back file
        existing_held = set(_load_held_refs())
        existing_held.update(extra_held_exts)
        _save_held_refs(list(existing_held))
        # Write combined held-back file
        held_all = pd.concat([held_df, extra_held], ignore_index=True)
        _write_round(rdir, "02 Clients_HELD_BACK", held_all)
    elif len(held_df) > 0:
        _write_round(rdir, "02 Clients_HELD_BACK", held_df)

    swapped = (upload_df.get("owners", pd.Series(dtype=str)).str.strip() != "").sum() if len(upload_df) > 0 else 0
    print(f"  {swapped} owners swapped to Plannr UUIDs")

    n_held_total = len(held_df) + len(extra_held)
    msg = (
        f"\nR2 done. Upload upload/R2/02 Clients BT.csv / xlsx to Plannr ({len(upload_df)} rows)."
    )
    if n_held_total > 0:
        msg += (
            f"\n  NOTE: {n_held_total} client(s) held back — see upload/R2/02 Clients_HELD_BACK.csv"
            f"\n    {len(held_df)} trust(s): no owner in Vision data"
            f"\n    {len(extra_held)} business(es): owner outside migration scope"
            f"\n  Supply the Plannr owner UUID for each and upload manually."
        )
    msg += (
        f"\nThen update (or re-drop) the Plannr client export in:\n"
        f"  {EXPORTS_DIR}\n"
        f"and run: python higgs_remap.py r3"
    )
    print(msg)


# ---------------------------------------------------------------------------
# Round 3 — Plans + I&E + Expenses + Notes + Circles
# ---------------------------------------------------------------------------

def run_r3():
    print("\n=== R3: Plans + I&E + Expenses + Notes + Circles ===")

    client_map  = _load_export("client")
    held_refs   = _load_held_refs()   # externalrefs of held-back trusts (if any)
    rdir        = os.path.join(UPLOAD_DIR, "R3")
    total_held  = 0

    # helper: filter rows whose sole/first client resolved to blank
    def _filter_held(df, col):
        blank = df[col].fillna("").str.strip() == ""
        n = blank.sum()
        return df[~blank].copy(), n

    def _clean_multi(df, col):
        """After a multi-col swap, strip empty parts from comma-joined values.
        Returns cleaned df and count of rows that became entirely blank."""
        new_vals = []
        gone_blank = 0
        for v in df[col]:
            parts = [p.strip() for p in str(v).split(",") if p.strip()]
            cleaned = ",".join(parts)
            new_vals.append(cleaned)
            if not cleaned:
                gone_blank += 1
        df = df.copy()
        df[col] = new_vals
        return df, gone_blank

    # --- Plans (individuals + BT combined) ---
    def _process_plans(path, label):
        if not os.path.isfile(path):
            return None, 0
        df = _read_csv(path)
        if "client_uuids" in df.columns:
            df, unres = _swap_multi_col(df, "client_uuids", client_map)
            if unres:
                unique_u = sorted(set(unres))
                print(f"  {label}: {len(unres)} owner ref(s) unresolvable (outside scope) — {len(unique_u)} unique, dropped from client_uuids")
            df, n_blank = _clean_multi(df, "client_uuids")
            n_ok = (df["client_uuids"].str.strip() != "").sum()
            print(f"  {label}: {n_ok}/{len(df)} rows have at least one resolved client_uuid")
            if n_blank:
                print(f"  {label}: {n_blank} rows became blank after drop — will be held back")
        if "assigned_adviser_uuid" in df.columns:
            df["assigned_adviser_uuid"] = DEST_ADVISER_UUID
        df, n_h = _filter_held(df, "client_uuids") if "client_uuids" in df.columns else (df, 0)
        return df, n_h

    plans_path = os.path.join(IMPORTER_DIR, "02 Plans.csv")
    if not os.path.isfile(plans_path):
        _stop(f"Plans file not found: {plans_path}")
    df_plans, n_hplan = _process_plans(plans_path, "Plans (ind)")
    total_held += n_hplan
    if n_hplan:
        print(f"  Plans (ind): {n_hplan} rows held back (blank client — held-back trust)")

    df_plans_bt, n_hplan_bt = _process_plans(os.path.join(IMPORTER_DIR, "02_BT Plans BT.csv"), "Plans (BT)")
    total_held += n_hplan_bt
    if n_hplan_bt:
        print(f"  Plans (BT): {n_hplan_bt} rows held back")

    if df_plans_bt is not None and len(df_plans_bt) > 0:
        df_plans_all = pd.concat([df_plans, df_plans_bt], ignore_index=True)
        print(f"  Plans combined: {len(df_plans)} ind + {len(df_plans_bt)} BT = {len(df_plans_all)}")
    else:
        df_plans_all = df_plans
    _write_round(rdir, "03 Plans", df_plans_all)

    # Collect Plannr plan UUIDs that ARE in this upload
    # (needed only if we were cross-referencing — not required here)

    # --- Income & Employment ---
    ie_path = os.path.join(IMPORTER_DIR, "06 Income & Employment.csv")
    if os.path.isfile(ie_path):
        df_ie = _read_csv(ie_path)
        for col in ("client_uuid",):
            if col in df_ie.columns:
                df_ie, unres = _swap_col(df_ie, col, client_map)
                if unres:
                    print(f"  I&E: {len(set(unres))} unresolvable client_uuid(s) — rows dropped")
        df_ie, n_h = _filter_held(df_ie, "client_uuid") if "client_uuid" in df_ie.columns else (df_ie, 0)
        total_held += n_h
        if n_h:
            print(f"  I&E: {n_h} rows held back (unresolvable client)")
        _write_round(rdir, "06 Income and Employment", df_ie)
    else:
        print(f"  06 Income & Employment.csv not found, skipping")

    # --- Expenses (split into 5,000-row files matching build output) ---
    exp_chunks = sorted(glob.glob(os.path.join(IMPORTER_DIR, "07 Expenses (*).csv")))
    if not exp_chunks:
        # fallback: single file from older build
        single = os.path.join(IMPORTER_DIR, "07 Expenses.csv")
        exp_chunks = [single] if os.path.isfile(single) else []
    if exp_chunks:
        total_exp = 0
        for _ec_path in exp_chunks:
            _ec_name = os.path.splitext(os.path.basename(_ec_path))[0]
            df_exp = _read_csv(_ec_path)
            if "client_uuid" in df_exp.columns:
                df_exp, unres = _swap_col(df_exp, "client_uuid", client_map)
                if unres:
                    print(f"  {_ec_name}: {len(set(unres))} unresolvable client_uuid(s) — rows dropped")
            df_exp, n_h = _filter_held(df_exp, "client_uuid") if "client_uuid" in df_exp.columns else (df_exp, 0)
            total_held += n_h
            if n_h:
                print(f"  {_ec_name}: {n_h} rows held back")
            _write_round(rdir, _ec_name, df_exp)
            total_exp += len(df_exp)
        print(f"  Expenses total: {total_exp} rows across {len(exp_chunks)} file(s)")
    else:
        print(f"  07 Expenses*.csv not found, skipping")

    # --- Notes ---
    notes_path = os.path.join(IMPORTER_DIR, "08 Notes.csv")
    if os.path.isfile(notes_path):
        df_notes = _read_csv(notes_path)
        # Notes primary ref column: account_uuid if present, else client_uuid
        ref_col = "account_uuid" if "account_uuid" in df_notes.columns else "client_uuid"
        if ref_col in df_notes.columns:
            df_notes, unres = _swap_col(df_notes, ref_col, client_map)
            if unres:
                print(f"  Notes: {len(set(unres))} unresolvable {ref_col}(s) — rows dropped")
        df_notes, n_h = _filter_held(df_notes, ref_col) if ref_col in df_notes.columns else (df_notes, 0)
        total_held += n_h
        if n_h:
            print(f"  Notes: {n_h} rows held back")
        _write_round(rdir, "08 Notes", df_notes)
    else:
        print(f"  08 Notes.csv not found, skipping")

    # --- Circles ---
    circles_path = os.path.join(IMPORTER_DIR, "09 Circles.csv")
    if os.path.isfile(circles_path):
        df_circles = _read_csv(circles_path)
        if len(df_circles) > 0 and "account_uuids" in df_circles.columns:
            df_circles, unres = _swap_multi_col(df_circles, "account_uuids", client_map)
            if unres:
                print(f"  Circles: {len(set(unres))} unresolvable account_uuid(s) — partial membership only")
            df_circles, _ = _clean_multi(df_circles, "account_uuids")
            # Drop circles that resolve to fewer than 2 members (no longer a valid circle)
            def _count_members(v):
                return sum(1 for p in str(v).split(",") if p.strip())
            valid = df_circles["account_uuids"].apply(_count_members) >= 2
            n_dropped = (~valid).sum()
            if n_dropped:
                print(f"  Circles: {n_dropped} circle(s) dropped (fewer than 2 resolved members)")
            df_circles = df_circles[valid].copy()
        _write_round(rdir, "09 Circles", df_circles)
    else:
        print(f"  09 Circles.csv not found or empty, skipping")

    if total_held:
        print(f"\n  Total rows held back (linked to ownerless trusts): {total_held}")

    print(
        f"\nR3 done. Upload R3 files to Plannr.\n"
        f"Then drop a Plannr plan export into:\n"
        f"  {EXPORTS_DIR}\n"
        f"and run: python higgs_remap.py r4"
    )


# ---------------------------------------------------------------------------
# Round 4 — Valuations + Charges
# ---------------------------------------------------------------------------

def run_r4():
    print("\n=== R4: Valuations + Charges ===")

    client_map = _load_export("client")
    plan_map   = _load_export("plan")
    rdir       = os.path.join(UPLOAD_DIR, "R4")

    # --- Valuations (individuals + BT combined) ---
    # Populate plan_uuid from review pack x_plan_ref; keep policy_number as fallback.
    def _swap_val_plan_uuids(df_v, review_xlsx, sheet="Valuations"):
        """Read x_plan_ref from review pack, swap to Plannr plan UUID, fill plan_uuid."""
        try:
            rp = _read_xlsx_sheet(review_xlsx, sheet)
            n = min(len(rp), len(df_v))
            x_refs = _get_col(rp.iloc[:n], "x_plan_ref") + [""] * max(0, len(df_v) - n)
        except Exception as e:
            print(f"  WARNING: could not read {sheet} x_plan_ref from {os.path.basename(review_xlsx)}: {e}")
            return df_v, 0, 0
        plan_uuid_vals = []
        unres = []
        for ref in x_refs:
            ref = str(ref).strip()
            mapped, bad = _swap(ref, plan_map)
            plan_uuid_vals.append(mapped)
            if bad and ref:
                unres.append(ref)
        df_v = df_v.copy()
        df_v["plan_uuid"] = plan_uuid_vals
        n_resolved = sum(1 for v in plan_uuid_vals if v)
        return df_v, n_resolved, len(set(unres))

    val_dfs = []
    val_path = os.path.join(IMPORTER_DIR, "03 Valuations.csv")
    if os.path.isfile(val_path):
        df_v = _read_csv(val_path)
        df_v, n_res, n_unres = _swap_val_plan_uuids(df_v, REVIEW_PACK)
        print(f"  Valuations (ind): {len(df_v)} rows, {n_res} plan_uuids resolved, {n_unres} unresolvable")
        val_dfs.append(df_v)
    val_bt_path = os.path.join(IMPORTER_DIR, "03_BT Valuations BT.csv")
    if os.path.isfile(val_bt_path):
        df_vbt = _read_csv(val_bt_path)
        df_vbt, n_res_bt, n_unres_bt = _swap_val_plan_uuids(df_vbt, REVIEW_BT)
        print(f"  Valuations (BT): {len(df_vbt)} rows, {n_res_bt} plan_uuids resolved, {n_unres_bt} unresolvable")
        val_dfs.append(df_vbt)
    if val_dfs:
        df_val = pd.concat(val_dfs, ignore_index=True) if len(val_dfs) > 1 else val_dfs[0]
        # Drop rows with neither plan_uuid nor policy_number (can't link to a plan)
        has_link = (df_val["plan_uuid"].str.strip() != "") | (df_val["policy_number"].str.strip() != "")
        n_dropped = (~has_link).sum()
        df_val = df_val[has_link].copy()
        print(f"  Valuations total: {len(df_val)} rows ({n_dropped} dropped — no plan_uuid or policy_number)")
        _write_round(rdir, "04 Valuations", df_val)
    else:
        print(f"  03 Valuations.csv not found, skipping")

    # --- Charges ---
    chg_path = os.path.join(IMPORTER_DIR, "05 Charges.csv")
    if not os.path.isfile(chg_path):
        print(f"  05 Charges.csv not found, skipping charges")
    else:
        df_chg = _read_csv(chg_path)
        print(f"  Charges source: {len(df_chg)} rows")

        # Read plan refs and client refs from review pack Charges helper columns.
        # The review pack Charges sheet has x_plan_ref (source plan UUID) and
        # x_client_refs (comma-separated source client UUIDs) inserted after
        # the adviser_uuid column via src_after in the builder.
        x_plan_refs   = [""] * len(df_chg)
        x_client_refs = [""] * len(df_chg)

        if os.path.isfile(REVIEW_PACK):
            try:
                df_rp = _read_xlsx_sheet(REVIEW_PACK, "Charges")
                # Align row count (review pack data rows == importer CSV rows)
                if len(df_rp) != len(df_chg):
                    print(
                        f"  WARNING: review pack Charges has {len(df_rp)} rows "
                        f"but importer CSV has {len(df_chg)}. "
                        f"Proceeding with min({len(df_rp)},{len(df_chg)}) aligned rows."
                    )
                n = min(len(df_rp), len(df_chg))
                x_plan_refs   = _get_col(df_rp.iloc[:n], "x_plan_ref")
                x_client_refs = _get_col(df_rp.iloc[:n], "x_client_refs")
                x_plan_refs   += [""] * max(0, len(df_chg) - n)
                x_client_refs += [""] * max(0, len(df_chg) - n)
                n_plan_refs = sum(1 for v in x_plan_refs if v.strip())
                print(f"  Review pack: {n_plan_refs} x_plan_ref values found")
            except Exception as e:
                print(f"  WARNING: Could not read review pack Charges sheet: {e}")
                print(f"  plan_uuid and client_uuid will be blank in output.")
        else:
            print(f"  WARNING: Review pack not found at {REVIEW_PACK}")
            print(f"  plan_uuid and client_uuid will be blank in charges output.")

        # Swap plan_uuid: x_plan_ref (source plan ref) -> Plannr plan UUID
        unres_plan = []
        plan_uuid_vals = []
        for ref in x_plan_refs:
            ref = str(ref).strip()
            mapped, bad = _swap(ref, plan_map)
            plan_uuid_vals.append(mapped)
            if bad and ref:
                unres_plan.append(ref)
        # Swap client_uuid: first source client ref from x_client_refs -> Plannr client UUID
        unres_client = []
        client_uuid_vals = []
        for refs_str in x_client_refs:
            refs_str = str(refs_str).strip()
            if not refs_str or refs_str.lower() in ("nan", ""):
                client_uuid_vals.append("")
                continue
            # x_client_refs may be ", " or "," or "|" separated
            first = re.split(r"\s*[,|]\s*", refs_str)[0].strip()
            mapped, bad = _swap(first, client_map)
            client_uuid_vals.append(mapped)
            if bad and first:
                unres_client.append(first)

        # Assign both UUID columns then filter rows that are fully resolved
        df_chg = df_chg.copy()
        df_chg["plan_uuid"]   = plan_uuid_vals
        df_chg["client_uuid"] = client_uuid_vals
        before = len(df_chg)
        df_chg = df_chg[df_chg["plan_uuid"].str.strip() != ""].copy()
        n_swapped = (df_chg["plan_uuid"].str.strip() != "").sum()
        if unres_plan:
            print(f"  Charges: {len(set(unres_plan))} unresolvable plan ref(s) (failed plans) — {before - len(df_chg)} rows dropped")
        print(f"  Charges: {n_swapped} plan_uuids resolved")
        if unres_client:
            print(f"  Charges: {len(set(unres_client))} unresolvable client ref(s)")
        n_cli = (df_chg["client_uuid"].str.strip() != "").sum()
        print(f"  Charges: {n_cli} client_uuids resolved")

        if "adviser_uuid" in df_chg.columns:
            df_chg["adviser_uuid"] = DEST_ADVISER_UUID

        # Append BT charges (if present)
        chg_bt_path = os.path.join(IMPORTER_DIR, "05_BT Charges BT.csv")
        if os.path.isfile(chg_bt_path):
            df_chg_bt = _read_csv(chg_bt_path)
            print(f"  Charges (BT): {len(df_chg_bt)} rows")
            x_plan_refs_bt   = [""] * len(df_chg_bt)
            x_client_refs_bt = [""] * len(df_chg_bt)
            if os.path.isfile(REVIEW_BT):
                try:
                    df_rp_bt = _read_xlsx_sheet(REVIEW_BT, "Charges")
                    n = min(len(df_rp_bt), len(df_chg_bt))
                    x_plan_refs_bt   = _get_col(df_rp_bt.iloc[:n], "x_plan_ref") + [""] * max(0, len(df_chg_bt) - n)
                    x_client_refs_bt = _get_col(df_rp_bt.iloc[:n], "x_client_refs") + [""] * max(0, len(df_chg_bt) - n)
                    print(f"  BT review pack: {sum(1 for v in x_plan_refs_bt if v.strip())} x_plan_ref values")
                except Exception as e:
                    print(f"  WARNING: Could not read BT review pack Charges sheet: {e}")
            bt_plan_uuids, bt_unres_plan = [], []
            for ref in x_plan_refs_bt:
                ref = str(ref).strip()
                mapped, bad = _swap(ref, plan_map)
                bt_plan_uuids.append(mapped)
                if bad and ref:
                    bt_unres_plan.append(ref)
            bt_client_uuids, bt_unres_cli = [], []
            for refs_str in x_client_refs_bt:
                refs_str = str(refs_str).strip()
                if not refs_str or refs_str.lower() in ("nan", ""):
                    bt_client_uuids.append(""); continue
                first = re.split(r"\s*[,|]\s*", refs_str)[0].strip()
                mapped, bad = _swap(first, client_map)
                bt_client_uuids.append(mapped)
                if bad and first:
                    bt_unres_cli.append(first)
            df_chg_bt = df_chg_bt.copy()
            df_chg_bt["plan_uuid"]   = bt_plan_uuids
            df_chg_bt["client_uuid"] = bt_client_uuids
            df_chg_bt = df_chg_bt[df_chg_bt["plan_uuid"].str.strip() != ""].copy()
            if bt_unres_plan:
                print(f"  Charges BT: {len(set(bt_unres_plan))} unresolvable plan ref(s) — {len(bt_plan_uuids) - len(df_chg_bt)} rows dropped")
            if bt_unres_cli:
                print(f"  Charges BT: {len(set(bt_unres_cli))} unresolvable client ref(s)")
            if "adviser_uuid" in df_chg_bt.columns:
                df_chg_bt["adviser_uuid"] = DEST_ADVISER_UUID
            df_chg = pd.concat([df_chg, df_chg_bt], ignore_index=True)
            print(f"  Charges combined: {len(df_chg)} rows")

        _write_round(rdir, "05 Charges", df_chg)

    # --- Premiums note ---
    prem_path = os.path.join(IMPORTER_DIR, "04 Premiums.csv")
    if os.path.isfile(prem_path):
        df_prem = _read_csv(prem_path)
        if len(df_prem) > 0:
            print(
                f"\n  NOTE: {len(df_prem)} premium row(s) exist in 04 Premiums.csv "
                f"but cannot be auto-remapped (plan_uuid not recoverable from importer CSV).\n"
                f"  Options:\n"
                f"    a) Link premiums manually in Plannr after R3 upload.\n"
                f"    b) Re-run the build step with x_plan_ref added as an extra_col "
                f"       for premiums, then add a premiums section to run_r4()."
            )

    print(f"\nR4 done.")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

ROUNDS = {"r1": run_r1, "r2": run_r2, "r3": run_r3, "r4": run_r4}

if __name__ == "__main__":
    if len(sys.argv) != 2 or sys.argv[1].lower() not in ROUNDS:
        print("Usage: python higgs_remap.py <r1|r2|r3|r4>\n")
        print("  r1  Individuals clients — no ref swap, sets adviser UUID + classifications")
        print("  r2  Business/trust clients — swaps owners; holds back ownerless trusts")
        print("  r3  Plans + I&E + Expenses + Notes + Circles — swaps client refs")
        print("  r4  Valuations + Charges — swaps plan and client refs\n")
        print(f"  Reads: {IMPORTER_DIR}")
        print(f"  Reads: {EXPORTS_DIR}  (Plannr exports dropped by user)")
        print(f"  Writes: {UPLOAD_DIR}/R<n>/")
        sys.exit(1)

    ROUNDS[sys.argv[1].lower()]()
