"""
HFP AR Migration — Analysis 26 Oct 2026
Steps 1b (valuation diff) + 2a-d (AR edits, CRM changes, conflicts).

Inputs:
  SENT:     Plannr_*.xlsx files in DATA_ROOT (as sent 14 Oct)
  RETURNED: DATA_ROOT/returned_20261026/*.xlsx (AR's copies)
  OLD_EXP:  DATA_ROOT/exports_20261014/
  NEW_EXP:  DATA_ROOT/exports_20261026/
  SNAP:     DATA_ROOT/snapshot_20261012T090906Z/

Output:
  DATA_ROOT/hfp_changes_20261026.xlsx
    Tabs: AR edits | Review entries | CRM new | CRM deleted | CRM changed | Conflicts

Rule 9: aggregates only in console. Names/UUIDs in Excel output only.
"""
import os, csv, json, re
from collections import defaultdict

DATA_ROOT   = r"C:\dev\AR Migration"
OLD_EXP     = os.path.join(DATA_ROOT, "exports_20261014")
NEW_EXP     = os.path.join(DATA_ROOT, "exports_20261026")
SNAP_DIR    = os.path.join(DATA_ROOT, "snapshot_20261012T090906Z")
RETURNED    = os.path.join(DATA_ROOT, "returned_20261026")
OUT_XLSX    = os.path.join(DATA_ROOT, "hfp_changes_20261026.xlsx")

def p(*a): print(*a, flush=True, file=__import__("sys").stdout)
import sys, io
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# ── helpers ───────────────────────────────────────────────────────────────────

def norm(v):
    """Return a normalised string for comparison. Strips whitespace, lowercases
    True/False, strips leading zeros from pure integers, normalises dates."""
    if v is None:
        return ""
    s = str(v).strip()
    # Bool-like
    if s.lower() in ("true","false","yes","no","1","0"):
        return s.lower().replace("yes","true").replace("no","false").replace("1","true").replace("0","false")
    # Numeric: strip trailing .0 and leading zeros
    try:
        f = float(s.replace(",",""))
        if f == int(f):
            return str(int(f))
        return str(f)
    except ValueError:
        pass
    # Date normalisation: DD/MM/YYYY → YYYY-MM-DD and strip time
    m = re.match(r'^(\d{1,2})/(\d{1,2})/(\d{4})$', s)
    if m:
        return f"{m.group(3)}-{int(m.group(2)):02d}-{int(m.group(1)):02d}"
    m = re.match(r'^(\d{4})-(\d{2})-(\d{2})(?:[ T].*)?$', s)
    if m:
        return f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
    return s


def is_info(col):
    return str(col).startswith("INFO:")


def is_original(col):
    return str(col).startswith("ORIGINAL_")


def comparable_cols(headers):
    """Return columns that should be compared (not INFO:, not ORIGINAL_)."""
    return [h for h in headers if not is_info(h) and not is_original(h)]


def load_xlsx_sheet(path, sheet):
    """Return (headers, [{col: val}]) for a sheet. Uses data_only=True."""
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    if not rows:
        return [], []
    headers = [str(c) if c is not None else "" for c in rows[0]]
    data = []
    for row in rows[1:]:
        if any(c is not None for c in row):
            data.append(dict(zip(headers, row)))
    return headers, data


def load_csv(path):
    """Return (headers, [{col: val}])."""
    with open(path, encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        headers = reader.fieldnames or []
        data = list(reader)
    return headers, data


def load_json_list(path):
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    if isinstance(data, list):
        return data
    if isinstance(data, dict):
        return list(data.values())
    return []


def info_name(row, sheet_label):
    """Extract a display name from INFO: columns for the Excel output."""
    info_cols = [k for k in row if str(k).startswith("INFO:")]
    parts = []
    for k in info_cols[:2]:
        v = row.get(k)
        if v:
            parts.append(str(v))
    return " / ".join(parts) if parts else ""


def uuid_col(headers, sheet_label):
    """Identify the reference UUID column for a given sheet."""
    if "uuid" in headers:
        return "uuid"
    return None


# ── compare two rows ──────────────────────────────────────────────────────────

def diff_rows(sent_row, ret_row, cols, skip_blank_added=False):
    """
    Compare sent vs returned for the given columns.
    Returns list of (col, sent_val, ret_val, format_only).
    format_only=True means values differ only in formatting.
    """
    diffs = []
    for col in cols:
        if col == "Review":
            continue  # handled separately in review entries
        s = sent_row.get(col) if sent_row else None
        r = ret_row.get(col) if ret_row else None
        s_str = str(s).strip() if s is not None else ""
        r_str = str(r).strip() if r is not None else ""
        if s_str == r_str:
            continue
        # Check format-only
        ns = norm(s_str)
        nr = norm(r_str)
        format_only = (ns == nr)
        if not format_only:
            diffs.append((col, s_str, r_str, False))
        else:
            diffs.append((col, s_str, r_str, True))
    return diffs


# ── Step 1b: Valuation diff ───────────────────────────────────────────────────

def step1b_valuations():
    p("\n── Step 1b: Valuation diff ──────────────────────────────────────────────")

    old_plan_path = os.path.join(OLD_EXP, "Plan Export HFP Migration 2026-10-14.csv")
    new_plan_path = os.path.join(NEW_EXP, "Plan Export HFP Migration 2026-10-27.csv")

    _, old_plans = load_csv(old_plan_path)
    _, new_plans = load_csv(new_plan_path)

    old_map = {r["uuid"]: r for r in old_plans if r.get("uuid")}
    new_map = {r["uuid"]: r for r in new_plans if r.get("uuid")}

    # Verify unit alignment: compare a handful of unchanged plans against snapshot
    with open(os.path.join(SNAP_DIR, "plans_dedup.json"), encoding="utf-8") as f:
        snap = json.load(f)  # dict keyed by uuid

    unit_checks = 0
    unit_ok = 0
    for uuid, plan in snap.items():
        if not isinstance(plan, dict): continue
        sv = plan.get("value")
        if not sv or not isinstance(sv, dict): continue
        snap_pence = sv.get("amount")
        if snap_pence is None: continue
        if uuid not in old_map or uuid not in new_map: continue
        old_val = old_map[uuid].get("current_value","")
        new_val = new_map[uuid].get("current_value","")
        if norm(old_val) != norm(new_val):
            continue  # changed plan, skip for unit check
        try:
            snap_pounds = float(snap_pence) / 100
            exp_pounds  = float(old_val) if old_val else 0.0
            if abs(snap_pounds - exp_pounds) < 0.01:
                unit_ok += 1
            unit_checks += 1
            if unit_checks >= 5:
                break
        except (ValueError, TypeError):
            pass

    p(f"  Unit check (unchanged plans, snap pence / 100 vs export pounds): {unit_ok}/{unit_checks} match")

    # Find changed plans
    changed = []
    for uuid, new_row in new_map.items():
        if uuid not in old_map:
            continue
        old_row = old_map[uuid]
        old_val = old_row.get("current_value","")
        new_val = new_row.get("current_value","")
        old_vat = old_row.get("current_value_at","")
        new_vat = new_row.get("current_value_at","")
        val_changed  = norm(old_val) != norm(new_val)
        date_changed = norm(old_vat) != norm(new_vat)
        if val_changed or date_changed:
            # Check for multiple owners
            client_uuids = new_row.get("client_uuids","")
            multi_owner = "|" in str(client_uuids) if client_uuids else False
            changed.append({
                "uuid": uuid,
                "old_value": old_val,
                "new_value": new_val,
                "old_date":  old_vat,
                "new_date":  new_vat,
                "multi_owner": multi_owner,
                "client_uuids": client_uuids,
            })

    val_only  = sum(1 for c in changed if norm(c["old_value"]) != norm(c["new_value"]) and norm(c["old_date"]) == norm(c["new_date"]))
    date_only = sum(1 for c in changed if norm(c["old_value"]) == norm(c["new_value"]) and norm(c["old_date"]) != norm(c["new_date"]))
    both      = sum(1 for c in changed if norm(c["old_value"]) != norm(c["new_value"]) and norm(c["old_date"]) != norm(c["new_date"]))
    multi     = sum(1 for c in changed if c["multi_owner"])

    p(f"  Plans with changed value or date: {len(changed)}")
    p(f"    value only changed:  {val_only}")
    p(f"    date only changed:   {date_only}")
    p(f"    both changed:        {both}")
    p(f"    multi-owner (no split): {multi}")

    # Totals before/after (pounds)
    def safe_float(s):
        try: return float(str(s).replace(",","").strip())
        except: return 0.0

    total_before = sum(safe_float(c["old_value"]) for c in changed)
    total_after  = sum(safe_float(c["new_value"]) for c in changed)
    p(f"  Total value before (changed plans, £): {total_before:,.2f}")
    p(f"  Total value after  (changed plans, £): {total_after:,.2f}")

    return changed


# ── merged returned Plans ─────────────────────────────────────────────────────

def merge_returned_plans():
    """
    Load the three returned Plans files. For each UUID, merge edits:
    - Take any non-blank value that differs from the sent value.
    - Track conflicts (same UUID+col different across files).
    Returns (merged_rows_by_uuid, merged_provider_rows, intra_conflicts).
    """
    import openpyxl

    ret_files = []
    for fname in sorted(os.listdir(RETURNED)):
        if fname.endswith(".xlsx") and fname.startswith("Plans"):
            ret_files.append(os.path.join(RETURNED, fname))

    # Load sent as baseline
    sent_plans_path = os.path.join(DATA_ROOT, "Plannr_Plans.xlsx")
    sent_hdrs, sent_rows = load_xlsx_sheet(sent_plans_path, "Plans")
    sent_prov_hdrs, sent_prov = load_xlsx_sheet(sent_plans_path, "Provider_plans_not_in_Plannr")

    sent_by_uuid = {}
    for row in sent_rows:
        uuid = str(row.get("uuid","")).strip()
        if uuid:
            sent_by_uuid[uuid] = row

    # Load all returned Plans sheets
    ret_plans_by_file = []
    ret_prov_by_file  = []
    for fpath in ret_files:
        _, plans_rows = load_xlsx_sheet(fpath, "Plans")
        _, prov_rows  = load_xlsx_sheet(fpath, "Provider_plans_not_in_Plannr")
        ret_plans_by_file.append({str(r.get("uuid","")).strip(): r for r in plans_rows if r.get("uuid")})
        ret_prov_by_file.append(prov_rows)

    # Merge Plans
    all_uuids = set()
    for d in ret_plans_by_file:
        all_uuids |= set(d.keys())

    merged = {}
    intra_conflicts = []  # (uuid, col, values across files)

    for uuid in all_uuids:
        rows = [d.get(uuid) for d in ret_plans_by_file]
        base = sent_by_uuid.get(uuid, {})
        merged_row = dict(base)  # start from sent

        for col in sent_hdrs:
            if is_info(col) or is_original(col):
                continue
            sent_val = str(base.get(col,"") or "").strip()
            file_vals = []
            for row in rows:
                if row is None: continue
                v = str(row.get(col,"") or "").strip()
                file_vals.append(v)

            distinct_changes = set(v for v in file_vals if norm(v) != norm(sent_val) and v != "")
            if len(distinct_changes) > 1:
                intra_conflicts.append({"uuid": uuid, "col": col, "values": list(file_vals)})
                # Use first non-blank change
                for v in file_vals:
                    if norm(v) != norm(sent_val) and v:
                        merged_row[col] = v
                        break
            elif distinct_changes:
                merged_row[col] = distinct_changes.pop()

        merged[uuid] = merged_row

    # Merge Provider (match by Portfolio number + Wrapper ref)
    def prov_key(r):
        return (str(r.get("Portfolio number","") or "").strip(),
                str(r.get("Wrapper ref","") or "").strip())

    sent_prov_map = {prov_key(r): r for r in sent_prov}

    # Collect all provider keys
    all_prov_keys = set()
    for rows in ret_prov_by_file:
        for r in rows:
            all_prov_keys.add(prov_key(r))

    merged_prov = {}
    for key in all_prov_keys:
        base = sent_prov_map.get(key, {})
        merged_row = dict(base)
        for col in sent_prov_hdrs:
            sent_val = str(base.get(col,"") or "").strip()
            file_vals = []
            for rows in ret_prov_by_file:
                row = next((r for r in rows if prov_key(r) == key), None)
                if row:
                    file_vals.append(str(row.get(col,"") or "").strip())
            distinct_changes = set(v for v in file_vals if norm(v) != norm(sent_val) and v != "")
            if distinct_changes:
                merged_row[col] = distinct_changes.pop()
        merged_prov[key] = merged_row

    p(f"  Merged Plans: {len(merged)} rows, intra-conflicts: {len(intra_conflicts)}")
    p(f"  Merged Provider: {len(merged_prov)} rows")

    return merged, merged_prov, intra_conflicts, sent_by_uuid, sent_hdrs, sent_prov_map, sent_prov_hdrs


# ── Step 2a: AR edits ─────────────────────────────────────────────────────────

def step2a_ar_edits(merged_plans, merged_prov, sent_by_uuid, sent_hdrs,
                    sent_prov_map, sent_prov_hdrs):
    """
    Compare RETURNED vs SENT for each sheet.
    Returns list of edit dicts (for Excel) and summary counts.
    """
    p("\n── Step 2a: AR edits ────────────────────────────────────────────────────")

    edits = []   # rows for AR edits tab
    review_entries = []  # rows for Review entries tab

    # ── Clients ──────────────────────────────────────────────────────────────
    sent_clients_path = os.path.join(DATA_ROOT, "Plannr_Clients.xlsx")
    ret_clients_path  = os.path.join(RETURNED, "Clients with no Adviser.xlsx")
    sent_c_hdrs, sent_c_rows = load_xlsx_sheet(sent_clients_path, "Clients")
    ret_c_hdrs,  ret_c_rows  = load_xlsx_sheet(ret_clients_path,  "Clients")

    sent_c_map = {str(r.get("uuid","")).strip(): r for r in sent_c_rows if r.get("uuid")}
    ret_c_map  = {str(r.get("uuid","")).strip(): r for r in ret_c_rows  if r.get("uuid")}

    c_added = [uuid for uuid in ret_c_map if uuid not in sent_c_map]
    c_deleted = [uuid for uuid in sent_c_map if uuid not in ret_c_map]
    c_no_uuid = [r for r in ret_c_rows if not str(r.get("uuid","")).strip()]

    c_cell_changes = 0; c_format_only = 0
    c_review_added = 0; c_review_changed = 0

    cmp_cols_c = comparable_cols(sent_c_hdrs)
    for uuid, ret_row in ret_c_map.items():
        if uuid not in sent_c_map:
            continue
        sent_row = sent_c_map[uuid]
        iname = info_name(sent_row, "Clients")

        # Review column
        sr = str(sent_row.get("Review","") or "").strip()
        rr = str(ret_row.get("Review","") or "").strip()
        if norm(rr) != norm(sr):
            if not sr:
                c_review_added += 1
            else:
                c_review_changed += 1
            review_entries.append({
                "sheet": "Clients", "uuid": uuid, "info_name": iname,
                "column": "Review", "old_value": sr, "new_value": rr,
            })

        diffs = diff_rows(sent_row, ret_row, cmp_cols_c)
        for col, old_v, new_v, fmt in diffs:
            if fmt:
                c_format_only += 1
            else:
                c_cell_changes += 1
                edits.append({
                    "sheet": "Clients", "uuid": uuid, "info_name": iname,
                    "column": col, "old_value": old_v, "new_value": new_v,
                })

    p(f"  Clients:  added={len(c_added)}  deleted={len(c_deleted)}  no_uuid={len(c_no_uuid)}")
    p(f"    data cells changed: {c_cell_changes}  format-only: {c_format_only}")
    p(f"    Review added: {c_review_added}  Review changed: {c_review_changed}")

    # ── Plans ─────────────────────────────────────────────────────────────────
    p_added   = [uuid for uuid in merged_plans if uuid not in sent_by_uuid]
    p_deleted = [uuid for uuid in sent_by_uuid if uuid not in merged_plans]

    p_cell_changes = 0; p_format_only = 0
    p_review_added = 0; p_review_changed = 0
    cmp_cols_p = comparable_cols(sent_hdrs)

    for uuid, ret_row in merged_plans.items():
        if uuid not in sent_by_uuid:
            continue
        sent_row = sent_by_uuid[uuid]
        iname = info_name(sent_row, "Plans")

        sr = str(sent_row.get("Review","") or "").strip()
        rr = str(ret_row.get("Review","") or "").strip()
        if norm(rr) != norm(sr):
            if not sr:
                p_review_added += 1
            else:
                p_review_changed += 1
            review_entries.append({
                "sheet": "Plans", "uuid": uuid, "info_name": iname,
                "column": "Review", "old_value": sr, "new_value": rr,
            })

        diffs = diff_rows(sent_row, ret_row, cmp_cols_p)
        for col, old_v, new_v, fmt in diffs:
            if fmt:
                p_format_only += 1
            else:
                p_cell_changes += 1
                edits.append({
                    "sheet": "Plans", "uuid": uuid, "info_name": iname,
                    "column": col, "old_value": old_v, "new_value": new_v,
                })

    p(f"  Plans:    added={len(p_added)}  deleted={len(p_deleted)}")
    p(f"    data cells changed: {p_cell_changes}  format-only: {p_format_only}")
    p(f"    Review added: {p_review_added}  Review changed: {p_review_changed}")

    # ── Provider_plans_not_in_Plannr ──────────────────────────────────────────
    def prov_key(r):
        return (str(r.get("Portfolio number","") or "").strip(),
                str(r.get("Wrapper ref","") or "").strip())

    prov_added   = [k for k in merged_prov if k not in sent_prov_map]
    prov_deleted = [k for k in sent_prov_map if k not in merged_prov]
    prov_changes = 0; prov_fmt = 0

    for key, ret_row in merged_prov.items():
        if key not in sent_prov_map:
            continue
        sent_row = sent_prov_map[key]
        cmp_cols = comparable_cols(sent_prov_hdrs)
        diffs = diff_rows(sent_row, ret_row, cmp_cols)
        for col, old_v, new_v, fmt in diffs:
            if fmt:
                prov_fmt += 1
            else:
                prov_changes += 1
                edits.append({
                    "sheet": "Provider_plans_not_in_Plannr",
                    "uuid": f"{key[0]}|{key[1]}", "info_name": "",
                    "column": col, "old_value": old_v, "new_value": new_v,
                })

    p(f"  Provider: added={len(prov_added)}  deleted={len(prov_deleted)}")
    p(f"    data cells changed: {prov_changes}  format-only: {prov_fmt}")

    # ── Charges ──────────────────────────────────────────────────────────────
    sent_ch_path = os.path.join(DATA_ROOT, "Plannr_Charges.xlsx")
    ret_ch_path  = os.path.join(RETURNED, "Plannr Charges.xlsx")
    sent_ch_hdrs, sent_ch_rows = load_xlsx_sheet(sent_ch_path, "Charges")
    ret_ch_hdrs,  ret_ch_rows  = load_xlsx_sheet(ret_ch_path,  "Charges")

    sent_ch_map = {str(r.get("uuid","")).strip(): r for r in sent_ch_rows if r.get("uuid")}
    ret_ch_map  = {str(r.get("uuid","")).strip(): r for r in ret_ch_rows  if r.get("uuid")}

    ch_added   = [uuid for uuid in ret_ch_map if uuid not in sent_ch_map]
    ch_deleted = [uuid for uuid in sent_ch_map if uuid not in ret_ch_map]
    ch_changes = 0; ch_fmt = 0

    cmp_cols_ch = comparable_cols(sent_ch_hdrs)
    for uuid, ret_row in ret_ch_map.items():
        if uuid not in sent_ch_map:
            continue
        sent_row = sent_ch_map[uuid]
        iname = info_name(sent_row, "Charges")
        diffs = diff_rows(sent_row, ret_row, cmp_cols_ch)
        for col, old_v, new_v, fmt in diffs:
            if fmt:
                ch_fmt += 1
            else:
                ch_changes += 1
                edits.append({
                    "sheet": "Charges", "uuid": uuid, "info_name": iname,
                    "column": col, "old_value": old_v, "new_value": new_v,
                })

    p(f"  Charges:  added={len(ch_added)}  deleted={len(ch_deleted)}")
    p(f"    data cells changed: {ch_changes}  format-only: {ch_fmt}")

    # ── Trail Fees (not in SENT) ──────────────────────────────────────────────
    p(f"  Trail Fees (Excluding Transact & Rathbones).xlsx: NOT in SENT — flagged as unmatched")

    return edits, review_entries, {
        "clients":  {"added": len(c_added), "deleted": len(c_deleted),
                     "no_uuid": len(c_no_uuid), "changed_cells": c_cell_changes,
                     "format_only": c_format_only,
                     "review_added": c_review_added, "review_changed": c_review_changed},
        "plans":    {"added": len(p_added), "deleted": len(p_deleted),
                     "changed_cells": p_cell_changes, "format_only": p_format_only,
                     "review_added": p_review_added, "review_changed": p_review_changed},
        "provider": {"added": len(prov_added), "deleted": len(prov_deleted),
                     "changed_cells": prov_changes, "format_only": prov_fmt},
        "charges":  {"added": len(ch_added), "deleted": len(ch_deleted),
                     "changed_cells": ch_changes, "format_only": ch_fmt},
    }, sent_ch_map, ret_ch_map, merged_plans, sent_by_uuid


# ── Step 2b: What the AR wrote ────────────────────────────────────────────────

def step2b_what_ar_wrote(review_entries, merged_plans, sent_by_uuid,
                          sent_ch_map, ret_ch_map, merged_prov, sent_prov_map,
                          sent_prov_hdrs):
    p("\n── Step 2b: What the AR wrote ───────────────────────────────────────────")

    # Review entries: split by sheet
    clients_reviews = [e for e in review_entries if e["sheet"] == "Clients"]
    plans_reviews   = [e for e in review_entries if e["sheet"] == "Plans"]
    p(f"  Review column entries on Clients: {len(clients_reviews)}")
    p(f"  Review column entries on Plans:   {len(plans_reviews)}")

    # Provider tab: action column counts
    def prov_key(r): return (str(r.get("Portfolio number","") or "").strip(), str(r.get("Wrapper ref","") or "").strip())
    action_counts = defaultdict(int)
    add_to_plannr_filled = 0
    add_to_plannr_missing = 0
    for key, ret_row in merged_prov.items():
        action = str(ret_row.get("action","") or "").strip()
        if action:
            action_counts[action] += 1
        if action.lower() == "add to plannr":
            client_col = str(ret_row.get("Client in your Plannr","") or "").strip()
            if client_col:
                add_to_plannr_filled += 1
            else:
                add_to_plannr_missing += 1

    p(f"  Provider tab action counts:")
    for action, cnt in sorted(action_counts.items()):
        p(f"    {action!r}: {cnt}")
    p(f"    'Add to Plannr' with Client filled: {add_to_plannr_filled}  missing: {add_to_plannr_missing}")

    # Charges: Rate needed rows filled in — compare by row index (no UUID on Rate needed rows)
    sent_ch_path = os.path.join(DATA_ROOT, "Plannr_Charges.xlsx")
    ret_ch_path  = os.path.join(RETURNED, "Plannr Charges.xlsx")
    _, sent_ch_all = load_xlsx_sheet(sent_ch_path, "Charges")
    _, ret_ch_all  = load_xlsx_sheet(ret_ch_path, "Charges")

    rate_needed_filled = 0
    rate_needed_blank  = 0
    protection_with_amount = 0

    for i, sent_row in enumerate(sent_ch_all):
        if i >= len(ret_ch_all):
            break
        ret_row = ret_ch_all[i]
        basis = str(sent_row.get("basis","") or "").strip()
        info_type = str(sent_row.get("INFO: Type","") or "").strip().lower()

        if basis == "Rate needed":
            sent_pct = str(sent_row.get("percentage","") or "").strip()
            ret_pct  = str(ret_row.get("percentage","") or "").strip()
            ret_inc  = str(ret_row.get("indicative_annual_income","") or "").strip()
            was_blank = norm(sent_pct) in ("", "0", "none")
            now_filled = (ret_pct and norm(ret_pct) not in ("", "0", "none")) or \
                         (ret_inc and norm(ret_inc) not in ("", "0", "none"))
            if was_blank and now_filled:
                rate_needed_filled += 1
            elif was_blank and not now_filled:
                rate_needed_blank += 1

        if "protection" in info_type:
            sent_pct = str(sent_row.get("percentage","") or "").strip()
            sent_inc = str(sent_row.get("indicative_annual_income","") or "").strip()
            ret_pct  = str(ret_row.get("percentage","") or "").strip()
            ret_inc  = str(ret_row.get("indicative_annual_income","") or "").strip()
            old_had_value = (norm(sent_pct) not in ("", "0", "none")) or \
                            (norm(sent_inc) not in ("", "0", "none"))
            new_has_value = (ret_pct and norm(ret_pct) not in ("", "0", "none")) or \
                            (ret_inc and norm(ret_inc) not in ("", "0", "none"))
            if not old_had_value and new_has_value:
                protection_with_amount += 1

    ch_deleted_count = sum(1 for uuid in sent_ch_map if uuid not in ret_ch_map)

    p(f"  Charges:")
    p(f"    Rate needed rows filled in: {rate_needed_filled}  still blank: {rate_needed_blank}")
    p(f"    Protection rows with new amount: {protection_with_amount}")
    p(f"    Rows deleted (in SENT, not in RETURNED): {ch_deleted_count}")

    return {
        "clients_review": len(clients_reviews),
        "plans_review": len(plans_reviews),
        "provider_actions": dict(action_counts),
        "add_to_plannr_filled": add_to_plannr_filled,
        "add_to_plannr_missing": add_to_plannr_missing,
        "rate_needed_filled": rate_needed_filled,
        "rate_needed_blank": rate_needed_blank,
        "protection_with_amount": protection_with_amount,
        "charges_deleted": ch_deleted_count,
    }


# ── Step 2c: CRM changes ──────────────────────────────────────────────────────

def load_export_as_map(path, key="uuid"):
    """Load CSV export as dict keyed by key column."""
    _, rows = load_csv(path)
    result = {}
    for r in rows:
        k = str(r.get(key,"") or "").strip()
        if k:
            result[k] = r
    return result


def diff_export_rows(old_row, new_row, skip_cols=None):
    skip = set(skip_cols or [])
    all_cols = set(old_row.keys()) | set(new_row.keys())
    diffs = []
    for col in sorted(all_cols):
        if col in skip: continue
        ov = str(old_row.get(col,"") or "").strip()
        nv = str(new_row.get(col,"") or "").strip()
        if norm(ov) != norm(nv):
            diffs.append((col, ov, nv))
    return diffs


def step2c_crm_changes(val_changed_uuids):
    p("\n── Step 2c: CRM changes (new exports vs 14 Oct) ─────────────────────────")

    SKIP_META = {"updated_at"}  # always changes; not meaningful

    # UUID-keyed exports
    uuid_exports = [
        ("clients",    "Client Export HFP Migration 2026-10-14.csv",
                       "Client Export HFP Migration 2026-10-27.csv"),
        ("plans",      "Plan Export HFP Migration 2026-10-14.csv",
                       "Plan Export HFP Migration 2026-10-27.csv"),
        ("charges",    "Charges Export HFP Migration 2026-10-14.csv",
                       "Charges Export HFP Migration 2026-10-27.csv"),
        ("premiums",   "Premium Export HFP Migration 2026-10-14.csv",
                       "Premium Export HFP Migration 2026-10-27.csv"),
        ("notes",      "Note Export HFP Migration 2026-10-14.csv",
                       "Note Export HFP Migration 2026-10-27.csv"),
    ]

    crm_new     = []
    crm_deleted = []
    crm_changed = []

    summary = {}

    for label, old_fname, new_fname in uuid_exports:
        old_path = os.path.join(OLD_EXP, old_fname)
        new_path = os.path.join(NEW_EXP, new_fname)
        if not os.path.exists(old_path) or not os.path.exists(new_path):
            p(f"  {label}: MISSING FILE — skip")
            continue

        old_map = load_export_as_map(old_path)
        new_map = load_export_as_map(new_path)

        added   = [uuid for uuid in new_map if uuid not in old_map]
        deleted = [uuid for uuid in old_map if uuid not in new_map]
        changed_count = 0

        for uuid in new_map:
            if uuid not in old_map: continue
            old_row = old_map[uuid]
            new_row = new_map[uuid]
            diffs = diff_export_rows(old_row, new_row, skip_cols=SKIP_META)
            if diffs:
                changed_count += 1
                iname = new_row.get("name","") or new_row.get("client_names","") or ""
                for col, ov, nv in diffs:
                    crm_changed.append({
                        "entity": label, "uuid": uuid, "info_name": str(iname),
                        "field": col, "old_value": ov, "new_value": nv,
                    })

        for uuid in added:
            r = new_map[uuid]
            iname = r.get("name","") or r.get("client_names","") or ""
            crm_new.append({"entity": label, "uuid": uuid, "info_name": str(iname)})

        for uuid in deleted:
            r = old_map[uuid]
            iname = r.get("name","") or r.get("client_names","") or ""
            crm_deleted.append({"entity": label, "uuid": uuid, "info_name": str(iname)})

        p(f"  {label}: old={len(old_map)}  new={len(new_map)}  added={len(added)}  deleted={len(deleted)}  changed={changed_count}")
        summary[label] = {"old": len(old_map), "new": len(new_map),
                          "added": len(added), "deleted": len(deleted),
                          "changed": changed_count}

    # Circles — one row per member; key = (circle_uuid, account_uuid)
    circles_old_path = os.path.join(OLD_EXP, "Circle Export HFP Migration 2026-10-14.csv")
    circles_new_path = os.path.join(NEW_EXP, "Circle Export HFP Migration 2026-10-27.csv")
    if os.path.exists(circles_old_path) and os.path.exists(circles_new_path):
        _, old_circ_rows = load_csv(circles_old_path)
        _, new_circ_rows = load_csv(circles_new_path)
        # Distinct circles
        old_circles = {r["circle_uuid"] for r in old_circ_rows if r.get("circle_uuid")}
        new_circles = {r["circle_uuid"] for r in new_circ_rows if r.get("circle_uuid")}
        circ_added   = new_circles - old_circles
        circ_deleted = old_circles - new_circles
        # Membership changes: (circle_uuid, account_uuid)
        old_members = {(r["circle_uuid"], r["account_uuid"]) for r in old_circ_rows if r.get("circle_uuid") and r.get("account_uuid")}
        new_members = {(r["circle_uuid"], r["account_uuid"]) for r in new_circ_rows if r.get("circle_uuid") and r.get("account_uuid")}
        mem_added   = new_members - old_members
        mem_deleted = old_members - new_members
        # Membership changes in existing circles
        mem_added_existing   = sum(1 for (c,_) in mem_added   if c in old_circles)
        mem_deleted_existing = sum(1 for (c,_) in mem_deleted if c in new_circles)
        for c in circ_added:
            crm_new.append({"entity": "circles", "uuid": c, "info_name": ""})
        for c in circ_deleted:
            crm_deleted.append({"entity": "circles", "uuid": c, "info_name": ""})
        for (c, a) in mem_added:
            if c in old_circles:
                crm_changed.append({"entity": "circles", "uuid": c, "info_name": "",
                                    "field": "member_added", "old_value": "", "new_value": a})
        for (c, a) in mem_deleted:
            if c in new_circles:
                crm_changed.append({"entity": "circles", "uuid": c, "info_name": "",
                                    "field": "member_removed", "old_value": a, "new_value": ""})
        changed_c = mem_added_existing + mem_deleted_existing
        p(f"  circles: old={len(old_circles)}  new={len(new_circles)}  added={len(circ_added)}  deleted={len(circ_deleted)}  membership_changes={changed_c}")
        summary["circles"] = {"old": len(old_circles), "new": len(new_circles),
                              "added": len(circ_added), "deleted": len(circ_deleted),
                              "changed": changed_c}

    # Expenditure — no per-row UUID; compare row counts per client
    exp_old_path = os.path.join(OLD_EXP, "Client Expenditure Export HFP Migration 2026-10-14.csv")
    exp_new_path = os.path.join(NEW_EXP, "Client Expenditure Export HFP Migration 2026-10-27.csv")
    if os.path.exists(exp_old_path) and os.path.exists(exp_new_path):
        _, old_exp_rows = load_csv(exp_old_path)
        _, new_exp_rows = load_csv(exp_new_path)
        from collections import Counter
        old_counts = Counter(r.get("client_uuid","") for r in old_exp_rows if r.get("client_uuid"))
        new_counts = Counter(r.get("client_uuid","") for r in new_exp_rows if r.get("client_uuid"))
        all_clients = set(old_counts) | set(new_counts)
        exp_changed_clients = sum(1 for c in all_clients if old_counts.get(c,0) != new_counts.get(c,0))
        exp_new_clients = sum(1 for c in new_counts if c not in old_counts)
        exp_del_clients = sum(1 for c in old_counts if c not in new_counts)
        if exp_changed_clients:
            crm_changed.append({"entity": "expenditure", "uuid": "(multiple)",
                                 "info_name": f"{exp_changed_clients} clients",
                                 "field": "row_count", "old_value": f"{len(old_exp_rows)} total",
                                 "new_value": f"{len(new_exp_rows)} total"})
        p(f"  expenditure: old={len(old_exp_rows)} rows  new={len(new_exp_rows)} rows  clients_changed={exp_changed_clients}  clients_new={exp_new_clients}  clients_lost={exp_del_clients}")
        summary["expenditure"] = {"old": len(old_exp_rows), "new": len(new_exp_rows),
                                  "added": exp_new_clients, "deleted": exp_del_clients,
                                  "changed": exp_changed_clients}

    # Dependants (JSON) — skip 'accounts' key (new export has it; old doesn't)
    DEP_SKIP = {"accounts", "updated_at"}
    old_dep_path = os.path.join(OLD_EXP, "dependants.json")
    new_dep_path = os.path.join(NEW_EXP, "dependants.json")
    old_deps = {str(r.get("uuid","")).strip(): r for r in load_json_list(old_dep_path) if r.get("uuid")}
    new_deps = {str(r.get("uuid","")).strip(): r for r in load_json_list(new_dep_path) if r.get("uuid")}

    dep_added   = [uuid for uuid in new_deps if uuid not in old_deps]
    dep_deleted = [uuid for uuid in old_deps if uuid not in new_deps]
    dep_changed_count = 0
    for uuid in new_deps:
        if uuid not in old_deps: continue
        old_r = {k: v for k, v in old_deps[uuid].items() if k not in DEP_SKIP}
        new_r = {k: v for k, v in new_deps[uuid].items() if k not in DEP_SKIP}
        if json.dumps(old_r, sort_keys=True) != json.dumps(new_r, sort_keys=True):
            dep_changed_count += 1
            crm_changed.append({
                "entity": "dependants", "uuid": uuid, "info_name": "",
                "field": "(any)", "old_value": "(see JSON)", "new_value": "(see JSON)",
            })
    for uuid in dep_added:
        crm_new.append({"entity": "dependants", "uuid": uuid, "info_name": ""})
    for uuid in dep_deleted:
        crm_deleted.append({"entity": "dependants", "uuid": uuid, "info_name": ""})
    p(f"  dependants: old={len(old_deps)}  new={len(new_deps)}  added={len(dep_added)}  deleted={len(dep_deleted)}  changed={dep_changed_count}")
    summary["dependants"] = {"old": len(old_deps), "new": len(new_deps),
                             "added": len(dep_added), "deleted": len(dep_deleted),
                             "changed": dep_changed_count}

    # Valuation summary (from Step 1b)
    val_uuids = set(val_changed_uuids)
    p(f"  plans (valuation): {len(val_uuids)} plans with changed value/date")

    return crm_new, crm_deleted, crm_changed, summary


# ── Step 2d: Conflicts ────────────────────────────────────────────────────────

def step2d_conflicts(ar_edits, crm_changed):
    p("\n── Step 2d: Conflicts ───────────────────────────────────────────────────")

    # Build index of AR edits: (entity, uuid, field) -> (old, new)
    ar_index = {}
    entity_map = {"Clients": "clients", "Plans": "plans", "Charges": "charges",
                  "Provider_plans_not_in_Plannr": "provider"}
    for e in ar_edits:
        entity = entity_map.get(e["sheet"], e["sheet"].lower())
        key = (entity, str(e["uuid"]), e["column"])
        ar_index[key] = (e["old_value"], e["new_value"])

    conflicts = []
    for c in crm_changed:
        key = (c["entity"], c["uuid"], c["field"])
        if key in ar_index:
            ar_old, ar_new = ar_index[key]
            conflicts.append({
                "entity": c["entity"],
                "uuid":   c["uuid"],
                "info_name": c["info_name"],
                "field":  c["field"],
                "ar_old": ar_old, "ar_new": ar_new,
                "crm_old": c["old_value"], "crm_new": c["new_value"],
            })

    p(f"  Conflicts (same record+field changed in AR edits AND CRM): {len(conflicts)}")
    return conflicts


# ── Write Excel ───────────────────────────────────────────────────────────────

def write_excel(ar_edits, review_entries, crm_new, crm_deleted, crm_changed,
                conflicts, val_changed):
    import openpyxl
    from openpyxl.styles import Font, PatternFill

    wb = openpyxl.Workbook()

    HDR_FILL = PatternFill("solid", fgColor="1F4E79")
    HDR_FONT = Font(bold=True, color="FFFFFF")

    def add_sheet(name, headers, rows):
        if name == "Sheet":
            ws = wb.active
            ws.title = name
        else:
            ws = wb.create_sheet(name)
        for ci, h in enumerate(headers, 1):
            cell = ws.cell(1, ci, h)
            cell.font = HDR_FONT
            cell.fill = HDR_FILL
        for ri, row in enumerate(rows, 2):
            for ci, h in enumerate(headers, 1):
                ws.cell(ri, ci, row.get(h,""))
        # Auto-width
        for col in ws.columns:
            max_len = max((len(str(c.value or "")) for c in col), default=10)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 50)
        return ws

    # AR edits
    add_sheet("AR edits",
        ["sheet","uuid","info_name","column","old_value","new_value"],
        ar_edits)

    # Review entries
    add_sheet("Review entries",
        ["sheet","uuid","info_name","column","old_value","new_value"],
        review_entries)

    # CRM new
    add_sheet("CRM new",
        ["entity","uuid","info_name"],
        crm_new)

    # CRM deleted
    add_sheet("CRM deleted",
        ["entity","uuid","info_name"],
        crm_deleted)

    # CRM changed
    add_sheet("CRM changed",
        ["entity","uuid","info_name","field","old_value","new_value"],
        crm_changed)

    # Conflicts
    add_sheet("Conflicts",
        ["entity","uuid","info_name","field","ar_old","ar_new","crm_old","crm_new"],
        conflicts)

    # Valuations sheet (Step 1b detail)
    val_rows = [{"uuid": v["uuid"], "old_value_gbp": v["old_value"],
                 "new_value_gbp": v["new_value"], "old_date": v["old_date"],
                 "new_date": v["new_date"],
                 "multi_owner": "Yes" if v["multi_owner"] else "No"}
                for v in val_changed]
    add_sheet("Valuations",
        ["uuid","old_value_gbp","new_value_gbp","old_date","new_date","multi_owner"],
        val_rows)

    # Remove default sheet if unused
    if "Sheet" in wb.sheetnames and len(wb.sheetnames) > 1:
        del wb["Sheet"]

    wb.save(OUT_XLSX)
    p(f"\n  Saved: {OUT_XLSX}")


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    p("=" * 68)
    p("HFP AR Migration — Analysis 26 Oct 2026 (Steps 1b + 2)")
    p("=" * 68)

    # Step 1b
    val_changed = step1b_valuations()
    val_changed_uuids = [c["uuid"] for c in val_changed]

    # Merge returned Plans (needed for 2a and 2b)
    p("\n── Merging returned Plans files ─────────────────────────────────────────")
    (merged_plans, merged_prov, intra_conflicts,
     sent_by_uuid, sent_hdrs, sent_prov_map, sent_prov_hdrs) = merge_returned_plans()
    if intra_conflicts:
        p(f"  WARNING: {len(intra_conflicts)} intra-file conflicts in returned Plans")

    # Step 2a
    (ar_edits, review_entries, ar_summary,
     sent_ch_map, ret_ch_map, merged_plans, sent_by_uuid) = step2a_ar_edits(
        merged_plans, merged_prov, sent_by_uuid, sent_hdrs,
        sent_prov_map, sent_prov_hdrs)

    # Step 2b
    step2b_what_ar_wrote(review_entries, merged_plans, sent_by_uuid,
                          sent_ch_map, ret_ch_map, merged_prov, sent_prov_map,
                          sent_prov_hdrs)

    # Step 2c
    crm_new, crm_deleted, crm_changed, crm_summary = step2c_crm_changes(val_changed_uuids)

    # Step 2d
    conflicts = step2d_conflicts(ar_edits, crm_changed)

    # Write Excel
    p("\n── Writing Excel output ─────────────────────────────────────────────────")
    write_excel(ar_edits, review_entries, crm_new, crm_deleted,
                crm_changed, conflicts, val_changed)

    # Summary tables
    p("\n" + "=" * 68)
    p("SUMMARY")
    p("=" * 68)

    p("\n[Step 1] Exports: 7 exports + dependants — all OK")

    p("\n[Step 1b] Valuations")
    p(f"  Plans with changed value or date: {len(val_changed)}")
    safe_float = lambda s: float(str(s).replace(",","").strip()) if s else 0.0
    tb = sum(safe_float(c["old_value"]) for c in val_changed)
    ta = sum(safe_float(c["new_value"]) for c in val_changed)
    multi = sum(1 for c in val_changed if c["multi_owner"])
    p(f"  Total value before (£): {tb:,.2f}  after (£): {ta:,.2f}")
    p(f"  Multi-owner plans (not split): {multi}")

    p("\n[Step 2a] AR edits (RETURNED vs SENT)")
    p(f"  {'Sheet':<38}  {'Added':>5}  {'Deleted':>7}  {'Changed':>7}  {'Fmt-only':>8}")
    p(f"  {'-'*38}  {'-----':>5}  {'-------':>7}  {'-------':>7}  {'--------':>8}")
    for sheet, label in [("clients","Clients"),("plans","Plans"),
                          ("provider","Provider"),("charges","Charges")]:
        s = ar_summary[sheet]
        p(f"  {label:<38}  {s['added']:>5}  {s['deleted']:>7}  {s.get('changed_cells',0):>7}  {s.get('format_only',0):>8}")

    p("\n[Step 2b] What the AR wrote")
    p(f"  Review entries added/changed on Clients: {ar_summary['clients']['review_added']+ar_summary['clients']['review_changed']}")
    p(f"  Review entries added/changed on Plans:   {ar_summary['plans']['review_added']+ar_summary['plans']['review_changed']}")

    p("\n[Step 2c] CRM changes (26 Oct export vs 14 Oct)")
    p(f"  {'Entity':<20}  {'Old':>6}  {'New':>6}  {'Added':>5}  {'Deleted':>7}  {'Changed':>7}")
    p(f"  {'-'*20}  {'----':>6}  {'----':>6}  {'-----':>5}  {'-------':>7}  {'-------':>7}")
    for label, s in crm_summary.items():
        p(f"  {label:<20}  {s['old']:>6}  {s['new']:>6}  {s['added']:>5}  {s['deleted']:>7}  {s['changed']:>7}")

    p(f"\n[Step 2d] Conflicts: {len(conflicts)}")

    p(f"\nDetail written to: {OUT_XLSX}")
    p("STOP — do not build import files yet.")


if __name__ == "__main__":
    main()
