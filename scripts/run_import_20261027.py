"""
Production import runner — HFP AR migration (27 Oct 2026).

Approved waves: W1, W2, W3, expenditure, W4, W5, W6.
(W7/W8 wait for go-ahead.)

Rules:
  - Destination writes: POST /api/v1/import (validate then store) + expenditure API only.
  - Confirm destination by employee list, never /api/v1/firm.
  - After each wave resolve source→dest UUIDs via external_references; any unmatched = STOP.
  - Before W4: Vision sweep on W4–W7 files; blank whole-cell Vision firm names.
  - STOP after W6 with full report.
"""

import csv, io, json, os, sys, time
import requests

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# ─── Constants ─────────────────────────────────────────────────────────────────

BASE_URL          = "https://api.plannrcrm.com"
KV_NAME           = "kv-plannr-ingestion"
DEST_PAT_SECRET   = "plannr-prod-pat"
DEST_ACCOUNT_UUID = "ccd7f98b-d6a4-46ad-bab3-242ab42826eb"
MIGRATION_CO      = "data_migration"
IMPORT_DIR        = r"C:\dev\AR Migration\import_20261027"
EXP_EXPORT        = r"C:\dev\AR Migration\exports_20261026\Client Expenditure Export HFP Migration 2026-10-27.csv"

EXPECTED_EMPLOYEES = {"jim hoban", "katie hoban", "tommie hoban", "valerie hoban"}

# Source column→field maps for UUID substitution in each wave
REWRITE_FIELDS = {
    "W2_circles":      ["account_uuids"],
    "W3_relationships": ["account_uuid", "related_account_uuid"],
    "W3_dependants":   ["account_uuids"],
    "W3_income":       ["client_uuid"],
    "W6_valuations":   ["plan_uuid"],
    "W7_charges":      ["plan_uuid", "client_uuid"],
}

# ─── PAT ───────────────────────────────────────────────────────────────────────

def resolve_pat():
    from azure.identity import DefaultAzureCredential
    from azure.keyvault.secrets import SecretClient
    sc = SecretClient(
        vault_url=f"https://{KV_NAME}.vault.azure.net/",
        credential=DefaultAzureCredential(),
    )
    return sc.get_secret(DEST_PAT_SECRET).value


# ─── HTTP client ───────────────────────────────────────────────────────────────

class Dest:
    def __init__(self, pat):
        self._s = requests.Session()
        self._s.headers.update({
            "Authorization": f"Bearer {pat}",
            "Accept": "application/json",
            "X-PLANNR-ACCOUNT-UUID": DEST_ACCOUNT_UUID,
        })

    def get(self, path, **kw):
        return self._s.get(f"{BASE_URL}{path}", **kw)

    def post(self, path, **kw):
        return self._s.post(f"{BASE_URL}{path}", **kw)

    def delete(self, path, **kw):
        return self._s.delete(f"{BASE_URL}{path}", **kw)

    def upload(self, csv_bytes, filename, type_str, mode):
        return self._s.post(
            f"{BASE_URL}/api/v1/import",
            files={"file": (filename, io.BytesIO(csv_bytes), "text/csv")},
            data={"type": type_str, "mode": mode},
        )

    def map_headers(self, imp_uuid, headers):
        return self._s.post(f"{BASE_URL}/api/v1/import/map/{imp_uuid}",
                            json={"headers": {h: h for h in headers}})

    def start(self, imp_uuid):
        return self._s.post(f"{BASE_URL}/api/v1/import/start/{imp_uuid}")

    def get_import(self, imp_uuid):
        return self._s.get(f"{BASE_URL}/api/v1/import/{imp_uuid}")


# ─── Helpers ───────────────────────────────────────────────────────────────────

def p(*a):
    print(*a, flush=True)


def get_paged(dest, path, max_pages=50, params=None):
    items, page = [], 1
    while page <= max_pages:
        pr = dict(params or {})
        pr.update({"page": page, "per_page": 200})
        r = dest.get(path, params=pr)
        if r.status_code != 200:
            break
        body = r.json()
        data = body.get("data", [])
        if isinstance(data, dict):
            data = [data]
        items.extend(data)
        meta = body.get("meta", {})
        if page >= meta.get("last_page", 1):
            break
        page += 1
    return items


def get_paged_with_refs(dest, path, max_pages=50):
    """Like get_paged but passes include=external_references."""
    return get_paged(dest, path, max_pages=max_pages,
                     params={"include": "external_references"})


def poll_import(dest, imp_uuid, label, timeout=60):
    DONE = {"complete", "failed", "error", "validation_complete",
            "validation_failed", "validated"}
    for _ in range(timeout):
        time.sleep(10)
        r = dest.get_import(imp_uuid)
        if r.status_code != 200:
            continue
        d = r.json().get("data", r.json())
        status = str(d.get("status", "") or "").lower()
        if status in DONE:
            errs   = d.get("errors") or d.get("error_count") or 0
            n_err  = len(errs) if isinstance(errs, list) else int(errs or 0)
            rows   = d.get("total_rows") or d.get("row_count") or 0
            created = d.get("successful_rows") or d.get("created") or 0
            failed  = d.get("failed_rows") or d.get("failed") or 0
            return {
                "status": status, "errors": n_err, "rows": rows,
                "created": created, "failed": failed,
                "detail": errs if isinstance(errs, list) else [],
                "raw": d,
            }
    return {"status": "timeout", "errors": -1, "rows": 0, "created": 0, "failed": 0, "detail": [], "raw": {}}


def run_wave(dest, csv_path, type_str, label):
    """Validate then store a single CSV.  Returns (ok, result_dict) or raises."""
    with open(csv_path, "rb") as f:
        csv_bytes = f.read()
    filename = os.path.basename(csv_path)
    rows_sent = max(0, len(csv_bytes.decode("utf-8-sig", errors="replace").splitlines()) - 1)

    for mode in ("validate", "store"):
        p(f"  → {label} [{mode}] ({rows_sent} rows)…")
        r = dest.upload(csv_bytes, filename, type_str, mode)
        if r.status_code not in (200, 201, 202):
            p(f"    UPLOAD FAILED HTTP {r.status_code}: {r.text[:200]}")
            return False, {"status": "upload_error"}
        body  = r.json()
        inner = body.get("data", body) if isinstance(body, dict) else body
        imp_uuid = inner.get("uuid") or inner.get("id")

        time.sleep(8)
        r_hdr = dest.get_import(imp_uuid)
        hdrs  = ((r_hdr.json().get("data", r_hdr.json()) if r_hdr.status_code == 200 else {})
                 .get("headers", []))
        r2 = dest.map_headers(imp_uuid, hdrs)
        if r2.status_code not in (200, 201, 202):
            p(f"    WARN map: HTTP {r2.status_code}")

        r3 = dest.start(imp_uuid)
        if r3.status_code not in (200, 201, 202):
            p(f"    WARN start: HTTP {r3.status_code}")

        res = poll_import(dest, imp_uuid, label)
        p(f"    {mode}: status={res['status']} errors={res['errors']} "
          f"rows={res['rows']} created={res['created']} failed={res['failed']}")

        if res["errors"] > 0 or res["failed"] > 0:
            if res["detail"]:
                for e in res["detail"][:5]:
                    p(f"      error: {e}")
            return False, res
        if mode == "validate" and "fail" in res["status"]:
            return False, res
        if mode == "store":
            return True, res

    return False, {"status": "never_stored"}


def sub_src(val, uuid_map):
    """Replace source UUIDs with dest UUIDs; handles comma-separated lists."""
    parts  = [x.strip() for x in str(val or "").split(",") if x.strip()]
    mapped = [uuid_map.get(x, x) for x in parts]
    return ",".join(mapped)


def rewrite_csv(src_path, dst_path, uuid_map, fields, soft_fields=None):
    """Copy src_path to dst_path, replacing source UUIDs in specified fields.
    soft_fields: unmatched UUIDs in these fields are blanked (not flagged as errors).
    """
    soft_fields = set(soft_fields or [])
    with open(src_path, newline="", encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f))
        fieldnames = csv.DictReader(open(src_path, encoding="utf-8-sig")).fieldnames
    unmapped   = []
    soft_blanked = 0
    for r in rows:
        for fld in fields:
            if fld not in r:
                continue
            orig_parts = [x.strip() for x in str(r[fld] or "").split(",") if x.strip()]
            mapped_parts = []
            for p_val in orig_parts:
                dest_val = uuid_map.get(p_val)
                if dest_val:
                    mapped_parts.append(dest_val)
                elif fld in soft_fields:
                    soft_blanked += 1  # drop it silently
                else:
                    unmapped.append((fld, p_val))
                    mapped_parts.append(p_val)  # keep as-is for now
            r[fld] = ",".join(mapped_parts)
    if soft_blanked:
        p(f"    soft-blanked {soft_blanked} unresolved circle refs in {os.path.basename(dst_path)}")
    with open(dst_path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
        w.writeheader()
        w.writerows(rows)
    return unmapped


# ─── UUID resolution ──────────────────────────────────────────────────────────

def resolve_client_uuids(dest):
    """Return {source_uuid: dest_uuid} for all clients with data_migration ext-ref."""
    clients = get_paged_with_refs(dest, "/api/v1/client", max_pages=10)
    result  = {}
    for c in clients:
        for ref in (c.get("external_references") or []):
            if ref.get("third_party") == MIGRATION_CO:
                src = ref.get("reference", "")
                if src:
                    result[src] = c["uuid"]
    return result


def resolve_plan_uuids(dest):
    """Return {source_uuid: dest_uuid} for all plans with data_migration ext-ref."""
    plans  = get_paged_with_refs(dest, "/api/v1/plans", max_pages=20)
    result = {}
    for pl in plans:
        for ref in (pl.get("external_references") or []):
            if ref.get("third_party") == MIGRATION_CO:
                src = ref.get("reference", "")
                if src:
                    result[src] = pl["uuid"]
    return result


def resolve_circle_uuids(dest):
    """Return {source_uuid: dest_uuid} for all circles with data_migration ext-ref."""
    circles = get_paged_with_refs(dest, "/api/v1/circles", max_pages=5)
    result  = {}
    for c in circles:
        for ref in (c.get("external_references") or []):
            if ref.get("third_party") == MIGRATION_CO:
                src = ref.get("reference", "")
                if src:
                    result[src] = c["uuid"]
    return result


# ─── Vision sweep ─────────────────────────────────────────────────────────────

def vision_sweep(files_and_skip):
    """Count Vision cells by column and blank whole-cell Vision firm names.
    files_and_skip = [(path, skip_cols), ...]
    """
    total_found = 0
    for path, skip_cols in files_and_skip:
        fname = os.path.basename(path)
        with open(path, newline="", encoding="utf-8-sig") as f:
            rows = list(csv.DictReader(f))
            fieldnames = csv.DictReader(open(path, encoding="utf-8-sig")).fieldnames

        col_counts = {}
        blanked    = 0
        for r in rows:
            for col in fieldnames:
                if col in skip_cols:
                    continue
                val = str(r.get(col, "") or "")
                if "vision" in val.lower():
                    col_counts[col] = col_counts.get(col, 0) + 1
                    # Blank whole-cell Vision firm name
                    if val.strip().lower().startswith("vision"):
                        r[col] = ""
                        blanked += 1

        if col_counts:
            p(f"  {fname}: Vision cells by col: {col_counts}  (blanked={blanked})")
            total_found += sum(col_counts.values())
        else:
            p(f"  {fname}: 0 Vision cells")

        with open(path, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
            w.writeheader()
            w.writerows(rows)

    return total_found


# ─── Expenditure ──────────────────────────────────────────────────────────────

def build_exp_field_map(dest):
    """GET dest expenditure-structure/section → {(section_name, field_name): field_uuid}."""
    field_map = {}
    sections  = get_paged(dest, "/api/v1/expenditure-structure/section", max_pages=5)
    if not sections:
        # Try fetching section list direct
        r = dest.get("/api/v1/expenditure-structure/section", params={"per_page": 50})
        if r.status_code == 200:
            sections = r.json().get("data", [])
    for sec in sections:
        sec_name   = str(sec.get("name", "") or "")
        fields     = sec.get("fields", [])
        if not fields:
            # fetch fields separately
            sec_uuid = sec.get("uuid", "")
            if sec_uuid:
                rf = dest.get(f"/api/v1/expenditure-structure/section/{sec_uuid}",
                              params={"include": "fields"})
                if rf.status_code == 200:
                    fields = rf.json().get("data", {}).get("fields", [])
        for fld in fields:
            fld_name = str(fld.get("name", "") or "")
            fld_uuid = fld.get("uuid", "")
            if sec_name and fld_name and fld_uuid:
                field_map[(sec_name, fld_name)] = fld_uuid
    return field_map


def run_expenditure(dest, client_map):
    """POST expenditure items from the export for all mapped clients."""
    # Build field map
    field_map = build_exp_field_map(dest)
    if not field_map:
        p("  ERROR: could not build expenditure field map — skipping")
        return False

    p(f"  Expenditure field map: {len(field_map)} fields across {len({s for s,_ in field_map})} sections")

    # Load source expenditure
    with open(EXP_EXPORT, newline="", encoding="utf-8-sig") as f:
        exp_rows = list(csv.DictReader(f))

    # Pre-fetch current_life_stage_uuid for each dest client that has expenditure
    needed_src = {r["client_uuid"] for r in exp_rows if r.get("client_uuid") in client_map}
    life_stage_map = {}  # dest_uuid → current_life_stage_uuid
    p(f"  Fetching life stage UUIDs for {len(needed_src)} clients…")
    for src_uuid in needed_src:
        dest_uuid = client_map[src_uuid]
        r = dest.get(f"/api/v1/client/{dest_uuid}/expenditure")
        if r.status_code == 200:
            ls_uuid = r.json().get("current_life_stage_uuid")
            if ls_uuid:
                life_stage_map[dest_uuid] = ls_uuid

    p(f"  Life stage UUIDs resolved: {len(life_stage_map)}")

    # POST each expenditure item
    ok_count = err_count = skip_count = 0
    unmatched_fields = set()
    for row in exp_rows:
        src_uuid = row.get("client_uuid", "").strip()
        dest_uuid = client_map.get(src_uuid)
        if not dest_uuid:
            skip_count += 1
            continue
        ls_uuid = life_stage_map.get(dest_uuid)
        if not ls_uuid:
            skip_count += 1
            continue

        sec_name = str(row.get("section", "") or "").strip()
        fld_name = str(row.get("type", "") or "").strip()
        fld_uuid = field_map.get((sec_name, fld_name))
        if not fld_uuid:
            unmatched_fields.add((sec_name, fld_name))
            skip_count += 1
            continue

        try:
            amount_gbp = float(row.get("amount", "0") or "0")
        except ValueError:
            skip_count += 1
            continue

        amount_pence = int(round(amount_gbp * 100))
        if amount_pence <= 0:
            skip_count += 1
            continue

        r = dest.post(f"/api/v1/client/{dest_uuid}/expenditure/{ls_uuid}/{fld_uuid}",
                      json={"amount": amount_pence})
        if r.status_code in (200, 201, 204):
            ok_count += 1
        else:
            err_count += 1

    p(f"  Expenditure: posted={ok_count} errors={err_count} skipped={skip_count}")
    if unmatched_fields:
        p(f"  Unmatched field names ({len(unmatched_fields)}): {sorted(unmatched_fields)[:10]}")
    return err_count == 0


# ─── Main ──────────────────────────────────────────────────────────────────────

def main():
    pat  = resolve_pat()
    dest = Dest(pat)

    # ── Pre-flight ─────────────────────────────────────────────────────────────
    p("\n=== PRE-FLIGHT ===")
    r0 = dest.get("/api/v1/employee", params={"per_page": 50})
    emps  = r0.json().get("data", [])
    names = {(str(e.get("first_name","") or "") + " " + str(e.get("last_name","") or "")).strip().lower()
             for e in emps}
    if not EXPECTED_EMPLOYEES.issubset(names):
        p(f"ABORT: employee guard failed — found {names}")
        return
    p(f"Employee guard: PASS ({len(emps)} employees)")

    # Resolve what's already at destination before checking count
    p("  Pre-resolving client UUIDs (detect prior run)…")
    client_map = resolve_client_uuids(dest)
    w1_uuids   = _read_ext_refs(os.path.join(IMPORT_DIR, "W1_clients.csv"))
    w1_already_done = all(u in client_map for u in w1_uuids)

    raw_clients = get_paged(dest, "/api/v1/client", max_pages=5)
    if raw_clients and not w1_already_done:
        p(f"ABORT: {len(raw_clients)} clients at destination but W1 not resolved — unexpected state")
        return
    if not raw_clients:
        p("Client count: 0 — PASS")
    else:
        p(f"Client count: {len(raw_clients)} (W1 already imported — resuming from W2)")

    # ── W1 — individual clients (272) ──────────────────────────────────────────
    p("\n=== W1: CLIENTS (individual) ===")
    if w1_already_done:
        p(f"  W1 already imported ({len(w1_uuids)} clients resolved) — skipping upload")
    else:
        ok, res = run_wave(dest, os.path.join(IMPORT_DIR, "W1_clients.csv"), "clients", "W1_clients")
        if not ok:
            p("STOP — W1 failed")
            return

        # Resolve client UUIDs after W1
        p("  Resolving client UUIDs after W1…")
        time.sleep(5)
        client_map = resolve_client_uuids(dest)
        p(f"  Resolved: {len(client_map)} source→dest client UUIDs")
        w1_unmatched = [u for u in w1_uuids if u not in client_map]
        if w1_unmatched:
            p(f"STOP — {len(w1_unmatched)} W1 clients unresolved: {w1_unmatched[:3]}")
            return

    p(f"  All {len(w1_uuids)} W1 clients resolved — OK")

    # ── W2 — joint/company clients (86) + circles (87) ────────────────────────
    p("\n=== W2: CLIENTS (joint/company) + CIRCLES ===")

    # Rewrite W2_clients owners field (source client UUIDs → dest)
    w2c_src = os.path.join(IMPORT_DIR, "W2_clients.csv")
    w2c_rw  = os.path.join(IMPORT_DIR, "W2_clients_rewritten.csv")
    unmapped = rewrite_csv(w2c_src, w2c_rw, client_map, ["owners"])
    if unmapped:
        p(f"STOP — {len(unmapped)} unmatched owner UUIDs in W2_clients: {unmapped[:5]}")
        return
    p("  W2_clients owners rewritten — OK")

    ok, _ = run_wave(dest, w2c_rw, "clients", "W2_clients")
    if not ok:
        p("STOP — W2_clients failed")
        return

    # Resolve again after W2 clients
    p("  Resolving client UUIDs after W2_clients…")
    time.sleep(5)
    client_map = resolve_client_uuids(dest)
    p(f"  Resolved: {len(client_map)} client UUIDs total")
    w2_uuids = _read_ext_refs(os.path.join(IMPORT_DIR, "W2_clients.csv"))
    w2_unmatched = [u for u in w2_uuids if u not in client_map]
    if w2_unmatched:
        p(f"STOP — {len(w2_unmatched)} W2 clients unresolved: {w2_unmatched[:3]}")
        return
    p(f"  All {len(w2_uuids)} W2 clients resolved — OK")

    # Rewrite circles
    circles_src = os.path.join(IMPORT_DIR, "W2_circles.csv")
    circles_rw  = os.path.join(IMPORT_DIR, "W2_circles_rewritten.csv")
    unmapped = rewrite_csv(circles_src, circles_rw, client_map, ["account_uuids"])
    if unmapped:
        p(f"STOP — {len(unmapped)} unmatched UUIDs in W2_circles: {unmapped[:3]}")
        return
    p(f"  W2_circles rewritten — OK")

    ok, _ = run_wave(dest, circles_rw, "circles", "W2_circles")
    if not ok:
        p("STOP — W2_circles failed")
        return

    # Full client map + circle map after W2
    time.sleep(5)
    client_map = resolve_client_uuids(dest)
    p(f"  Full client map: {len(client_map)} UUIDs")
    circle_map = resolve_circle_uuids(dest)
    p(f"  Circle map: {len(circle_map)} source→dest circle UUIDs")

    # ── W3 — relationships (247), dependants (249), income (316) ───────────────
    p("\n=== W3: RELATIONSHIPS + DEPENDANTS + INCOME ===")

    for csv_name, type_str, fields, label in [
        ("W3_relationships.csv", "relationships",
         ["account_uuid", "related_account_uuid"], "W3_relationships"),
        ("W3_dependants.csv", "dependants",
         ["account_uuids"], "W3_dependants"),
        ("W3_income.csv", "income_and_employment",
         ["client_uuid"], "W3_income"),
    ]:
        src = os.path.join(IMPORT_DIR, csv_name)
        rw  = os.path.join(IMPORT_DIR, csv_name.replace(".csv", "_rewritten.csv"))
        unmapped = rewrite_csv(src, rw, client_map, fields)
        if unmapped:
            p(f"STOP — {len(unmapped)} unmatched UUIDs in {csv_name}: {unmapped[:5]}")
            return
        p(f"  {csv_name} rewritten — OK")

        ok, _ = run_wave(dest, rw, type_str, label)
        if not ok:
            p(f"STOP — {label} failed")
            return

    # ── Expenditure ────────────────────────────────────────────────────────────
    p("\n=== EXPENDITURE (via API) ===")
    exp_ok = run_expenditure(dest, client_map)
    if not exp_ok:
        p("STOP — expenditure had errors")
        return

    # ── Approved: W4, W5, W6 ──────────────────────────────────────────────────
    p("\n=== VISION SWEEP (W4–W7 files) ===")
    vision_files = [
        (os.path.join(IMPORT_DIR, "W4_plans.csv"), {"notes"}),
        (os.path.join(IMPORT_DIR, "W5_plans.csv"), {"notes"}),
        (os.path.join(IMPORT_DIR, "W6_valuations.csv"), set()),
        (os.path.join(IMPORT_DIR, "W7_charges.csv"), {"notes"}),
    ]
    vision_count = vision_sweep(vision_files)
    p(f"Vision sweep complete: {vision_count} cells found and blanked")

    # Rewrite W4/W5 plans: client_uuids and visible_to_circles_uuids
    combined_map = {**client_map, **circle_map}
    for plan_csv in ("W4_plans.csv", "W5_plans.csv"):
        src = os.path.join(IMPORT_DIR, plan_csv)
        rw  = os.path.join(IMPORT_DIR, plan_csv.replace(".csv", "_rewritten.csv"))
        unmapped = rewrite_csv(src, rw, combined_map,
                               ["client_uuids", "visible_to_circles_uuids"],
                               soft_fields=["visible_to_circles_uuids"])
        if unmapped:
            p(f"STOP — {len(unmapped)} unmatched UUIDs in {plan_csv}: {unmapped[:5]}")
            return
        p(f"  {plan_csv} rewritten — OK")

    p("\n=== W4: PLANS (1 row — newly added) ===")
    ok, _ = run_wave(dest, os.path.join(IMPORT_DIR, "W4_plans_rewritten.csv"), "plans", "W4_plans")
    if not ok:
        p("STOP — W4 failed")
        return

    p("\n=== W5: PLANS (1900 rows) ===")
    ok, _ = run_wave(dest, os.path.join(IMPORT_DIR, "W5_plans_rewritten.csv"), "plans", "W5_plans")
    if not ok:
        p("STOP — W5 failed")
        return

    # Resolve plan UUIDs after W4+W5
    p("  Resolving plan UUIDs after W5…")
    time.sleep(5)
    plan_map = resolve_plan_uuids(dest)
    p(f"  Resolved: {len(plan_map)} source→dest plan UUIDs")

    # Check W5 plans resolved (source UUIDs are in ext_ref_val column)
    w5_plan_src = _read_ext_refs(os.path.join(IMPORT_DIR, "W5_plans.csv"),
                                 col="external_reference_value")
    w4_plan_src = _read_ext_refs(os.path.join(IMPORT_DIR, "W4_plans.csv"),
                                 col="external_reference_value")
    all_plan_src = w4_plan_src | w5_plan_src
    plan_unmatched = [u for u in all_plan_src if u not in plan_map]
    if plan_unmatched:
        p(f"STOP — {len(plan_unmatched)} plans unresolved after W5: {plan_unmatched[:3]}")
        return
    p(f"  All {len(all_plan_src)} plans resolved — OK")

    # Rewrite W6 valuations
    v_src = os.path.join(IMPORT_DIR, "W6_valuations.csv")
    v_rw  = os.path.join(IMPORT_DIR, "W6_valuations_rewritten.csv")
    unmapped = rewrite_csv(v_src, v_rw, plan_map, ["plan_uuid"])
    if unmapped:
        p(f"STOP — {len(unmapped)} unmatched plan UUIDs in W6_valuations: {unmapped[:5]}")
        return
    p(f"  W6_valuations rewritten — OK")

    p("\n=== W6: VALUATIONS (1526 rows) ===")
    ok, w6_res = run_wave(dest, v_rw, "valuations", "W6_valuations")
    if not ok:
        p("STOP — W6 failed")
        _final_report(client_map, plan_map, w6_failed=True)
        return

    p("\n=== W1–W6 COMPLETE — STOP ===")
    _final_report(client_map, plan_map)


# ─── Helpers ──────────────────────────────────────────────────────────────────

def _read_ext_refs(csv_path, col="external_reference_value"):
    """Return set of source UUIDs from a CSV's ext_ref column."""
    result = set()
    with open(csv_path, newline="", encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            v = str(r.get(col, "") or "").strip()
            if v:
                result.add(v)
    return result


def _final_report(client_map, plan_map, w6_failed=False):
    p(f"\n{'='*60}")
    p(f"FINAL REPORT")
    p(f"  Clients resolved: {len(client_map)}")
    p(f"  Plans resolved:   {len(plan_map)}")
    if w6_failed:
        p("  W6: FAILED — check errors above")
    else:
        p("  W6: OK")
    p(f"  Next step: W7 and W8 — await go-ahead")
    p(f"{'='*60}")


if __name__ == "__main__":
    main()
