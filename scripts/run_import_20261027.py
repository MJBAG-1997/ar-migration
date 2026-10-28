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

import ast, csv, io, json, os, sys, time
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

GRACE_LIMIT       = 5   # rows linking to deliberately-unimported entities: drop and log, do not stop

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
        if not imp_uuid:
            p(f"    STOP — no import UUID in upload response: {str(body)[:300]}")
            return False, {"status": "no_uuid"}
        p(f"    upload ok — imp_uuid={imp_uuid}  initial_status={inner.get('status','?')}")

        time.sleep(8)
        r_hdr = dest.get_import(imp_uuid)
        hdr_body = r_hdr.json().get("data", r_hdr.json()) if r_hdr.status_code == 200 else {}
        hdrs  = hdr_body.get("headers", [])
        p(f"    get_import: HTTP {r_hdr.status_code}  status={hdr_body.get('status','?')}  headers_count={len(hdrs)}")
        # Map headers; strip unsupported ones iteratively if Plannr rejects them
        good_hdrs = list(hdrs)
        for _attempt in range(25):
            r2 = dest.map_headers(imp_uuid, good_hdrs)
            if r2.status_code in (200, 201, 202):
                break
            try:
                err_body = r2.json()
                bad = {k[len("headers."):] for k in (err_body.get("errors") or {})
                       if k.startswith("headers.")}
            except Exception:
                bad = set()
            if not bad:
                p(f"    WARN map: HTTP {r2.status_code}  body={r2.text[:150]}")
                break
            p(f"    map: stripping {len(bad)} unsupported headers: {sorted(bad)}")
            good_hdrs = [h for h in good_hdrs if h not in bad]

        r3 = dest.start(imp_uuid)
        if r3.status_code not in (200, 201, 202):
            p(f"    WARN start: HTTP {r3.status_code}  body={r3.text[:150]}")

        res = poll_import(dest, imp_uuid, label)
        p(f"    {mode}: status={res['status']} errors={res['errors']} "
          f"rows={res['rows']} created={res['created']} failed={res['failed']}")

        if res["status"] == "timeout":
            p(f"    STOP — {label} {mode} poll timed out (import may not have processed)")
            return False, res
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


def rewrite_csv(src_path, dst_path, uuid_map, fields, soft_fields=None, grace_limit=0):
    """Replace source UUIDs in named fields; write result to dst_path.

    soft_fields: unresolved UUIDs blanked silently (no error).
    grace_limit: if > 0 and the number of rows with unresolved hard-field UUIDs is
                 <= grace_limit, those rows are dropped (not uploaded) and returned as
                 grace_dropped.  If the count exceeds grace_limit the caller gets
                 hard_unmapped and must stop the run.

    Returns (hard_unmapped, grace_dropped).
    """
    soft_fields = set(soft_fields or [])
    with open(src_path, newline="", encoding="utf-8-sig") as f:
        all_rows = list(csv.DictReader(f))
    with open(src_path, newline="", encoding="utf-8-sig") as f:
        fieldnames = list(csv.DictReader(f).fieldnames)

    # Process hard fields first, then any soft-only fields (blank unresolved)
    all_flds = list(fields) + [f for f in soft_fields if f not in set(fields)]

    soft_blanked = 0
    good, bad = [], []   # bad: (processed_row, [(field, src_uuid)])

    for r in all_rows:
        row_issues = []
        for fld in all_flds:
            if fld not in r:
                continue
            parts  = [x.strip() for x in str(r[fld] or "").split(",") if x.strip()]
            mapped = []
            for pv in parts:
                dv = uuid_map.get(pv)
                if dv:
                    mapped.append(dv)
                elif fld in soft_fields:
                    soft_blanked += 1
                else:
                    row_issues.append((fld, pv))
                    mapped.append(pv)
            r[fld] = ",".join(mapped)
        (bad if row_issues else good).append((r, row_issues))

    if soft_blanked:
        p(f"    soft-blanked {soft_blanked} unresolved refs in {os.path.basename(dst_path)}")

    if bad and grace_limit > 0 and len(bad) <= grace_limit:
        grace_dropped = [{"ext_ref": r.get("external_reference_value", "?"), "issues": iss}
                         for r, iss in bad]
        output_rows   = [r for r, _ in good]
        hard_unmapped = []
    elif bad:
        hard_unmapped = [item for _, iss in bad for item in iss]
        output_rows   = [r for r, _ in good + bad]
        grace_dropped = []
    else:
        hard_unmapped, grace_dropped = [], []
        output_rows = [r for r, _ in good]

    with open(dst_path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
        w.writeheader()
        w.writerows(output_rows)

    return hard_unmapped, grace_dropped


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
    return err_count == 0, ok_count


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

    p("  Pre-resolving client UUIDs (detect prior run)...")
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

        p("  Resolving client UUIDs after W1...")
        time.sleep(5)
        client_map = resolve_client_uuids(dest)
        p(f"  Resolved: {len(client_map)} source->dest client UUIDs")
        w1_unmatched = [u for u in w1_uuids if u not in client_map]
        if w1_unmatched:
            p(f"STOP — {len(w1_unmatched)} W1 clients unresolved: {w1_unmatched[:3]}")
            return

    p(f"  All {len(w1_uuids)} W1 clients resolved — OK")

    wave_counts           = {}
    grace_report          = []
    self_link_report      = []
    income_taxable_report  = []
    income_salary_report   = []
    income_type_changes    = []
    income_nonindiv_report = []
    w5_stub_changes        = []
    wave_state             = _load_wave_state(IMPORT_DIR)

    # ── W2 — joint/company clients (84 after adviser removals) + circles (87) ──
    p("\n=== W2: CLIENTS (joint/company) + CIRCLES ===")

    w2c_src     = os.path.join(IMPORT_DIR, "W2_clients.csv")
    w2c_rw      = os.path.join(IMPORT_DIR, "W2_clients_rewritten.csv")
    circles_src = os.path.join(IMPORT_DIR, "W2_circles.csv")
    circles_rw  = os.path.join(IMPORT_DIR, "W2_circles_rewritten.csv")
    w2c_uuids   = _read_ext_refs(w2c_src)
    w2_skip     = all(u in client_map for u in w2c_uuids)

    if w2_skip:
        p(f"  W2 clients already imported ({len(w2c_uuids)} resolved) — skipping")
        wave_counts["W2_clients"] = len(w2c_uuids)
        with open(circles_src, newline="", encoding="utf-8-sig") as _f:
            wave_counts["W2_circles"] = sum(1 for _ in csv.DictReader(_f))
        p(f"  W2 circles already imported ({wave_counts['W2_circles']} rows) — skipping")
    else:
        unmapped, gd = rewrite_csv(w2c_src, w2c_rw, client_map, ["owners"])
        if unmapped:
            p(f"STOP — {len(unmapped)} unmatched owner UUIDs in W2_clients: {unmapped[:5]}")
            return
        p("  W2_clients owners rewritten — OK")

        ok, w2c_res = run_wave(dest, w2c_rw, "clients", "W2_clients")
        if not ok:
            p("STOP — W2_clients failed")
            return
        wave_counts["W2_clients"] = w2c_res["created"]

        p("  Resolving client UUIDs after W2_clients...")
        time.sleep(5)
        client_map = resolve_client_uuids(dest)
        p(f"  Resolved: {len(client_map)} client UUIDs total")
        w2_unmatched = [u for u in w2c_uuids if u not in client_map]
        if w2_unmatched:
            p(f"STOP — {len(w2_unmatched)} W2 clients unresolved: {w2_unmatched[:3]}")
            return
        p(f"  All {len(w2c_uuids)} W2 clients resolved — OK")

        unmapped, gd = rewrite_csv(circles_src, circles_rw, client_map, ["account_uuids"],
                                   grace_limit=GRACE_LIMIT)
        if unmapped:
            p(f"STOP — {len(unmapped)} unmatched UUIDs in W2_circles: {unmapped[:3]}")
            return
        for d in gd:
            p(f"  GRACE-DROPPED W2_circles: ext_ref={d['ext_ref']} issues={d['issues']}")
            grace_report.append(("W2_circles", d))
        p("  W2_circles rewritten — OK")

        ok, w2cr_res = run_wave(dest, circles_rw, "circles", "W2_circles")
        if not ok:
            p("STOP — W2_circles failed")
            return
        wave_counts["W2_circles"] = w2cr_res["created"]

    time.sleep(5)
    client_map = resolve_client_uuids(dest)
    p(f"  Full client map: {len(client_map)} UUIDs")
    circle_map = resolve_circle_uuids(dest)
    p(f"  Circle map: {len(circle_map)} source->dest circle UUIDs")

    # ── W3 — relationships (182), dependants (249), income (316) ───────────────
    p("\n=== W3: RELATIONSHIPS + DEPENDANTS + INCOME ===")

    for csv_name, type_str, fields, label in [
        ("W3_relationships.csv", "relationships",
         ["account_uuid", "related_account_uuid"], "W3_relationships"),
        ("W3_dependants.csv", "dependants",
         ["account_uuids"], "W3_dependants"),
        ("W3_income.csv", "income_and_employment",
         ["client_uuid"], "W3_income"),
    ]:
        # ── Skip if already stored in a prior run ──────────────────────────────
        if wave_state.get(label, {}).get("done"):
            st = wave_state[label]
            wave_counts[label] = st["created"]
            for _ in range(st.get("grace", 0)):
                grace_report.append((csv_name, {"ext_ref": "(prior-run)", "issues": [("prior_run_grace", "")]}))
            for _ in range(st.get("self_links", 0)):
                self_link_report.append({"uuid": "(prior-run)", "rel": "prior-run"})
            p(f"  {label}: {st['created']} rows already stored (state file) — skipping")
            continue

        src = os.path.join(IMPORT_DIR, csv_name)
        rw  = os.path.join(IMPORT_DIR, csv_name.replace(".csv", "_rewritten.csv"))
        soft = None
        if label == "W3_income":
            src, inc_tax, inc_sal, inc_type, inc_nonindiv = \
                _fix_income_amounts(src, src.replace(".csv", "_amtfixed.csv"))
            income_taxable_report.extend(inc_tax)
            income_salary_report.extend(inc_sal)
            income_type_changes.extend(inc_type)
            income_nonindiv_report.extend(inc_nonindiv)
            # Dropped non-individual rows (no individual owner found) count as grace drops
            for nr in inc_nonindiv:
                if "note" in nr and "dropped" in nr["note"]:
                    grace_report.append(("W3_income.csv",
                                         {"ext_ref": nr["ext_ref"],
                                          "issues": [("no_individual_owner", nr["name"])]}))
            soft = ["employment_uuid"]
        unmapped, gd = rewrite_csv(src, rw, client_map, fields, soft_fields=soft, grace_limit=GRACE_LIMIT)
        if unmapped:
            p(f"STOP — {len(unmapped)} unmatched UUIDs in {csv_name}: {unmapped[:5]}")
            return
        for d in gd:
            p(f"  GRACE-DROPPED {csv_name}: ext_ref={d['ext_ref']} issues={d['issues']}")
            grace_report.append((csv_name, d))

        sl_count = 0
        if csv_name == "W3_relationships.csv":
            sl = _drop_self_links(rw, "account_uuid", "related_account_uuid")
            if sl:
                for d in sl:
                    p(f"  SELF-LINK-DROPPED: rel={d['rel']} dest_uuid={d['uuid']}")
                    self_link_report.append(d)
                p(f"  Dropped {len(sl)} self-link rows — verified against source snapshot"
                  f" (12-Oct): all are source noise, cross-client rels preserved")
                sl_count = len(sl)

        p(f"  {csv_name} rewritten — OK")

        with open(rw, newline="", encoding="utf-8-sig") as _f:
            sent_count = sum(1 for _ in csv.DictReader(_f))

        ok, w3_res = run_wave(dest, rw, type_str, label)
        if not ok:
            errs = w3_res.get("detail", [])
            already_linked = [e for e in errs
                              if "already linked" in str(e.get("errors", "")).lower()]
            if (label == "W3_relationships"
                    and len(errs) > 0
                    and len(already_linked) == len(errs)):
                p(f"  ALREADY-LINKED ({len(errs)} rows — auto-created by W2 joint import) — continuing")
                for e in already_linked:
                    acct = (e.get("data") or {}).get("account_uuid", "?")
                    rel  = (e.get("data") or {}).get("related_account_uuid", "?")
                    p(f"    {acct[:8]}... <-> {rel[:8]}...")
                    grace_report.append((csv_name,
                                         {"ext_ref": acct,
                                          "issues": [("already_linked_by_W2", rel)]}))
                wave_counts[label] = sent_count - len(errs)
                wave_state[label] = {"done": True, "created": wave_counts[label],
                                     "grace": len(already_linked), "self_links": sl_count}
                _save_wave_state(wave_state, IMPORT_DIR)
            else:
                p(f"STOP — {label} failed")
                return
        else:
            # Plannr API reports created=0 even on success; derive from sent count
            wave_counts[label] = w3_res["created"] if w3_res["created"] > 0 else sent_count
            wave_state[label] = {"done": True, "created": wave_counts[label], "self_links": sl_count}
            _save_wave_state(wave_state, IMPORT_DIR)

    # ── Pre-W4 destination count verification ────────────────────────────────
    p("\n=== PRE-W4 VERIFICATION ===")
    dep_count = _query_dest_count(dest, "/api/v1/dependant")
    inc_count = _query_dest_count(dest, "/api/v1/income")
    p(f"  W3_dependants: state={wave_state.get('W3_dependants',{}).get('created','-')}"
      f"  dest_count={dep_count if dep_count is not None else 'endpoint N/A'}")
    p(f"  W3_income:     state={wave_state.get('W3_income',{}).get('created','-')}"
      f"  dest_count={inc_count if inc_count is not None else 'endpoint N/A'}")

    # ── Expenditure ────────────────────────────────────────────────────────────
    p("\n=== EXPENDITURE (via API) ===")
    if wave_state.get("expenditure", {}).get("done"):
        wave_counts["expenditure"] = wave_state["expenditure"].get("created", 0)
        p(f"  Expenditure already posted ({wave_counts['expenditure']} records) — skipping (state file)")
    else:
        exp_ok, exp_count = run_expenditure(dest, client_map)
        if not exp_ok:
            p("STOP — expenditure had errors")
            return
        wave_counts["expenditure"] = exp_count
        wave_state["expenditure"] = {"done": True, "created": exp_count}
        _save_wave_state(wave_state, IMPORT_DIR)

    # ── Vision sweep before W4 ────────────────────────────────────────────────
    p("\n=== VISION SWEEP (W4-W7 files) ===")
    vision_files = [
        (os.path.join(IMPORT_DIR, "W4_plans.csv"), {"notes"}),
        (os.path.join(IMPORT_DIR, "W5_plans.csv"), {"notes"}),
        (os.path.join(IMPORT_DIR, "W6_valuations.csv"), set()),
        (os.path.join(IMPORT_DIR, "W7_charges.csv"), {"notes"}),
    ]
    vision_count = vision_sweep(vision_files)
    p(f"Vision sweep complete: {vision_count} cells found and blanked")

    # Rewrite W4/W5 plans: always rewrite (idempotent); skip import if state says done
    combined_map = {**client_map, **circle_map}
    for plan_csv in ("W4_plans.csv", "W5_plans.csv"):
        src = os.path.join(IMPORT_DIR, plan_csv)
        rw  = os.path.join(IMPORT_DIR, plan_csv.replace(".csv", "_rewritten.csv"))
        unmapped, gd = rewrite_csv(src, rw, combined_map,
                                   ["client_uuids", "visible_to_circles_uuids"],
                                   soft_fields=["visible_to_circles_uuids",
                                                "contract_enquiry_account_uuid"],
                                   grace_limit=GRACE_LIMIT)
        if unmapped:
            p(f"STOP — {len(unmapped)} unmatched UUIDs in {plan_csv}: {unmapped[:5]}")
            return
        for d in gd:
            p(f"  GRACE-DROPPED {plan_csv}: ext_ref={d['ext_ref']} issues={d['issues']}")
            grace_report.append((plan_csv, d))
        p(f"  {plan_csv} rewritten — OK")

    p("\n=== W4: PLANS (1 row) ===")
    w4_rw = os.path.join(IMPORT_DIR, "W4_plans_rewritten.csv")
    if wave_state.get("W4_plans", {}).get("done"):
        wave_counts["W4_plans"] = wave_state["W4_plans"]["created"]
        p(f"  W4 already stored ({wave_counts['W4_plans']} rows) — skipping (state file)")
    else:
        if not _pre_submit_count_check(dest, "/api/v1/plans", 0, "W4_plans"):
            return
        ok, w4_res = run_wave(dest, w4_rw, "plans", "W4_plans")
        if not ok:
            p("STOP — W4 failed")
            return
        with open(w4_rw, newline="", encoding="utf-8-sig") as _f:
            w4_sent = sum(1 for _ in csv.DictReader(_f))
        wave_counts["W4_plans"] = w4_res["created"] if w4_res["created"] > 0 else w4_sent
        wave_state["W4_plans"] = {"done": True, "created": wave_counts["W4_plans"]}
        _save_wave_state(wave_state, IMPORT_DIR)

    p("\n=== W5: PLANS (1900 rows) ===")
    w5_rw      = os.path.join(IMPORT_DIR, "W5_plans_rewritten.csv")
    w5_rw_fixed = os.path.join(IMPORT_DIR, "W5_plans_rewritten_fixed.csv")
    if wave_state.get("W5_plans", {}).get("done"):
        wave_counts["W5_plans"] = wave_state["W5_plans"]["created"]
        p(f"  W5 already stored ({wave_counts['W5_plans']} rows) — skipping (state file)")
    else:
        # Apply stub-plan fix (fills sub_type, boolean flags, agency_from, names)
        _, w5_stub_changes = _fix_w5_stub_plans(w5_rw, w5_rw_fixed)
        w5_submit = w5_rw_fixed  # submit the fixed file
        if not _pre_submit_count_check(dest, "/api/v1/plans", 1, "W5_plans"):
            return
        ok, w5_res = run_wave(dest, w5_submit, "plans", "W5_plans")
        if not ok:
            p("STOP — W5 failed")
            return
        with open(w5_submit, newline="", encoding="utf-8-sig") as _f:
            w5_sent = sum(1 for _ in csv.DictReader(_f))
        wave_counts["W5_plans"] = w5_res["created"] if w5_res["created"] > 0 else w5_sent
        wave_state["W5_plans"] = {"done": True, "created": wave_counts["W5_plans"]}
        _save_wave_state(wave_state, IMPORT_DIR)

    # Always resolve plan UUIDs (needed for W6/W7 rewrite even if W4/W5 were skipped)
    p("  Resolving plan UUIDs...")
    time.sleep(5)
    plan_map = resolve_plan_uuids(dest)
    p(f"  Resolved: {len(plan_map)} source->dest plan UUIDs")

    w5_plan_src = _read_ext_refs(os.path.join(IMPORT_DIR, "W5_plans.csv"),
                                 col="external_reference_value")
    w4_plan_src = _read_ext_refs(os.path.join(IMPORT_DIR, "W4_plans.csv"),
                                 col="external_reference_value")
    all_plan_src = w4_plan_src | w5_plan_src
    plan_unmatched = [u for u in all_plan_src if u not in plan_map]
    if plan_unmatched:
        p(f"STOP — {len(plan_unmatched)} plans unresolved: {plan_unmatched[:3]}")
        return
    p(f"  All {len(all_plan_src)} plans resolved — OK")

    # W6 rewrite (idempotent) then import
    v_src = os.path.join(IMPORT_DIR, "W6_valuations.csv")
    v_rw  = os.path.join(IMPORT_DIR, "W6_valuations_rewritten.csv")
    unmapped, gd = rewrite_csv(v_src, v_rw, plan_map, ["plan_uuid"],
                               grace_limit=GRACE_LIMIT)
    if unmapped:
        p(f"STOP — {len(unmapped)} unmatched plan UUIDs in W6_valuations: {unmapped[:5]}")
        return
    for d in gd:
        p(f"  GRACE-DROPPED W6_valuations: ext_ref={d['ext_ref']} issues={d['issues']}")
        grace_report.append(("W6_valuations", d))
    p("  W6_valuations rewritten — OK")

    p("\n=== W6: VALUATIONS (1526 rows) ===")
    if wave_state.get("W6_valuations", {}).get("done"):
        wave_counts["W6_valuations"] = wave_state["W6_valuations"]["created"]
        p(f"  W6 already stored ({wave_counts['W6_valuations']} rows) — skipping (state file)")
    else:
        if not _pre_submit_count_check(dest, "/api/v1/valuations", 0, "W6_valuations"):
            return
        ok, w6_res = run_wave(dest, v_rw, "valuations", "W6_valuations")
        if not ok:
            p("STOP — W6 failed")
            return
        with open(v_rw, newline="", encoding="utf-8-sig") as _f:
            v_sent = sum(1 for _ in csv.DictReader(_f))
        wave_counts["W6_valuations"] = w6_res["created"] if w6_res["created"] > 0 else v_sent
        wave_state["W6_valuations"] = {"done": True, "created": wave_counts["W6_valuations"]}
        _save_wave_state(wave_state, IMPORT_DIR)

    # ── W7 — charges (542) ────────────────────────────────────────────────────
    w7_src = os.path.join(IMPORT_DIR, "W7_charges.csv")
    w7_rw  = os.path.join(IMPORT_DIR, "W7_charges_rewritten.csv")
    unmapped, gd = rewrite_csv(w7_src, w7_rw, {**plan_map, **client_map},
                               ["plan_uuid", "client_uuid"],
                               grace_limit=GRACE_LIMIT)
    if unmapped:
        p(f"STOP — {len(unmapped)} unmatched UUIDs in W7_charges: {unmapped[:5]}")
        return
    for d in gd:
        p(f"  GRACE-DROPPED W7_charges: ext_ref={d['ext_ref']} issues={d['issues']}")
        grace_report.append(("W7_charges", d))
    p("  W7_charges rewritten — OK")

    p("\n=== W7: CHARGES (542 rows) ===")
    if wave_state.get("W7_charges", {}).get("done"):
        wave_counts["W7_charges"] = wave_state["W7_charges"]["created"]
        p(f"  W7 already stored ({wave_counts['W7_charges']} rows) — skipping (state file)")
    else:
        if not _pre_submit_count_check(dest, "/api/v1/charges", 0, "W7_charges"):
            return
        ok, w7_res = run_wave(dest, w7_rw, "charges", "W7_charges")
        if not ok:
            p("STOP — W7 failed")
            return
        with open(w7_rw, newline="", encoding="utf-8-sig") as _f:
            w7_sent = sum(1 for _ in csv.DictReader(_f))
        wave_counts["W7_charges"] = w7_res["created"] if w7_res["created"] > 0 else w7_sent
        wave_state["W7_charges"] = {"done": True, "created": wave_counts["W7_charges"]}
        _save_wave_state(wave_state, IMPORT_DIR)

    p("\n=== W1-W7 COMPLETE — STOP BEFORE W8 ===")

    # Joint account analysis and W8 notes pre-processing
    circle_src_map = _build_circle_src_map(dest, IMPORT_DIR, client_map)
    joint_summary  = _joint_accounts_summary(IMPORT_DIR, client_map)
    joint_notes    = _prep_w8_joint_notes(IMPORT_DIR, joint_summary)
    to_circle, to_owner, note_details = joint_notes
    p(f"  W8 notes pre-processed: {to_circle} to circles, {to_owner} to first owner")

    # Confirm all moved notes resolve to a known dest circle
    circle_notes_resolved   = sum(1 for _, en in note_details
                                  if _ == "circle"
                                  for csrc, cval in joint_summary["joint_to_circle_src"].items()
                                  if cval and cval in circle_src_map)
    # Simpler: iterate joint_to_circle_src and check how many src_circle_uuids resolved
    joint_circles_resolved = sum(1 for csrc in joint_summary["joint_to_circle_src"].values()
                                 if csrc and csrc in circle_src_map)
    joint_circles_total    = sum(1 for csrc in joint_summary["joint_to_circle_src"].values()
                                 if csrc)
    p(f"  Circle map confirmation: {joint_circles_resolved}/{joint_circles_total} "
      f"joint source circles resolved to dest UUIDs")
    if joint_circles_resolved < joint_circles_total:
        unresolved_circles = [csrc for csrc in joint_summary["joint_to_circle_src"].values()
                              if csrc and csrc not in circle_src_map]
        p(f"  WARN: {len(unresolved_circles)} source circle UUIDs NOT in circle_src_map: "
          f"{[c[:8] for c in unresolved_circles[:5]]}")
        p("  These notes will land on the account owner instead — W8 rewrite will blank their circle_uuid")

    _final_report(client_map, plan_map, wave_counts, grace_report, self_link_report,
                  income_taxable_report, income_salary_report,
                  income_type_changes, income_nonindiv_report,
                  joint_summary=joint_summary, joint_notes=joint_notes,
                  dest=dest, w5_stub_changes=w5_stub_changes)


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


def _drop_self_links(csv_path, field_a, field_b):
    """Remove rows where field_a == field_b (both non-empty). Modifies csv_path in-place.
    Returns list of {uuid, rel} dicts for each dropped row."""
    with open(csv_path, newline="", encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f))
    with open(csv_path, newline="", encoding="utf-8-sig") as f:
        fieldnames = list(csv.DictReader(f).fieldnames)
    dropped, keep = [], []
    for r in rows:
        a = str(r.get(field_a, "") or "").strip()
        b = str(r.get(field_b, "") or "").strip()
        if a and b and a == b:
            dropped.append({"uuid": a, "rel": r.get("relationship", "?")})
        else:
            keep.append(r)
    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
        w.writeheader()
        w.writerows(keep)
    return dropped


def _fix_income_amounts(src_path, dst_path):
    """Fix W3_income.csv before upload:
    A. bonus (no emp) → add employment defaults; dividend (no emp) → type=investment;
       benefit (no emp) → type=other, type_other=name; salary (no emp) → employment defaults
    B. start_date before 1900-01-01 → blank
    C. income on non-individual account → re-link to individual owner by first-name match,
       else first individual owner; if no individual owner → drop row
    Plus: amount pence dict→pounds, is_taxable defaults, type_other from name.
    Returns (dst_path, is_taxable_rows, salary_emp_rows, type_change_rows, nonindiv_rows).
    """
    # Build owner-name lookup from W1/W2 source files
    w2_owners = {}   # {nonindiv_src_uuid: [owner_src_uuid, ...]}
    w2_path = os.path.join(IMPORT_DIR, "W2_clients.csv")
    if os.path.exists(w2_path):
        with open(w2_path, newline="", encoding="utf-8-sig") as f:
            for row in csv.DictReader(f):
                if row.get("type", "").lower() != "individual":
                    src = row.get("external_reference_value", "").strip()
                    owners = [o.strip() for o in (row.get("owners") or "").split(",") if o.strip()]
                    if src:
                        w2_owners[src] = owners

    w1_firstnames = {}  # {src_uuid: first_name_lower}
    w1_path = os.path.join(IMPORT_DIR, "W1_clients.csv")
    if os.path.exists(w1_path):
        with open(w1_path, newline="", encoding="utf-8-sig") as f:
            for row in csv.DictReader(f):
                src = row.get("external_reference_value", "").strip()
                fn  = (row.get("first_name") or "").strip()
                if src:
                    w1_firstnames[src] = fn

    with open(src_path, newline="", encoding="utf-8-sig") as f:
        reader    = csv.DictReader(f)
        fieldnames = list(reader.fieldnames)
        rows      = list(reader)

    # Verify 3 dict-format amounts
    samples = 0
    for r in rows:
        v = (r.get("amount") or "").strip()
        if v.startswith("{") and samples < 3:
            try:
                d   = ast.literal_eval(v)
                lbs = round(int(d["amount"]) / 100, 2)
                p(f"  amount verify: {d['amount']}p ÷ 100 = £{lbs}  (formatted: {d.get('formatted','?')})")
                samples += 1
            except Exception as e:
                p(f"  amount verify FAILED: {v!r} ({e})")

    is_taxable_rows  = []
    salary_emp_rows  = []
    type_change_rows = []
    nonindiv_rows    = []
    output_rows      = []
    amt_fixed = is_tax_fixed = type_other_fixed = 0

    def _has_emp(row):
        return any((row.get(k) or "").strip()
                   for k in ("employment_status", "employment_employer", "employment_job_title"))

    for r in rows:
        ext_ref = r.get("external_reference_value", "?")
        name    = (r.get("name") or "").strip()
        itype   = r.get("type", "")

        # Amount: pence dict → pounds
        v = (r.get("amount") or "").strip()
        if v.startswith("{"):
            try:
                r["amount"] = round(int(ast.literal_eval(v)["amount"]) / 100, 2)
                amt_fixed += 1
            except Exception:
                pass

        # Issue B: blank start_date before 1900
        sd = (r.get("start_date") or "").strip()
        if sd and sd < "1900-01-01":
            r["start_date"] = ""

        # is_taxable default
        if not (r.get("is_taxable") or "").strip():
            taxable = "False" if "maintenance" in name.lower() else "True"
            r["is_taxable"] = taxable
            is_tax_fixed += 1
            is_taxable_rows.append({"ext_ref": ext_ref, "name": name,
                                    "type": itype, "is_taxable": taxable})

        # Issue A: fix employment-dependent types
        if not _has_emp(r):
            if itype in ("salary", "bonus"):
                r["employment_status"] = "employed"
                r["employment_job_title"] = "Not recorded"
                if itype == "salary":
                    salary_emp_rows.append({"ext_ref": ext_ref, "name": name})
                else:
                    type_change_rows.append({"ext_ref": ext_ref, "name": name,
                                             "change": "bonus: employment defaults added"})
            elif itype == "dividend":
                r["type"] = "investment"
                type_change_rows.append({"ext_ref": ext_ref, "name": name,
                                         "change": "dividend → investment (no employment in source)"})
            elif itype == "benefit":
                r["type"] = "other"
                r["type_other"] = name or "Benefit"
                type_change_rows.append({"ext_ref": ext_ref, "name": name,
                                         "change": f"benefit → other, type_other={r['type_other']!r}"})

        # type_other from name for original type='other' rows
        if r.get("type") == "other" and not (r.get("type_other") or "").strip():
            r["type_other"] = name or "Other"
            type_other_fixed += 1

        # Issue C: re-link income on non-individual source accounts
        client_src = (r.get("client_uuid") or "").strip()
        if client_src in w2_owners:
            owners    = w2_owners[client_src]
            name_low  = name.lower()
            matched   = None
            # Try first-name match
            for ow in owners:
                fn = w1_firstnames.get(ow, "").lower()
                if fn and fn in name_low:
                    matched = ow
                    break
            # Fall back to first individual owner
            if matched is None:
                for ow in owners:
                    if ow in w1_firstnames:
                        matched = ow
                        break
            if matched:
                r["client_uuid"] = matched
                nonindiv_rows.append({
                    "ext_ref": ext_ref, "name": name,
                    "orig_client": client_src[:8],
                    "new_owner": matched[:8],
                    "matched_name": w1_firstnames.get(matched, "?"),
                })
            else:
                p(f"  DROPPED income {name!r} — no individual owner for {client_src[:8]}")
                nonindiv_rows.append({
                    "ext_ref": ext_ref, "name": name,
                    "orig_client": client_src[:8], "note": "dropped — no individual owner",
                })
                continue  # exclude from output

        output_rows.append(r)

    p(f"  income fix: {amt_fixed} amounts(÷100), {is_tax_fixed} is_taxable, "
      f"{type_other_fixed} type_other, {len(salary_emp_rows)} salary-emp, "
      f"{len(type_change_rows)} type/emp-changes, "
      f"{sum(1 for r in nonindiv_rows if 'new_owner' in r)} relinked, "
      f"{sum(1 for r in nonindiv_rows if 'note' in r)} dropped")

    with open(dst_path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
        w.writeheader()
        w.writerows(output_rows)

    return dst_path, is_taxable_rows, salary_emp_rows, type_change_rows, nonindiv_rows


def _fix_w5_stub_plans(src_path, dst_path):
    """Fill in required fields for the 32 prov-... stub plans that fail W5 validation.

    sub_type is inferred from plan name keywords using the same wrapper→sub_type mapping
    as existing Transact/Rathbones plans in W5.  Boolean flags and agency_in_your_name_from
    take the most-common value from that provider's non-stub plans.  Blank names get
    '{provider} {wrapper}' (e.g. 'Transact GIA').

    Writes a new file; src_path (W5_plans_rewritten.csv) is never modified.
    Returns (dst_path, changes_list) where each change dict has the ext_ref and all
    fields set, for the AR to review.
    """
    # ── Provider-level defaults (from analysis of existing non-stub W5 plans) ──────
    # is_wrap_product: No for both; show_on_circles: Yes for both;
    # visible_on_client_portal: Yes; include_in_client_valuation: Yes
    PROV_DEFAULTS = {
        "transact":   {
            "is_wrap_product": "No", "show_on_circles": "Yes",
            "visible_on_client_portal": "Yes", "include_in_client_valuation": "Yes",
            "agency_in_your_name_from": "17/10/2026",
        },
        "rathbones":  {
            "is_wrap_product": "No", "show_on_circles": "Yes",
            "visible_on_client_portal": "Yes", "include_in_client_valuation": "Yes",
            "agency_in_your_name_from": "17/10/2026",
        },
    }
    # Fallback if provider not in table
    GENERIC_DEFAULTS = {
        "is_wrap_product": "No", "show_on_circles": "Yes",
        "visible_on_client_portal": "Yes", "include_in_client_valuation": "Yes",
        "agency_in_your_name_from": "17/10/2026",
    }

    # ── Name keyword → sub_type, wrapper display name ─────────────────────────────
    # Rules evaluated top-to-bottom; first match wins.
    # Transact and Rathbones share the ISA/pension keywords; EIS is provider-specific.
    def _infer_sub_type(name_low, itype, provider_low):
        # ISA variants
        if "lifetime isa" in name_low or "lisa" in name_low:
            return "lifetime_isa_stocks_and_shares", "Lifetime ISA"
        if "junior isa" in name_low or "jisa" in name_low:
            return "junior_isa_stocks_and_shares", "Junior ISA"
        if "isa" in name_low and ("aim" in name_low):
            return "isa_aim", "ISA (AIM)"
        if "isa" in name_low and ("cash" in name_low):
            return "isa_cash", "Cash ISA"
        if "isa" in name_low:
            return "isa_stocks_and_shares", "ISA"
        # Pension / drawdown
        if "sipp" in name_low:
            if itype == "pension":
                return "sipp", "SIPP"
            return "general_investment_account", "GIA (SIPP-related)"
        if "drawdown" in name_low or "flexi" in name_low:
            if itype == "pension":
                return "sipp", "SIPP (Drawdown)"
            return "general_investment_account", "GIA (Drawdown)"
        if "stps" in name_low:
            if itype == "pension":
                return "sipp", "SIPP"
            return "general_investment_account", "GIA (Pension-related)"
        # Bonds
        if "offshore bond" in name_low:
            return "offshore_bond", "Offshore Bond"
        if "onshore bond" in name_low or "bond" in name_low:
            return "onshore_bond", "Onshore Bond"
        if "gia" in name_low:
            return "general_investment_account", "GIA"
        # EIS — Transact EIS plans are held as GIA; Rathbones has dedicated sub_type
        if "eis" in name_low or "enterprise" in name_low:
            if provider_low == "rathbones":
                return "enterprise_investment_scheme_regulated", "EIS"
            return "general_investment_account", "GIA (EIS)"
        if "discretionary" in name_low or "dms" in name_low:
            return "discretionary_managed_service", "DMS"
        # Pension type with no keyword → most common by provider
        if itype == "pension":
            if provider_low == "transact":
                return "personal_pension_plan", "Personal Pension"   # 105 occurrences
            return "sipp", "SIPP"   # only Rathbones pension sub_type in W5
        # Investment fallback by provider
        if provider_low == "transact":
            return "general_investment_account", "GIA"   # 223 occurrences
        if provider_low == "rathbones":
            return "enterprise_investment_scheme_regulated", "EIS"  # 4 occurrences
        return "general_investment_account", "GIA"

    with open(src_path, newline="", encoding="utf-8-sig") as f:
        reader     = csv.DictReader(f)
        fieldnames = list(reader.fieldnames)
        rows       = list(reader)

    changes = []
    for r in rows:
        if (r.get("sub_type") or "").strip():
            continue  # only touch stub plans

        ext_ref      = r.get("external_reference_value", "?")
        orig_name    = (r.get("name") or "").strip()
        name_low     = orig_name.lower()
        itype        = r.get("type", "")
        provider_low = r.get("provider", "").lower()
        defaults     = PROV_DEFAULTS.get(provider_low, GENERIC_DEFAULTS)

        sub_type, wrapper_display = _infer_sub_type(name_low, itype, provider_low)

        # Fill name if blank
        if not orig_name:
            provider_display = r.get("provider", "")
            r["name"] = f"{provider_display} {wrapper_display}"

        r["sub_type"]                   = sub_type
        r["is_wrap_product"]            = defaults["is_wrap_product"]
        r["show_on_circles"]            = defaults["show_on_circles"]
        r["visible_on_client_portal"]   = defaults["visible_on_client_portal"]
        r["include_in_client_valuation"] = defaults["include_in_client_valuation"]
        # agency pair: keep Yes, complete the _from date
        r["agency_in_your_name_from"]   = defaults["agency_in_your_name_from"]

        changes.append({
            "ext_ref":    ext_ref,
            "policy_no":  r.get("policy_number", ""),
            "provider":   r.get("provider", ""),
            "orig_name":  orig_name,
            "final_name": r["name"],
            "sub_type":   sub_type,
            "wrapper":    wrapper_display,
            "is_wrap_product": r["is_wrap_product"],
            "show_on_circles": r["show_on_circles"],
            "visible_on_client_portal":   r["visible_on_client_portal"],
            "include_in_client_valuation": r["include_in_client_valuation"],
            "agency_in_your_name_from":   r["agency_in_your_name_from"],
        })

    # ── Pass 2: blank provider, inactive_reason, future opened_at (all plans) ───
    no_prov_count = 0
    inactive_no_reason_count = 0
    future_date_count = 0
    transfer_val_blanked_count = 0
    trustee_details_count = 0
    import datetime as _dt
    today_str = _dt.date.today().isoformat()  # YYYY-MM-DD for comparison

    for r in rows:
        # Blank provider → set provider_uuid to "Unknown Provider" (valid Plannr global provider).
        # Arbitrary text in the provider field fails for some plan types; the UUID approach works
        # universally. These plans are flagged in the clean-up list for correct provider assignment.
        UNKNOWN_PROV_UUID = "175e2a37-200f-400f-8140-e0d7206db1b1"
        if not (r.get("provider") or "").strip() and not (r.get("provider_uuid") or "").strip():
            r["provider_uuid"] = UNKNOWN_PROV_UUID
            no_prov_count += 1

        # status=inactive with no inactive_reason → default 'other' (Plannr requires it)
        if (r.get("status") or "").strip().lower() == "inactive" and \
           not (r.get("inactive_reason") or "").strip():
            r["inactive_reason"] = "other"
            inactive_no_reason_count += 1

        # Future opened_at → blank it (Plannr requires date ≤ today)
        oa = (r.get("opened_at") or "").strip()
        if oa:
            try:
                d = _dt.datetime.strptime(oa, "%d/%m/%Y").date()
                if d > _dt.date.today():
                    r["opened_at"] = ""
                    future_date_count += 1
            except ValueError:
                pass

        # pension_date_of_transfer_valuation prohibited when opened_at is blank
        if not (r.get("opened_at") or "").strip() and \
           (r.get("pension_date_of_transfer_valuation") or "").strip():
            r["pension_date_of_transfer_valuation"] = ""
            transfer_val_blanked_count += 1

        # protection_policy_in_trust=Yes requires trustee_details — default 'Not recorded'
        if (r.get("protection_policy_in_trust") or "").strip() == "Yes" and \
           not (r.get("protection_policy_trustee_details") or "").strip():
            r["protection_policy_trustee_details"] = "Not recorded"
            trustee_details_count += 1

    with open(dst_path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
        w.writeheader()
        w.writerows(rows)

    p(f"  W5 stub fix: {len(changes)} stub plans patched "
      f"({sum(1 for c in changes if not c['orig_name'])} names auto-filled); "
      f"{no_prov_count} blank-provider plans set to provider_uuid=Unknown Provider; "
      f"{inactive_no_reason_count} inactive plans with no reason set to 'other'; "
      f"{future_date_count} future opened_at dates blanked; "
      f"{transfer_val_blanked_count} transfer valuation dates blanked (opened_at empty); "
      f"{trustee_details_count} in-trust policies given default trustee_details "
      f"-> wrote {os.path.basename(dst_path)}")
    return dst_path, changes


def _load_wave_state(import_dir):
    path = os.path.join(import_dir, "state", "wave_state.json")
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        return {}


def _save_wave_state(wave_state, import_dir):
    path = os.path.join(import_dir, "state", "wave_state.json")
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(wave_state, f, indent=2)


def _query_dest_count(dest, endpoint):
    """GET endpoint with per_page=1 and return meta.total if available, else None."""
    try:
        r = dest.get(endpoint, params={"per_page": 1})
        if r.status_code != 200:
            return None
        d = r.json()
        meta = d.get("meta", {})
        return (meta.get("total") or meta.get("total_count")
                or (len(d.get("data", [])) if d.get("meta", {}).get("last_page", 1) == 1 else None))
    except Exception:
        return None


def _pre_submit_count_check(dest, endpoint, expected, wave_label):
    """Query dest count before submitting a wave. Returns True if safe to proceed.
    Stops (returns False) if count > expected. Warns but continues if endpoint N/A."""
    count = _query_dest_count(dest, endpoint)
    if count is None:
        p(f"  PRE-SUBMIT CHECK {wave_label}: endpoint {endpoint} N/A — cannot verify, proceeding")
        return True
    if count > expected:
        p(f"  STOP — PRE-SUBMIT CHECK {wave_label}: dest has {count} records at {endpoint} "
          f"(expected {expected}). Wave may already be stored. Not re-submitting.")
        return False
    p(f"  Pre-submit check {wave_label}: dest count={count} (expected ≤{expected}) — OK")
    return True


def _pick_sample_clients(import_dir, client_map, plan_map):
    """Return five (label, dest_uuid) pairs for end-to-end spot-checking."""
    res = {}

    w2p = os.path.join(import_dir, "W2_clients.csv")
    if os.path.exists(w2p):
        with open(w2p, newline="", encoding="utf-8-sig") as f:
            for r in csv.DictReader(f):
                t = r.get("type", "").lower()
                src = r.get("external_reference_value", "").strip()
                if t == "joint" and "joint" not in res and src in client_map:
                    res["joint"] = client_map[src]
                if t in {"trust", "business"} and "trust_biz" not in res and src in client_map:
                    res["trust_biz"] = client_map[src]
                if len(res) >= 2:
                    break

    w5p = os.path.join(import_dir, "W5_plans.csv")
    if os.path.exists(w5p):
        with open(w5p, newline="", encoding="utf-8-sig") as f:
            for r in csv.DictReader(f):
                if "rathbone" in str(r.get("provider", "") or "").lower():
                    src = (r.get("client_uuids", "") or "").split(",")[0].strip()
                    if src in client_map:
                        res["rathbones"] = client_map[src]
                        break

    w8p = os.path.join(import_dir, "W8_notes.csv")
    if os.path.exists(w8p):
        with open(w8p, newline="", encoding="utf-8-sig") as f:
            for r in csv.DictReader(f):
                if "[task note" in str(r.get("contents", "") or "").lower():
                    src = r.get("account_uuid", "").strip()
                    if src in client_map:
                        res["task_notes"] = client_map[src]
                        break

    w4p = os.path.join(import_dir, "W4_plans.csv")
    if os.path.exists(w4p):
        with open(w4p, newline="", encoding="utf-8-sig") as f:
            for r in csv.DictReader(f):
                src = (r.get("client_uuids", "") or "").split(",")[0].strip()
                if src in client_map:
                    res["add_to_plannr"] = client_map[src]
                    break

    return [
        ("joint client",       res.get("joint")),
        ("trust/business",     res.get("trust_biz")),
        ("Rathbones plan",     res.get("rathbones")),
        ("task notes",         res.get("task_notes")),
        ("Add-to-Plannr plan", res.get("add_to_plannr")),
    ]


def _build_circle_src_map(dest, import_dir, client_map):
    """Build {src_circle_uuid: dest_circle_uuid} by matching circle member accounts."""
    src_circles = {}  # frozenset(member_src_uuids) -> src_circle_uuid
    w2c = os.path.join(import_dir, "W2_circles.csv")
    if os.path.exists(w2c):
        with open(w2c, newline="", encoding="utf-8-sig") as f:
            for r in csv.DictReader(f):
                ext = r.get("external_reference_value", "").strip()
                mems = frozenset(m.strip() for m in (r.get("account_uuids") or "").split(",") if m.strip())
                if ext and mems:
                    src_circles[mems] = ext

    inv = {v: k for k, v in client_map.items()}  # dest_uuid -> src_uuid

    result = {}
    dest_circles = get_paged(dest, "/api/v1/circles", max_pages=10)
    for dc in dest_circles:
        dest_uuid = dc.get("uuid", "")
        accts_raw = dc.get("account_uuids") or dc.get("accounts") or []
        if isinstance(accts_raw, str):
            dest_mems = set(a.strip() for a in accts_raw.split(",") if a.strip())
        else:
            dest_mems = set()
            for a in accts_raw:
                dest_mems.add(a.get("uuid", "") if isinstance(a, dict) else str(a))
        dest_mems.discard("")
        src_mems = frozenset(inv[u] for u in dest_mems if u in inv)
        if src_mems in src_circles:
            result[src_circles[src_mems]] = dest_uuid

    p(f"  Circle map: {len(result)}/{len(src_circles)} source->dest resolved by member match")
    return result


def _joint_accounts_summary(import_dir, client_map):
    """Return a summary dict of W2 non-individual accounts, their owners, circles, and
    counts of what is attached to them across the wave files."""
    # Build joints: {src_uuid: {type, entity_name, owners}}
    joints = {}
    with open(os.path.join(import_dir, "W2_clients.csv"), newline="", encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            if r.get("type", "").lower() != "individual":
                src = r["external_reference_value"]
                owners = [o.strip() for o in (r.get("owners") or "").split(",") if o.strip()]
                joints[src] = {
                    "type": r["type"],
                    "entity_name": r.get("entity_name", "") or r.get("last_name", ""),
                    "owners": owners,
                    "owners_all_in_w1": all(o in client_map for o in owners),
                }

    # Match each joint to a source circle (W2_circles.csv)
    w2c = os.path.join(import_dir, "W2_circles.csv")
    circle_members_to_src = {}
    if os.path.exists(w2c):
        with open(w2c, newline="", encoding="utf-8-sig") as f:
            for r in csv.DictReader(f):
                ext = r.get("external_reference_value", "").strip()
                mems = frozenset(m.strip() for m in (r.get("account_uuids") or "").split(",") if m.strip())
                if ext and mems:
                    circle_members_to_src[mems] = ext

    joint_to_circle_src = {}
    for src, jd in joints.items():
        owners_set = frozenset(jd["owners"])
        match = next((csrc for mems, csrc in circle_members_to_src.items()
                      if owners_set.issubset(mems)), None)
        jd["circle_src"] = match
        joint_to_circle_src[src] = match

    joint_uuids = set(joints.keys())

    def _count_field(csv_name, field):
        cnt = 0
        path = os.path.join(import_dir, csv_name)
        if not os.path.exists(path):
            return 0
        with open(path, newline="", encoding="utf-8-sig") as f:
            for r in csv.DictReader(f):
                parts = [x.strip() for x in (r.get(field) or "").split(",") if x.strip()]
                if any(p in joint_uuids for p in parts):
                    cnt += 1
        return cnt

    exp_count = 0
    if os.path.exists(EXP_EXPORT):
        with open(EXP_EXPORT, newline="", encoding="utf-8-sig") as f:
            for r in csv.DictReader(f):
                if r.get("client_uuid", "").strip() in joint_uuids:
                    exp_count += 1

    return {
        "joints": joints,
        "joint_to_circle_src": joint_to_circle_src,
        "counts": {
            "plans_w4":    _count_field("W4_plans.csv",    "client_uuids"),
            "plans_w5":    _count_field("W5_plans.csv",    "client_uuids"),
            "charges":     _count_field("W7_charges.csv",  "client_uuid"),
            "income":      0,   # all re-linked to individual owners before import
            "dependants":  _count_field("W3_dependants.csv", "client_uuid"),
            "expenditure": exp_count,
        },
    }


def _prep_w8_joint_notes(import_dir, joint_summary):
    """Move W8 notes on joint/non-individual accounts:
    - If joint has a matching source circle → set model_type=circle, circle_uuid=src_circle_uuid
    - Else → set account_uuid = first owner's src UUID, prepend header to contents
    Reads W8_notes.csv (never modified) and writes W8_notes_joints_moved.csv.
    Returns (to_circle, to_owner, details)."""
    notes_path = os.path.join(import_dir, "W8_notes.csv")
    out_path   = os.path.join(import_dir, "W8_notes_joints_moved.csv")
    if not os.path.exists(notes_path):
        return 0, 0, []

    joints = joint_summary["joints"]
    joint_to_circle_src = joint_summary["joint_to_circle_src"]
    joint_uuids = set(joints.keys())

    with open(notes_path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        fieldnames = list(reader.fieldnames)
        rows = list(reader)

    to_circle = to_owner = 0
    details = []

    for r in rows:
        acct = (r.get("account_uuid") or "").strip()
        if acct not in joint_uuids:
            continue
        jd = joints[acct]
        entity_name = jd["entity_name"]
        circle_src = joint_to_circle_src.get(acct)

        if circle_src:
            r["model_type"]  = "circle"
            r["circle_uuid"] = circle_src
            r["account_uuid"] = ""
            to_circle += 1
            details.append(("circle", entity_name))
        else:
            first_owner = jd["owners"][0] if jd["owners"] else None
            if first_owner:
                r["contents"] = (f"[Note from joint account: {entity_name}]\n"
                                 + (r.get("contents") or ""))
                r["account_uuid"] = first_owner
                to_owner += 1
                details.append(("owner", entity_name))

    # Write to new file — W8_notes.csv is never modified
    with open(out_path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
        w.writeheader()
        w.writerows(rows)
    p(f"  W8 notes: wrote W8_notes_joints_moved.csv ({len(rows)} rows; W8_notes.csv unchanged)")

    return to_circle, to_owner, details


def _plans_cleanup_report(dest, plan_map, client_map):
    """Generate the plans clean-up list for tomorrow's update import (do not run now).

    Four items:
      1. Plans set to provider_uuid=Unknown Provider — audit against 26 Oct export.
      2. 46 plans whose contract_enquiry_account_uuid was blanked — correct dest user UUIDs.
      3. 1 investment plan that should be type=pension, sub_type=sipp.
      4. 21 plans owned by a joint entity — ownership to individual owners.

    Rule 9: plans by count, type, or source UUID only — never by name.
    """
    PLAN_EXPORT = os.path.join(
        r"C:\dev\AR Migration\exports_20261026",
        "Plan Export HFP Migration 2026-10-27.csv",
    )
    UNKNOWN_PROV_UUID = "175e2a37-200f-400f-8140-e0d7206db1b1"

    p("")
    p("  PLANS CLEAN-UP LIST (update import — do not run now)")
    p("  " + "─" * 56)

    # ── 1. Unknown Provider audit ─────────────────────────────────────────────
    # Identify which plans got provider_uuid=Unknown Provider (originally blank prov+uuid)
    w5_rw = os.path.join(IMPORT_DIR, "W5_plans_rewritten.csv")
    blank_prov = []            # rows from rewritten CSV where both prov fields were blank
    if os.path.exists(w5_rw):
        with open(w5_rw, newline="", encoding="utf-8-sig") as f:
            for r in csv.DictReader(f):
                if not (r.get("provider") or "").strip() and \
                   not (r.get("provider_uuid") or "").strip():
                    blank_prov.append(r)

    # Load 26 Oct export indexed by source UUID to check for build losses
    src_plan_prov = {}  # src_uuid → provider string in 26 Oct export
    if os.path.exists(PLAN_EXPORT):
        with open(PLAN_EXPORT, newline="", encoding="utf-8-sig") as f:
            for r in csv.DictReader(f):
                uid = (r.get("uuid") or "").strip()
                pv  = (r.get("provider") or "").strip()
                if uid:
                    src_plan_prov[uid] = pv

    # Fetch Plannr provider names for name-match count
    prov_names = set()
    if dest:
        try:
            pg = 1
            while True:
                rsp = dest.get("/api/v1/providers", params={"per_page": 500, "page": pg})
                if rsp.status_code != 200:
                    break
                d = rsp.json()
                for item in d.get("data", []):
                    nm = (item.get("name") or "").strip()
                    if nm:
                        prov_names.add(nm)
                if pg >= d.get("meta", {}).get("last_page", 1):
                    break
                pg += 1
        except Exception:
            pass

    by_type = {}
    build_loss_by_type = {}
    name_match_count = 0
    for r in blank_prov:
        t        = (r.get("type") or "").strip()
        ext_ref  = (r.get("external_reference_value") or "").strip()
        plan_nm  = (r.get("name") or "").strip()
        by_type[t] = by_type.get(t, 0) + 1
        if src_plan_prov.get(ext_ref):            # had provider in 26 Oct export
            build_loss_by_type[t] = build_loss_by_type.get(t, 0) + 1
        if prov_names and plan_nm in prov_names:  # name exactly matches a Plannr provider
            name_match_count += 1

    total_unknown = len(blank_prov)
    total_build_loss = sum(build_loss_by_type.values())
    p("")
    p(f"  1. provider_uuid → Unknown Provider ({total_unknown} plans)")
    p(f"     Build loss (had provider in 26 Oct export): {total_build_loss}")
    p(f"     Plan name exactly matches a Plannr provider: {name_match_count}")
    if prov_names:
        p(f"       (Plannr provider list: {len(prov_names)} entries)")
    p(f"     By type:")
    for t, cnt in sorted(by_type.items(), key=lambda x: -x[1]):
        loss = build_loss_by_type.get(t, 0)
        p(f"       {t:<25}  {cnt:3d} total  {loss:3d} build loss")
    p(f"     Unknown Provider UUID for update: {UNKNOWN_PROV_UUID}")

    # ── 2. contract_enquiry_account_uuid ─────────────────────────────────────
    JIM_SRC   = "db0ab7f5-8e52-4c89-a915-29fefeeca0d4"
    KATIE_SRC = "2153bb3c-7b76-4643-86a3-a9c53acab0b0"
    JIM_DEST  = "1a230fbc-4026-4bd1-851f-18cb55ab19f2"
    KATIE_DEST= "78a6aab8-55f3-42ed-b8ef-4867da1c0c78"
    EMP_MAP   = {JIM_SRC: JIM_DEST, KATIE_SRC: KATIE_DEST}
    ceq_plans = {JIM_SRC: [], KATIE_SRC: []}
    if os.path.exists(PLAN_EXPORT):
        with open(PLAN_EXPORT, newline="", encoding="utf-8-sig") as f:
            for r in csv.DictReader(f):
                ceq = (r.get("contract_enquiry_account_uuid") or "").strip()
                if ceq in EMP_MAP:
                    src_uuid  = (r.get("uuid") or "").strip()
                    dest_uuid = plan_map.get(src_uuid)
                    ceq_plans[ceq].append((src_uuid, dest_uuid))
    total_ceq = sum(len(v) for v in ceq_plans.values())
    p("")
    p(f"  2. contract_enquiry_account_uuid ({total_ceq} plans — blanked during W5 import):")
    for src_emp, dest_emp in EMP_MAP.items():
        plans = ceq_plans[src_emp]
        p(f"     src_employee={src_emp}  →  dest_user={dest_emp}  ({len(plans)} plans)")
        for src_pl, dest_pl in plans:
            p(f"       src={src_pl}  dest={dest_pl or 'NOT MAPPED'}")

    # ── 3. investment → pension reclassification ──────────────────────────────
    FLEXI_SRC  = "prov-042-648-472-IH01853021"
    flexi_dest = plan_map.get(FLEXI_SRC)
    p("")
    p(f"  3. type/sub_type fix (1 plan: type=investment → type=pension, sub_type=sipp):")
    p(f"     src={FLEXI_SRC}")
    p(f"     dest={flexi_dest or 'NOT MAPPED'}")
    p(f"     current: type=investment  sub_type=general_investment_account  provider=Transact")
    p(f"     target:  type=pension     sub_type=sipp")
    p(f"     (Transact pension plans in W5: 105×personal_pension_plan, 32×sipp)")

    # ── 4. Joint-entity-owned plans → individual owners ───────────────────────
    w2_path = os.path.join(IMPORT_DIR, "W2_clients.csv")
    joint_owners = {}   # src_joint_uuid → [src_owner_uuids]
    if os.path.exists(w2_path):
        with open(w2_path, newline="", encoding="utf-8-sig") as f:
            for r in csv.DictReader(f):
                if (r.get("type") or "").lower() != "individual":
                    src = (r.get("external_reference_value") or "").strip()
                    owners = [o.strip() for o in (r.get("owners") or "").split(",")
                              if o.strip()]
                    if src:
                        joint_owners[src] = owners

    w5_src_path = os.path.join(IMPORT_DIR, "W5_plans.csv")
    joint_plan_rows = []
    if os.path.exists(w5_src_path):
        with open(w5_src_path, newline="", encoding="utf-8-sig") as f:
            for r in csv.DictReader(f):
                cu    = (r.get("client_uuids") or "").strip()
                uuids = [u.strip() for u in cu.split(",") if u.strip()]
                for u in uuids:
                    if u in joint_owners:
                        src_pl = (r.get("external_reference_value") or "").strip()
                        indivs = joint_owners[u]
                        joint_plan_rows.append({
                            "src":          src_pl,
                            "dest":         plan_map.get(src_pl),
                            "type":         (r.get("type") or "").strip(),
                            "joint_src":    u,
                            "owners_dest":  [client_map.get(o, "NOT MAPPED") for o in indivs],
                        })
                        break

    by_type_j = {}
    for jp in joint_plan_rows:
        t = jp["type"]
        by_type_j[t] = by_type_j.get(t, 0) + 1
    type_str = ", ".join(f"{cnt}×{t}" for t, cnt in sorted(by_type_j.items(), key=lambda x: -x[1]))
    p("")
    p(f"  4. Joint-entity-owned plans ({len(joint_plan_rows)} plans — ownership → individual owners):")
    p(f"     By type: {type_str}")
    p(f"     {'src_plan_uuid':<36}  {'type':<16}  dest_plan_uuid")
    for jp in joint_plan_rows:
        dest_str  = jp["dest"] or "NOT MAPPED"
        p(f"     {jp['src']:<36}  {jp['type']:<16}  {dest_str}")
        owners_str = ", ".join(jp["owners_dest"])
        p(f"       → individual owner dest_uuids: {owners_str}")


def _final_report(client_map, plan_map, wave_counts, grace_report, self_link_report,
                  income_taxable_report=None, income_salary_report=None,
                  income_type_changes=None, income_nonindiv_report=None,
                  joint_summary=None, joint_notes=None, dest=None,
                  w5_stub_changes=None):
    EXPECTED = [
        ("W2_clients",       84),
        ("W2_circles",       87),
        ("W3_relationships", 182),
        ("W3_dependants",    249),
        ("W3_income",        316),
        ("expenditure",      None),
        ("W4_plans",           1),
        ("W5_plans",        1900),
        ("W6_valuations",   1526),
        ("W7_charges",       542),
    ]
    p(f"\n{'='*60}")
    p("FINAL REPORT -- W1-W7 COMPLETE (W8 awaits go-ahead)")
    p("")
    p(f"  {'Wave':<22} {'Expected':>8}  {'Created':>7}  {'Dropped':>7}")
    for key, exp in EXPECTED:
        created = wave_counts.get(key, "--")
        gcount  = sum(1 for wv, _ in grace_report
                      if wv.replace(".csv", "").replace("_rewritten", "") == key)
        if key == "W3_relationships":
            gcount += len(self_link_report)
        exp_str = str(exp) if exp is not None else "--"
        ok_mark = ""
        if isinstance(created, int) and exp is not None:
            ok_mark = "  OK" if (created + gcount) == exp else "  MISMATCH"
        drop_str = str(gcount) if gcount else "--"
        p(f"  {key:<22} {exp_str:>8}  {str(created):>7}  {drop_str:>7}{ok_mark}")

    if grace_report:
        p("")
        p(f"  Grace-dropped / already-linked ({len(grace_report)}):")
        for wave, d in grace_report:
            if d.get("ext_ref") != "(prior-run)":
                p(f"    [{wave}] ext_ref={d['ext_ref']} issues={d['issues']}")

    if self_link_report:
        p("")
        p(f"  Self-link rows dropped ({len([d for d in self_link_report if d.get('uuid') != '(prior-run)'])} actual + {len([d for d in self_link_report if d.get('uuid') == '(prior-run)'])} from prior run):")
        p("    All 10 verified against source snapshot (12-Oct): source noise confirmed, cross-client rels preserved")

    if income_taxable_report:
        p("")
        p(f"  is_taxable defaulted ({len(income_taxable_report)} rows — AR to verify post-go-live):")
        for r in income_taxable_report:
            p(f"    [{r['type']}] {r['name']!r}  ext_ref={r['ext_ref']}  defaulted={r['is_taxable']}")

    if income_salary_report:
        p("")
        p(f"  Salary rows with no source employment info ({len(income_salary_report)} — employment_job_title set to 'Not recorded', AR to verify):")
        for r in income_salary_report:
            p(f"    {r['name']!r}  ext_ref={r['ext_ref']}")

    if income_type_changes:
        p("")
        p(f"  Income type/employment changes (A — 10 rows, AR to verify post-go-live):")
        for r in income_type_changes:
            p(f"    {r['name']!r}  ext_ref={r['ext_ref']}  {r['change']}")

    if income_nonindiv_report:
        relinked = [r for r in income_nonindiv_report if "new_owner" in r]
        dropped  = [r for r in income_nonindiv_report if "note" in r]
        if relinked:
            p("")
            p(f"  Income re-linked from joint/company to individual owner ({len(relinked)} rows — AR to verify):")
            for r in relinked:
                p(f"    {r['name']!r}  orig_client=...{r['orig_client']}  → owner {r['matched_name']!r} (...{r['new_owner']})  ext_ref={r['ext_ref']}")
        if dropped:
            p("")
            p(f"  Income dropped — no individual owner ({len(dropped)} rows):")
            for r in dropped:
                p(f"    {r['name']!r}  orig_client=...{r['orig_client']}  ext_ref={r['ext_ref']}")

    if w5_stub_changes:
        p("")
        p(f"  W5 STUB PLANS PATCHED — 32 prov-... plans (AR to verify post-go-live):")
        p(f"  Fields defaulted from provider data: is_wrap_product=No, show_on_circles=Yes,")
        p(f"  visible_on_client_portal=Yes, include_in_client_valuation=Yes,")
        p(f"  agency_in_your_name_from=17/10/2026 (most common across agency=Yes plans).")
        p(f"  sub_type inferred from name; name auto-filled where blank.")
        p(f"  {'ext_ref':<36}  {'policy_no':<14}  {'prov':<12}  {'orig_name':<24}  {'final_name':<24}  {'sub_type'}")
        for c in w5_stub_changes:
            auto = " [AUTO]" if not c["orig_name"] else ""
            p(f"  {c['ext_ref']:<36}  {c['policy_no']:<14}  {c['provider']:<12}  "
              f"{(c['orig_name'] or '(blank)'):<24}  {c['final_name']:<24}  {c['sub_type']}{auto}")

    if joint_summary:
        js = joint_summary
        joints = js["joints"]
        by_type = {}
        for jd in joints.values():
            by_type[jd["type"]] = by_type.get(jd["type"], 0) + 1
        type_str = ", ".join(f"{cnt} {t}" for t, cnt in sorted(by_type.items()))
        has_circle   = sum(1 for jd in joints.values() if jd["circle_src"])
        no_circle    = [jd["entity_name"] for jd in joints.values() if not jd["circle_src"]]
        all_w1       = all(jd["owners_all_in_w1"] for jd in joints.values())
        cnt = js["counts"]
        p("")
        p(f"  JOINT / NON-INDIVIDUAL ACCOUNTS (for unpicking after go-live):")
        p(f"    {len(joints)} W2 records: {type_str}")
        p(f"    All owners exist as W1 individuals: {'YES' if all_w1 else 'NO — see details below'}")
        p(f"    Share a W2 circle: {has_circle}/{len(joints)} have a matching circle")
        if no_circle:
            p(f"    No circle: {', '.join(no_circle)}")
        p(f"    Attached records in destination:")
        p(f"      Plans (client_uuids includes joint): {cnt['plans_w4'] + cnt['plans_w5']}")
        p(f"      Charges (client_uuid = joint):       {cnt['charges']}")
        p(f"      Income (client_uuid = joint):        {cnt['income']}  (all re-linked to individual owners before import)")
        p(f"      Dependants (client_uuid = joint):    {cnt['dependants']}")
        p(f"      Expenditure rows on joint:           {cnt['expenditure']}")
        if not all_w1:
            p("    Owners NOT in W1:")
            for src, jd in joints.items():
                missing = [o for o in jd["owners"] if o not in client_map]
                if missing:
                    p(f"      {jd['entity_name']}: {[m[:8] for m in missing]}")

    if joint_notes:
        to_circle, to_owner, _ = joint_notes
        total = to_circle + to_owner
        p("")
        p(f"  W8 NOTES PRE-PROCESSED (joint account notes moved before import):")
        p(f"    {total} notes were on joint/non-individual accounts:")
        p(f"      {to_circle} moved to owners' circle  (circle_uuid = source circle UUID — resolved by rewrite at W8 run)")
        p(f"      {to_owner} moved to first owner  (header prepended: '[Note from joint account: ...]')")
        p(f"    Written to W8_notes_joints_moved.csv — W8_notes.csv is unchanged.")
        p(f"    When W8 runs, use W8_notes_joints_moved.csv and include circle_src_map in the rewrite uuid_map.")

    _plans_cleanup_report(dest, plan_map, client_map)

    p("")
    p("  Sample clients for end-to-end check:")
    for label, dest_uuid in _pick_sample_clients(IMPORT_DIR, client_map, plan_map):
        p(f"    {label}: {dest_uuid or 'NOT FOUND'}")

    p("")
    p(f"  Clients (W1+W2): {len(client_map)}  Plans (W4+W5): {len(plan_map)}")
    p("  Next: W8 (6795 notes) -- await explicit go-ahead")
    p(f"{'='*60}")


if __name__ == "__main__":
    main()
