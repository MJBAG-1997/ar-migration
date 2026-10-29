"""
W8 notes import and post-import read-only checks (2026-10-29).

Steps:
  1. Pre-check W8_notes_final_v2.csv: blank created_at rows; earliest/latest.
     STOP if any blank.
  2. Source all-notes fetch (all 356 clients) to find notes in fetch but not in file.
  3. Rewrite account_uuid + plan_uuid (source→dest) into a temp rewritten file.
  4. Submit temp file: POST /api/v1/import type=notes mode=store.
  5. Report created vs file.

Read-only checks (destination, no writes after step 5):
  6. Import history: GET /api/v1/import all pages.
  7. Dependants (v2 endpoint): site total + C1, C4, C6.
  8. Premiums: all 171 plans in W6_premiums.csv.
  9. Expenditure: all clients — non-zero source vs dest lines.
  10. Unknown Provider: plan count by type.
  11. under_advice raw check: 3 plans filed as 'no'.

Rule 9: no client names, no note contents. UUIDs, counts, types only.
Source is GET-only. No destination writes after step 5.
"""

import csv, io, json, os, sys, time
import requests

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# ── Constants ─────────────────────────────────────────────────────────────────
BASE_URL          = "https://api.plannrcrm.com"
KV_NAME           = "kv-plannr-ingestion"
SRC_PAT_SECRET    = "plannr-hfp-pat"
DEST_PAT_SECRET   = "plannr-prod-pat"
DEST_ACCOUNT_UUID = "ccd7f98b-d6a4-46ad-bab3-242ab42826eb"
MIGRATION_CO      = "data_migration"
IMPORT_DIR        = r"C:\dev\AR Migration\import_20261027"
EXPORTS_DIR       = r"C:\dev\AR Migration\exports_20261026"

EXPECTED_EMPLOYEES = {"jim hoban", "katie hoban", "tommie hoban", "valerie hoban"}

PROVIDER_REQUIRED_TYPES = {
    "annuity", "general_insurance", "investment", "mortgage",
    "pension", "protection_policy",
}

DEPENDANTS_CHECK = {
    "C1": "587bbc91-6c99-4114-a6c0-8416c3a68cb6",
    "C4": "c80da0e1-f368-4d66-b97a-6d88bd27e736",
    "C6": "26697c8a-e94b-4eeb-b021-d338793e63f7",
}

V2_PATH  = os.path.join(IMPORT_DIR, "W8_notes_final_v2.csv")
RW_PATH  = os.path.join(IMPORT_DIR, "W8_notes_final_v2_rewritten.csv")
WS_PATH  = os.path.join(IMPORT_DIR, "state", "wave_state.json")

# ── Helpers ───────────────────────────────────────────────────────────────────
def p(*a): print(*a, flush=True)

def resolve_pats():
    from azure.identity import DefaultAzureCredential
    from azure.keyvault.secrets import SecretClient
    sc = SecretClient(
        vault_url=f"https://{KV_NAME}.vault.azure.net/",
        credential=DefaultAzureCredential(),
    )
    return sc.get_secret(SRC_PAT_SECRET).value, sc.get_secret(DEST_PAT_SECRET).value

def load_json(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)

def load_csv(path):
    with open(path, newline="", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))

def load_wave_state():
    if os.path.exists(WS_PATH):
        with open(WS_PATH, encoding="utf-8") as f:
            return json.load(f)
    return {}

def save_wave_state(ws):
    with open(WS_PATH, "w", encoding="utf-8") as f:
        json.dump(ws, f, indent=2)


# ── API clients ───────────────────────────────────────────────────────────────
class SourceClient:
    def __init__(self, pat):
        self._s = requests.Session()
        self._s.headers.update({
            "Authorization": f"Bearer {pat}",
            "Accept":        "application/json",
        })
        r     = self._s.get(f"{BASE_URL}/api/v1/me")
        inner = r.json().get("data", r.json())
        acct  = (inner.get("current_account") or {})
        uuid  = acct.get("uuid") or inner.get("uuid")
        self._s.headers["X-PLANNR-ACCOUNT-UUID"] = uuid
        p(f"  Source account UUID: {uuid}")

    def get(self, path, **kw):
        return self._s.get(f"{BASE_URL}{path}", **kw)


class Dest:
    def __init__(self, pat):
        self._s = requests.Session()
        self._s.headers.update({
            "Authorization":          f"Bearer {pat}",
            "Accept":                 "application/json",
            "X-PLANNR-ACCOUNT-UUID":  DEST_ACCOUNT_UUID,
        })

    def get(self, path, **kw):
        return self._s.get(f"{BASE_URL}{path}", **kw)

    def post(self, path, **kw):
        return self._s.post(f"{BASE_URL}{path}", **kw)

    def upload(self, csv_bytes, filename, type_str, mode):
        return self._s.post(
            f"{BASE_URL}/api/v1/import",
            files={"file": (filename, io.BytesIO(csv_bytes), "text/csv")},
            data={"type": type_str, "mode": mode},
        )

    def map_headers(self, imp_uuid, headers):
        return self._s.post(
            f"{BASE_URL}/api/v1/import/map/{imp_uuid}",
            json={"headers": {h: h for h in headers}},
        )

    def start(self, imp_uuid):
        return self._s.post(f"{BASE_URL}/api/v1/import/start/{imp_uuid}")

    def get_import(self, imp_uuid):
        return self._s.get(f"{BASE_URL}/api/v1/import/{imp_uuid}")


# ── Pagination ────────────────────────────────────────────────────────────────
def get_paged(client, path, max_pages=500, params=None):
    items, page = [], 1
    while page <= max_pages:
        pr = dict(params or {}); pr.update({"page": page, "per_page": 200})
        r = client.get(path, params=pr)
        if r.status_code != 200:
            break
        body = r.json()
        data = body.get("data", [])
        if isinstance(data, dict): data = [data]
        items.extend(data)
        meta = body.get("meta", {}) if isinstance(body, dict) else {}
        if page >= (meta.get("last_page") or 1):
            break
        page += 1
    return items

def get_paged_meta_total(client, path, params=None):
    """Single-page request; return (items, total) using meta.total."""
    pr = dict(params or {}); pr.update({"page": 1, "per_page": 200})
    r = client.get(path, params=pr)
    if r.status_code != 200:
        return [], 0
    body = r.json()
    data = body.get("data", [])
    if isinstance(data, dict): data = [data]
    meta  = body.get("meta", {}) if isinstance(body, dict) else {}
    total = meta.get("total", len(data))
    return data, total


# ── Import poll ───────────────────────────────────────────────────────────────
DONE_STATUSES = {
    "complete", "failed", "error",
    "validation_complete", "validation_failed", "validated",
}

def poll_import(dest, imp_uuid, label, timeout=90):
    for _ in range(timeout):
        time.sleep(10)
        r = dest.get_import(imp_uuid)
        if r.status_code != 200:
            continue
        d      = r.json().get("data", r.json())
        status = str(d.get("status", "") or "").lower()
        if status in DONE_STATUSES:
            errs    = d.get("errors") or d.get("error_count") or 0
            n_err   = len(errs) if isinstance(errs, list) else int(errs or 0)
            rows    = d.get("total_rows") or d.get("row_count") or 0
            created = d.get("successful_rows") or d.get("created") or 0
            failed  = d.get("failed_rows") or d.get("failed") or 0
            return {
                "status": status, "errors": n_err, "rows": rows,
                "created": created, "failed": failed,
                "detail": errs if isinstance(errs, list) else [],
                "raw": d,
            }
    return {"status": "timeout", "errors": -1, "rows": 0, "created": 0, "failed": 0, "detail": [], "raw": {}}


# ── W8 submit ─────────────────────────────────────────────────────────────────
def run_notes(dest, csv_path):
    """Submit notes CSV in store mode (no validate — notes importer validates on store)."""
    with open(csv_path, "rb") as f:
        csv_bytes = f.read()
    filename  = os.path.basename(csv_path)
    rows_sent = max(0, len(csv_bytes.decode("utf-8-sig", errors="replace").splitlines()) - 1)

    p(f"  → notes [store] ({rows_sent} rows)…")
    r = dest.upload(csv_bytes, filename, "notes", "store")
    if r.status_code not in (200, 201, 202):
        p(f"    UPLOAD FAILED HTTP {r.status_code}: {r.text[:300]}")
        return False, {"status": "upload_error"}

    body     = r.json()
    inner    = body.get("data", body) if isinstance(body, dict) else body
    imp_uuid = inner.get("uuid") or inner.get("id")
    if not imp_uuid:
        p(f"    STOP — no import UUID in upload response: {str(body)[:300]}")
        return False, {"status": "no_uuid"}
    p(f"    upload ok — imp_uuid={imp_uuid}  initial_status={inner.get('status','?')}")

    time.sleep(8)
    r_hdr    = dest.get_import(imp_uuid)
    hdr_body = r_hdr.json().get("data", r_hdr.json()) if r_hdr.status_code == 200 else {}
    hdrs     = hdr_body.get("headers", [])
    p(f"    get_import: HTTP {r_hdr.status_code}  status={hdr_body.get('status','?')}  headers={len(hdrs)}")

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

    res = poll_import(dest, imp_uuid, "notes")
    p(f"    store: status={res['status']} errors={res['errors']} "
      f"rows={res['rows']} created={res['created']} failed={res['failed']}")

    if res["status"] == "timeout":
        p("    STOP — poll timed out (import may still be running)")
        return False, res
    if res["errors"] > 0 or res["failed"] > 0:
        if res["detail"]:
            for e in res["detail"][:10]:
                p(f"      error: {e}")
        p("    STOP — errors/failures in store")
        return False, res
    return True, res


# ── Main ──────────────────────────────────────────────────────────────────────
def main():

    # ── Step 1: Pre-check ────────────────────────────────────────────────────
    p("\n=== STEP 1: PRE-CHECK W8_notes_final_v2.csv ===")
    v2_rows = load_csv(V2_PATH)
    blanks  = [r for r in v2_rows if not (r.get("created_at") or "").strip()]
    dates   = sorted(r["created_at"] for r in v2_rows if (r.get("created_at") or "").strip())
    p(f"  Total rows:       {len(v2_rows)}")
    p(f"  Blank created_at: {len(blanks)}")
    if blanks:
        p("  Blank rows (first 5):")
        for br in blanks[:5]:
            p(f"    ext_ref={br.get('external_reference_value','?')}  acct={br.get('account_uuid','?')}")
        p("STOP — blank created_at rows found")
        return
    p(f"  Earliest created_at: {dates[0]}")
    p(f"  Latest created_at:   {dates[-1]}")
    p("  Pre-check PASS")

    # ── Setup ────────────────────────────────────────────────────────────────
    p("\n=== SETUP: KEYS + EMPLOYEE GUARD ===")
    src_pat, dest_pat = resolve_pats()
    p(f"  KV: OK — src len={len(src_pat)}, dest len={len(dest_pat)}")
    src  = SourceClient(src_pat)
    dest = Dest(dest_pat)

    r0    = dest.get("/api/v1/employee", params={"per_page": 50})
    emps  = r0.json().get("data", [])
    names = {((e.get("first_name") or "") + " " + (e.get("last_name") or "")).strip().lower()
             for e in emps}
    if not EXPECTED_EMPLOYEES.issubset(names):
        p(f"ABORT: employee guard failed — found {names}")
        return
    p(f"  Employee guard: PASS ({len(emps)} employees)")

    client_map = load_json(os.path.join(IMPORT_DIR, "state", "client_map.json"))
    plan_map   = load_json(os.path.join(IMPORT_DIR, "state", "plan_map.json"))
    p(f"  client_map: {len(client_map)}, plan_map: {len(plan_map)}")

    # ── Step 2: Source all-notes fetch ───────────────────────────────────────
    p("\n=== STEP 2: SOURCE ALL-NOTES FETCH (find notes not in file) ===")
    src_uuids    = list(client_map.keys())
    note_meta    = {}   # {note_uuid: created_at}
    appearances  = 0

    for i, su in enumerate(src_uuids, 1):
        notes = get_paged(src, f"/api/v1/client/{su}/all-notes", max_pages=500)
        for note in notes:
            nuid = str(note.get("uuid") or "").strip()
            if nuid and nuid not in note_meta:
                note_meta[nuid] = str(note.get("created_at") or "")
        appearances += len(notes)
        if i % 50 == 0 or i == len(src_uuids):
            p(f"  {i}/{len(src_uuids)} clients — {appearances} appearances, "
              f"{len(note_meta)} unique note UUIDs")

    v2_ext_refs = {
        (r.get("external_reference_value") or "").strip()
        for r in v2_rows
        if (r.get("external_reference_value") or "").strip()
    }
    missing_in_file = {uid: note_meta[uid] for uid in note_meta if uid not in v2_ext_refs}

    p(f"\n  Notes in source fetch:     {len(note_meta)}")
    p(f"  Notes in v2 file:          {len(v2_ext_refs)}")
    p(f"  In fetch but NOT in file:  {len(missing_in_file)}")
    if missing_in_file:
        p("  (UUID, created_at) — Rule 9: UUIDs only, no contents:")
        for uid, cat in sorted(missing_in_file.items(), key=lambda x: x[1]):
            p(f"    {uid}  created={cat}")

    # ── Step 3: Rewrite UUIDs ────────────────────────────────────────────────
    p("\n=== STEP 3: REWRITE SOURCE UUIDs → DEST ===")
    with open(V2_PATH, newline="", encoding="utf-8-sig") as f:
        reader     = csv.DictReader(f)
        fieldnames = list(reader.fieldnames)
        all_rows   = list(reader)

    unmapped  = []
    out_rows  = []
    for row in all_rows:
        new = dict(row)

        # account_uuid: source client → dest client (required)
        src_acct = (row.get("account_uuid") or "").strip()
        dst_acct = client_map.get(src_acct)
        if not dst_acct:
            unmapped.append(("account_uuid", src_acct))
            out_rows.append(new)
            continue
        new["account_uuid"] = dst_acct

        # plan_uuid: source plan → dest plan (optional — only model_type=plan rows)
        src_plan = (row.get("plan_uuid") or "").strip()
        if src_plan:
            dst_plan = plan_map.get(src_plan)
            if not dst_plan:
                unmapped.append(("plan_uuid", src_plan))
                out_rows.append(new)
                continue
            new["plan_uuid"] = dst_plan

        out_rows.append(new)

    if unmapped:
        p(f"STOP — {len(unmapped)} unmatched UUIDs:")
        for fld, uid in unmapped[:10]:
            p(f"  {fld}: {uid}")
        return

    with open(RW_PATH, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
        w.writeheader()
        w.writerows(out_rows)
    p(f"  Rewrite OK — {len(out_rows)} rows → {os.path.basename(RW_PATH)}")
    p(f"  account_uuid sample: {out_rows[0].get('account_uuid','')[:8]}…")

    # ── Step 4: Submit ───────────────────────────────────────────────────────
    p("\n=== STEP 4: SUBMIT W8 NOTES ===")
    ok, res = run_notes(dest, RW_PATH)
    if not ok:
        p("STOP — W8 submit failed (see errors above)")
        return

    file_rows = len(out_rows)
    p(f"\n  W8 REPORT:")
    p(f"  File rows sent:  {file_rows}")
    p(f"  Importer rows:   {res['rows']}")
    p(f"  Created:         {res['created']}")
    p(f"  Failed:          {res['failed']}")
    if res["created"] == file_rows:
        p(f"  ✓ created matches file ({file_rows})")
    else:
        p(f"  !! MISMATCH: created {res['created']} vs file {file_rows}")

    # Update wave state
    ws = load_wave_state()
    ws["W8_notes"] = {"done": True, "created": res["created"], "file": "W8_notes_final_v2.csv"}
    save_wave_state(ws)
    p("  Wave state updated: W8_notes done")

    # ═══════════════════════════════════════════════════════════════════════════
    # READ-ONLY CHECKS (no destination writes after this point)
    # ═══════════════════════════════════════════════════════════════════════════

    # ── Check 6: Import history ───────────────────────────────────────────────
    p("\n=== CHECK 6: IMPORT HISTORY ===")
    imports = get_paged(dest, "/api/v1/import", max_pages=20)
    p(f"  Total import jobs: {len(imports)}")
    p(f"  {'type':<20} {'mode':<10} {'status':<25} {'sent':>6} {'created':>8} {'failed':>7}  date")
    p(f"  {'-'*20} {'-'*10} {'-'*25} {'-'*6} {'-'*8} {'-'*7}  {'-'*10}")
    notes_ever_stored = False
    for imp in imports:
        itype   = str(imp.get("type") or "")
        imode   = str(imp.get("mode") or "")
        istatus = str(imp.get("status") or "")
        irows   = imp.get("total_rows") or imp.get("row_count") or 0
        icreated= imp.get("successful_rows") or imp.get("created") or 0
        ifailed = imp.get("failed_rows") or imp.get("failed") or 0
        idate   = str(imp.get("created_at") or "")[:10]
        p(f"  {itype:<20} {imode:<10} {istatus:<25} {irows:>6} {icreated:>8} {ifailed:>7}  {idate}")
        if itype == "notes" and imode == "store" and "complete" in istatus.lower():
            notes_ever_stored = True
        if itype == "premiums" and imode == "store" and "complete" in istatus.lower():
            p(f"    ^ premiums store completed")
    prem_stored = any(
        str(imp.get("type") or "") == "premiums"
        and "complete" in str(imp.get("status") or "").lower()
        for imp in imports
    )
    p(f"\n  Premiums ever stored: {'YES' if prem_stored else 'NO — never stored'}")

    # ── Check 7: Dependants (v2 endpoint) ────────────────────────────────────
    p("\n=== CHECK 7: DEPENDANTS (GET /api/v2/dependants) ===")
    # Site total
    r_all = dest.get("/api/v2/dependants", params={"per_page": 1, "page": 1})
    if r_all.status_code == 200:
        meta_all   = r_all.json().get("meta", {}) if isinstance(r_all.json(), dict) else {}
        site_total = meta_all.get("total", len(r_all.json().get("data", [])))
    else:
        site_total = f"HTTP {r_all.status_code}"
    p(f"  Site total dependants: {site_total}  (expected 249)")

    for label, duuid in DEPENDANTS_CHECK.items():
        r_dep = dest.get("/api/v2/dependants", params={
            "filter[account_uuids]": duuid,
            "per_page": 200, "page": 1,
        })
        if r_dep.status_code == 200:
            body    = r_dep.json()
            data    = body.get("data", [])
            meta_d  = body.get("meta", {}) if isinstance(body, dict) else {}
            cnt     = meta_d.get("total", len(data))
        else:
            cnt = f"HTTP {r_dep.status_code}"
        p(f"  {label} ({duuid[:8]}…): {cnt}")

    # ── Check 8: Premiums ────────────────────────────────────────────────────
    p("\n=== CHECK 8: PREMIUMS ===")
    w6p = load_csv(os.path.join(IMPORT_DIR, "W6_premiums.csv"))
    p(f"  W6_premiums.csv rows: {len(w6p)}")
    total_prem_actual = 0
    plans_with_prem   = 0
    plans_no_dest_uuid = 0
    plans_api_error    = 0
    for pr_row in w6p:
        src_plan = (pr_row.get("plan_uuid") or "").strip()
        dst_plan = plan_map.get(src_plan)
        if not dst_plan:
            plans_no_dest_uuid += 1
            continue
        r_pr = dest.get(f"/api/v1/plans/{dst_plan}/premiums",
                        params={"per_page": 200, "page": 1})
        if r_pr.status_code == 200:
            pr_data = r_pr.json().get("data", [])
            if isinstance(pr_data, dict): pr_data = [pr_data]
            meta_pr = r_pr.json().get("meta", {})
            cnt_pr  = meta_pr.get("total", len(pr_data))
            total_prem_actual += cnt_pr
            if cnt_pr > 0:
                plans_with_prem += 1
        else:
            plans_api_error += 1

    p(f"  Plans in W6_premiums:    {len(w6p)}")
    p(f"  Plans unmapped to dest:  {plans_no_dest_uuid}")
    p(f"  Plans with API error:    {plans_api_error}")
    p(f"  Plans with premiums > 0: {plans_with_prem}")
    p(f"  Total premiums actual:   {total_prem_actual}  (expected 171)")

    # ── Check 9: Expenditure ─────────────────────────────────────────────────
    p("\n=== CHECK 9: EXPENDITURE (all clients) ===")
    exp_raw = load_csv(os.path.join(
        EXPORTS_DIR, "Client Expenditure Export HFP Migration 2026-10-27.csv"))

    # Source: non-zero lines per source client UUID
    src_exp_by_client = {}
    for row in exp_raw:
        sc    = (row.get("client_uuid") or "").strip()
        amt_s = (row.get("amount") or "").strip()
        try:
            amt = float(amt_s)
        except ValueError:
            amt = 0.0
        if sc and amt != 0.0:
            src_exp_by_client.setdefault(sc, 0)
            src_exp_by_client[sc] += 1

    total_src_lines   = sum(src_exp_by_client.values())
    total_src_clients = len(src_exp_by_client)
    p(f"  Source: {total_src_lines} non-zero lines across {total_src_clients} clients")

    # Destination: fetch for each source client
    total_dst_lines   = 0
    total_dst_clients = 0
    diff_clients      = []   # dest UUIDs that differ

    n_checked = 0
    for src_c, src_cnt in src_exp_by_client.items():
        dst_c = client_map.get(src_c)
        if not dst_c:
            diff_clients.append((f"src:{src_c[:8]}…", src_cnt, "?no_dest_map"))
            continue
        r_exp = dest.get(f"/api/v1/client/{dst_c}/expenditure")
        if r_exp.status_code != 200:
            diff_clients.append((dst_c[:8] + "…", src_cnt, f"HTTP{r_exp.status_code}"))
            continue
        exp_body  = r_exp.json()
        exp_data  = exp_body.get("data", exp_body) if isinstance(exp_body, dict) else exp_body
        if isinstance(exp_data, dict):
            dst_cnt = sum(
                1 for sec in (exp_data.get("sections") or [])
                for fld in (sec.get("fields") or [])
                if fld.get("amount") not in (None, 0, "0", "0.0", "0.00", "")
            )
        elif isinstance(exp_data, list):
            dst_cnt = len(exp_data)
        else:
            dst_cnt = 0
        total_dst_lines += dst_cnt
        if dst_cnt > 0:
            total_dst_clients += 1
        if dst_cnt != src_cnt:
            diff_clients.append((dst_c[:8] + "…", src_cnt, dst_cnt))
        n_checked += 1
        if n_checked % 25 == 0:
            p(f"  …checked {n_checked}/{total_src_clients} clients")

    p(f"  Destination: {total_dst_lines} non-zero lines across {total_dst_clients} clients")
    if total_src_lines == total_dst_lines:
        p(f"  ✓ line totals match ({total_src_lines})")
    else:
        p(f"  !! line total mismatch: source={total_src_lines} dest={total_dst_lines}")
    if diff_clients:
        p(f"  Clients with different counts ({len(diff_clients)}):")
        for dc_id, s, d in diff_clients:
            p(f"    dest {dc_id}  source={s}  dest={d}")
    else:
        p(f"  ✓ all {total_src_clients} clients match")

    # ── Check 10: Unknown Provider ───────────────────────────────────────────
    p("\n=== CHECK 10: UNKNOWN PROVIDER PLANS ===")
    all_plans = get_paged(dest, "/api/v1/plans", max_pages=20)
    unknown_plans = []
    for pl in all_plans:
        prov = pl.get("provider") or {}
        if isinstance(prov, dict):
            prov_name = prov.get("name") or ""
        else:
            prov_name = str(prov or "")
        if prov_name.strip().lower() == "unknown provider":
            unknown_plans.append(pl)

    by_type = {}
    for pl in unknown_plans:
        pt = str(pl.get("type") or "")
        by_type[pt] = by_type.get(pt, 0) + 1

    required = {t: c for t, c in by_type.items() if t in PROVIDER_REQUIRED_TYPES}
    other    = {t: c for t, c in by_type.items() if t not in PROVIDER_REQUIRED_TYPES}

    p(f"  Total plans with Unknown Provider: {len(unknown_plans)}")
    p(f"  Types where provider is required (annuity/general_insurance/investment/mortgage/pension/protection_policy):")
    if required:
        for t, c in sorted(required.items()):
            p(f"    {t}: {c}")
    else:
        p(f"    none")
    p(f"  Other types:")
    if other:
        for t, c in sorted(other.items()):
            p(f"    {t}: {c}")
    else:
        p(f"    none")

    # ── Check 11: under_advice raw ───────────────────────────────────────────
    p("\n=== CHECK 11: under_advice raw API value (3 plans filed as 'no') ===")
    w5 = load_csv(os.path.join(IMPORT_DIR, "W5_plans_rewritten_fixed.csv"))
    ua_no_plans = [
        r for r in w5
        if (r.get("under_advice") or "").strip().lower() == "no"
        and plan_map.get((r.get("external_reference_value") or "").strip())
    ]
    p(f"  Plans in W5 with under_advice=no: {len(ua_no_plans)} (showing first 3)")
    for row in ua_no_plans[:3]:
        src_p = (row.get("external_reference_value") or "").strip()
        dst_p = plan_map.get(src_p, "?")
        r_pl  = dest.get(f"/api/v1/plans/{dst_p}")
        if r_pl.status_code == 200:
            pl_data = r_pl.json().get("data", r_pl.json())
            if isinstance(pl_data, dict):
                raw_ua = pl_data.get("under_advice")
            else:
                raw_ua = "?"
        else:
            raw_ua = f"HTTP {r_pl.status_code}"
        p(f"  plan {dst_p[:8]}… (src {src_p[:8]}…)"
          f"  type={row.get('type','')}  under_advice raw={raw_ua!r}")

    # ─────────────────────────────────────────────────────────────────────────
    p("\n" + "=" * 60)
    p("W8 + READ-ONLY CHECKS COMPLETE")
    p("=" * 60)


if __name__ == "__main__":
    main()
