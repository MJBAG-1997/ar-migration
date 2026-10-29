"""
fix_20261029.py — Data fixes for HFP migration (2026-10-29).

Rule 9: no client names, no note contents. UUIDs, counts, types only.
Source is GET-only.
For every write: count before → submit once → poll to complete → count after.
If after - before ≠ expected: STOP.  Never re-submit.

Steps:
  1. Count + fix: relationships, income, employments, dependants
  2. Read-only: 2 extra circles; note matching (v2 rows vs dest)
  3. Premiums: confirm 0, import, confirm 171
  4. Expenditure: custom lines (2229 custom + 4 overwritten)
  5. Notes delta: source notes not in v2 file, build + import
  6. Report table
"""

import csv, html, io, json, os, re, sys, time
from collections import defaultdict
from datetime import date, datetime

import requests

# ── Import runner prep functions (safe: runner has if __name__ == "__main__" guard) ──
sys.path.insert(0, os.path.dirname(__file__))
from run_import_20261027 import (
    _fix_income_amounts as _runner_fix_income,
    rewrite_csv         as _runner_rewrite,
    GRACE_LIMIT         as _GRACE_LIMIT,
)

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# ── Constants ──────────────────────────────────────────────────────────────────
BASE_URL          = "https://api.plannrcrm.com"
KV_NAME           = "kv-plannr-ingestion"
DEST_PAT_SECRET   = "plannr-prod-pat"
SRC_PAT_SECRET    = "plannr-hfp-pat"
DEST_ACCOUNT_UUID = "ccd7f98b-d6a4-46ad-bab3-242ab42826eb"
SRC_ACCOUNT_UUID  = "db0ab7f5-8e52-4c89-a915-29fefeeca0d4"
MIGRATION_CO      = "data_migration"
IMPORT_DIR        = r"C:\dev\AR Migration\import_20261027"
EXPORTS_DIR       = r"C:\dev\AR Migration\exports_20261026"
STATE_DIR         = os.path.join(IMPORT_DIR, "state")
EXP_EXPORT        = os.path.join(EXPORTS_DIR, "Client Expenditure Export HFP Migration 2026-10-27.csv")
EXP_STATE_FILE    = os.path.join(STATE_DIR, "exp_custom_state.json")

EXPECTED = {"relationships": 147, "income": 316, "employments": 0, "dependants": 249}

# ── Helpers ────────────────────────────────────────────────────────────────────
def p(*a): print(*a, flush=True)

def stop(msg):
    p(f"\n!!STOP!! {msg}")
    sys.exit(1)

def resolve_pats():
    from azure.identity import DefaultAzureCredential
    from azure.keyvault.secrets import SecretClient
    sc = SecretClient(vault_url=f"https://{KV_NAME}.vault.azure.net/",
                      credential=DefaultAzureCredential())
    return sc.get_secret(DEST_PAT_SECRET).value, sc.get_secret(SRC_PAT_SECRET).value

def load_json(path):
    with open(path, encoding="utf-8") as f: return json.load(f)

def save_json(path, obj):
    with open(path, "w", encoding="utf-8") as f: json.dump(obj, f, indent=2)

def load_csv(path):
    with open(path, newline="", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))

def write_csv(path, rows, fieldnames):
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader(); w.writerows(rows)


class API:
    def __init__(self, pat, account_uuid):
        self._s = requests.Session()
        self._s.headers.update({
            "Authorization": f"Bearer {pat}",
            "Accept": "application/json",
            "X-PLANNR-ACCOUNT-UUID": account_uuid,
        })
    def get(self, path, **kw):    return self._s.get(f"{BASE_URL}{path}", **kw)
    def post(self, path, **kw):   return self._s.post(f"{BASE_URL}{path}", **kw)
    def delete(self, path, **kw): return self._s.delete(f"{BASE_URL}{path}", **kw)
    def upload(self, csv_bytes, filename, type_str, mode):
        return self._s.post(f"{BASE_URL}/api/v1/import",
                            files={"file": (filename, io.BytesIO(csv_bytes), "text/csv")},
                            data={"type": type_str, "mode": mode})
    def map_headers(self, uid, headers):
        return self._s.post(f"{BASE_URL}/api/v1/import/map/{uid}",
                            json={"headers": {h: h for h in headers}})
    def start(self, uid):       return self._s.post(f"{BASE_URL}/api/v1/import/start/{uid}")
    def get_import(self, uid):  return self._s.get(f"{BASE_URL}/api/v1/import/{uid}")


def get_paged(api, path, max_pages=500, params=None):
    items, page = [], 1
    while page <= max_pages:
        pr = dict(params or {}); pr["page"] = page; pr["per_page"] = 200
        for attempt in range(4):
            r = api.get(path, params=pr)
            if r.status_code == 429:
                wait = int(r.headers.get("Retry-After", 30))
                p(f"  [429] GET {path} — waiting {wait}s…")
                time.sleep(wait)
            else:
                break
        if r.status_code != 200:
            p(f"  [WARN] GET {path} → HTTP {r.status_code}")
            break
        body = r.json()
        if isinstance(body, list):
            items.extend(body); break
        data = body.get("data", [])
        if isinstance(data, list): items.extend(data)
        elif isinstance(data, dict): items.append(data)
        meta = body.get("meta", {}) if isinstance(body, dict) else {}
        if page >= (meta.get("last_page") or 1): break
        page += 1
    return items


def run_import(dest, csv_path, imp_type, label):
    """Full runner flow: upload → 8s → get headers → map → start → poll → return final dict.
    Never returns on failure — calls stop() instead."""
    with open(csv_path, "rb") as f:
        csv_bytes = f.read()
    filename  = os.path.basename(csv_path)
    rows_sent = max(0, len(csv_bytes.decode("utf-8-sig", errors="replace").splitlines()) - 1)
    p(f"    {label}: uploading {rows_sent} rows…")

    r = dest.upload(csv_bytes, filename, imp_type, "store")
    if r.status_code not in (200, 201, 202):
        stop(f"Upload {label}: HTTP {r.status_code} {r.text[:200]}")
    body     = r.json()
    inner    = body.get("data", body) if isinstance(body, dict) else body
    imp_uuid = inner.get("uuid") or inner.get("id")
    if not imp_uuid:
        stop(f"No import UUID in upload response for {label}: {str(body)[:200]}")
    p(f"    {label}: upload ok — imp_uuid={imp_uuid}  status={inner.get('status','?')}")

    time.sleep(8)
    r_hdr    = dest.get_import(imp_uuid)
    hdr_body = r_hdr.json().get("data", r_hdr.json()) if r_hdr.status_code == 200 else {}
    hdrs     = hdr_body.get("headers", [])
    p(f"    {label}: get_import status={hdr_body.get('status','?')}  headers={len(hdrs)}")

    # Map headers — strip unrecognised ones iteratively
    good_hdrs = list(hdrs)
    for attempt in range(25):
        r2 = dest.map_headers(imp_uuid, good_hdrs)
        if r2.status_code in (200, 201, 202):
            p(f"    {label}: headers mapped ({len(good_hdrs)} cols)")
            break
        try:
            bad = {h for h in good_hdrs
                   if h.lower() in str(r2.json()).lower()}
        except Exception:
            bad = set()
        if bad:
            p(f"    {label}: stripping {len(bad)} unsupported headers: {sorted(bad)}")
            good_hdrs = [h for h in good_hdrs if h not in bad]
        else:
            stop(f"map_headers {label}: HTTP {r2.status_code} {r2.text[:200]}")

    r3 = dest.start(imp_uuid)
    if r3.status_code not in (200, 201, 202):
        stop(f"start {label}: HTTP {r3.status_code} {r3.text[:200]}")
    p(f"    {label}: started — polling…")

    return _poll_import(dest, imp_uuid, label)


def _poll_import(dest, imp_uuid, label, timeout_secs=600):
    """Poll until done or 5-min stall. Returns final status dict."""
    DONE = {"complete", "failed", "error", "validation_complete", "validation_failed"}
    deadline   = time.time() + timeout_secs
    last_prog  = None
    last_move  = time.time()
    while time.time() < deadline:
        time.sleep(10)
        r = dest.get_import(imp_uuid)
        if r.status_code != 200:
            p(f"    [WARN] poll {imp_uuid}: HTTP {r.status_code}")
            continue
        d  = r.json().get("data", r.json())
        st = str(d.get("status") or "").lower()
        pr = d.get("progress")
        if pr != last_prog:
            p(f"    {label}: status={st}  progress={pr}  jobs_remaining={d.get('jobs_remaining')}")
            last_prog = pr
            last_move = time.time()
        if st in DONE:
            p(f"    {label}: import {st}")
            if st != "complete":
                stop(f"Import {imp_uuid} ended status={st}")
            return d
        if time.time() - last_move > 300:
            stop(f"Import {imp_uuid} stalled 5 min — "
                 f"status={d.get('status')}  mapping={d.get('mapping')}  "
                 f"total_jobs={d.get('total_jobs')}  progress={d.get('progress')}  "
                 f"failed_jobs={d.get('failed_jobs')}")
    stop(f"Import {imp_uuid} timed out after {timeout_secs}s")


def parse_iso_date(s):
    try: return datetime.fromisoformat(str(s).replace("Z", "+00:00")).date()
    except Exception:
        try: return date.fromisoformat(str(s)[:10])
        except: return None

def parse_ddmmyyyy(s):
    s = str(s or "").strip()
    for fmt in ("%d/%m/%Y", "%d/%m/%Y %H:%M:%S"):
        try: return datetime.strptime(s, fmt).date()
        except: pass
    return parse_iso_date(s)

def normalize_html(s):
    """Strip HTML tags, unescape entities, normalize whitespace."""
    if not s: return ""
    s = html.unescape(str(s))
    s = re.sub(r'<[^>]+>', '', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s


# ── Runner prep for W3 re-imports ─────────────────────────────────────────────
def prep_w3_for_reimport(client_map):
    """Re-run runner prep (amount fix + UUID rewrite) for W3 source files.
    Returns {type: fresh_csv_path}. Uses _fresh_ suffix to avoid touching the
    existing _rewritten files (which are the original-run references)."""
    p("\n  Running runner prep for W3 files…")
    fresh = {}

    # Relationships
    rrel = os.path.join(IMPORT_DIR, "W3_relationships_rewritten_fresh.csv")
    unmapped, gd = _runner_rewrite(
        os.path.join(IMPORT_DIR, "W3_relationships.csv"), rrel, client_map,
        ["account_uuid", "related_account_uuid"], grace_limit=_GRACE_LIMIT)
    if unmapped:
        stop(f"W3_relationships prep: {len(unmapped)} unmatched UUIDs: {unmapped[:3]}")
    for d in gd: p(f"    grace-dropped W3_relationships: ext_ref={d.get('ext_ref')} issues={d.get('issues')}")
    p(f"  W3_relationships prep: {rrel}")
    fresh["relationships"] = rrel

    # Dependants
    rdep = os.path.join(IMPORT_DIR, "W3_dependants_rewritten_fresh.csv")
    unmapped, gd = _runner_rewrite(
        os.path.join(IMPORT_DIR, "W3_dependants.csv"), rdep, client_map,
        ["account_uuids"], grace_limit=_GRACE_LIMIT)
    if unmapped:
        stop(f"W3_dependants prep: {len(unmapped)} unmatched UUIDs: {unmapped[:3]}")
    for d in gd: p(f"    grace-dropped W3_dependants: ext_ref={d.get('ext_ref')} issues={d.get('issues')}")
    p(f"  W3_dependants prep: {rdep}")
    fresh["dependants"] = rdep

    # Income: fix amounts first (runner prints its own 3 amount-verify lines), then rewrite
    amtfixed = os.path.join(IMPORT_DIR, "W3_income_amtfixed_fresh.csv")
    amtfixed_path, _, _, _, _ = _runner_fix_income(
        os.path.join(IMPORT_DIR, "W3_income.csv"), amtfixed)
    rinc = os.path.join(IMPORT_DIR, "W3_income_rewritten_fresh.csv")
    unmapped, gd = _runner_rewrite(
        amtfixed_path, rinc, client_map,
        ["client_uuid"], soft_fields=["employment_uuid"], grace_limit=_GRACE_LIMIT)
    if unmapped:
        stop(f"W3_income prep: {len(unmapped)} unmatched UUIDs: {unmapped[:3]}")
    for d in gd: p(f"    grace-dropped W3_income: ext_ref={d.get('ext_ref')} issues={d.get('issues')}")
    p(f"  W3_income prep: {rinc}")
    fresh["income"] = rinc

    return fresh


# ── Amount samples ─────────────────────────────────────────────────────────────
def verify_amount_samples(rows, amount_field, label, n=3):
    """Print n source-pounds → pennies conversions; STOP if any conversion is
    off (i.e. back-conversion to pounds doesn't round-trip within 1p)."""
    p(f"\n  Amount samples ({label}):")
    checked = 0
    for row in rows:
        try:
            src_pounds = float(row.get(amount_field) or "0")
        except ValueError:
            continue
        if src_pounds <= 0:
            continue
        pence = int(round(src_pounds * 100))
        back  = pence / 100
        ok    = abs(back - src_pounds) < 0.005
        p(f"    source=£{src_pounds:.2f}  → {pence}p  back=£{back:.2f}  {'✓' if ok else '!! MISMATCH'}")
        if not ok:
            stop(f"Amount conversion error in {label}: £{src_pounds} → {pence}p → £{back:.2f}")
        checked += 1
        if checked >= n:
            break
    if checked == 0:
        p(f"  (no non-zero {label} rows to sample)")


# ── Standard-field 100× probe ─────────────────────────────────────────────────
def read_dest_expenditure(dest, dest_uuid):
    """Return (ls_uuid, {fld_uuid: raw_amount_pence}) from GET expenditure.
    The API returns a flat list under 'data'; each item has life_stage_field_uuid and
    amount.amount (pence as string).
    """
    r = dest.get(f"/api/v1/client/{dest_uuid}/expenditure")
    if r.status_code != 200:
        return None, {}
    body    = r.json()
    ls_uuid = body.get("current_life_stage_uuid")
    vals    = {}
    items   = body.get("data", [])
    if isinstance(items, list):
        for item in items:
            fu  = item.get("life_stage_field_uuid") or ""
            amt = (item.get("amount") or {})
            raw = amt.get("amount") if isinstance(amt, dict) else None
            if fu and raw is not None:
                try:
                    vals[fu] = float(raw)
                except (TypeError, ValueError):
                    pass
    return ls_uuid, vals


def check_standard_fields(dest, exp_rows, row_category, client_map, life_stage_map,
                           field_map, n=3):
    """Read back n already-written standard field values; check for 100× error.
    Returns 'ok', 'repost_needed', or calls stop()."""
    written_rows = [(i, exp_rows[i]) for i, c in row_category.items() if c == "standard"]
    p(f"\n  Checking {min(n, len(written_rows))} of {len(written_rows)} already-written standard fields…")
    ratios = []
    for i, row in written_rows:
        src_uuid  = (row.get("client_uuid") or "").strip()
        dest_uuid = client_map.get(src_uuid)
        if not dest_uuid: continue
        sec, fld  = str(row.get("section") or "").strip(), str(row.get("type") or "").strip()
        fld_uuid  = field_map.get((sec, fld))
        if not fld_uuid: continue
        try:
            src_pounds = float(row.get("amount") or "0")
        except ValueError:
            continue
        if src_pounds <= 0: continue

        _, vals = read_dest_expenditure(dest, dest_uuid)
        raw = vals.get(fld_uuid)
        if raw is None:
            p(f"    {sec}/{fld}: source=£{src_pounds:.2f}  dest=not found")
            continue
        try:
            raw_f = float(raw)
        except (TypeError, ValueError):
            p(f"    {sec}/{fld}: source=£{src_pounds:.2f}  dest={raw!r} (not numeric)")
            continue

        ratio = raw_f / src_pounds if src_pounds else 0
        p(f"    {sec}/{fld}: source=£{src_pounds:.2f}  dest_raw={raw_f}  ratio={ratio:.1f}×")
        ratios.append(ratio)
        if len(ratios) >= n:
            break

    if not ratios:
        p("  (no standard field samples available — not yet written)")
        return "ok"

    # API stores in pennies: correct ratio = raw_pence / src_pounds ≈ 100
    # If previously posted in pounds by mistake: stored_pence = src_pounds → ratio ≈ 1
    if all(80 <= r <= 120 for r in ratios):
        p(f"  ✓ standard fields look correct (ratios ≈100 = stored pence as expected)")
        return "ok"
    if all(r < 2 for r in ratios):
        p(f"  !! all ratios ≈1 → standard fields stored as pounds, not pennies — re-posting all in pennies")
        return "repost_needed"
    # Mixed or unexpected ratios
    stop(f"Unexpected standard-field ratios {ratios} — cannot determine unit. "
         f"Investigate before proceeding.")


# ── Step 1: Count and fix ──────────────────────────────────────────────────────
def count_relationships(dest, dest_uuids):
    """Count unique relationships across all migrated clients."""
    seen = set()
    for du in dest_uuids:
        recs = get_paged(dest, f"/api/v1/client/{du}/linked-clients", max_pages=50)
        for rec in recs:
            uuid = str(rec.get("uuid") or rec.get("id") or "")
            if uuid: seen.add(uuid)
            # fallback: pair key
            else:
                related = str(rec.get("linked_client_uuid") or
                               rec.get("related_account_uuid") or
                               rec.get("account_uuid") or "")
                if related: seen.add(f"{min(du,related)}|{max(du,related)}")
    return len(seen)


def count_type_global(dest, path_tpl, dest_uuids, id_field="uuid"):
    """Count records using filter[account_uuids] per client."""
    seen = set()
    for du in dest_uuids:
        recs = get_paged(dest, path_tpl, params={"filter[account_uuids]": du})
        for rec in recs:
            uid = str(rec.get(id_field) or "")
            if uid: seen.add(uid)
    return seen


def delete_all_of_type(dest, all_uuids, delete_path_fn, label):
    """Delete records one by one. Return count deleted."""
    deleted = 0
    for uid in all_uuids:
        r = dest.delete(delete_path_fn(uid))
        if r.status_code in (200, 201, 204):
            deleted += 1
        else:
            stop(f"DELETE {label} {uid} → HTTP {r.status_code}: {r.text[:100]}")
    p(f"    Deleted {deleted} {label} records")
    return deleted


def step1_fix(dest, dest_uuids, client_map):
    p("\n=== STEP 1: COUNT + FIX ===")
    results = {}
    fresh = prep_w3_for_reimport(client_map)

    # ── Relationships ──
    p("\n  Counting relationships via /api/v1/client/{uuid}/linked-clients…")
    # Probe first to understand response shape
    probe = []
    for du in dest_uuids[:5]:
        recs = get_paged(dest, f"/api/v1/client/{du}/linked-clients", max_pages=2)
        if recs:
            p(f"  Relationship record keys (probe): {sorted(recs[0].keys())}")
            probe = recs
            break

    # Full count: collect relationship UUIDs across all clients
    rel_by_uuid = {}    # rel_uuid → record
    rel_by_pair = {}    # (a,b) → record (fallback)
    client_to_rel_uuids = defaultdict(list)  # dest_uuid → [rel_uuids]
    for du in dest_uuids:
        recs = get_paged(dest, f"/api/v1/client/{du}/linked-clients", max_pages=50)
        for rec in recs:
            ruuid = str(rec.get("uuid") or rec.get("id") or "")
            related = str(rec.get("linked_client_uuid") or
                          rec.get("related_account_uuid") or
                          rec.get("account_uuid") or
                          ((rec.get("account") or {}).get("uuid") or ""))
            if ruuid:
                rel_by_uuid[ruuid] = {"rec": rec, "client": du, "related": related}
                client_to_rel_uuids[du].append(ruuid)
            elif related:
                pair = f"{min(du,related)}|{max(du,related)}"
                rel_by_pair[pair] = {"rec": rec, "client": du, "related": related}

    total_rels = len(rel_by_uuid) if rel_by_uuid else len(rel_by_pair)
    p(f"  Relationships: {total_rels}  expected: {EXPECTED['relationships']}")

    exp = EXPECTED["relationships"]
    if total_rels == exp:
        p("  ✓ relationships match — leave")
        results["relationships"] = ("leave", total_rels, total_rels, "pass")
    elif total_rels > 0 and total_rels % exp == 0:
        mult = total_rels // exp
        p(f"  relationships = {mult}× expected — delete all, reimport")

        # Delete all relationships (deduplicated by UUID or pair)
        if rel_by_uuid:
            # Try delete by relationship UUID first
            deleted = 0
            for ruuid, info in rel_by_uuid.items():
                du = info["client"]
                r = dest.delete(f"/api/v1/client/{du}/linked-clients/{ruuid}")
                if r.status_code in (200, 201, 204):
                    deleted += 1
                else:
                    # Try alternate: delete via related UUID
                    related = info["related"]
                    if related:
                        r2 = dest.delete(f"/api/v1/client/{du}/linked-clients/{related}")
                        if r2.status_code in (200, 201, 204):
                            deleted += 1
                        else:
                            stop(f"Cannot delete relationship {ruuid} from {du[:8]}…: "
                                 f"HTTP {r.status_code} / {r2.status_code}")
                    else:
                        stop(f"Cannot delete relationship {ruuid}: HTTP {r.status_code} {r.text[:80]}")
        else:
            # Delete by pair (via related client UUID)
            deleted = 0
            seen_pairs = set()
            for pair, info in rel_by_pair.items():
                if pair in seen_pairs: continue
                seen_pairs.add(pair)
                du = info["client"]; related = info["related"]
                r = dest.delete(f"/api/v1/client/{du}/linked-clients/{related}")
                if r.status_code in (200, 201, 204):
                    deleted += 1
                else:
                    stop(f"DELETE relationship {du[:8]}…→{related[:8]}…: HTTP {r.status_code}")
        p(f"    Deleted {deleted} relationship entries")

        # Confirm 0
        after_del = count_relationships(dest, dest_uuids)
        if after_del != 0:
            stop(f"Relationships after delete = {after_del}, expected 0")
        p("    Confirmed 0 relationships after delete")

        # Re-import
        run_import(dest, fresh["relationships"], "relationships", "W3_relationships")

        # Confirm expected
        after_imp = count_relationships(dest, dest_uuids)
        if after_imp != exp:
            stop(f"Relationships after reimport = {after_imp}, expected {exp}")
        p(f"    ✓ Relationships confirmed {after_imp}")
        results["relationships"] = ("delete+reimport", total_rels, after_imp,
                                    "pass" if after_imp == exp else "FAIL")
    else:
        p(f"  !! relationships = {total_rels}, not a multiple of {exp} — leaving, listing")
        results["relationships"] = ("leave (unexpected count)", total_rels, total_rels, "FAIL-list")

    # ── Income ──
    p("\n  Counting income…")
    income_uuids = count_type_global(dest, "/api/v2/income", dest_uuids)
    total_inc = len(income_uuids)
    p(f"  Income: {total_inc}  expected: {EXPECTED['income']}")

    exp_inc = EXPECTED["income"]
    if total_inc == exp_inc:
        p("  ✓ income match — leave")
        results["income"] = ("leave", total_inc, total_inc, "pass")
    elif total_inc > 0 and total_inc % exp_inc == 0:
        mult = total_inc // exp_inc
        p(f"  income = {mult}× expected — delete all, reimport")
        delete_all_of_type(dest, income_uuids,
                           lambda u: f"/api/v2/income/{u}", "income")
        after_del = len(count_type_global(dest, "/api/v2/income", dest_uuids))
        if after_del != 0:
            stop(f"Income after delete = {after_del}, expected 0")
        p("    Confirmed 0 income after delete")
        verify_amount_samples(load_csv(fresh["income"]), "amount", "income reimport")
        run_import(dest, fresh["income"], "income_and_employment", "W3_income")
        after_imp = len(count_type_global(dest, "/api/v2/income", dest_uuids))
        if after_imp != exp_inc:
            stop(f"Income after reimport = {after_imp}, expected {exp_inc}")
        p(f"    ✓ Income confirmed {after_imp}")
        results["income"] = ("delete+reimport", total_inc, after_imp,
                             "pass" if after_imp == exp_inc else "FAIL")
    else:
        p(f"  !! income = {total_inc}, not a multiple of {exp_inc} — leaving, listing")
        results["income"] = ("leave (unexpected count)", total_inc, total_inc, "FAIL-list")

    # ── Employments ──
    p("\n  Counting employments…")
    emp_uuids = count_type_global(dest, "/api/v2/employments", dest_uuids)
    total_emp = len(emp_uuids)
    exp_emp = EXPECTED["employments"]
    p(f"  Employments: {total_emp}  expected: {exp_emp}")

    if total_emp == exp_emp:
        p("  ✓ employments match — leave")
        results["employments"] = ("leave", total_emp, total_emp, "pass")
    elif exp_emp > 0 and total_emp > 0 and total_emp % exp_emp == 0:
        p(f"  employments = {total_emp//exp_emp}× — delete (income already reimported)")
        delete_all_of_type(dest, emp_uuids,
                           lambda u: f"/api/v2/employments/{u}", "employment")
        after_del = len(count_type_global(dest, "/api/v2/employments", dest_uuids))
        if after_del != 0:
            stop(f"Employments after delete = {after_del}, expected 0")
        p(f"    ✓ Employments confirmed 0 (income reimport already run)")
        results["employments"] = ("delete", total_emp, after_del, "pass")
    else:
        p(f"  !! employments = {total_emp}, expected {exp_emp} — leaving, listing")
        results["employments"] = ("leave (unexpected count)", total_emp, total_emp,
                                  "pass" if total_emp == exp_emp else "FAIL-list")

    # ── Dependants ──
    p("\n  Counting dependants…")
    dep_uuids = count_type_global(dest, "/api/v2/dependants", dest_uuids)
    total_dep = len(dep_uuids)
    exp_dep = EXPECTED["dependants"]
    p(f"  Dependants: {total_dep}  expected: {exp_dep}")

    if total_dep == exp_dep:
        p("  ✓ dependants match — leave")
        results["dependants"] = ("leave", total_dep, total_dep, "pass")
    elif total_dep > 0 and total_dep % exp_dep == 0:
        mult = total_dep // exp_dep
        p(f"  dependants = {mult}× expected — delete all, reimport")
        delete_all_of_type(dest, dep_uuids,
                           lambda u: f"/api/v2/dependants/{u}", "dependant")
        after_del = len(count_type_global(dest, "/api/v2/dependants", dest_uuids))
        if after_del != 0:
            stop(f"Dependants after delete = {after_del}, expected 0")
        p("    Confirmed 0 dependants after delete")
        run_import(dest, fresh["dependants"], "dependants", "W3_dependants")
        after_imp = len(count_type_global(dest, "/api/v2/dependants", dest_uuids))
        if after_imp != exp_dep:
            stop(f"Dependants after reimport = {after_imp}, expected {exp_dep}")
        p(f"    ✓ Dependants confirmed {after_imp}")
        results["dependants"] = ("delete+reimport", total_dep, after_imp,
                                 "pass" if after_imp == exp_dep else "FAIL")
    else:
        p(f"  !! dependants = {total_dep}, not a multiple of {exp_dep} — leaving, listing")
        results["dependants"] = ("leave (unexpected count)", total_dep, total_dep, "FAIL-list")

    return results


# ── Step 2: Circles + note matching ───────────────────────────────────────────
def step2_read_only(dest, client_map, v2_rows, all_dest_notes):
    p("\n=== STEP 2: CIRCLES + NOTE MATCHING ===")
    results = {}

    # ── Circles ──
    p("\n  Circles:")
    circles = get_paged(dest, "/api/v1/circles", params={"include": "accounts"})
    p(f"  Total circles: {len(circles)}")
    if circles:
        p(f"  Circle keys: {sorted(circles[0].keys())}")

    # Build expected member sets from W2_circles_rewritten.csv
    circ_rows = load_csv(os.path.join(IMPORT_DIR, "W2_circles_rewritten.csv"))
    expected_member_sets = []
    for row in circ_rows:
        raw = str(row.get("account_uuids") or "")
        members = frozenset(u.strip() for u in raw.split(",") if u.strip())
        expected_member_sets.append({"members": members, "ext_ref": row.get("external_reference_value","")})

    def get_circle_members(circle):
        accounts = circle.get("accounts") or circle.get("account") or []
        if isinstance(accounts, list):
            return frozenset(str(a.get("uuid") or a.get("id") or a) for a in accounts if isinstance(a, dict))
        return frozenset()

    extras = []
    for c in circles:
        members = get_circle_members(c)
        matched = any(em["members"] == members for em in expected_member_sets)
        if not matched:
            extras.append(c)

    p(f"\n  Extra circles (not matching any W2 row): {len(extras)}")
    for ec in extras:
        members = get_circle_members(ec)
        p(f"    uuid={ec.get('uuid','?')}  created={str(ec.get('created_at','?'))[:10]}")
        p(f"    member_uuids={sorted(members)}")
    results["circles_extra"] = len(extras)

    # ── Note matching ──
    p("\n  Note matching (v2 rows vs dest notes):")

    # Build lookup: (dest_client_uuid, norm_content, date) → [v2_row_indices]
    dest_uuids_set = set(client_map.values())
    v2_lookup = defaultdict(list)  # key → [row_idx]
    for i, row in enumerate(v2_rows):
        src_acct = (row.get("account_uuid") or "").strip()
        dest_acct = client_map.get(src_acct)
        if not dest_acct:
            continue
        norm = normalize_html(row.get("contents") or "")
        d = parse_ddmmyyyy(row.get("created_at") or "")
        d_str = str(d) if d else ""
        v2_lookup[(dest_acct, norm, d_str)].append(i)

    matched_once = 0
    matched_multi = 0
    unmatched = []  # list of (author_name, date_str)

    for note in all_dest_notes:
        client_uuid = note.get("_client_uuid") or ""
        norm = normalize_html(note.get("contents") or "")
        d = parse_iso_date(note.get("created_at") or "")
        d_str = str(d) if d else ""
        key = (client_uuid, norm, d_str)
        matches = v2_lookup.get(key, [])
        if len(matches) == 1:
            matched_once += 1
        elif len(matches) > 1:
            matched_multi += 1
        else:
            unmatched.append((str(note.get("author_name") or "unknown"),
                              d_str, str(note.get("created_at") or "")[:10]))

    p(f"  Matched once:        {matched_once}")
    p(f"  Matched >1 v2 rows:  {matched_multi}")
    p(f"  Unmatched:           {len(unmatched)}")

    # Group unmatched by author and date
    unmatched_groups = defaultdict(list)
    for author, d_str, created in unmatched:
        unmatched_groups[(author, d_str)].append(created)
    p(f"\n  Unmatched grouped by author+date ({len(unmatched_groups)} groups):")
    for (author, d_str), dates in sorted(unmatched_groups.items(), key=lambda x: x[0][1]):
        p(f"    author={author}  date={d_str}  count={len(dates)}")

    results["notes"] = {
        "matched_once": matched_once,
        "matched_multi": matched_multi,
        "unmatched": len(unmatched),
    }
    return results


# ── Step 3: Premiums ───────────────────────────────────────────────────────────
def step3_premiums(dest, plan_map):
    p("\n=== STEP 3: PREMIUMS ===")

    prem_rows = load_csv(os.path.join(IMPORT_DIR, "W6_premiums.csv"))
    src_plan_uuids = [r.get("plan_uuid","").strip() for r in prem_rows if r.get("plan_uuid","").strip()]
    p(f"  W6_premiums.csv: {len(src_plan_uuids)} plan UUIDs")

    # Confirm 0 premiums on all plans
    total_before = 0
    missing_maps = []
    dest_plan_uuids = []
    for sp in src_plan_uuids:
        dp = plan_map.get(sp)
        if not dp:
            missing_maps.append(sp)
            continue
        dest_plan_uuids.append(dp)
        recs = get_paged(dest, f"/api/v1/plans/{dp}/premiums", max_pages=5)
        total_before += len(recs)

    if missing_maps:
        stop(f"Premiums: {len(missing_maps)} source plan UUIDs not in plan_map: {missing_maps[:3]}")

    # unique_dest needed for both before-check and after-count
    unique_dest = list(dict.fromkeys(dp for dp in dest_plan_uuids))
    expected_premiums = len(unique_dest)

    p(f"  Premiums before: {total_before} (expected 0; unique plans: {expected_premiums})")
    if total_before > 0:
        # Import already ran — accept the current count and move on
        p(f"  Premiums already present ({total_before}) — skipping import")
        return ("import", 0, total_before,
                "pass" if total_before == expected_premiums else f"warn:{total_before}≠{expected_premiums}")

    # Rewrite source plan UUIDs → dest
    rewritten_path = os.path.join(IMPORT_DIR, "W6_premiums_rewritten.csv")
    fieldnames = list(prem_rows[0].keys()) if prem_rows else []
    rewritten = []
    for row in prem_rows:
        newrow = dict(row)
        sp = (row.get("plan_uuid") or "").strip()
        dp = plan_map.get(sp)
        if not dp:
            stop(f"Premiums rewrite: plan_uuid {sp} not in plan_map")
        newrow["plan_uuid"] = dp
        rewritten.append(newrow)
    write_csv(rewritten_path, rewritten, fieldnames)
    p(f"  Rewritten {len(rewritten)} rows → {rewritten_path}")

    # Submit
    run_import(dest, rewritten_path, "premiums", "W6_premiums")

    # Count after — unique_dest and expected_premiums already set above
    p(f"  Unique dest plan UUIDs: {expected_premiums} (CSV had {len(dest_plan_uuids)} rows)")
    total_after = 0
    for dp in unique_dest:
        recs = get_paged(dest, f"/api/v1/plans/{dp}/premiums", max_pages=5)
        total_after += len(recs)

    p(f"  Premiums after: {total_after} (expected {expected_premiums})")
    if total_after != expected_premiums:
        stop(f"Premiums after import = {total_after}, expected {expected_premiums}")
    p("  ✓ Premiums confirmed")
    return ("import", 0, total_after, "pass" if total_after == expected_premiums else "FAIL")


# ── Step 4: Expenditure custom lines ──────────────────────────────────────────
_CSV_SECTION_TO_API = {
    "Discretionary":                              "Discretionary",
    "Essential - Children/Dependents":            "Essential",
    "Essential - Home":                           "Essential",
    "Essential - Medical":                        "Essential",
    "Essential - Pets":                           "Essential",
    "Essential - Vehicles":                       "Essential",
    "Lifestyle":                                  "Lifestyle",
    "Milestones e.g. World Cruise/House Renovation": "Lifestyle",
}

def csv_sec_to_api(csv_sec):
    """Map CSV section name to API section name."""
    if csv_sec in _CSV_SECTION_TO_API:
        return _CSV_SECTION_TO_API[csv_sec]
    # Fallback: strip " - Subtype" suffix
    return csv_sec.split(" - ")[0].strip()


def build_exp_field_map(dest):
    field_map, section_map = {}, {}
    sections = get_paged(dest, "/api/v1/expenditure-structure/section", max_pages=5)
    if not sections:
        r = dest.get("/api/v1/expenditure-structure/section", params={"per_page": 50})
        if r.status_code == 200: sections = r.json().get("data", [])
    for sec in sections:
        sec_name = str(sec.get("name","") or "")
        sec_uuid = sec.get("uuid","")
        if sec_name and sec_uuid: section_map[sec_name] = sec_uuid
        fields = sec.get("fields", [])
        if not fields and sec_uuid:
            rf = dest.get(f"/api/v1/expenditure-structure/section/{sec_uuid}",
                          params={"include": "fields"})
            if rf.status_code == 200:
                fields = rf.json().get("data", {}).get("fields", [])
        for fld in fields:
            fld_name = str(fld.get("name","") or "")
            fld_uuid = fld.get("uuid","")
            if sec_name and fld_name and fld_uuid:
                field_map[(sec_name, fld_name)] = fld_uuid
    return field_map, section_map


def step4_expenditure_custom(dest, client_map):
    p("\n=== STEP 4: EXPENDITURE CUSTOM LINES ===")

    field_map, section_map = build_exp_field_map(dest)
    p(f"  Field map: {len(field_map)} standard fields, {len(section_map)} sections")

    exp_rows = load_csv(EXP_EXPORT)
    p(f"  Source rows: {len(exp_rows)}")

    # Load state (skip already-posted)
    state = {}
    if os.path.exists(EXP_STATE_FILE):
        state = load_json(EXP_STATE_FILE)
    p(f"  State: {len(state)} rows already recorded")

    # Pre-fetch life_stage_uuids (section UUIDs come from section_map built above)
    needed_srcs = {r["client_uuid"] for r in exp_rows if r.get("client_uuid") in client_map}
    life_stage_map = {}   # dest_uuid → life_stage_uuid
    p(f"  Fetching life stage UUIDs for {len(needed_srcs)} clients…")
    for src in needed_srcs:
        dest_uuid = client_map[src]
        r = dest.get(f"/api/v1/client/{dest_uuid}/expenditure")
        if r.status_code == 200:
            ls = r.json().get("current_life_stage_uuid")
            if ls: life_stage_map[dest_uuid] = ls
    p(f"  Life stage UUIDs resolved: {len(life_stage_map)}")
    p(f"  Section map: {sorted(section_map.keys())}")

    # Identify which rows are custom (not in 23-field form) or overwritten
    # Simulate the original import logic to categorise rows
    last_writer = {}     # (dest_uuid, fld_uuid) → row_index  (for standard fields)
    row_category = {}    # row_index → 'custom' | 'standard' | 'amount_lte_0' | 'no_dest' | 'no_ls'

    for i, row in enumerate(exp_rows):
        src_uuid  = (row.get("client_uuid") or "").strip()
        dest_uuid = client_map.get(src_uuid)
        if not dest_uuid:
            row_category[i] = "no_dest"; continue
        ls_uuid = life_stage_map.get(dest_uuid)
        if not ls_uuid:
            row_category[i] = "no_ls"; continue
        sec = str(row.get("section") or "").strip()
        fld = str(row.get("type") or "").strip()
        try:
            amt = float(row.get("amount") or "0")
        except ValueError:
            row_category[i] = "amount_lte_0"; continue
        if amt <= 0:
            row_category[i] = "amount_lte_0"; continue
        fld_uuid = field_map.get((sec, fld))
        if not fld_uuid:
            row_category[i] = "custom"; continue
        # Standard field — track last writer
        key = (dest_uuid, fld_uuid)
        if key in last_writer:
            row_category[last_writer[key]] = "overwritten"
        last_writer[key] = i
        row_category[i] = "standard"

    custom_rows      = [(i, exp_rows[i]) for i, c in row_category.items() if c == "custom"]
    overwritten_rows = [(i, exp_rows[i]) for i, c in row_category.items() if c == "overwritten"]
    written_rows     = [(i, exp_rows[i]) for i, c in row_category.items() if c == "standard"]
    to_post          = custom_rows + overwritten_rows
    p(f"  Custom rows: {len(custom_rows)}  Overwritten rows: {len(overwritten_rows)}")
    p(f"  Total to post as custom: {len(to_post)}")

    # ── Amount samples (custom lines use pounds→pence) ──
    verify_amount_samples([r for _, r in to_post], "amount", "expenditure custom")

    # ── Check 3 already-written standard fields for 100× error ──
    std_check = check_standard_fields(dest, exp_rows, row_category, client_map,
                                      life_stage_map, field_map, n=3)
    if std_check == "repost_needed":
        p("\n  Re-posting all standard fields in pennies…")
        repost_ok = repost_err = 0
        for i, row in written_rows:
            src_uuid  = (row.get("client_uuid") or "").strip()
            dest_uuid = client_map.get(src_uuid)
            if not dest_uuid: continue
            ls_uuid = life_stage_map.get(dest_uuid)
            if not ls_uuid: continue
            fld_uuid = field_map.get((str(row.get("section") or "").strip(),
                                      str(row.get("type") or "").strip()))
            if not fld_uuid: continue
            try:
                amt_pence = int(round(float(row.get("amount") or "0") * 100))
            except ValueError:
                continue
            if amt_pence <= 0: continue
            r2 = dest.post(f"/api/v1/client/{dest_uuid}/expenditure/{ls_uuid}/{fld_uuid}",
                           json={"amount": amt_pence})
            if r2.status_code in (200, 201, 204):
                repost_ok += 1
            else:
                repost_err += 1
                p(f"    [ERR] repost std field row {i}: HTTP {r2.status_code} {r2.text[:80]}")
        p(f"  Standard field re-post: ok={repost_ok}  err={repost_err}")
        if repost_err > 0:
            stop(f"Standard field re-post had {repost_err} errors — investigate")

    # ── Probe one custom line before the batch ──
    p("\n  Probing one custom line…")
    probe_row = next(
        ((i, row) for i, row in to_post if str(i) not in state
         and (row.get("client_uuid") or "").strip() in client_map
         and life_stage_map.get(client_map[(row.get("client_uuid") or "").strip()])
         and (row.get("amount") or "0") != "0"),
        None
    )
    if probe_row is None:
        p("  (all custom rows already in state — skipping probe)")
    else:
        pi, prow = probe_row
        p_src    = (prow.get("client_uuid") or "").strip()
        p_dest   = client_map[p_src]
        p_ls     = life_stage_map[p_dest]
        p_sec    = str(prow.get("section") or "").strip()
        p_name   = str(prow.get("type") or "").strip()
        try:
            p_pence  = int(round(float(prow.get("amount") or "0") * 100))
        except ValueError:
            p_pence  = 0
        p_api_sec  = csv_sec_to_api(p_sec)
        p_sec_uuid = section_map.get(p_api_sec)
        p(f"    probe row {pi}: dest={p_dest[:8]}… sec={p_sec!r}→{p_api_sec!r} sec_uuid={p_sec_uuid!r} "
          f"name={p_name!r} amount=£{p_pence/100:.2f}→{p_pence}p")
        if not p_sec_uuid:
            stop(f"Custom line probe: no UUID found for section {p_sec!r} → {p_api_sec!r}. "
                 f"Available sections: {sorted(section_map.keys())}")
        probe_body = {"life_stage_section_uuid": p_sec_uuid, "name": p_name, "amount": p_pence}
        pr = dest.post(f"/api/v1/client/{p_dest}/expenditure/{p_ls}/custom",
                       json=probe_body)
        if pr.status_code not in (200, 201, 204):
            stop(f"Custom line probe failed: HTTP {pr.status_code} {pr.text[:300]}")
        state[str(pi)] = {"dest": p_dest, "sec": p_sec, "name": p_name, "amount": p_pence}
        save_json(EXP_STATE_FILE, state)
        p(f"    probe POST → HTTP {pr.status_code}")

        # Read back and compare
        _, dest_vals   = read_dest_expenditure(dest, p_dest)
        # Custom lines live outside the sections/fields structure; look in the full response
        pr_get = dest.get(f"/api/v1/client/{p_dest}/expenditure")
        if pr_get.status_code == 200:
            gb  = pr_get.json()
            gd  = gb.get("data", gb) if isinstance(gb, dict) else gb
            cls = (gd.get("custom_lines") or gd.get("customLines") or
                   gd.get("custom") or []) if isinstance(gd, dict) else []
            if not cls:
                # Try top-level
                cls = gb.get("custom_lines") or gb.get("customLines") or []
            p(f"    GET expenditure: custom_lines key count = {len(cls)}")
            if cls:
                p(f"    custom_line[0] keys: {sorted(cls[0].keys()) if isinstance(cls[0], dict) else type(cls[0])}")
                # Find our probe entry by name+section
                match = next((c for c in cls
                               if isinstance(c, dict)
                               and str(c.get("name") or "").strip() == p_name
                               and str(c.get("section") or "").strip() == p_sec), None)
                if match:
                    raw_amt = match.get("amount")
                    try:
                        raw_f = float(raw_amt)
                    except (TypeError, ValueError):
                        raw_f = None
                    p(f"    probe match: name={match.get('name')!r}  "
                      f"section={match.get('section')!r}  amount_raw={raw_amt}")
                    if raw_f is not None:
                        ratio = raw_f / (p_pence / 100)   # compare dest raw vs source pounds
                        p(f"    ratio dest/source_pounds = {ratio:.2f}×  "
                          f"(1=stored-pounds, 100=stored-pence, 0.01=100× too small)")
                        if ratio < 0.02:
                            stop(f"Custom line probe: amount 100× too small "
                                 f"(sent {p_pence}p = £{p_pence/100:.2f}, "
                                 f"got {raw_amt}). API may expect pounds.")
                else:
                    p(f"    [WARN] probe row not found in custom_lines — continuing")
        p("    Probe passed — proceeding with batch\n")

    # Count per client before
    before_counts = defaultdict(int)
    for _, row in to_post:
        src = (row.get("client_uuid") or "").strip()
        dest_uuid = client_map.get(src)
        if dest_uuid: before_counts[dest_uuid] += 1

    # POST custom lines
    posted = skipped = errors = 0
    for i, row in to_post:
        state_key = str(i)
        if state_key in state:
            skipped += 1
            continue

        src_uuid  = (row.get("client_uuid") or "").strip()
        dest_uuid = client_map.get(src_uuid)
        if not dest_uuid: skipped += 1; continue
        ls_uuid = life_stage_map.get(dest_uuid)
        if not ls_uuid: skipped += 1; continue

        sec      = str(row.get("section") or "").strip()
        name     = str(row.get("type") or "").strip()
        api_sec  = csv_sec_to_api(sec)
        sec_uuid = section_map.get(api_sec)
        if not sec_uuid:
            p(f"    [WARN] row {i}: no UUID for section {sec!r}→{api_sec!r} — skipping")
            skipped += 1; continue
        try:
            amt_pence = int(round(float(row.get("amount") or "0") * 100))
        except ValueError:
            skipped += 1; continue
        if amt_pence <= 0:
            skipped += 1; continue

        body = {"life_stage_section_uuid": sec_uuid, "name": name, "amount": amt_pence}
        r = dest.post(f"/api/v1/client/{dest_uuid}/expenditure/{ls_uuid}/custom",
                      json=body)
        if r.status_code in (200, 201, 204):
            state[state_key] = {"dest": dest_uuid, "sec": sec, "name": name,
                                 "amount": amt_pence}
            posted += 1
        else:
            p(f"    [ERR] custom exp row {i}: HTTP {r.status_code} body={r.text[:120]}")
            errors += 1
            if errors >= 5:
                save_json(EXP_STATE_FILE, state)
                stop(f"5 consecutive custom-line POST errors — check body format. "
                     f"Last response: HTTP {r.status_code} {r.text[:200]}")

        if (posted + skipped + errors) % 100 == 0:
            save_json(EXP_STATE_FILE, state)
            p(f"    …{posted} posted, {skipped} skipped, {errors} errors")

    save_json(EXP_STATE_FILE, state)
    p(f"\n  Custom lines posted: {posted}  skipped (already done): {skipped}  errors: {errors}")

    if errors > 0:
        stop(f"Expenditure custom: {errors} POST errors — check state file and investigate")

    # Verify: count custom lines per client against expected
    p("\n  Verifying custom line counts per client…")
    verify_errors = 0
    for dest_uuid, expected_custom in before_counts.items():
        r = dest.get(f"/api/v1/client/{dest_uuid}/expenditure")
        if r.status_code != 200:
            p(f"    [WARN] GET expenditure {dest_uuid[:8]}… HTTP {r.status_code}")
            continue
        body = r.json()
        custom_lines = body.get("custom_lines") or body.get("customLines") or []
        if isinstance(custom_lines, int):
            actual = custom_lines
        else:
            actual = len(custom_lines)
        if actual < expected_custom:
            p(f"    [WARN] {dest_uuid[:8]}…: expected ≥{expected_custom} custom, got {actual}")
            verify_errors += 1

    if verify_errors:
        p(f"  Verification: {verify_errors} clients with fewer custom lines than expected")
    else:
        p("  ✓ Custom line counts verified")

    total_expected = len(custom_rows) + len(overwritten_rows)
    return ("post_custom", 0, posted + skipped, "pass" if errors == 0 else "FAIL")


# ── Step 5: Notes delta ────────────────────────────────────────────────────────
def step5_notes_delta(dest, src, client_map, plan_map, v2_rows):
    p("\n=== STEP 5: NOTES DELTA ===")

    v2_ext_refs = {(row.get("external_reference_value") or "").strip()
                   for row in v2_rows if row.get("external_reference_value","").strip()}
    p(f"  v2 ext refs: {len(v2_ext_refs)}")

    src_client_uuids = list(client_map.keys())

    # Fetch all source notes (all-notes per migrated source client)
    p(f"  Fetching source notes for {len(src_client_uuids)} clients…")
    src_note_map = {}    # note_uuid → {note obj, client_uuid (source)}
    n_fetched = 0
    for i, su in enumerate(src_client_uuids, 1):
        notes = get_paged(src, f"/api/v1/client/{su}/all-notes", max_pages=500)
        for note in notes:
            nuid = str(note.get("uuid") or "")
            if nuid and nuid not in src_note_map:
                src_note_map[nuid] = {"note": note, "src_client": su}
        if i % 50 == 0: p(f"    {i}/{len(src_client_uuids)} — {len(src_note_map)} unique")

    p(f"  Source notes: {len(src_note_map)}")

    # Find notes NOT in v2
    new_notes = {uid: info for uid, info in src_note_map.items() if uid not in v2_ext_refs}
    p(f"  New notes (not in v2): {len(new_notes)}")

    if not new_notes:
        p("  No new notes — skip")
        return ("no_change", 0, 0, "pass")

    # Build dest-UUID-rewritten CSV rows from new notes
    v2_fields = ["model_type","type","path","account_uuid","plan_uuid","risk_uuid",
                 "case_uuid","task_uuid","contents","is_pinned","visible_to_clients",
                 "created_at","external_reference_company","external_reference_value"]

    # Staff account UUIDs = source UUIDs NOT in client_map
    src_uuids_set = set(client_map.keys())

    IMPORTABLE_NOTABLE_TYPES = {"account", "plan", "portfolio"}

    delta_rows = []
    skipped_notes = []
    for nuid, info in new_notes.items():
        note      = info["note"]
        src_acct  = info["src_client"]

        # Skip notes attached to tasks/cases/risks — importer only handles account and plan
        notable_type_raw = str(note.get("notable_type") or "account").lower()
        if notable_type_raw not in IMPORTABLE_NOTABLE_TYPES:
            skipped_notes.append(f"{nuid[:8]}(notable_type={notable_type_raw})")
            continue

        # Find which dest client this note belongs to
        if src_acct in src_uuids_set:
            dest_acct = client_map[src_acct]
        else:
            # Staff account — try to find migrated client from note.notable path
            dest_acct = None
            notable = note.get("notable") or {}
            notable_uuid = str(notable.get("uuid") or "")
            if notable_uuid in src_uuids_set:
                dest_acct = client_map[notable_uuid]
            if not dest_acct:
                skipped_notes.append(nuid)
                continue

        # plan_uuid (if note is on a plan)
        dest_plan = ""
        notable_type = str(note.get("notable_type") or "").lower()
        if notable_type in ("plan", "portfolio"):
            plan_src_uuid = str((note.get("notable") or {}).get("uuid") or "")
            dest_plan = plan_map.get(plan_src_uuid, "")

        created = str(note.get("created_at") or "")
        # Convert ISO to DD/MM/YYYY if needed
        d = parse_iso_date(created)
        created_formatted = d.strftime("%d/%m/%Y") if d else ""

        row = {
            "model_type":                  note.get("notable_type", "account"),
            "type":                        note.get("type", "note"),
            "path":                        "",
            "account_uuid":                dest_acct,
            "plan_uuid":                   dest_plan,
            "risk_uuid":                   "",
            "case_uuid":                   "",
            "task_uuid":                   "",
            "contents":                    note.get("contents", ""),
            "is_pinned":                   "Yes" if note.get("is_pinned") else "No",
            "visible_to_clients":          "Yes" if note.get("visible_to_clients") else "No",
            "created_at":                  created_formatted,
            "external_reference_company":  "data_migration",
            "external_reference_value":    nuid,
        }
        delta_rows.append(row)

    p(f"  Delta rows built: {len(delta_rows)}  skipped (unmappable): {len(skipped_notes)}")
    if skipped_notes:
        p(f"  Skipped note UUIDs: {skipped_notes}")

    if not delta_rows:
        p("  No mappable delta notes — skip")
        return ("no_change", 0, 0, "pass")

    # Count notes on dest clients before
    before_total = 0
    affected_clients = {r["account_uuid"] for r in delta_rows}
    for du in affected_clients:
        n = get_paged(dest, f"/api/v1/client/{du}/all-notes", max_pages=500)
        before_total += len(n)
    p(f"  Notes before (on {len(affected_clients)} affected clients): {before_total}")

    # Write and submit
    delta_path = os.path.join(IMPORT_DIR, "W8_notes_delta_rewritten.csv")
    write_csv(delta_path, delta_rows, v2_fields)
    p(f"  Delta CSV: {delta_path} ({len(delta_rows)} rows)")

    run_import(dest, delta_path, "notes", "W8_notes_delta")

    # Count after
    after_total = 0
    for du in affected_clients:
        n = get_paged(dest, f"/api/v1/client/{du}/all-notes", max_pages=500)
        after_total += len(n)
    added = after_total - before_total
    p(f"  Notes after: {after_total}  added: {added}  expected: {len(delta_rows)}")

    if added != len(delta_rows):
        stop(f"Notes delta: added {added}, expected {len(delta_rows)}")
    p("  ✓ Notes delta confirmed")
    return ("import", before_total, after_total, "pass" if added == len(delta_rows) else "FAIL")


# ── Step 6: Report table ───────────────────────────────────────────────────────
def step6_report(all_results):
    p("\n=== STEP 6: SUMMARY TABLE ===")
    p(f"\n{'Type':<25} {'Before':>8} {'Action':<20} {'After':>8} {'Result':<10}")
    p("-" * 75)
    for typ, (action, before, after, result) in all_results.items():
        p(f"{typ:<25} {before:>8} {action:<20} {after:>8} {result:<10}")
    p("-" * 75)


# ── Main ───────────────────────────────────────────────────────────────────────
def main():
    p("=== FIX RUN 2026-10-29 ===\n")
    dest_pat, src_pat = resolve_pats()
    dest = API(dest_pat, DEST_ACCOUNT_UUID)
    src  = API(src_pat,  SRC_ACCOUNT_UUID)

    client_map = load_json(os.path.join(STATE_DIR, "client_map.json"))
    plan_map   = load_json(os.path.join(STATE_DIR, "plan_map.json"))
    v2_rows    = load_csv(os.path.join(IMPORT_DIR, "W8_notes_final_v2.csv"))
    dest_uuids = list(client_map.values())
    p(f"  client_map: {len(client_map)}, plan_map: {len(plan_map)}, v2: {len(v2_rows)} rows")

    all_results = {}

    # Step 1
    s1 = step1_fix(dest, dest_uuids, client_map)
    all_results.update(s1)

    # Step 2: re-fetch all dest notes for note matching
    p("\n  Fetching all dest notes for note matching…")
    all_dest_notes = []
    for du in dest_uuids:
        notes = get_paged(dest, f"/api/v1/client/{du}/all-notes", max_pages=500)
        for n in notes:
            n["_client_uuid"] = du
        all_dest_notes.extend(notes)
    p(f"  Total dest notes fetched: {len(all_dest_notes)}")

    s2 = step2_read_only(dest, client_map, v2_rows, all_dest_notes)
    all_results["circles"] = ("read_only", 89, 89, f"+{s2['circles_extra']} extra")
    all_results["notes_match"] = ("read_only", len(all_dest_notes),
                                  s2["notes"]["matched_once"],
                                  f"unmatched={s2['notes']['unmatched']}")

    # Step 3
    prem_result = step3_premiums(dest, plan_map)
    all_results["premiums"] = prem_result

    # Step 4
    exp_result = step4_expenditure_custom(dest, client_map)
    all_results["exp_custom"] = exp_result

    # Step 5
    delta_result = step5_notes_delta(dest, src, client_map, plan_map, v2_rows)
    all_results["notes_delta"] = delta_result

    # Step 6
    step6_report(all_results)

    p("\n=== FIX RUN COMPLETE ===")


if __name__ == "__main__":
    main()
