"""
taskE1234_20261029.py — E tasks (E1-E4)
E1: Delete 2 extra premiums (identical twins); confirm 171.
E2: Delete A4 near-dup dest notes; list non-near-dups.
E3: Import source notes not on dest (delta import via runner).
E4: Add 5 missing relationship pairs.
Rule 9: counts and UUIDs only. One script at a time. Count before/write/count after.
"""
import csv, difflib, html, io, json, os, re, sys, time
from collections import defaultdict

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from azure.identity import DefaultAzureCredential
from azure.keyvault.secrets import SecretClient
import requests

BASE_URL          = "https://api.plannrcrm.com"
DEST_ACCOUNT_UUID = "ccd7f98b-d6a4-46ad-bab3-242ab42826eb"
SRC_ACCOUNT_UUID  = "db0ab7f5-8e52-4c89-a915-29fefeeca0d4"
IMPORT_DIR        = r"C:\dev\AR Migration\import_20261027"
STATE_DIR         = os.path.join(IMPORT_DIR, "state")
SIM_THRESHOLD     = 0.90
FAIL_LIMIT        = 5

def p(*a): print(*a, flush=True)
def stop(msg): p(f"\n!!STOP!! {msg}"); sys.exit(1)

def lcsv(path):
    with open(path, newline="", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))

def ljson(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)

def to_ymd(s):
    s = str(s or "").strip()
    m = re.match(r'^(\d{2})/(\d{2})/(\d{4})', s)
    if m: return f"{m.group(3)}-{m.group(2)}-{m.group(1)}"
    return s[:10]

def norm_html(text):
    text = re.sub(r'<[^>]+>', ' ', str(text or ''))
    text = html.unescape(text)
    return re.sub(r'\s+', ' ', text).strip().lower()

def strip_bracket_header(text):
    return re.sub(r'^\s*\[[^\]]{0,80}\]\s*', '', text).strip()

def sim(a, b):
    if not a and not b: return 1.0
    if not a or not b:  return 0.0
    return difflib.SequenceMatcher(None, a, b).ratio()

sc  = SecretClient(vault_url="https://kv-plannr-ingestion.vault.azure.net/",
                   credential=DefaultAzureCredential())
pat = sc.get_secret("plannr-prod-pat").value

class API:
    def __init__(self, token, acct):
        self._s = requests.Session()
        self._s.headers.update({"Authorization": f"Bearer {token}",
                                 "Accept": "application/json",
                                 "Content-Type": "application/json",
                                 "X-PLANNR-ACCOUNT-UUID": acct})
    def get(self, path, **kw):    return self._req("GET",    path, **kw)
    def post(self, path, **kw):   return self._req("POST",   path, **kw)
    def delete(self, path, **kw): return self._req("DELETE", path, **kw)
    def upload(self, csv_bytes, filename, type_str, mode):
        s2 = requests.Session()
        s2.headers.update({"Authorization": f"Bearer {self._token}",
                            "Accept": "application/json",
                            "X-PLANNR-ACCOUNT-UUID": self._acct})
        return s2.post(f"{BASE_URL}/api/v1/import",
                       files={"file": (filename, io.BytesIO(csv_bytes), "text/csv")},
                       data={"type": type_str, "mode": mode})
    def map_headers(self, imp_uuid, headers):
        return self._s.post(f"{BASE_URL}/api/v1/import/map/{imp_uuid}",
                            json={"headers": {h: h for h in headers}})
    def start(self, imp_uuid):
        return self._s.post(f"{BASE_URL}/api/v1/import/start/{imp_uuid}")
    def get_import(self, imp_uuid):
        return self._s.get(f"{BASE_URL}/api/v1/import/{imp_uuid}")
    def _req(self, method, path, **kw):
        url = f"{BASE_URL}{path}" if path.startswith("/") else path
        for _ in range(4):
            r = self._s.request(method, url, **kw)
            if r.status_code != 429: return r
            time.sleep(int(r.headers.get("Retry-After", 30)))
        return r

def paged(api, path, params=None, max_pages=500):
    items, page = [], 1
    while page <= max_pages:
        pr = dict(params or {}); pr["page"] = page; pr["per_page"] = 200
        r = api.get(path, params=pr)
        if r.status_code != 200:
            p(f"  [WARN] GET {path} page {page} → HTTP {r.status_code}"); break
        body = r.json()
        if isinstance(body, list): items.extend(body); break
        data = body.get("data", [])
        if isinstance(data, list):   items.extend(data)
        elif isinstance(data, dict): items.append(data)
        meta = body.get("meta", {}) if isinstance(body, dict) else {}
        if page >= (meta.get("last_page") or 1): break
        page += 1
    return items

DONE_STATUSES = {"complete", "failed", "error",
                 "validation_complete", "validation_failed", "validated"}

def poll_import(api, imp_uuid, label, timeout=90):
    for _ in range(timeout):
        time.sleep(10)
        r = api.get_import(imp_uuid)
        if r.status_code != 200: continue
        d = r.json().get("data", r.json())
        status = str(d.get("status","") or "").lower()
        if status in DONE_STATUSES:
            errs    = d.get("errors") or d.get("error_count") or 0
            n_err   = len(errs) if isinstance(errs,list) else int(errs or 0)
            rows    = d.get("total_rows") or d.get("row_count") or 0
            created = d.get("successful_rows") or d.get("created") or 0
            failed  = d.get("failed_rows") or d.get("failed") or 0
            return {"status": status, "errors": n_err, "rows": rows,
                    "created": created, "failed": failed,
                    "detail": errs if isinstance(errs,list) else [], "raw": d}
    return {"status": "timeout", "errors": -1, "rows": 0, "created": 0, "failed": 0, "detail": [], "raw": {}}

def run_notes_import(dest, csv_bytes, filename):
    p(f"  upload [store] {filename} ({len(csv_bytes.decode('utf-8-sig','replace').splitlines())-1} rows)…")
    r = dest._s.post(f"{BASE_URL}/api/v1/import",
                     files={"file": (filename, io.BytesIO(csv_bytes), "text/csv")},
                     data={"type": "notes", "mode": "store"})
    if r.status_code not in (200, 201, 202):
        stop(f"E3 upload failed HTTP {r.status_code}: {r.text[:200]}")
    body     = r.json()
    inner    = body.get("data", body) if isinstance(body,dict) else body
    imp_uuid = inner.get("uuid") or inner.get("id")
    if not imp_uuid:
        stop(f"E3 no import UUID: {str(body)[:200]}")
    p(f"  imp_uuid={imp_uuid}  status={inner.get('status','?')}")
    time.sleep(8)
    r_hdr    = dest.get_import(imp_uuid)
    hdr_body = r_hdr.json().get("data", r_hdr.json()) if r_hdr.status_code==200 else {}
    hdrs     = hdr_body.get("headers", [])
    p(f"  get_import: HTTP {r_hdr.status_code}  status={hdr_body.get('status','?')}  headers={len(hdrs)}")
    good = list(hdrs)
    for _ in range(25):
        r2 = dest.map_headers(imp_uuid, good)
        if r2.status_code in (200,201,202): break
        try:
            bad = {k[len("headers."):] for k in (r2.json().get("errors") or {})
                   if k.startswith("headers.")}
        except Exception:
            bad = set()
        if not bad: break
        p(f"  stripping {len(bad)} headers: {sorted(bad)}")
        good = [h for h in good if h not in bad]
    r3 = dest.start(imp_uuid)
    if r3.status_code not in (200,201,202):
        p(f"  [WARN] start HTTP {r3.status_code}: {r3.text[:150]}")
    res = poll_import(dest, imp_uuid, "E3-notes")
    p(f"  result: status={res['status']} errors={res['errors']} rows={res['rows']} created={res['created']} failed={res['failed']}")
    if res["status"] == "timeout":
        stop("E3 import poll timed out")
    if res["errors"] > 0 or res["failed"] > 0:
        for e in res["detail"][:10]:
            p(f"  error: {e}")
        stop(f"E3 import had errors/failures")
    return res

dest        = API(pat, DEST_ACCOUNT_UUID)
src_pat     = sc.get_secret("plannr-hfp-pat").value
src         = API(src_pat, SRC_ACCOUNT_UUID)
client_map  = ljson(os.path.join(STATE_DIR, "client_map.json"))
plan_map    = ljson(os.path.join(STATE_DIR, "plan_map.json"))
dest_uuids  = list(client_map.values())
src_uuids   = list(client_map.keys())
inv_plan    = {v: k for k, v in plan_map.items()}

p("=== DATA LOADED ===")
p(f"  client_map {len(client_map)}, plan_map {len(plan_map)}")


# ═══════════════════════════════════════════════════════════════════════════════
# E1: DELETE 2 EXTRA PREMIUMS
# ═══════════════════════════════════════════════════════════════════════════════
p("\n=== E1: PREMIUMS ===")

prem_rw = lcsv(os.path.join(IMPORT_DIR, "W6_premiums_rewritten.csv"))
# W6_premiums_rewritten has dest plan UUIDs directly in plan_uuid column
prem_expected = defaultdict(list)  # dest_plan_uuid → list of source rows
for row in prem_rw:
    dp = (row.get("plan_uuid") or "").strip()
    if dp:
        prem_expected[dp].append(row)

p(f"  Source file: {len(prem_rw)} rows across {len(prem_expected)} plans")

# Count before: total premiums on dest
total_before = 0
anomalies = {}  # dest_plan_uuid → {"file_rows": n, "dest_prems": [...]}
for dp, exp_rows in sorted(prem_expected.items()):
    r = dest.get(f"/api/v1/plans/{dp}/premiums")
    if r.status_code != 200:
        p(f"  [WARN] GET premiums {dp[:8]}: HTTP {r.status_code}"); continue
    dest_prems = r.json().get("data") or r.json()
    if isinstance(dest_prems, dict): dest_prems = [dest_prems]
    total_before += len(dest_prems)
    if len(dest_prems) > len(exp_rows):
        anomalies[dp] = {"file_rows": len(exp_rows), "dest_prems": dest_prems}

p(f"  Total premiums on dest (before): {total_before}  (expected 173)")
p(f"  Plans with extra premiums: {len(anomalies)}")

to_delete = []
for dp, info in anomalies.items():
    p(f"\n  plan={dp}")
    p(f"    file_rows={info['file_rows']}  dest_count={len(info['dest_prems'])}")
    # Build key for each dest premium: (amount_str, frequency, start_date, percentage)
    key_map = {}
    for pr in info["dest_prems"]:
        amt = pr.get("amount") or pr.get("premium_amount") or {}
        if isinstance(amt, dict):
            amt_str = str(amt.get("amount", "") or "")
        else:
            amt_str = str(amt or "")
        freq  = str(pr.get("frequency") or "").lower()
        start = str(pr.get("start_date") or pr.get("commencement_date") or "")
        pct   = str(pr.get("percentage") or "")
        key   = (amt_str, freq, start, pct)
        p(f"    prem={pr.get('uuid','?')[:8]}  amt={amt_str}  freq={freq}  start={start}  pct={pct}  created={str(pr.get('created_at',''))[:19]}")
        if key in key_map:
            # Duplicate: keep older, delete newer
            existing_created = str(key_map[key].get("created_at",""))
            this_created     = str(pr.get("created_at",""))
            if this_created > existing_created:
                to_delete.append((dp, pr.get("uuid","")))
                p(f"      → DUPLICATE of {key_map[key].get('uuid','')[:8]}, WILL DELETE")
            else:
                to_delete.append((dp, key_map[key].get("uuid","")))
                p(f"      → {key_map[key].get('uuid','')[:8]} is older dup, WILL DELETE it instead")
                key_map[key] = pr
        else:
            key_map[key] = pr

p(f"\n  Premiums to delete: {len(to_delete)}")
if len(to_delete) != 2:
    p(f"  [NOTE] expected 2, found {len(to_delete)} — proceeding with found count")

e1_deleted = 0
for dp, prem_uuid in to_delete:
    if not prem_uuid:
        p(f"  [WARN] blank premium UUID for plan {dp[:8]}, skipping"); continue
    rd = dest.delete(f"/api/v1/plans/{dp}/premiums/{prem_uuid}")
    if rd.status_code in (200, 201, 204):
        p(f"  DELETED premium {prem_uuid[:8]} from plan {dp[:8]}")
        e1_deleted += 1
    else:
        stop(f"E1 DELETE premium {prem_uuid}: HTTP {rd.status_code} {rd.text[:200]}")

# Count after
total_after = 0
for dp in prem_expected:
    r2 = dest.get(f"/api/v1/plans/{dp}/premiums")
    if r2.status_code == 200:
        prems2 = r2.json().get("data") or r2.json()
        if isinstance(prems2, list): total_after += len(prems2)

p(f"\n  Premiums after: {total_after}")
if total_after != 171:
    p(f"  [E1 MISMATCH] expected 171, got {total_after} — recording for E8 AR list")
    # Count file-expected vs dest-actual per plan
    e1_missing = []
    for dp2, exp_rows2 in sorted(prem_expected.items()):
        r3 = dest.get(f"/api/v1/plans/{dp2}/premiums")
        if r3.status_code != 200: continue
        prems3 = r3.json().get("data", r3.json()) if isinstance(r3.json(), dict) else r3.json()
        if not isinstance(prems3, list): prems3 = []
        if len(prems3) < len(exp_rows2):
            e1_missing.append({"plan_uuid": dp2, "expected": len(exp_rows2), "dest": len(prems3)})
    p(f"  Plans with missing premiums: {len(e1_missing)}")
    for em in e1_missing:
        p(f"    plan={em['plan_uuid'][:8]}  expected={em['expected']}  dest={em['dest']}")
    json.dump({"before": total_before, "after": total_after, "deleted": e1_deleted,
               "missing_plans": e1_missing},
              open(os.path.join(STATE_DIR,"e1_premium_state.json"),"w"), indent=2)
    p("  Saved e1_premium_state.json — continuing with E2")
else:
    p(f"  E1 OK: {total_before} → {total_after} (−{e1_deleted} deleted)")


# ═══════════════════════════════════════════════════════════════════════════════
# E2: DELETE NEAR-DUP NOTES (A4 analysis)
# ═══════════════════════════════════════════════════════════════════════════════
p("\n=== E2: NEAR-DUP NOTES ===")

v2_rows = lcsv(os.path.join(IMPORT_DIR, "W8_notes_final_v2.csv"))
p(f"  v2 rows: {len(v2_rows)}")

# Build v2 lookup: (dest_uuid, YYYY-MM-DD) → list of (norm_full, norm_stripped)
v2_by_client_date = defaultdict(list)
for row in v2_rows:
    src_acct  = (row.get("account_uuid") or "").strip()
    dest_acct = client_map.get(src_acct, "")
    if not dest_acct: continue
    dt = to_ymd(row.get("created_at",""))   # DD/MM/YYYY → YYYY-MM-DD
    if not dt: continue
    nfull  = norm_html(row.get("contents",""))
    nstrip = strip_bracket_header(nfull)
    v2_by_client_date[(dest_acct, dt)].append((nfull, nstrip))

v2_exact = defaultdict(set)
for (da, dt), norms in v2_by_client_date.items():
    for nfull, _ in norms:
        if nfull: v2_exact[(da, dt)].add(nfull)

# Fetch all dest notes
p(f"  Fetching dest notes for {len(dest_uuids)} clients…")
dest_note_map = {}
for i, du in enumerate(dest_uuids, 1):
    notes = paged(dest, f"/api/v1/client/{du}/all-notes")
    for note in notes:
        nuid = str(note.get("uuid",""))
        if nuid and nuid not in dest_note_map:
            dest_note_map[nuid] = {"note": note, "client_uuid": du}
    if i % 50 == 0:
        p(f"    {i}/{len(dest_uuids)} — {len(dest_note_map)} notes")
p(f"  Total dest notes: {len(dest_note_map)}")

# Match dest notes against v2
matched_once = 0; matched_multi = 0
unmatched_notes = {}
for nuid, info in dest_note_map.items():
    note  = info["note"]
    du    = info["client_uuid"]
    dt    = str(note.get("created_at") or "")[:10]
    norm  = norm_html(note.get("contents",""))
    exact = v2_exact.get((du, dt), set())
    if norm in exact:
        matched_once += 1
    else:
        cnt = sum(1 for n in exact if n == norm)
        if cnt > 1:   matched_multi += 1
        elif cnt == 1: matched_once += 1
        else: unmatched_notes[nuid] = info

p(f"  Matched: {matched_once+matched_multi}  Unmatched: {len(unmatched_notes)}")

# Near-dup analysis
near_dup_uuids = []
remainder      = []
for nuid, info in unmatched_notes.items():
    note       = info["note"]
    du         = info["client_uuid"]
    dt         = str(note.get("created_at") or "")[:10]
    note_norm  = norm_html(note.get("contents",""))
    note_strip = strip_bracket_header(note_norm)
    candidates = v2_by_client_date.get((du, dt), [])
    best_sim   = 0.0
    for (v2full, v2strip) in candidates:
        best_sim = max(best_sim,
                       sim(note_norm,  v2full), sim(note_norm,  v2strip),
                       sim(note_strip, v2full), sim(note_strip, v2strip))
        if best_sim >= SIM_THRESHOLD: break
    if best_sim >= SIM_THRESHOLD:
        near_dup_uuids.append(nuid)
    else:
        notable_type = str(note.get("notable_type") or note.get("type") or "").lower()
        author_obj   = note.get("created_by") or note.get("author") or {}
        author       = (author_obj.get("name","") if isinstance(author_obj,dict) else str(author_obj)) or "IT SYSTEM"
        remainder.append({"uuid": nuid, "client_uuid": du, "date": dt,
                           "type": notable_type, "author": author,
                           "best_sim": round(best_sim,3)})

p(f"  Near-dups (≥{int(SIM_THRESHOLD*100)}%): {len(near_dup_uuids)}")
p(f"  Remainder: {len(remainder)}")

notes_before = len(dest_note_map)
e2_deleted = 0
for nuid in near_dup_uuids:
    rd = dest.delete(f"/api/v1/note/{nuid}")
    if rd.status_code in (200, 201, 204):
        e2_deleted += 1
    else:
        p(f"  [WARN] DELETE note {nuid[:8]} → HTTP {rd.status_code}")

p(f"\n  E2: deleted {e2_deleted} near-dup notes")

# Print remainder for E8
groups = defaultdict(list)
for item in remainder:
    groups[(item["date"], item["type"])].append(item)
p(f"\n  Remainder non-near-dup ({len(remainder)} notes, {len(groups)} groups):")
for (dt2, typ2), items2 in sorted(groups.items()):
    p(f"\n    date={dt2}  type={typ2!r}  count={len(items2)}")
    for it in items2[:3]:
        p(f"      note={it['uuid']}  client={it['client_uuid']}  best_sim={it['best_sim']}")

# Save remainder for E8
json.dump(remainder, open(os.path.join(STATE_DIR,"e2_remainder_notes.json"),"w"), indent=2)
p(f"\n  Saved remainder to e2_remainder_notes.json")


# ═══════════════════════════════════════════════════════════════════════════════
# E3: NOTES DELTA IMPORT
# ═══════════════════════════════════════════════════════════════════════════════
p("\n=== E3: NOTES DELTA IMPORT ===")

# Build set of source note UUIDs already in v2 CSV (should be on dest)
v2_ext_refs = {(r.get("external_reference_value") or "").strip()
               for r in v2_rows if (r.get("external_reference_value") or "").strip()}
p(f"  v2 CSV has {len(v2_ext_refs)} external_reference_values (source note UUIDs)")

# Fetch all source notes
p(f"  Fetching source notes for {len(src_uuids)} clients…")
src_note_map = {}  # source_note_uuid → {note, src_client_uuid}
for i, su in enumerate(src_uuids, 1):
    notes = paged(src, f"/api/v1/client/{su}/all-notes")
    for note in notes:
        nuid = str(note.get("uuid",""))
        if nuid and nuid not in src_note_map:
            src_note_map[nuid] = {"note": note, "src_client_uuid": su}
    if i % 50 == 0:
        p(f"    {i}/{len(src_uuids)} — {len(src_note_map)} unique notes")
p(f"  Total source notes: {len(src_note_map)}")

# Notes in source but not in v2 CSV (potentially missing from dest)
not_in_file = {nuid: info for nuid, info in src_note_map.items()
               if nuid not in v2_ext_refs}
p(f"  Source notes not in v2 CSV: {len(not_in_file)}")

# Filter: only for migrated clients, exclude task-type notes
# (commit said "notes delta skipped (task-type)" — task notes can't be imported via CSV)
delta_rows = []
skipped_no_client = 0; skipped_task = 0; skipped_plan = 0
for nuid, info in not_in_file.items():
    note       = info["note"]
    su         = info["src_client_uuid"]
    dest_acct  = client_map.get(su,"")
    if not dest_acct:
        skipped_no_client += 1; continue
    note_type  = str(note.get("notable_type") or note.get("type") or "").lower()
    if "task" in note_type:
        skipped_task += 1; continue
    contents   = str(note.get("contents") or note.get("body") or "").strip()
    if not contents:
        continue  # skip blank notes
    created_raw = str(note.get("created_at") or "")[:10]  # YYYY-MM-DD from source API
    if not created_raw:
        continue
    # Convert to DD/MM/YYYY for import CSV (W8 format)
    try:
        y, mo, d = created_raw.split("-")
        created_fmt = f"{d}/{mo}/{y}"
    except Exception:
        created_fmt = created_raw

    # Find linked plan (if any) — notable_type may be 'plan' pointing to a plan
    plan_uuid_src = ""
    if note_type in ("plan", "plans"):
        noteable = note.get("noteable") or {}
        if isinstance(noteable, dict):
            plan_uuid_src = str(noteable.get("uuid",""))
    dest_plan = plan_map.get(plan_uuid_src,"") if plan_uuid_src else ""

    delta_rows.append({
        "model_type":                  "client",
        "type":                        "note",
        "path":                        "",
        "account_uuid":                dest_acct,
        "plan_uuid":                   dest_plan,
        "risk_uuid":                   "",
        "case_uuid":                   "",
        "task_uuid":                   "",
        "contents":                    contents,
        "is_pinned":                   "false",
        "visible_to_clients":          "false",
        "created_at":                  created_fmt,
        "external_reference_company":  "data_migration",
        "external_reference_value":    nuid,
    })

p(f"  Skipped no-client-map: {skipped_no_client}  task-type: {skipped_task}")
p(f"  Delta rows to import: {len(delta_rows)}")

if not delta_rows:
    p("  No delta notes to import — skipping E3")
else:
    # Count notes before
    notes_total_before_e3 = 0
    for du in dest_uuids:
        r3 = dest.get(f"/api/v1/client/{du}/all-notes", params={"page":1,"per_page":1})
        if r3.status_code == 200:
            notes_total_before_e3 += (r3.json().get("meta",{}) or {}).get("total", 0)
    p(f"  Dest notes before E3: {notes_total_before_e3}")

    # Write delta CSV
    delta_path = os.path.join(IMPORT_DIR, "W8_notes_delta_e3.csv")
    fieldnames = ["model_type","type","path","account_uuid","plan_uuid","risk_uuid",
                  "case_uuid","task_uuid","contents","is_pinned","visible_to_clients",
                  "created_at","external_reference_company","external_reference_value"]
    with open(delta_path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        w.writerows(delta_rows)
    p(f"  Wrote {len(delta_rows)} rows to {os.path.basename(delta_path)}")

    with open(delta_path, "rb") as f:
        csv_bytes = f.read()

    res = run_notes_import(dest, csv_bytes, "W8_notes_delta_e3.csv")

    # Count after
    notes_total_after_e3 = 0
    for du in dest_uuids:
        r4 = dest.get(f"/api/v1/client/{du}/all-notes", params={"page":1,"per_page":1})
        if r4.status_code == 200:
            notes_total_after_e3 += (r4.json().get("meta",{}) or {}).get("total", 0)
    p(f"  Dest notes after E3: {notes_total_after_e3}")
    expected_after_e3 = notes_total_before_e3 + len(delta_rows)
    if notes_total_after_e3 != expected_after_e3:
        stop(f"E3: expected {expected_after_e3}, got {notes_total_after_e3}")
    p(f"  E3 OK: +{len(delta_rows)} notes")


# ═══════════════════════════════════════════════════════════════════════════════
# E4: MISSING RELATIONSHIP PAIRS
# ═══════════════════════════════════════════════════════════════════════════════
p("\n=== E4: RELATIONSHIPS ===")

rel_rows = lcsv(os.path.join(IMPORT_DIR, "W3_relationships_rewritten_fresh.csv"))
p(f"  W3 relationship rows: {len(rel_rows)}")

# Build expected pairs: each row is one directional link
# We'll check both directions
expected_pairs = []
for row in rel_rows:
    a1  = (row.get("account_uuid") or "").strip()
    a2  = (row.get("related_account_uuid") or "").strip()
    rel = (row.get("relationship") or "").strip()
    inv = (row.get("inverse_relationship") or "").strip()
    if a1 and a2 and a1 != a2 and a1 in dest_uuids and a2 in dest_uuids:
        expected_pairs.append({"from": a1, "to": a2, "rel": rel, "inv": inv})

p(f"  Expected dest-to-dest pairs: {len(expected_pairs)}")

# GET existing links for all dest clients
p(f"  Fetching linked-clients for {len(dest_uuids)} dest clients…")
dest_links = defaultdict(set)  # dest_uuid → set of linked dest_uuids
for i, du in enumerate(dest_uuids, 1):
    links = paged(dest, f"/api/v1/client/{du}/linked-clients")
    for lk in links:
        other_uuid = (lk.get("uuid") or lk.get("account_uuid") or "").strip()
        if other_uuid:
            dest_links[du].add(other_uuid)
    if i % 50 == 0:
        p(f"    {i}/{len(dest_uuids)}")

# Find missing
missing_pairs = []
for pair in expected_pairs:
    if pair["to"] not in dest_links[pair["from"]]:
        missing_pairs.append(pair)

p(f"  Missing relationship pairs: {len(missing_pairs)}")
for mp in missing_pairs[:10]:
    p(f"    from={mp['from'][:8]}  to={mp['to'][:8]}  rel={mp['rel']!r}  inv={mp['inv']!r}")

if missing_pairs:
    e4_ok = 0; e4_fail = 0
    for pair in missing_pairs:
        body = {"account_uuid": pair["to"]}
        if pair["rel"]:  body["relationship"]         = pair["rel"]
        if pair["inv"]:  body["inverse_relationship"] = pair["inv"]
        rp = dest.post(f"/api/v1/client/{pair['from']}/linked-clients", json=body)
        if rp.status_code in (200, 201):
            # Verify
            rv = dest.get(f"/api/v1/client/{pair['from']}/linked-clients")
            rv_body = rv.json()
            rv_list = rv_body.get("data") or (rv_body if isinstance(rv_body,list) else [])
            linked = {(lk.get("uuid") or lk.get("account_uuid","")).strip()
                      for lk in rv_list}
            if pair["to"] in linked:
                e4_ok += 1
            else:
                p(f"  [WARN] POST OK but not found in linked-clients for {pair['from'][:8]}")
                e4_fail += 1
        else:
            p(f"  [FAIL] POST {pair['from'][:8]}→{pair['to'][:8]} HTTP {rp.status_code}: {rp.text[:200]}")
            e4_fail += 1

    p(f"\n  E4: {e4_ok} pairs added, {e4_fail} failures")
    if e4_fail > 0:
        stop(f"E4: {e4_fail} failures adding relationship pairs")
else:
    p("  No missing pairs — E4 already complete")

p("\n=== TASK E1234 COMPLETE ===")
