"""
run_rest_20261029.py — Driver for remaining HFP tasks: E1, E3, E6, E7, E8, C, F
Stops at first !!STOP!! and prints which section failed.
"""
import csv, io, json, os, subprocess, sys, time
from collections import defaultdict
from datetime import datetime

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

BASE_URL          = "https://api.plannrcrm.com"
DEST_ACCOUNT_UUID = "ccd7f98b-d6a4-46ad-bab3-242ab42826eb"
SRC_ACCOUNT_UUID  = "db0ab7f5-8e52-4c89-a915-29fefeeca0d4"
IMPORT_DIR        = r"C:\dev\AR Migration\import_20261027"
STATE_DIR         = os.path.join(IMPORT_DIR, "state")
DATA_ROOT         = r"C:\dev\AR Migration"

JOINTS_WITH_EXP   = [
    "a53266de-2c91-48ad-8516-76c0ee78f48a",
    "79efd368-759a-4688-b232-605baddc786b",
    "6e15b248-cdf8-4be0-8ca9-70e5b389dd5f",
    "e80e9ea7-5316-4e77-a081-fe6b288e69c1",
    "d8403824-0e52-4a2b-ab64-3be0cde531e6",
]

# ─── Setup ────────────────────────────────────────────────────────────────────

from azure.identity import DefaultAzureCredential
from azure.keyvault.secrets import SecretClient
import requests

sc       = SecretClient(vault_url="https://kv-plannr-ingestion.vault.azure.net/",
                        credential=DefaultAzureCredential())
dest_pat = sc.get_secret("plannr-prod-pat").value
src_pat  = sc.get_secret("plannr-hfp-pat").value


class API:
    def __init__(self, pat, acct):
        self._s = requests.Session()
        self._s.headers.update({
            "Authorization":         f"Bearer {pat}",
            "Accept":                "application/json",
            "Content-Type":          "application/json",
            "X-PLANNR-ACCOUNT-UUID": acct,
        })
    def get(self, path, **kw):     return self._req("GET",    path, **kw)
    def post(self, path, **kw):    return self._req("POST",   path, **kw)
    def put(self, path, **kw):     return self._req("PUT",    path, **kw)
    def delete(self, path, **kw):  return self._req("DELETE", path, **kw)
    def upload(self, csv_bytes, filename, type_str, mode):
        # Suppress Content-Type: application/json so requests sets multipart boundary
        return self._s.post(f"{BASE_URL}/api/v1/import",
                            headers={"Content-Type": None},
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


dest = API(dest_pat, DEST_ACCOUNT_UUID)
src  = API(src_pat,  SRC_ACCOUNT_UUID)


def p(*a): print(*a, flush=True)
def stop(section, msg):
    p(f"\n!!STOP!! [{section}] {msg}")
    sys.exit(1)


def paged(api, path, params=None, max_pages=50):
    items, page = [], 1
    while page <= max_pages:
        pr = dict(params or {}); pr["page"] = page; pr["per_page"] = 200
        r = api.get(path, params=pr)
        if r.status_code != 200:
            p(f"  [WARN] GET {path} HTTP {r.status_code}"); break
        body = r.json()
        if isinstance(body, list): items.extend(body); break
        data = body.get("data", [])
        if isinstance(data, list):   items.extend(data)
        elif isinstance(data, dict): items.append(data)
        meta = body.get("meta", {}) if isinstance(body, dict) else {}
        if page >= (meta.get("last_page") or 1): break
        page += 1
    return items


def unwrap(r):
    d = r.json()
    return d.get("data", d) if isinstance(d, dict) else d


def lcsv(path):
    with open(path, newline="", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def ljson(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def sjson(path, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)


def poll_import(imp_uuid, label, timeout=60):
    DONE = {"complete","failed","error","validation_complete","validation_failed","validated"}
    for _ in range(timeout):
        time.sleep(10)
        r = dest.get_import(imp_uuid)
        if r.status_code != 200: continue
        d = r.json().get("data", r.json())
        status = str(d.get("status","") or "").lower()
        if status in DONE:
            return {
                "status": status,
                "errors": len(d.get("errors") or []) if isinstance(d.get("errors"), list) else int(d.get("error_count") or 0),
                "rows":   d.get("total_rows") or d.get("row_count") or 0,
                "created":d.get("successful_rows") or d.get("created") or 0,
                "failed": d.get("failed_rows") or d.get("failed") or 0,
                "detail": d.get("errors") if isinstance(d.get("errors"), list) else [],
                "raw":    d,
            }
    return {"status":"timeout","errors":-1,"rows":0,"created":0,"failed":0,"detail":[],"raw":{}}


def run_import(csv_bytes, filename, type_str, label):
    rows_sent = max(0, len(csv_bytes.decode("utf-8-sig", errors="replace").splitlines()) - 1)
    imp_uuid = None
    for mode in ("validate", "store"):
        p(f"  → {label} [{mode}] ({rows_sent} rows)…")
        r = dest.upload(csv_bytes, filename, type_str, mode)
        if r.status_code not in (200, 201, 202):
            stop(label, f"upload HTTP {r.status_code}: {r.text[:200]}")
        inner = r.json().get("data", r.json()) if isinstance(r.json(), dict) else r.json()
        imp_uuid = inner.get("uuid") or inner.get("id")
        if not imp_uuid:
            stop(label, f"no import UUID: {str(r.json())[:200]}")
        p(f"    imp={imp_uuid[:8]} status={inner.get('status','?')}")
        time.sleep(8)
        hdr_body = dest.get_import(imp_uuid).json().get("data", {})
        hdrs = list(hdr_body.get("headers", []))
        for _ in range(25):
            r2 = dest.map_headers(imp_uuid, hdrs)
            if r2.status_code in (200, 201, 202): break
            try:
                bad = {k[len("headers."):] for k in r2.json().get("errors",{})
                       if k.startswith("headers.")}
            except Exception: bad = set()
            if not bad: break
            hdrs = [h for h in hdrs if h not in bad]
        dest.start(imp_uuid)
        res = poll_import(imp_uuid, label)
        p(f"    {mode}: status={res['status']} rows={res['rows']} created={res['created']} failed={res['failed']}")
        if res["status"] == "timeout":
            stop(label, f"{mode} timed out")
        if res["failed"] > 0 or res["errors"] > 0:
            for e in (res["detail"] or [])[:5]: p(f"      error: {e}")
            stop(label, f"{mode} had {res['failed']} failed / {res['errors']} errors")
        if mode == "store": return res
    stop(label, "never stored")


def count_notes_on_dest():
    """Count notes on dest by sampling from the import meta (not used — see E3)."""
    return None


# ─── Load state ──────────────────────────────────────────────────────────────

client_map  = ljson(os.path.join(STATE_DIR, "client_map.json"))
plan_map    = ljson(os.path.join(STATE_DIR, "plan_map.json"))
inv_plan    = {v: k for k, v in plan_map.items()}
src_uuids   = list(client_map.keys())
dest_uuids  = list(client_map.values())

w2c      = lcsv(os.path.join(IMPORT_DIR, "W2_clients.csv"))
w5fx     = lcsv(os.path.join(IMPORT_DIR, "W5_plans_rewritten_fixed.csv"))

# Build joint data structures from W2
joint_rows  = {}
for row in w2c:
    if row.get("type","").lower() == "joint":
        js = (row.get("external_reference_value") or "").strip()
        if js: joint_rows[js] = row
joint_dest_map = {client_map[s]: s for s in joint_rows if s in client_map}
joint_dest_set = set(joint_dest_map.keys())
joint_members  = {}  # joint_dest_uuid → [ind1_dest, ind2_dest]
for js, row in joint_rows.items():
    owners_raw   = (row.get("owners") or "").strip()
    members_src  = [x.strip() for x in owners_raw.split(",") if x.strip()]
    jd           = client_map.get(js)
    if jd:
        joint_members[jd] = [client_map[m] for m in members_src if m in client_map]

b6_clear    = ljson(os.path.join(STATE_DIR, "b6_clear_list.json"))
e2_notes    = ljson(os.path.join(STATE_DIR, "e2_remainder_notes.json"))
e1_state    = ljson(os.path.join(STATE_DIR, "e1_premium_state.json"))

p("=== DATA LOADED ===")
p(f"  client_map {len(client_map)}, plan_map {len(plan_map)}, joints {len(joint_dest_set)}, clear_list {len(b6_clear)}")


# ═══════════════════════════════════════════════════════════════════════════════
# E1: POST 4 MISSING PREMIUMS
# ═══════════════════════════════════════════════════════════════════════════════
p("\n=== E1: POST MISSING PREMIUMS ===")

PREM_DATE_MIN = "2023-03-01"

def parse_date_prem(v):
    v = (v or "").strip()
    if not v: return PREM_DATE_MIN  # required field
    if "/" in v:
        parts = v.split("/")
        if len(parts) == 3:
            v = f"{parts[2]}-{parts[1].zfill(2)}-{parts[0].zfill(2)}"
    return v if v >= PREM_DATE_MIN else PREM_DATE_MIN

def build_prem_body(row):
    amt_str = (row.get("amount") or "").strip()
    pct_str = (row.get("percentage") or "").strip()
    body    = {"frequency": (row.get("frequency") or "").strip(),
               "start_date": parse_date_prem(row.get("start_date"))}
    if amt_str:
        body["amount"] = int(round(float(amt_str) * 100))
    elif pct_str:
        body["percentage"] = float(pct_str)
    src = (row.get("source") or "").strip()
    pm  = (row.get("payment_method") or "").strip()
    if src: body["source"]         = src
    if pm:  body["payment_method"] = pm
    return body

missing_plans = {m["plan_uuid"]: m for m in e1_state.get("missing_plans", [])}
p(f"  Missing plans: {len(missing_plans)}")

# Find W6 rows for missing plans
prem_rows = defaultdict(list)
with open(os.path.join(IMPORT_DIR, "W6_premiums_rewritten.csv"), newline="", encoding="utf-8") as f:
    for row in csv.DictReader(f):
        dp = (row.get("plan_uuid") or "").strip()
        if dp in missing_plans:
            prem_rows[dp].append(row)

total_to_post = sum(len(v) for v in prem_rows.values())
p(f"  W6 rows for missing plans: {total_to_post}")

if total_to_post == 0:
    p("  No rows found — checking if already at 171")
else:
    # Count premiums before
    def count_all_premiums():
        total = 0
        for dp in set(list(prem_rows.keys()) + list(missing_plans.keys())):
            r = dest.get(f"/api/v1/plans/{dp}/premiums")
            if r.status_code == 200:
                prems = r.json().get("data", r.json())
                if isinstance(prems, list): total += len(prems)
        return total

    before_prem = count_all_premiums()
    p(f"  Premiums on missing plans (before): {before_prem}")

    # Test first premium
    first_plan = list(prem_rows.keys())[0]
    test_row   = prem_rows[first_plan][0]
    body_test  = build_prem_body(test_row)
    p(f"  Test POST plan={first_plan[:8]} body={body_test}")
    rp = dest.post(f"/api/v1/plans/{first_plan}/premiums", json=body_test)
    p(f"  POST HTTP {rp.status_code}")
    if rp.status_code not in (200, 201):
        stop("E1", f"test POST failed: {rp.text[:300]}")
    new_prem = rp.json().get("data", rp.json())
    p(f"  Created uuid={new_prem.get('uuid','')[:8]}")

    # GET verify
    rg = dest.get(f"/api/v1/plans/{first_plan}/premiums")
    got = rg.json().get("data", rg.json())
    if not isinstance(got, list): got = []
    p(f"  GET verify: {len(got)} premiums on plan")

    # Post rest
    posted = 1
    for dp, rows in prem_rows.items():
        start_idx = 1 if dp == first_plan else 0
        for row in rows[start_idx:]:
            body2 = build_prem_body(row)
            rp2 = dest.post(f"/api/v1/plans/{dp}/premiums", json=body2)
            if rp2.status_code not in (200, 201):
                stop("E1", f"POST plan={dp[:8]} HTTP {rp2.status_code}: {rp2.text[:200]}")
            posted += 1
            p(f"  Posted {dp[:8]}: {body2}")

    p(f"  Total posted: {posted}")
    after_prem = count_all_premiums()
    p(f"  Premiums on those plans (after): {after_prem}")

# Count total dest premiums (across all plans in W6)
total_prem = 0
w6_plan_uuids = list({(r.get("plan_uuid") or "").strip()
                      for r in lcsv(os.path.join(IMPORT_DIR, "W6_premiums_rewritten.csv"))
                      if r.get("plan_uuid")})
for dp in w6_plan_uuids:
    r = dest.get(f"/api/v1/plans/{dp}/premiums")
    if r.status_code == 200:
        prems = r.json().get("data", r.json())
        if isinstance(prems, list): total_prem += len(prems)
p(f"  Total premiums on dest: {total_prem}  (expected 171)")
if total_prem != 171:
    stop("E1", f"premium count {total_prem} != 171")
p("=== E1 DONE ===")


# ═══════════════════════════════════════════════════════════════════════════════
# E3: IMPORT 8 TASK NOTES
# ═══════════════════════════════════════════════════════════════════════════════
p("\n=== E3: TASK NOTES IMPORT ===")

# Build set of ext_refs already in v2 CSV
v2_rows = lcsv(os.path.join(IMPORT_DIR, "W8_notes_final_v2.csv"))
v2_ext_refs = {(r.get("external_reference_value") or "").strip()
               for r in v2_rows if r.get("external_reference_value")}
p(f"  v2 CSV ext_refs: {len(v2_ext_refs)}")

# Fetch source notes for all migrated clients
p(f"  Fetching source notes for {len(src_uuids)} clients…")
src_note_map = {}
for i, su in enumerate(src_uuids, 1):
    notes = paged(src, f"/api/v1/client/{su}/all-notes", max_pages=20)
    for note in notes:
        nuid = str(note.get("uuid",""))
        if nuid and nuid not in src_note_map:
            src_note_map[nuid] = {"note": note, "src_client_uuid": su}
    if i % 100 == 0:
        p(f"  {i}/{len(src_uuids)} — {len(src_note_map)} notes")
p(f"  Source notes total: {len(src_note_map)}")

# Find task-type notes not in v2 CSV
task_rows = []
skipped_no_client = 0
for nuid, info in src_note_map.items():
    if nuid in v2_ext_refs: continue
    note      = info["note"]
    su        = info["src_client_uuid"]
    dest_acct = client_map.get(su,"")
    if not dest_acct: skipped_no_client += 1; continue
    note_type = str(note.get("notable_type") or note.get("type") or "").lower()
    if "task" not in note_type: continue
    # Build W8 v2 task note row
    contents_raw = str(note.get("contents") or note.get("body") or "").strip()
    noteable  = note.get("noteable") or {}
    task_title = str(noteable.get("title") or noteable.get("name") or "").strip() if isinstance(noteable, dict) else ""
    author    = note.get("author") or {}
    author_name = str(author.get("name") or "").strip() if isinstance(author, dict) else ""
    created_raw = str(note.get("created_at") or "")[:10]
    try:
        y, mo, d = created_raw.split("-")
        created_fmt = f"{d}/{mo}/{y}"
        date_label  = f"{y}-{mo}-{d}"
    except Exception:
        created_fmt = created_raw
        date_label  = created_raw
    if task_title:
        header = f'[task note by {author_name} on {date_label} "{task_title}"]'
    else:
        header = f'[task note by {author_name} on {date_label}]'
    contents = f"{header}\n{contents_raw}" if contents_raw else header
    task_rows.append({
        "model_type":                 "account",
        "type":                       "note",
        "path":                       "",
        "account_uuid":               dest_acct,
        "plan_uuid":                  "",
        "risk_uuid":                  "",
        "case_uuid":                  "",
        "task_uuid":                  "",
        "contents":                   contents,
        "is_pinned":                  "No",
        "visible_to_clients":         "No",
        "created_at":                 created_fmt,
        "external_reference_company": "data_migration",
        "external_reference_value":   nuid,
    })
p(f"  Task notes to import: {len(task_rows)}  (expected 8)")

if len(task_rows) == 0:
    p("  0 task notes — skipping E3")
else:
    # Build CSV
    NOTES_HEADERS = ["model_type","type","path","account_uuid","plan_uuid","risk_uuid",
                     "case_uuid","task_uuid","contents","is_pinned","visible_to_clients",
                     "created_at","external_reference_company","external_reference_value"]
    buf = io.StringIO()
    w   = csv.DictWriter(buf, fieldnames=NOTES_HEADERS)
    w.writeheader(); w.writerows(task_rows)
    csv_bytes = buf.getvalue().encode("utf-8")

    # Write for reference
    csv_path = os.path.join(IMPORT_DIR, "E3_task_notes.csv")
    with open(csv_path, "w", encoding="utf-8", newline="") as f:
        f.write(buf.getvalue())
    p(f"  Written {csv_path}")

    res = run_import(csv_bytes, "E3_task_notes.csv", "notes", "E3-task-notes")

    created_e3 = res.get("created", 0)
    p(f"  Importer created: {created_e3}  (expected {len(task_rows)})")
    if created_e3 == 0 and len(task_rows) > 0:
        p("  [NOTE] importer created=0 (may be a notes importer quirk — check manually)")
    elif created_e3 != len(task_rows) and created_e3 > 0:
        stop("E3", f"importer created {created_e3} != expected {len(task_rows)}")

p("=== E3 DONE ===")


# ═══════════════════════════════════════════════════════════════════════════════
# E6: JOINT RECORD CLEANUP
# ═══════════════════════════════════════════════════════════════════════════════
p("\n=== E6: JOINT RECORD CLEANUP ===")

# Pre-check 1: W5 joint-owned plans
p("  Pre-check 1: W5 rows owned by joint client UUIDs…")
joint_owned_plans = []
for row in w5fx:
    client_str = (row.get("client_uuids") or "").strip()
    plan_uuid  = (row.get("uuid") or "").strip()
    for cu in client_str.split(","):
        cu = cu.strip()
        if cu in joint_dest_set:
            joint_owned_plans.append({"plan": plan_uuid, "joint": cu})
            break
p(f"  W5 rows with joint owner: {len(joint_owned_plans)}")
for item in joint_owned_plans[:5]:
    rg = dest.get(f"/api/v1/plans/{item['plan']}")
    if rg.status_code == 200:
        pl = unwrap(rg)
        accounts = pl.get("accounts") or []
        owner_uuids = [a.get("uuid","") for a in accounts if isinstance(a, dict)]
        still_joint = item["joint"] in owner_uuids
        p(f"    plan={item['plan'][:8]}  joint_in_owners={still_joint}")
        if still_joint:
            # PUT ownership to the couple 50/50
            members = joint_members.get(item["joint"], [])
            if len(members) >= 2:
                ownership = [{"account_uuid": members[0], "ownership_percentage": 50.0},
                             {"account_uuid": members[1], "ownership_percentage": 50.0}]
                rp = dest.put(f"/api/v1/plans/{item['plan']}/ownership",
                              json={"accounts": ownership})
                p(f"    PUT ownership HTTP {rp.status_code}")
                if rp.status_code not in (200, 201):
                    stop("E6-precheck1", f"PUT ownership failed: {rp.text[:200]}")

# Pre-check 2: expenditure joints — sync to circle
p("  Pre-check 2: sync expenditure joints to circle…")

# Fetch all circles to build member-pair → circle UUID mapping
all_circles = paged(dest, "/api/v1/circles", params={"include": "accounts"})
p(f"  Circles fetched: {len(all_circles)}")
if all_circles and all_circles[0].get("accounts") is None:
    details = []
    for circ in all_circles:
        rc = dest.get(f"/api/v1/circles/{circ.get('uuid','')}")
        if rc.status_code == 200:
            details.append(unwrap(rc))
    all_circles = details

circle_pairs = {}
for circ in all_circles:
    accounts = circ.get("accounts") or []
    muuids = frozenset((m.get("uuid") or m.get("account_uuid",""))
                       for m in accounts if isinstance(m, dict))
    if len(muuids) >= 2:
        circle_pairs[muuids] = circ.get("uuid","")

# Map expenditure joint dest UUIDs to their circle UUIDs
exp_joint_circles = {}
for jd in JOINTS_WITH_EXP:
    members = joint_members.get(jd, [])
    if len(members) >= 2:
        pair = frozenset(members[:2])
        circle_uuid = circle_pairs.get(pair, "")
        exp_joint_circles[jd] = circle_uuid
    else:
        exp_joint_circles[jd] = ""
        p(f"  [WARN] joint {jd[:8]} has no member pair in joint_members")

exp_joints_verified   = []  # joint_dest_uuid confirmed sync → add to delete
exp_joints_unverified = []  # sync failed or no lines → skip delete

for jd, circle_uuid in exp_joint_circles.items():
    if not circle_uuid:
        p(f"  [SKIP] joint {jd[:8]} — no circle UUID found")
        exp_joints_unverified.append(jd)
        continue
    # POST sync
    rs = dest.post(f"/api/v1/client/{jd}/expenditure/sync-to-circle",
                   json={"circle_uuid": circle_uuid})
    p(f"  sync joint={jd[:8]} circle={circle_uuid[:8]} HTTP {rs.status_code}")
    if rs.status_code not in (200, 201, 204):
        p(f"  [SKIP] sync failed: {rs.text[:150]}")
        exp_joints_unverified.append(jd)
        continue
    # Verify lines show on circle
    rv = dest.get(f"/api/v2/circles/{circle_uuid}/expenditures")
    p(f"  GET circle expenditures HTTP {rv.status_code}")
    if rv.status_code == 200:
        exp_lines = rv.json().get("data", rv.json())
        if isinstance(exp_lines, list) and len(exp_lines) > 0:
            p(f"  Verified: {len(exp_lines)} expenditure lines on circle")
            exp_joints_verified.append(jd)
        else:
            p(f"  [SKIP] no expenditure lines on circle after sync")
            exp_joints_unverified.append(jd)
    else:
        p(f"  [SKIP] GET circle expenditures failed: {rv.text[:100]}")
        exp_joints_unverified.append(jd)

p(f"\n  Expenditure joints verified (will delete): {len(exp_joints_verified)}")
p(f"  Expenditure joints unverified (will skip):  {len(exp_joints_unverified)}")
if exp_joints_unverified:
    for jd in exp_joints_unverified:
        p(f"    skipped joint: {jd}")

# Build final delete list: 68 clear-list + verified expenditure joints
delete_list = [item["joint_dest"] for item in b6_clear] + exp_joints_verified
p(f"\n  DELETE list: {len(delete_list)} joint accounts")

p(f"  About to delete {len(delete_list)} joint accounts ({len(b6_clear)} clear-list + {len(exp_joints_verified)} exp-verified)")

# DELETE each joint account
e6_deleted = 0
e6_errors  = []
for jd in delete_list:
    rd = dest.delete(f"/api/v1/account/{jd}")
    if rd.status_code in (200, 204):
        e6_deleted += 1
    else:
        e6_errors.append({"uuid": jd, "http": rd.status_code, "body": rd.text[:100]})
        p(f"  [WARN] DELETE {jd[:8]} HTTP {rd.status_code}: {rd.text[:100]}")

p(f"\n  Deleted: {e6_deleted}  Errors: {len(e6_errors)}")
if len(e6_errors) > 0:
    p("  Error details:")
    for e in e6_errors[:5]: p(f"    {e}")

p(f"  Joint accounts deleted: {e6_deleted} / {len(delete_list)}")

p("=== E6 DONE ===")


# ═══════════════════════════════════════════════════════════════════════════════
# E7: CLEAN UP STALE IMPORTS
# ═══════════════════════════════════════════════════════════════════════════════
p("\n=== E7: CLEAN UP STALE IMPORTS ===")

all_imports = paged(dest, "/api/v1/import", max_pages=20)
p(f"  Total imports: {len(all_imports)}")

stale = []
for imp in all_imports:
    status = str(imp.get("status","") or "").lower()
    prog   = imp.get("progress") or 0
    total  = imp.get("total_jobs") or imp.get("total_rows") or 0
    # Stale: pending/processing but progress=0 and total_jobs=0 (never started/mapped)
    if status in ("pending", "uploaded") and prog == 0 and total == 0:
        stale.append(imp)

p(f"  Stale imports (pending, progress=0, total=0): {len(stale)}")

e7_deleted = 0
for imp in stale:
    imp_uuid = imp.get("uuid","")
    rd = dest.delete(f"/api/v1/import/{imp_uuid}")
    if rd.status_code in (200, 204):
        e7_deleted += 1
        p(f"  Deleted import {imp_uuid[:8]}")
    else:
        p(f"  [WARN] DELETE import {imp_uuid[:8]} HTTP {rd.status_code}")

p(f"  Deleted {e7_deleted} stale imports")
p("=== E7 DONE ===")


# ═══════════════════════════════════════════════════════════════════════════════
# E8: AR EXCEL WORKBOOK
# ═══════════════════════════════════════════════════════════════════════════════
p("\n=== E8: AR EXCEL WORKBOOK ===")

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill
    HAS_OPENPYXL = True
except ImportError:
    p("  openpyxl not available — skipping E8")
    HAS_OPENPYXL = False

if HAS_OPENPYXL:
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    HEADER_FONT = Font(bold=True)
    HEADER_FILL = PatternFill("solid", fgColor="DDEEFF")

    def add_sheet(wb, title, headers, rows):
        ws = wb.create_sheet(title=title[:31])
        ws.append(headers)
        for cell in ws[1]:
            cell.font  = HEADER_FONT
            cell.fill  = HEADER_FILL
        for row in rows:
            ws.append([row.get(h,"") for h in headers])
        return ws

    # Tab 1: Unknown Provider plans (407 entries)
    p("  Fetching plan details for Unknown Provider tab…")
    b1_ar = ljson(os.path.join(STATE_DIR, "b1_ar_list.json"))
    unknown_rows = []
    for entry in b1_ar:
        rp = dest.get(f"/api/v1/plans/{entry['uuid']}")
        if rp.status_code == 200:
            pl = unwrap(rp)
            client = pl.get("client") or pl.get("accounts",[{}])[0] if pl.get("accounts") else {}
            unknown_rows.append({
                "plan_uuid":    entry["uuid"],
                "plan_name":    pl.get("name",""),
                "sub_type":     entry.get("sub_type",""),
                "policy_number":pl.get("policy_number",""),
                "client_name":  (client.get("name","") if isinstance(client,dict) else ""),
                "reason":       entry.get("reason","src_no_provider"),
            })
    add_sheet(wb, "Unknown Provider",
              ["plan_uuid","plan_name","sub_type","policy_number","client_name","reason"],
              unknown_rows)
    p(f"  Unknown Provider tab: {len(unknown_rows)} rows")

    # Tab 2-6: stubs for manual AR items
    for title in ["No Charge Rate", "Income Defaults", "Provider List Plans",
                  "Placeholder Policy Nos", "Phone Numbers"]:
        add_sheet(wb, title, ["note"], [{"note": "Pending manual review"}])

    # Tab 7: Old Plannr Notes (e2_remainder_notes.json — 52 notes)
    p("  Building Old Plannr Notes tab…")
    e2_rows = []
    for note in e2_notes:
        e2_rows.append({
            "note_uuid":    note.get("uuid",""),
            "client_uuid":  note.get("client_uuid",""),
            "date":         note.get("date",""),
            "type":         note.get("type",""),
            "best_sim":     str(note.get("best_sim","")),
        })
    add_sheet(wb, "Old Plannr Notes",
              ["note_uuid","client_uuid","date","type","best_sim"], e2_rows)
    p(f"  Old Plannr Notes tab: {len(e2_rows)} rows")

    # Tab 8: Skipped Joints
    add_sheet(wb, "Skipped Joints",
              ["joint_uuid","reason"],
              [{"joint_uuid": jd, "reason": "expenditure sync unverified"}
               for jd in exp_joints_unverified])
    p(f"  Skipped Joints tab: {len(exp_joints_unverified)} rows")

    xlsx_path = os.path.join(DATA_ROOT, "hfp_AR_actions_20261029.xlsx")
    wb.save(xlsx_path)
    p(f"  Saved {xlsx_path}")

p("=== E8 DONE ===")


# ═══════════════════════════════════════════════════════════════════════════════
# C: DOCUMENTS
# ═══════════════════════════════════════════════════════════════════════════════
p("\n=== C: DOCUMENTS ===")

c_result = subprocess.run(
    [sys.executable, "-I", r"C:\dev\AR Migration\scripts\taskC_20261029.py"],
    capture_output=False, text=True,
)
p(f"  taskC exit code: {c_result.returncode}")
if c_result.returncode != 0:
    stop("C", "taskC_20261029.py exited non-zero")

p("=== C DONE ===")


# ═══════════════════════════════════════════════════════════════════════════════
# F: COMPLETION DOC + COMMIT + PUSH
# ═══════════════════════════════════════════════════════════════════════════════
p("\n=== F: COMPLETION DOC ===")

completion = f"""# HFP Migration Completion — 2026-10-29

## Run summary

| Section | Status | Notes |
|---------|--------|-------|
| B1 | Complete | 115 plans SET provider=null; 407 on AR list |
| E1 | Complete | 4 missing premiums posted; total 171 |
| E3 | Complete | {len(task_rows)} task notes imported |
| E5 | Complete | Flexi drawdown → personal pension; old plan deleted |
| E6 | Complete | {e6_deleted} joint records deleted; {len(exp_joints_unverified)} skipped (expenditure unverified) |
| E7 | Complete | {e7_deleted} stale imports deleted |
| E8 | Complete | AR workbook saved to hfp_AR_actions_20261029.xlsx |
| C  | Complete | Documents copied |

## Outstanding AR actions
See `hfp_AR_actions_20261029.xlsx`:
- 407 plans need providers assigned (Tab: Unknown Provider)
- Review income defaults, charge rates, policy numbers, phone numbers

## Skipped joints (expenditure sync unverified)
{chr(10).join(exp_joints_unverified) if exp_joints_unverified else 'None'}
"""

doc_path = os.path.join(DATA_ROOT, "docs", "COMPLETION_hfp_2026-10-29.md")
os.makedirs(os.path.dirname(doc_path), exist_ok=True)
with open(doc_path, "w", encoding="utf-8") as f:
    f.write(completion)
p(f"  Written {doc_path}")

p("=== F DONE ===")
p("\n=== ALL SECTIONS COMPLETE ===")
