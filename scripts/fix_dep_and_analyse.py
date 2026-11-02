"""
fix_dep_and_analyse.py
1. Map + start existing dependants import f4983389-5e72-40fb-b1a3-165fefffba92
   then poll and confirm count = 249.
2. Read-only: other imports waiting to be mapped/started.
3. Read-only: employments with no income record on the same client.
4. Read-only: 207 relationships split into categories.
Rule 9: counts and UUIDs only.
"""

import json, sys, time
from collections import defaultdict

sys.path.insert(0, r"C:\dev\AR Migration\scripts")
from azure.identity import DefaultAzureCredential
from azure.keyvault.secrets import SecretClient
import requests

BASE_URL          = "https://api.plannrcrm.com"
DEST_ACCOUNT_UUID = "ccd7f98b-d6a4-46ad-bab3-242ab42826eb"
KV_NAME           = "kv-plannr-ingestion"
IMPORT_DIR        = r"C:\dev\AR Migration\import_20261027"
DEP_IMP_UUID      = "f4983389-5e72-40fb-b1a3-165fefffba92"

def p(*a): print(*a, flush=True)

def stop(msg):
    p(f"\n!!STOP!! {msg}")
    sys.exit(1)

def resolve_pat():
    sc = SecretClient(vault_url=f"https://{KV_NAME}.vault.azure.net/",
                      credential=DefaultAzureCredential())
    return sc.get_secret("plannr-prod-pat").value

class Dest:
    def __init__(self, pat):
        self._s = requests.Session()
        self._s.headers.update({"Authorization": f"Bearer {pat}", "Accept": "application/json",
                                 "X-PLANNR-ACCOUNT-UUID": DEST_ACCOUNT_UUID})
    def get(self, path, **kw):   return self._s.get(f"{BASE_URL}{path}", **kw)
    def post(self, path, **kw):  return self._s.post(f"{BASE_URL}{path}", **kw)
    def delete(self, path, **kw):return self._s.delete(f"{BASE_URL}{path}", **kw)
    def map_headers(self, uid, headers):
        return self._s.post(f"{BASE_URL}/api/v1/import/map/{uid}",
                            json={"headers": {h: h for h in headers}})
    def start(self, uid): return self._s.post(f"{BASE_URL}/api/v1/import/start/{uid}")
    def get_import(self, uid): return self._s.get(f"{BASE_URL}/api/v1/import/{uid}")

def get_paged(dest, path, max_pages=500, params=None):
    items, page = [], 1
    while page <= max_pages:
        pr = dict(params or {}); pr["page"] = page; pr["per_page"] = 200
        r = dest.get(path, params=pr)
        if r.status_code != 200:
            p(f"  [WARN] GET {path} → {r.status_code}")
            break
        body = r.json()
        if isinstance(body, list): items.extend(body); break
        data = body.get("data", [])
        if isinstance(data, list): items.extend(data)
        elif isinstance(data, dict): items.append(data)
        meta = body.get("meta", {}) if isinstance(body, dict) else {}
        if page >= (meta.get("last_page") or 1): break
        page += 1
    return items

def poll_import(dest, uid, label, timeout_secs=600):
    DONE = {"complete", "failed", "error", "validation_complete", "validation_failed"}
    deadline = time.time() + timeout_secs
    last_progress = None
    last_move = time.time()
    while time.time() < deadline:
        time.sleep(10)
        r = dest.get_import(uid)
        if r.status_code != 200:
            p(f"  [WARN] poll {uid}: HTTP {r.status_code}")
            continue
        d = r.json().get("data", r.json())
        st  = str(d.get("status") or "").lower()
        prg = d.get("progress")
        rem = d.get("jobs_remaining")
        if prg != last_progress:
            p(f"  {label}: status={st}  progress={prg}  jobs_remaining={rem}")
            last_progress = prg
            last_move = time.time()
        if st in DONE:
            return d
        if time.time() - last_move > 300:
            # 5 minutes without progress movement
            d2 = r.json().get("data", r.json())
            stop(f"Import {uid} stalled 5 min — "
                 f"status={d2.get('status')}  mapping={d2.get('mapping')}  "
                 f"total_jobs={d2.get('total_jobs')}  progress={d2.get('progress')}  "
                 f"failed_jobs={d2.get('failed_jobs')}")
    stop(f"Import {uid} timed out after {timeout_secs}s")

# ── 1. Fix dependants import ───────────────────────────────────────────────────
p("=== 1. FIX DEPENDANTS IMPORT ===\n")
pat = resolve_pat()
dest = Dest(pat)
client_map = json.load(open(f"{IMPORT_DIR}/state/client_map.json"))

# Guard: dependants must be 0 before we start
p("  Checking current dependant count…")
seen_before = set()
for du in client_map.values():
    pg = 1
    while True:
        r = dest.get("/api/v2/dependants", params={"filter[account_uuids]": du,
                                                    "page": pg, "per_page": 200})
        if r.status_code != 200: break
        b = r.json()
        for rec in b.get("data", []): seen_before.add(rec.get("uuid",""))
        if pg >= (b.get("meta",{}).get("last_page") or 1): break
        pg += 1
p(f"  Dependants before start: {len(seen_before)}")
if len(seen_before) != 0:
    stop(f"Dependants = {len(seen_before)}, expected 0 before starting the import")

# GET the import
r = dest.get_import(DEP_IMP_UUID)
if r.status_code != 200:
    stop(f"GET import {DEP_IMP_UUID}: HTTP {r.status_code}")
d = r.json().get("data", r.json())
p(f"  Import status={d.get('status')}  mapping={d.get('mapping')!r}  "
  f"total_jobs={d.get('total_jobs')}  progress={d.get('progress')}  "
  f"failed_jobs={d.get('failed_jobs')}")
headers = d.get("headers", [])
p(f"  headers ({len(headers)}): {headers}")

mapping = d.get("mapping")
if not mapping:
    p("\n  Not mapped — mapping now…")
    good_hdrs = list(headers)
    for attempt in range(25):
        r2 = dest.map_headers(DEP_IMP_UUID, good_hdrs)
        p(f"  map_headers attempt {attempt+1}: HTTP {r2.status_code}  body={r2.text[:200]}")
        if r2.status_code in (200, 201, 202):
            break
        # Strip unrecognised headers
        try:
            err = r2.json()
            bad = set()
            for msg in (err.get("errors") or err.get("message") or []):
                msg_s = str(msg)
                for h in good_hdrs:
                    if h.lower() in msg_s.lower():
                        bad.add(h)
            if bad:
                p(f"  stripping {len(bad)} bad headers: {sorted(bad)}")
                good_hdrs = [h for h in good_hdrs if h not in bad]
                continue
        except Exception:
            pass
        stop(f"map_headers failed after {attempt+1} attempts: HTTP {r2.status_code} {r2.text[:200]}")
else:
    p("  Already mapped — skipping map step")

# Start
p("\n  Starting import…")
r3 = dest.start(DEP_IMP_UUID)
p(f"  start: HTTP {r3.status_code}  body={r3.text[:200]}")
if r3.status_code not in (200, 201, 202):
    stop(f"start failed: HTTP {r3.status_code} {r3.text[:300]}")

# Poll
p("  Polling…")
final = poll_import(dest, DEP_IMP_UUID, "dependants", timeout_secs=600)
p(f"  Final: status={final.get('status')}  errors={final.get('errors')}  "
  f"failed_jobs={final.get('failed_jobs')}")

# Count after
p("  Counting dependants after import…")
seen_after = set()
for du in client_map.values():
    pg = 1
    while True:
        r = dest.get("/api/v2/dependants", params={"filter[account_uuids]": du,
                                                    "page": pg, "per_page": 200})
        if r.status_code != 200: break
        b = r.json()
        for rec in b.get("data", []): seen_after.add(rec.get("uuid",""))
        if pg >= (b.get("meta",{}).get("last_page") or 1): break
        pg += 1
p(f"  Dependants after: {len(seen_after)}  expected: 249")
if len(seen_after) != 249:
    stop(f"Dependants after import = {len(seen_after)}, expected 249")
p("  ✓ Dependants confirmed 249")

# ── 2. Other pending imports ───────────────────────────────────────────────────
p("\n=== 2. OTHER PENDING IMPORTS ===\n")
PENDING = {"mapping", "pending", "created", "not_started", "queued"}
imports = get_paged(dest, "/api/v1/import", max_pages=5)
p(f"  Total imports: {len(imports)}")
waiting = []
for imp in imports:
    st = str(imp.get("status") or "").lower()
    if st in PENDING or (st == "mapping" and not imp.get("mapping")):
        waiting.append(imp)
        p(f"  WAITING: uuid={imp.get('uuid')}  type={imp.get('type')}  "
          f"mode={imp.get('mode')}  status={st}  "
          f"mapping={imp.get('mapping')!r}  "
          f"total_jobs={imp.get('total_jobs')}  progress={imp.get('progress')}")
if not waiting:
    p("  None — all imports are complete or failed")

# ── 3. Employments without linked income ──────────────────────────────────────
p("\n=== 3. EMPLOYMENTS VS INCOME ===\n")
emp_by_client = defaultdict(list)
inc_by_client = defaultdict(list)
for du in client_map.values():
    # Employments
    emps = get_paged(dest, "/api/v2/employments", params={"filter[account_uuids]": du})
    for e in emps: emp_by_client[du].append(e.get("uuid",""))
    # Income
    incs = get_paged(dest, "/api/v2/income", params={"filter[account_uuids]": du})
    for i in incs: inc_by_client[du].append(i.get("uuid",""))

total_emp = sum(len(v) for v in emp_by_client.values())
total_inc = sum(len(v) for v in inc_by_client.values())
p(f"  Total employments: {total_emp}  total income: {total_inc}")

# Check employments: look for employment_uuid field on each income record
# and vice versa
inc_emp_links = defaultdict(set)  # inc_uuid → {emp_uuids}
for du in client_map.values():
    incs = get_paged(dest, "/api/v2/income", params={"filter[account_uuids]": du,
                                                      "include": "employment"})
    for i in incs:
        iuid = i.get("uuid","")
        emp  = i.get("employment") or {}
        if isinstance(emp, dict) and emp.get("uuid"):
            inc_emp_links[iuid].add(emp["uuid"])

income_with_emp = sum(1 for v in inc_emp_links.values() if v)
p(f"  Income records with a linked employment: {income_with_emp} / {total_inc}")
p(f"  Income records with NO linked employment: {total_inc - income_with_emp} / {total_inc}")

all_linked_emp_uuids = {u for s in inc_emp_links.values() for u in s}
all_emp_uuids = {u for v in emp_by_client.values() for u in v}
orphan_emps = all_emp_uuids - all_linked_emp_uuids
p(f"  Employments not linked from any income record: {len(orphan_emps)}")
if orphan_emps:
    for eu in sorted(orphan_emps)[:10]:
        p(f"    orphan employment uuid={eu}")

# ── 4. Relationship analysis ───────────────────────────────────────────────────
p("\n=== 4. RELATIONSHIP ANALYSIS ===\n")
# Load W3_relationships.csv to get expected pairs
import csv
w3_rows = []
with open(f"{IMPORT_DIR}/W3_relationships.csv", newline="", encoding="utf-8-sig") as f:
    w3_rows = list(csv.DictReader(f))
p(f"  W3_relationships.csv rows: {len(w3_rows)}")
file_pairs = set()
for row in w3_rows:
    a = (row.get("account_uuid") or "").strip()
    b = (row.get("related_account_uuid") or "").strip()
    if a and b: file_pairs.add((min(a,b), max(a,b)))
p(f"  Unique pairs in file: {len(file_pairs)}")

# Fetch all dest relationships
rel_records = []  # list of {uuid, client, related, relationship_type, name}
for du in client_map.values():
    recs = get_paged(dest, f"/api/v1/client/{du}/linked-clients")
    for rec in recs:
        related_data = rec.get("account") or {}
        related_uuid = (str(related_data.get("uuid") or "") or
                        str(rec.get("linked_client_uuid") or "") or
                        str(rec.get("related_account_uuid") or ""))
        rel_records.append({
            "uuid":     rec.get("uuid",""),
            "client":   du,
            "related":  related_uuid,
            "rel_type": rec.get("relationship",""),
            "name":     rec.get("name",""),
        })

p(f"  Total relationship records fetched: {len(rel_records)}")

# Deduplicate by relationship UUID
seen_rel_uuids = set()
unique_rels = []
for r in rel_records:
    if r["uuid"] not in seen_rel_uuids:
        seen_rel_uuids.add(r["uuid"])
        unique_rels.append(r)
p(f"  Unique by UUID: {len(unique_rels)}")

# Load W2_clients to identify joint/trust/business dest UUIDs
w2_non_indiv = set()
with open(f"{IMPORT_DIR}/W2_clients.csv", newline="", encoding="utf-8-sig") as f:
    for row in csv.DictReader(f):
        if str(row.get("type","")).lower() != "individual":
            src = (row.get("external_reference_value") or "").strip()
            du = client_map.get(src)
            if du: w2_non_indiv.add(du)
p(f"  Non-individual dest UUIDs (joint/trust/business): {len(w2_non_indiv)}")

dest_uuids_set = set(client_map.values())

# Categorise each unique relationship
cat_in_file          = []  # pair exists in W3_relationships.csv
cat_joint            = []  # involves a joint/trust/business
cat_duplicate_pair   = []  # same two clients appear more than once
cat_external         = []  # one client not in our 356
cat_other            = []

pair_counts = defaultdict(list)  # (a,b) → [rel_uuids]
for r in unique_rels:
    a, b = r["client"], r["related"]
    pair = (min(a,b), max(a,b))
    pair_counts[pair].append(r["uuid"])

for r in unique_rels:
    a, b  = r["client"], r["related"]
    pair  = (min(a,b), max(a,b))
    src_a = next((s for s, d in client_map.items() if d == a), None)
    src_b = next((s for s, d in client_map.items() if d == b), None) if b else None
    in_file = (src_a, src_b) and (min(src_a or "",src_b or ""), max(src_a or "",src_b or "")) in file_pairs
    involves_non_indiv = a in w2_non_indiv or b in w2_non_indiv
    one_external = (b not in dest_uuids_set) if b else True
    dup_pair = len(pair_counts[pair]) > 1

    if dup_pair:
        cat_duplicate_pair.append(r)
    elif involves_non_indiv:
        cat_joint.append(r)
    elif one_external:
        cat_external.append(r)
    elif in_file:
        cat_in_file.append(r)
    else:
        cat_other.append(r)

p(f"\n  Relationship breakdown (207 unique by UUID):")
p(f"    In W3_relationships.csv file:        {len(cat_in_file)}")
p(f"    Involve joint/trust/business:        {len(cat_joint)}")
p(f"    Same two clients linked > once:      {len(set((min(r['client'],r['related']),max(r['client'],r['related'])) for r in cat_duplicate_pair))}")
p(f"      (duplicate records total):         {len(cat_duplicate_pair)}")
p(f"    One party outside our 356 clients:   {len(cat_external)}")
p(f"    Other:                               {len(cat_other)}")

if cat_duplicate_pair:
    p(f"\n  Duplicate pairs:")
    dup_pairs_seen = set()
    for r in cat_duplicate_pair:
        pair = (min(r["client"],r["related"]), max(r["client"],r["related"]))
        if pair not in dup_pairs_seen:
            dup_pairs_seen.add(pair)
            uuids = pair_counts[pair]
            p(f"    pair=({pair[0][:8]}…,{pair[1][:8]}…)  rel_uuids={[u[:8] for u in uuids]}  type={r['rel_type']!r}")

if cat_other:
    p(f"\n  Other (first 5):")
    for r in cat_other[:5]:
        p(f"    rel_uuid={r['uuid'][:8]}…  client={r['client'][:8]}…  "
          f"related={r['related'][:8]}…  type={r['rel_type']!r}")

p("\n=== DONE ===")
