"""Read-only analysis: pending imports, employments vs income, relationship breakdown."""
import csv, json, sys
from collections import defaultdict

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

sys.path.insert(0, r"C:\dev\AR Migration\scripts")
from azure.identity import DefaultAzureCredential
from azure.keyvault.secrets import SecretClient
import requests

BASE_URL          = "https://api.plannrcrm.com"
DEST_ACCOUNT_UUID = "ccd7f98b-d6a4-46ad-bab3-242ab42826eb"
IMPORT_DIR        = r"C:\dev\AR Migration\import_20261027"

def p(*a): print(*a, flush=True)

sc  = SecretClient(vault_url="https://kv-plannr-ingestion.vault.azure.net/",
                   credential=DefaultAzureCredential())
pat = sc.get_secret("plannr-prod-pat").value
s   = requests.Session()
s.headers.update({"Authorization": f"Bearer {pat}", "Accept": "application/json",
                  "X-PLANNR-ACCOUNT-UUID": DEST_ACCOUNT_UUID})

client_map  = json.load(open(f"{IMPORT_DIR}/state/client_map.json"))
dest_uuids  = list(client_map.values())

def get_paged(path, max_pages=500, params=None):
    items, page = [], 1
    while page <= max_pages:
        pr = dict(params or {}); pr["page"] = page; pr["per_page"] = 200
        r = s.get(f"{BASE_URL}{path}", params=pr)
        if r.status_code != 200:
            p(f"  [WARN] GET {path} -> {r.status_code}")
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

# ── 2. Other pending imports ───────────────────────────────────────────────────
p("=== 2. OTHER PENDING IMPORTS ===")
PENDING = {"mapping", "pending", "created", "not_started", "queued"}
imports = get_paged("/api/v1/import", max_pages=5)
p(f"  Total imports fetched: {len(imports)}")
waiting = []
for imp in imports:
    st = str(imp.get("status") or "").lower()
    if st in PENDING:
        waiting.append(imp)
        p(f"  WAITING: uuid={imp.get('uuid')}  type={imp.get('type')}  mode={imp.get('mode')}"
          f"  status={st}  mapping={imp.get('mapping')!r}"
          f"  total_jobs={imp.get('total_jobs')}  progress={imp.get('progress')}")
if not waiting:
    p("  None -- all other imports are complete or failed")

# ── 3. Employments vs income ───────────────────────────────────────────────────
p("\n=== 3. EMPLOYMENTS VS INCOME ===")
emp_by_client = defaultdict(list)
inc_by_client = defaultdict(list)
for i, du in enumerate(dest_uuids, 1):
    for e in get_paged("/api/v2/employments", params={"filter[account_uuids]": du}):
        emp_by_client[du].append(e.get("uuid",""))
    for inc in get_paged("/api/v2/income", params={"filter[account_uuids]": du}):
        inc_by_client[du].append(inc.get("uuid",""))
    if i % 50 == 0:
        p(f"  {i}/{len(dest_uuids)} clients scanned...")

total_emp = sum(len(v) for v in emp_by_client.values())
total_inc = sum(len(v) for v in inc_by_client.values())
p(f"  Total employments: {total_emp}  total income: {total_inc}")

# Fetch income with employment included
inc_emp_links = {}   # inc_uuid -> emp_uuid or None
for du in dest_uuids:
    incs = get_paged("/api/v2/income",
                     params={"filter[account_uuids]": du, "include": "employment"})
    for inc in incs:
        iuid = inc.get("uuid","")
        emp  = inc.get("employment") or {}
        inc_emp_links[iuid] = emp.get("uuid") if isinstance(emp, dict) else None

income_with_emp  = sum(1 for v in inc_emp_links.values() if v)
income_no_emp    = sum(1 for v in inc_emp_links.values() if not v)
p(f"  Income with linked employment:    {income_with_emp} / {total_inc}")
p(f"  Income with NO linked employment: {income_no_emp} / {total_inc}")

all_linked_emp = {v for v in inc_emp_links.values() if v}
all_emp        = {u for v in emp_by_client.values() for u in v}
orphan_emps    = all_emp - all_linked_emp
p(f"  Employments not referenced by any income: {len(orphan_emps)}")
for eu in sorted(orphan_emps)[:10]:
    p(f"    orphan emp uuid={eu}")

# ── 4. Relationship analysis ───────────────────────────────────────────────────
p("\n=== 4. RELATIONSHIP ANALYSIS ===")
w3_rows = list(csv.DictReader(open(f"{IMPORT_DIR}/W3_relationships.csv",
                                    newline="", encoding="utf-8-sig")))
p(f"  W3_relationships.csv: {len(w3_rows)} rows")
file_src_pairs = set()
for row in w3_rows:
    a = (row.get("account_uuid") or "").strip()
    b = (row.get("related_account_uuid") or "").strip()
    if a and b:
        file_src_pairs.add((min(a, b), max(a, b)))
p(f"  Unique source pairs in file: {len(file_src_pairs)}")

# Build reverse map dest->src for the 356 clients
dest_to_src = {v: k for k, v in client_map.items()}

# Fetch all linked-client records
rel_records = []
for du in dest_uuids:
    recs = get_paged(f"/api/v1/client/{du}/linked-clients")
    for rec in recs:
        # linked-clients response: {"name": ..., "relationship": ..., "uuid": <linked_client_uuid>}
        # "uuid" IS the linked client's UUID — there is no separate relationship record UUID
        related   = rec.get("uuid", "")
        # synthetic key to deduplicate directed pairs (client -> related)
        rel_key   = f"{du}|{related}"
        rel_records.append({
            "uuid":     rel_key,
            "client":   du,
            "related":  related,
            "rel_type": rec.get("relationship",""),
        })

# Deduplicate by relationship UUID
seen_uuids = set()
unique_rels = []
for r in rel_records:
    uid = r["uuid"]
    if uid and uid not in seen_uuids:
        seen_uuids.add(uid)
        unique_rels.append(r)
p(f"  Unique relationships by UUID: {len(unique_rels)}")

# Load W2 to identify non-individual dest UUIDs
w2_non_indiv = set()
for row in csv.DictReader(open(f"{IMPORT_DIR}/W2_clients.csv",
                                newline="", encoding="utf-8-sig")):
    if str(row.get("type","")).lower() != "individual":
        src = (row.get("external_reference_value") or "").strip()
        du  = client_map.get(src)
        if du:
            w2_non_indiv.add(du)
p(f"  Non-individual dest UUIDs: {len(w2_non_indiv)}")

# Count pair occurrences
pair_counts = defaultdict(list)
for r in unique_rels:
    pair = (min(r["client"], r["related"]), max(r["client"], r["related"]))
    pair_counts[pair].append(r["uuid"])

# Categorise
cat_file   = []
cat_joint  = []
cat_dup    = []
cat_ext    = []
cat_other  = []

dest_uuids_set = set(dest_uuids)

for r in unique_rels:
    a, b = r["client"], r["related"]
    pair = (min(a, b), max(a, b))
    src_a = dest_to_src.get(a, "")
    src_b = dest_to_src.get(b, "")
    src_pair = (min(src_a, src_b), max(src_a, src_b)) if src_a and src_b else None
    in_file        = src_pair and src_pair in file_src_pairs
    non_indiv      = a in w2_non_indiv or b in w2_non_indiv
    one_external   = b not in dest_uuids_set
    is_dup_pair    = len(pair_counts[pair]) > 1

    if is_dup_pair:
        cat_dup.append(r)
    elif non_indiv:
        cat_joint.append(r)
    elif one_external:
        cat_ext.append(r)
    elif in_file:
        cat_file.append(r)
    else:
        cat_other.append(r)

p(f"\n  Relationship breakdown ({len(unique_rels)} unique):")
p(f"    In W3_relationships.csv file:         {len(cat_file)}")
p(f"    Involve joint/trust/business client:  {len(cat_joint)}")
dup_pair_set = set()
for r in cat_dup:
    dup_pair_set.add((min(r["client"],r["related"]), max(r["client"],r["related"])))
p(f"    Same two clients linked >once:        {len(dup_pair_set)} pairs ({len(cat_dup)} records)")
p(f"    One party outside our 356:            {len(cat_ext)}")
p(f"    Other (unexpected):                   {len(cat_other)}")
p(f"    Sum check:                            {len(cat_file)+len(cat_joint)+len(cat_dup)+len(cat_ext)+len(cat_other)}")

if dup_pair_set:
    p(f"\n  Duplicate pairs (same two clients linked more than once):")
    for pair in sorted(dup_pair_set):
        uuids = pair_counts[pair]
        p(f"    ({pair[0][:8]}..., {pair[1][:8]}...)  {len(uuids)} records  "
          f"rel_uuids={[u[:8] for u in uuids]}")

if cat_other:
    p(f"\n  Unexpected relationships (first 5):")
    for r in cat_other[:5]:
        p(f"    rel_uuid={r['uuid'][:8]}  client={r['client'][:8]}  "
          f"related={r['related'][:8]}  type={r['rel_type']!r}")

p("\n=== DONE ===")
