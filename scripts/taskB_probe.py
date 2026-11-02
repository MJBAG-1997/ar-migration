"""
taskB_probe.py — Probe API structure before re-running taskB.
Checks:
1. Circle member structure (list vs individual GET)
2. Plans filter for joint client UUIDs
3. Charges endpoint (v1 vs v2, filter param)
4. Plan provider structure
5. W5_plans.csv contract_enquiry and adviser UUIDs for the 46 plans
6. Whether W5_plans.csv has any joint-owned plan rows
"""
import csv, json, os, sys, time
from collections import defaultdict

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from azure.identity import DefaultAzureCredential
from azure.keyvault.secrets import SecretClient
import requests

BASE_URL          = "https://api.plannrcrm.com"
DEST_ACCOUNT_UUID = "ccd7f98b-d6a4-46ad-bab3-242ab42826eb"
IMPORT_DIR        = r"C:\dev\AR Migration\import_20261027"
STATE_DIR         = os.path.join(IMPORT_DIR, "state")

def p(*a): print(*a, flush=True)
def load_csv(path):
    with open(path, newline="", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))
def load_json(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)

sc = SecretClient(vault_url="https://kv-plannr-ingestion.vault.azure.net/",
                  credential=DefaultAzureCredential())
pat = sc.get_secret("plannr-prod-pat").value

s = requests.Session()
s.headers.update({"Authorization": f"Bearer {pat}", "Accept": "application/json",
                  "X-PLANNR-ACCOUNT-UUID": DEST_ACCOUNT_UUID})

def g(path, params=None):
    r = s.get(f"{BASE_URL}{path}", params=params or {})
    return r

client_map = load_json(os.path.join(STATE_DIR, "client_map.json"))
plan_map   = load_json(os.path.join(STATE_DIR, "plan_map.json"))
dest_uuids = list(client_map.values())

w2c = load_csv(os.path.join(IMPORT_DIR, "W2_clients.csv"))
w5  = load_csv(os.path.join(IMPORT_DIR, "W5_plans.csv"))

# Build joint record info
joint_src_uuids = set()
joint_info = {}
for row in w2c:
    if str(row.get("type","")).lower() == "joint":
        src = (row.get("external_reference_value") or "").strip()
        if src:
            joint_src_uuids.add(src)
            joint_info[src] = row

p(f"Joint records: {len(joint_src_uuids)}")

# Build joint dest UUIDs
joint_dest_uuids = [client_map[s2] for s2 in joint_src_uuids if s2 in client_map]
p(f"Joint dest UUIDs mapped: {len(joint_dest_uuids)}")

# ─── 1. Circle member structure ───────────────────────────────────────────────
p("\n=== 1. CIRCLE STRUCTURE ===")
circles_r = g("/api/v1/circles", params={"page": 1, "per_page": 5})
p(f"  GET /api/v1/circles (page=1, per_page=5): HTTP {circles_r.status_code}")
if circles_r.status_code == 200:
    body = circles_r.json()
    data = body.get("data", body)
    if isinstance(data, list) and data:
        c0 = data[0]
        p(f"  Circle 0 keys: {sorted(c0.keys())}")
        accounts = c0.get("accounts", "NOT_PRESENT")
        p(f"  Circle 0 accounts field: {accounts!r}")
        # Now try individual GET
        circ_uuid = c0.get("uuid","")
        if circ_uuid:
            cr = g(f"/api/v1/circles/{circ_uuid}")
            p(f"  GET /api/v1/circles/{circ_uuid[:8]}…: HTTP {cr.status_code}")
            if cr.status_code == 200:
                cd = (cr.json().get("data") or cr.json())
                p(f"  Individual circle keys: {sorted(cd.keys())}")
                p(f"  Individual circle accounts: {cd.get('accounts','NOT_PRESENT')!r}")

# ─── 2. Plans filter for joint UUID ──────────────────────────────────────────
p("\n=== 2. PLANS FILTER VARIANTS ===")
if joint_dest_uuids:
    jdu = joint_dest_uuids[0]
    p(f"  Testing with joint dest UUID: {jdu}")
    for param_name in ["filter[client_uuid]", "filter[account_uuid]", "account_uuid", "client_uuid"]:
        r2 = g("/api/v1/plans", params={param_name: jdu, "per_page": 5})
        p(f"  {param_name!r}: HTTP {r2.status_code}")
        if r2.status_code == 200:
            body2 = r2.json()
            d2 = body2.get("data", body2) if isinstance(body2, dict) else body2
            p(f"    → {len(d2) if isinstance(d2, list) else 'dict'} records")
    # Also try no filter (all plans)
    r3 = g("/api/v1/plans", params={"per_page": 5})
    p(f"  no filter: HTTP {r3.status_code}")
    if r3.status_code == 200:
        b3 = r3.json()
        d3 = b3.get("data", b3) if isinstance(b3, dict) else b3
        total = b3.get("meta", {}).get("total", "?") if isinstance(b3, dict) else "?"
        p(f"    → page 1: {len(d3)} records, total={total}")
        if isinstance(d3, list) and d3:
            p0 = d3[0]
            p(f"    plan keys: {sorted(p0.keys())[:15]}…")
            prov = p0.get("provider") or p0.get("provider_uuid") or "N/A"
            p(f"    sample provider: {prov!r}")

# ─── 3. Charges endpoint ──────────────────────────────────────────────────────
p("\n=== 3. CHARGES ENDPOINT ===")
if joint_dest_uuids:
    jdu = joint_dest_uuids[0]
    for path_ver in ["/api/v2/charges", "/api/v1/charges"]:
        r4 = g(path_ver, params={"filter[account_uuids]": jdu, "per_page": 5})
        p(f"  GET {path_ver} filter[account_uuids]={jdu[:8]}…: HTTP {r4.status_code}")
        if r4.status_code == 200:
            b4 = r4.json()
            d4 = b4.get("data", b4) if isinstance(b4, dict) else b4
            p(f"    → {len(d4) if isinstance(d4, list) else 'dict'} records")
    # Try individual client charges
    for path_ver in ["/api/v2/charges", "/api/v1/charges"]:
        r5 = g(path_ver, params={"filter[account_uuids]": dest_uuids[0], "per_page": 5})
        p(f"  GET {path_ver} individual client: HTTP {r5.status_code}")

# ─── 4. Plan provider structure ──────────────────────────────────────────────
p("\n=== 4. PLAN PROVIDER STRUCTURE ===")
# Get all plans (page 1) and check provider
r6 = g("/api/v1/plans", params={"per_page": 5})
if r6.status_code == 200:
    plans6 = (r6.json().get("data") or [])
    if isinstance(plans6, list):
        for pl in plans6[:3]:
            puuid = pl.get("uuid","")
            prov_uuid = pl.get("provider_uuid","")
            prov_obj  = pl.get("provider") or {}
            prov_name = (str(pl.get("provider_name","") or "") or
                         (prov_obj.get("name","") if isinstance(prov_obj, dict) else ""))
            p(f"  plan={puuid[:8]}  provider_uuid={prov_uuid[:8] if prov_uuid else 'null'}…  "
              f"provider_name={prov_name!r}")

# Also find Unknown Provider UUID
r7 = g("/api/v1/providers", params={"per_page": 200, "page": 1})
p(f"\n  GET /api/v1/providers: HTTP {r7.status_code}")
if r7.status_code == 200:
    provs = (r7.json().get("data") or r7.json())
    if isinstance(provs, list):
        for pv in provs:
            if "unknown" in str(pv.get("name","")).lower():
                p(f"  Found unknown provider: uuid={pv.get('uuid')}  name={pv.get('name')!r}")
        p(f"  Total providers page 1: {len(provs)}")

# ─── 5. Contract enquiry — W5_plans.csv analysis ─────────────────────────────
p("\n=== 5. CONTRACT ENQUIRY SOURCE ANALYSIS ===")
SRC_ACCT = "db0ab7f5-8e52-4c89-a915-29fefeeca0d4"
JIM   = "1a230fbc-4026-4bd1-851f-18cb55ab19f2"
KATIE = "78a6aab8-55f3-42ed-b8ef-4867da1c0c78"

plans_with_ce = []
for row in w5:
    ce = (row.get("contract_enquiry_account_uuid") or "").strip()
    if ce:
        plans_with_ce.append(row)

p(f"  W5 rows with non-blank contract_enquiry_account_uuid: {len(plans_with_ce)}")

# What values appear?
ce_values = defaultdict(int)
for row in plans_with_ce:
    ce_values[(row.get("contract_enquiry_account_uuid") or "").strip()] += 1
for val, cnt in sorted(ce_values.items(), key=lambda x: -x[1])[:5]:
    p(f"  ce_uuid={val!r}: {cnt} plans")

# For plans with ce=SRC_ACCT, check active_adviser_uuid and seller_uuid
p(f"\n  For plans with ce=SRC_ACCT, adviser breakdown:")
adviser_counts = defaultdict(int)
for row in plans_with_ce:
    ce = (row.get("contract_enquiry_account_uuid") or "").strip()
    if ce == SRC_ACCT:
        adv = (row.get("active_adviser_uuid") or "").strip()
        sel = (row.get("seller_uuid") or "").strip()
        adviser_counts[(adv, sel)] += 1
for (adv, sel), cnt in sorted(adviser_counts.items(), key=lambda x: -x[1])[:5]:
    p(f"  active_adviser={adv[:8]}  seller={sel[:8]}  count={cnt}")

# ─── 6. Joint-owned plans in W5 ──────────────────────────────────────────────
p("\n=== 6. JOINT-OWNED PLANS IN W5 ===")
# Check how many plans have client_uuids in joint_src_uuids
joint_plans_in_w5 = [row for row in w5
                     if (row.get("client_uuids") or "").strip().split(",")[0].strip()
                     in joint_src_uuids]
p(f"  Plans in W5 with client_uuids = joint src UUID: {len(joint_plans_in_w5)}")

# Also check: are there any plans in W5 with comma-separated client_uuids (multiple owners)?
multi_owner = [row for row in w5
               if "," in (row.get("client_uuids") or "")]
p(f"  Plans in W5 with multiple client_uuids: {len(multi_owner)}")

# Also check dest-side: for joint dest UUIDs, fetch their plans on dest
p(f"\n  Checking dest plans for 3 joint dest UUIDs…")
for jdu in joint_dest_uuids[:3]:
    # Try filter[client_uuid]
    rj = g("/api/v1/plans", params={"filter[client_uuid]": jdu, "per_page": 200})
    if rj.status_code == 200:
        plans_j = (rj.json().get("data") or [])
        p(f"  joint {jdu[:8]}…: {len(plans_j)} plans via filter[client_uuid]")
    else:
        p(f"  joint {jdu[:8]}…: HTTP {rj.status_code} for filter[client_uuid]")
    # Also try client/{uuid}/plans
    rj2 = g(f"/api/v1/client/{jdu}/plans", params={"per_page": 200})
    p(f"    GET /api/v1/client/{jdu[:8]}/plans: HTTP {rj2.status_code}")
    if rj2.status_code == 200:
        plans_j2 = (rj2.json().get("data") or rj2.json())
        p(f"    → {len(plans_j2) if isinstance(plans_j2, list) else 'dict'} plans")

p("\n=== PROBE COMPLETE ===")
