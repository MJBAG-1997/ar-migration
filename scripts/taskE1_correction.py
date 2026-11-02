"""E1 correction: POST 4 missing premiums to 3 plans; confirm 171 total."""
import csv, json, requests, sys
from azure.identity import DefaultAzureCredential
from azure.keyvault.secrets import SecretClient

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

BASE     = "https://api.plannrcrm.com"
ACCT     = "ccd7f98b-d6a4-46ad-bab3-242ab42826eb"
DATA_ROOT = "C:/dev/AR Migration/import_20261027"
STATE    = f"{DATA_ROOT}/state"
EXPECTED_TOTAL = 171

sc  = SecretClient(vault_url="https://kv-plannr-ingestion.vault.azure.net/",
                   credential=DefaultAzureCredential())
pat = sc.get_secret("plannr-prod-pat").value
s   = requests.Session()
s.headers.update({"Authorization": f"Bearer {pat}", "Accept": "application/json",
                  "Content-Type": "application/json", "X-PLANNR-ACCOUNT-UUID": ACCT})

def to_pennies(v):
    try: return int(round(float(v) * 100))
    except (TypeError, ValueError): return 0

PREM_DATE_MIN = "2023-03-01"

def parse_date(v):
    """Convert DD/MM/YYYY to YYYY-MM-DD; return None if blank or pre-minimum."""
    v = (v or "").strip()
    if not v:
        return None
    # DD/MM/YYYY → YYYY-MM-DD
    if "/" in v:
        parts = v.split("/")
        if len(parts) == 3:
            v = f"{parts[2]}-{parts[1].zfill(2)}-{parts[0].zfill(2)}"
    return v if v >= PREM_DATE_MIN else None

def build_prem_body(row):
    amt_str = (row.get("amount") or "").strip()
    pct_str = (row.get("percentage") or "").strip()
    freq    = (row.get("frequency") or "").strip()
    start   = parse_date(row.get("start_date"))
    src     = (row.get("source") or "").strip()
    pm      = (row.get("payment_method") or "").strip()
    body = {"frequency": freq}
    if amt_str:
        body["amount"] = to_pennies(amt_str)
    elif pct_str:
        body["percentage"] = float(pct_str)
    if start:   body["start_date"]      = start
    if src:     body["source"]          = src
    if pm:      body["payment_method"]  = pm
    return body

# Load missing plan state
with open(f"{STATE}/e1_premium_state.json") as f:
    state = json.load(f)
missing_plans = {m["plan_uuid"]: m for m in state["missing_plans"]}
print(f"Missing plans: {list(missing_plans.keys())}")

# Load W6 CSV (dest plan UUIDs in plan_uuid column)
prem_rows = {}  # plan_uuid → list of rows
with open(f"{DATA_ROOT}/W6_premiums_rewritten.csv", newline="", encoding="utf-8") as f:
    for row in csv.DictReader(f):
        dp = (row.get("plan_uuid") or "").strip()
        if dp in missing_plans:
            prem_rows.setdefault(dp, []).append(row)

print(f"W6 rows found for missing plans:")
for dp, rows in prem_rows.items():
    print(f"  plan={dp[:8]}  rows={len(rows)}")

# Count current total before
def count_all_premiums():
    total = 0
    page = 1
    # Use plan_map to iterate only plans that have premiums
    with open(f"{DATA_ROOT}/W6_premiums_rewritten.csv", newline="", encoding="utf-8") as f:
        plan_uuids = list({row["plan_uuid"] for row in csv.DictReader(f) if row.get("plan_uuid")})
    for dp in plan_uuids:
        r = s.get(f"{BASE}/api/v1/plans/{dp}/premiums")
        if r.status_code == 200:
            prems = r.json().get("data", r.json())
            if isinstance(prems, list): total += len(prems)
    return total

before = count_all_premiums()
print(f"\nPremiums before: {before}")

# Test: POST the first missing premium from first plan
first_plan = list(prem_rows.keys())[0]
test_row = prem_rows[first_plan][0]
body = build_prem_body(test_row)
print(f"\nTest POST: plan={first_plan[:8]}  body={body}")

rp = s.post(f"{BASE}/api/v1/plans/{first_plan}/premiums", json=body)
print(f"POST HTTP {rp.status_code}")
if rp.status_code not in (200, 201):
    print(f"!!STOP!! {rp.text[:300]}"); sys.exit(1)
new_prem = rp.json().get("data", rp.json())
new_uuid = new_prem.get("uuid", "")
print(f"  Created premium uuid={new_uuid[:8]}")

# GET verify
rg = s.get(f"{BASE}/api/v1/plans/{first_plan}/premiums")
got = rg.json().get("data", rg.json())
if not isinstance(got, list): got = []
print(f"  GET verify: {len(got)} premium(s) on plan")

# POST remaining
posted = 1  # count the test one
for dp, rows in prem_rows.items():
    start_idx = 1 if dp == first_plan else 0
    for row in rows[start_idx:]:
        body2 = build_prem_body(row)
        rp2 = s.post(f"{BASE}/api/v1/plans/{dp}/premiums", json=body2)
        if rp2.status_code in (200, 201):
            posted += 1
            print(f"  Posted premium to plan={dp[:8]}: {body2}  HTTP {rp2.status_code}")
        else:
            print(f"  !!STOP!! POST plan={dp[:8]} HTTP {rp2.status_code}: {rp2.text[:200]}")
            sys.exit(1)

print(f"\nTotal posted: {posted}")

# Count after
after = count_all_premiums()
print(f"Premiums after: {after}  (expected {EXPECTED_TOTAL})")
if after != EXPECTED_TOTAL:
    print(f"!!STOP!! count mismatch: got {after}, expected {EXPECTED_TOTAL}")
    sys.exit(1)
print("\n=== E1 DONE ===")
