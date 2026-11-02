"""Delete old E5 flexi plan: expectation → charge → plan."""
import requests, sys
from azure.identity import DefaultAzureCredential
from azure.keyvault.secrets import SecretClient

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

BASE = "https://api.plannrcrm.com"
ACCT = "ccd7f98b-d6a4-46ad-bab3-242ab42826eb"
OLD  = "fff4250d-a0c6-4d97-8fec-b81b3483904d"

sc  = SecretClient(vault_url="https://kv-plannr-ingestion.vault.azure.net/",
                   credential=DefaultAzureCredential())
pat = sc.get_secret("plannr-prod-pat").value
s   = requests.Session()
s.headers.update({"Authorization": f"Bearer {pat}", "Accept": "application/json",
                  "Content-Type": "application/json", "X-PLANNR-ACCOUNT-UUID": ACCT})

# Step 1: get full charge UUID
rc = s.get(f"{BASE}/api/v1/plans/{OLD}/charges")
chgs = rc.json().get("data", rc.json()) if isinstance(rc.json(), dict) else rc.json()
if not isinstance(chgs, list): chgs = []
print(f"Charges on old plan: {len(chgs)}")
for ch in chgs:
    full_uuid = ch.get("uuid", "")
    print(f"  charge uuid={full_uuid}  next_exp={ch.get('next_expectation_date')}")

# Step 2: get expectations and delete each
re = s.get(f"{BASE}/api/v1/plans/{OLD}/expectations")
print(f"\nGET expectations HTTP {re.status_code}")
exps = []
if re.status_code == 200:
    body = re.json()
    exps = body.get("data", body) if isinstance(body, dict) else body
    if not isinstance(exps, list): exps = []
    print(f"  Count: {len(exps)}")
    for exp in exps:
        exp_uuid = exp.get("uuid", "")
        print(f"  exp uuid={exp_uuid}  type={exp.get('type')}  cat={exp.get('category')}")

for exp in exps:
    exp_uuid = exp.get("uuid", "")
    rd = s.delete(f"{BASE}/api/v1/plans/{OLD}/expectations/{exp_uuid}")
    print(f"\nDELETE expectation {exp_uuid[:8]} HTTP {rd.status_code}")
    if rd.status_code not in (200, 204):
        print(f"  Body: {rd.text[:200]}")

# Step 3: delete each charge
for ch in chgs:
    full_uuid = ch.get("uuid", "")
    rdc = s.delete(f"{BASE}/api/v1/plans/{OLD}/charges/{full_uuid}")
    print(f"\nDELETE charge {full_uuid[:8]} HTTP {rdc.status_code}")
    if rdc.status_code not in (200, 204):
        print(f"  Body: {rdc.text[:200]}")

# Verify charges gone
rc2 = s.get(f"{BASE}/api/v1/plans/{OLD}/charges")
chgs2 = rc2.json().get("data", rc2.json()) if isinstance(rc2.json(), dict) else rc2.json()
if not isinstance(chgs2, list): chgs2 = []
print(f"\nCharges remaining: {len(chgs2)}")

# Step 4: delete plan
rdp = s.delete(f"{BASE}/api/v1/plans/{OLD}")
print(f"\nDELETE plan HTTP {rdp.status_code}")
if rdp.status_code in (200, 204):
    print("Old plan deleted OK")
else:
    print(f"  Body: {rdp.text[:300]}")
    print("!!STOP!! Could not delete old plan")
    sys.exit(1)

# Count plans
total = 0
page = 1
while True:
    r = s.get(f"{BASE}/api/v1/plans", params={"page": page, "per_page": 200})
    if r.status_code != 200: break
    body = r.json(); data = body.get("data", [])
    total += len(data)
    if page >= body.get("meta", {}).get("last_page", 1): break
    page += 1
print(f"\nTotal plans after delete: {total}  (expected 1901)")

print("\n=== DONE ===")
