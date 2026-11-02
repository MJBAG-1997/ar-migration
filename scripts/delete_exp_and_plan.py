"""Delete expectation, charge, then old plan using correct singular paths."""
import requests, sys
from azure.identity import DefaultAzureCredential
from azure.keyvault.secrets import SecretClient

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

BASE     = "https://api.plannrcrm.com"
ACCT     = "ccd7f98b-d6a4-46ad-bab3-242ab42826eb"
OLD      = "fff4250d-a0c6-4d97-8fec-b81b3483904d"
CH_UUID  = "16367e0b-b65f-4126-bb4d-97c679842f6e"
EXP_UUID = "e7c6b80b-c709-4a46-9af0-74a809b6c6f1"

sc  = SecretClient(vault_url="https://kv-plannr-ingestion.vault.azure.net/",
                   credential=DefaultAzureCredential())
pat = sc.get_secret("plannr-prod-pat").value
s   = requests.Session()
s.headers.update({"Authorization": f"Bearer {pat}", "Accept": "application/json",
                  "Content-Type": "application/json", "X-PLANNR-ACCOUNT-UUID": ACCT})

# 1) Delete expectation
r1 = s.delete(f"{BASE}/api/v1/expectation/{EXP_UUID}")
print(f"DELETE expectation HTTP {r1.status_code}  {r1.text[:150]}")

# 2) Delete charge
r2 = s.delete(f"{BASE}/api/v1/charge/{CH_UUID}")
print(f"DELETE charge HTTP {r2.status_code}  {r2.text[:150]}")

# 3) Verify 0 expectations on plan
re = s.get(f"{BASE}/api/v1/plans/{OLD}/expectations")
exps = re.json().get("data", re.json())
if not isinstance(exps, list): exps = []
print(f"Expectations remaining: {len(exps)}")

# 4) Delete plan
r3 = s.delete(f"{BASE}/api/v1/plans/{OLD}")
print(f"DELETE plan HTTP {r3.status_code}  {r3.text[:200]}")
if r3.status_code not in (200, 204):
    print("!!STOP!! DELETE plan failed"); import sys; sys.exit(1)

# 5) Count plans
total = 0; page = 1
while True:
    r = s.get(f"{BASE}/api/v1/plans", params={"page": page, "per_page": 200})
    if r.status_code != 200: break
    body = r.json(); total += len(body.get("data", []))
    if page >= body.get("meta", {}).get("last_page", 1): break
    page += 1
print(f"\nTotal plans after delete: {total}  (expected 1901)")
print("\n=== E5 DONE ===")
