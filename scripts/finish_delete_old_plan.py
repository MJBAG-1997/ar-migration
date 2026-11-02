"""Mark plan inactive then delete it."""
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

# Set fully inactive
r1 = s.put(f"{BASE}/api/v1/plans/{OLD}", json={
    "status": "inactive",
    "inactive_reason": "other",
    "inactive_at": "2026-10-29"
})
print(f"PUT inactive HTTP {r1.status_code}")
if r1.status_code == 200:
    pl = r1.json().get("data", r1.json())
    print(f"  status={pl.get('status')}  inactive_at={pl.get('inactive_at')}")

# Check expectations now
re = s.get(f"{BASE}/api/v1/plans/{OLD}/expectations")
body = re.json()
exps = body.get("data", body) if isinstance(body, dict) else body
if not isinstance(exps, list): exps = []
print(f"\nExpectations after inactive: {len(exps)}")

# Try DELETE
rdp = s.delete(f"{BASE}/api/v1/plans/{OLD}")
print(f"\nDELETE plan HTTP {rdp.status_code}")
if rdp.status_code in (200, 204):
    print("Old plan deleted OK")
else:
    print(f"  Body: {rdp.text[:300]}")

# Count plans
total = 0; page = 1
while True:
    r = s.get(f"{BASE}/api/v1/plans", params={"page": page, "per_page": 200})
    if r.status_code != 200: break
    body2 = r.json(); total += len(body2.get("data", []))
    if page >= body2.get("meta", {}).get("last_page", 1): break
    page += 1
print(f"\nTotal plans: {total}  (expected 1901 if deleted, 1902 if blocked)")
print("\n=== DONE ===")
