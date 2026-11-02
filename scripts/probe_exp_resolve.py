"""Try v2 and mark-paid approaches to clear the expectation."""
import requests, sys
from azure.identity import DefaultAzureCredential
from azure.keyvault.secrets import SecretClient

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

BASE = "https://api.plannrcrm.com"
ACCT = "ccd7f98b-d6a4-46ad-bab3-242ab42826eb"
OLD      = "fff4250d-a0c6-4d97-8fec-b81b3483904d"
CH_UUID  = "16367e0b-b65f-4126-bb4d-97c679842f6e"
EXP_UUID = "e7c6b80b-c709-4a46-9af0-74a809b6c6f1"

sc  = SecretClient(vault_url="https://kv-plannr-ingestion.vault.azure.net/",
                   credential=DefaultAzureCredential())
pat = sc.get_secret("plannr-prod-pat").value
s   = requests.Session()
s.headers.update({"Authorization": f"Bearer {pat}", "Accept": "application/json",
                  "Content-Type": "application/json", "X-PLANNR-ACCOUNT-UUID": ACCT})

# Show full expectation object
re_list = s.get(f"{BASE}/api/v1/plans/{OLD}/expectations")
exps = re_list.json().get("data", re_list.json())
if isinstance(exps, list) and exps:
    print("Expectation keys:", list(exps[0].keys()))
    print("Full:", exps[0])

# v2 plan delete
r2 = s.delete(f"{BASE}/api/v2/plans/{OLD}")
print(f"\nDELETE /api/v2/plans HTTP {r2.status_code}  {r2.text[:150]}")

# Try mark-as-paid POST variants
for path in [
    f"/api/v1/plans/{OLD}/expectations/{EXP_UUID}/pay",
    f"/api/v1/plans/{OLD}/expectations/{EXP_UUID}/waive",
    f"/api/v1/plans/{OLD}/expectations/{EXP_UUID}/mark-paid",
    f"/api/v1/expectations/{EXP_UUID}/pay",
]:
    r = s.post(f"{BASE}{path}", json={})
    print(f"POST {path}  HTTP {r.status_code}  {r.text[:100]}")

# Try PUT on expectation
for body in [{"status": "paid"}, {"is_paid": True}, {"waived": True}]:
    r = s.put(f"{BASE}/api/v1/plans/{OLD}/expectations/{EXP_UUID}", json=body)
    print(f"PUT expectations {body}  HTTP {r.status_code}  {r.text[:100]}")

print("\n=== DONE ===")
