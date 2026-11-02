"""Probe correct delete endpoint for charges/expectations on old plan."""
import requests, sys
from azure.identity import DefaultAzureCredential
from azure.keyvault.secrets import SecretClient

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

BASE = "https://api.plannrcrm.com"
ACCT = "ccd7f98b-d6a4-46ad-bab3-242ab42826eb"
OLD       = "fff4250d-a0c6-4d97-8fec-b81b3483904d"
CH_UUID   = "16367e0b-b65f-4126-bb4d-97c679842f6e"
EXP_UUID  = "e7c6b80b-c709-4a46-9af0-74a809b6c6f1"

sc  = SecretClient(vault_url="https://kv-plannr-ingestion.vault.azure.net/",
                   credential=DefaultAzureCredential())
pat = sc.get_secret("plannr-prod-pat").value
s   = requests.Session()
s.headers.update({"Authorization": f"Bearer {pat}", "Accept": "application/json",
                  "Content-Type": "application/json", "X-PLANNR-ACCOUNT-UUID": ACCT})

# Try GET on the charge directly
for url in [
    f"/api/v1/charges/{CH_UUID}",
    f"/api/v1/plans/{OLD}/charges/{CH_UUID}",
]:
    r = s.get(f"{BASE}{url}")
    print(f"GET {url}  HTTP {r.status_code}")

# Try GET on expectation directly
for url in [
    f"/api/v1/expectations/{EXP_UUID}",
    f"/api/v1/plans/{OLD}/expectations/{EXP_UUID}",
]:
    r = s.get(f"{BASE}{url}")
    print(f"GET {url}  HTTP {r.status_code}")

# Try DELETE directly
print()
for url in [
    f"/api/v1/charges/{CH_UUID}",
    f"/api/v1/expectations/{EXP_UUID}",
]:
    r = s.delete(f"{BASE}{url}")
    print(f"DELETE {url}  HTTP {r.status_code}  {r.text[:120]}")

print("\n=== DONE ===")
