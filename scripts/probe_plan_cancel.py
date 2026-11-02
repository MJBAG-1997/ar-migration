"""Try to close/cancel old plan so expectations clear, then delete."""
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

# Try POST close/cancel
for path in [
    f"/api/v1/plans/{OLD}/close",
    f"/api/v1/plans/{OLD}/cancel",
    f"/api/v1/plans/{OLD}/archive",
    f"/api/v1/plans/{OLD}/deactivate",
]:
    r = s.post(f"{BASE}{path}", json={})
    print(f"POST {path.split('/')[-1]}  HTTP {r.status_code}  {r.text[:120]}")

# Try PUT status=inactive or ceased
for body in [
    {"status": "inactive"},
    {"status": "ceased"},
    {"status": "cancelled"},
    {"inactive_at": "2026-10-29"},
]:
    r = s.put(f"{BASE}/api/v1/plans/{OLD}", json=body)
    print(f"PUT {body}  HTTP {r.status_code}  {r.text[:150]}")

print("\n=== DONE ===")
