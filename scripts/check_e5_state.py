"""Check if E5 partial plan was created and current plan count."""
import json, requests, sys
from azure.identity import DefaultAzureCredential
from azure.keyvault.secrets import SecretClient

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

BASE = "https://api.plannrcrm.com"
ACCT = "ccd7f98b-d6a4-46ad-bab3-242ab42826eb"
EXT_REF = "e5_flexi_drawdown_replacement_20261029"
sc  = SecretClient(vault_url="https://kv-plannr-ingestion.vault.azure.net/",
                   credential=DefaultAzureCredential())
pat = sc.get_secret("plannr-prod-pat").value
s   = requests.Session()
s.headers.update({"Authorization": f"Bearer {pat}", "Accept": "application/json",
                  "X-PLANNR-ACCOUNT-UUID": ACCT})

# Count all plans
total = 0
page = 1
while True:
    r = s.get(f"{BASE}/api/v1/plans", params={"page": page, "per_page": 200})
    if r.status_code != 200:
        break
    body = r.json()
    data = body.get("data", [])
    total += len(data)
    meta = body.get("meta", {})
    if page >= meta.get("last_page", 1):
        break
    page += 1
print(f"Total plans: {total}")

# Check import status
ri = s.get(f"{BASE}/api/v1/import/ade57807-bb46-4ff7-a799-41df458ac6fd")
d  = ri.json().get("data", ri.json())
print(f"Import status={d.get('status')} created={d.get('successful_rows')} failed={d.get('failed_rows')}")

# Look for plan by ext ref
found = None
page2 = 1
while True:
    r2 = s.get(f"{BASE}/api/v1/plans", params={"page": page2, "per_page": 200, "include": "external_references"})
    if r2.status_code != 200:
        break
    body2 = r2.json()
    for pl in body2.get("data", []):
        for ref in pl.get("external_references") or []:
            if ref.get("value") == EXT_REF:
                found = pl.get("uuid")
                print(f"Found plan by ext_ref: {found}")
    meta2 = body2.get("meta", {})
    if page2 >= meta2.get("last_page", 1):
        break
    page2 += 1
if not found:
    print(f"No plan found with ext_ref={EXT_REF}")
print("=== DONE ===")
