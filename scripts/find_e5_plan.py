"""Find the newly created E5 plan by policy number or recent creation."""
import json, requests, sys
from azure.identity import DefaultAzureCredential
from azure.keyvault.secrets import SecretClient

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

BASE = "https://api.plannrcrm.com"
ACCT = "ccd7f98b-d6a4-46ad-bab3-242ab42826eb"
POLICY = "IH01853021"
OLD_UUID = "fff4250d-a0c6-4d97-8fec-b81b3483904d"

sc  = SecretClient(vault_url="https://kv-plannr-ingestion.vault.azure.net/",
                   credential=DefaultAzureCredential())
pat = sc.get_secret("plannr-prod-pat").value
s   = requests.Session()
s.headers.update({"Authorization": f"Bearer {pat}", "Accept": "application/json",
                  "X-PLANNR-ACCOUNT-UUID": ACCT})

# Find plans with matching policy number
matches = []
page = 1
while True:
    r = s.get(f"{BASE}/api/v1/plans", params={"page": page, "per_page": 200})
    if r.status_code != 200:
        print(f"HTTP {r.status_code}"); break
    body = r.json()
    for pl in body.get("data", []):
        if pl.get("policy_number") == POLICY:
            matches.append(pl)
    meta = body.get("meta", {})
    if page >= meta.get("last_page", 1):
        break
    page += 1

print(f"Plans with policy {POLICY}: {len(matches)}")
for pl in matches:
    uuid = pl.get("uuid","")
    sub  = pl.get("sub_type","?")
    typ  = pl.get("type","?")
    crat = pl.get("created_at","?")
    print(f"  uuid={uuid}  type={typ}/{sub}  created={str(crat)[:19]}  old={uuid==OLD_UUID}")

# Also check import bbd15f14
ri = s.get(f"{BASE}/api/v1/import/bbd15f14-85f3-407a-be3f-2daceffcb751")
d  = ri.json().get("data", ri.json())
print(f"\nImport bbd15f14: status={d.get('status')} created={d.get('successful_rows')} failed={d.get('failed_rows')} total_rows={d.get('total_rows')}")

print("=== DONE ===")
