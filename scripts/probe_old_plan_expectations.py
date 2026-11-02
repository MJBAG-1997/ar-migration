"""Probe what's blocking deletion of old flexi plan."""
import json, requests, sys
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

# GET old plan full details
r = s.get(f"{BASE}/api/v1/plans/{OLD}")
print(f"GET plan HTTP {r.status_code}")
if r.status_code == 200:
    pl = r.json().get("data", r.json())
    for k, v in sorted(pl.items()):
        if v is not None and v != "" and v != [] and v != {}:
            if not isinstance(v, (dict, list)):
                print(f"  {k}: {str(v)[:80]}")

# GET charges
rc = s.get(f"{BASE}/api/v1/plans/{OLD}/charges")
print(f"\nGET charges HTTP {rc.status_code}")
if rc.status_code == 200:
    body = rc.json()
    chgs = body.get("data") or body
    if isinstance(chgs, list):
        print(f"  Charges: {len(chgs)}")
        for ch in chgs:
            print(f"  uuid={ch.get('uuid','')[:8]}  next_expectation={ch.get('next_expectation_date')}")

# Try expectations endpoint
for ep in ["/api/v1/plans/{}/expectations", "/api/v1/charge-expectations"]:
    if "{}" in ep:
        ep = ep.format(OLD)
    re = s.get(f"{BASE}{ep}", params={"page":1,"per_page":50})
    print(f"\nGET {ep} HTTP {re.status_code}")
    if re.status_code == 200:
        body2 = re.json()
        data2 = body2.get("data") or body2
        if isinstance(data2, list):
            print(f"  Count: {len(data2)}")
            for item in data2[:3]:
                print(f"  {list(item.keys())[:6]}")

print("\n=== DONE ===")
