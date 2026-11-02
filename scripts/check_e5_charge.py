"""Check charges on E5 new plan and verify import result."""
import json, requests, sys
from azure.identity import DefaultAzureCredential
from azure.keyvault.secrets import SecretClient

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

BASE = "https://api.plannrcrm.com"
ACCT = "ccd7f98b-d6a4-46ad-bab3-242ab42826eb"
NEW_PLAN  = "45510cc2-84c0-465e-a02e-557b0d2d2799"
OLD_PLAN  = "fff4250d-a0c6-4d97-8fec-b81b3483904d"
IMP_UUID  = "60a636ae-dcf3-426b-a29c-c5c9819c4551"

sc  = SecretClient(vault_url="https://kv-plannr-ingestion.vault.azure.net/",
                   credential=DefaultAzureCredential())
pat = sc.get_secret("plannr-prod-pat").value
s   = requests.Session()
s.headers.update({"Authorization": f"Bearer {pat}", "Accept": "application/json",
                  "X-PLANNR-ACCOUNT-UUID": ACCT})

# Check new plan charges
r1 = s.get(f"{BASE}/api/v1/plans/{NEW_PLAN}/charges")
print(f"New plan charges HTTP {r1.status_code}")
if r1.status_code == 200:
    body = r1.json()
    chgs = body.get("data") or body
    if isinstance(chgs, list):
        print(f"  Charge count: {len(chgs)}")
        for ch in chgs:
            print(f"  uuid={ch.get('uuid','')[:8]}  type={ch.get('type')}  freq={ch.get('frequency')}  pct={ch.get('percentage')}")

# Check old plan charges
r2 = s.get(f"{BASE}/api/v1/plans/{OLD_PLAN}/charges")
print(f"Old plan charges HTTP {r2.status_code}")
if r2.status_code == 200:
    body2 = r2.json()
    chgs2 = body2.get("data") or body2
    if isinstance(chgs2, list):
        print(f"  Charge count: {len(chgs2)}")

# Check import
ri = s.get(f"{BASE}/api/v1/import/{IMP_UUID}")
d  = ri.json().get("data", ri.json())
print(f"\nImport {IMP_UUID[:8]}: status={d.get('status')} created={d.get('successful_rows')} total_rows={d.get('total_rows')}")
print("=== DONE ===")
