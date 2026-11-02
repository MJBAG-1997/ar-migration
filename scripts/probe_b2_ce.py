"""Probe B2: find the correct plan keys and contract_enquiry field."""
import csv, json, os, sys, requests
from azure.identity import DefaultAzureCredential
from azure.keyvault.secrets import SecretClient

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

IMP   = r"C:\dev\AR Migration\import_20261027"
STATE = os.path.join(IMP, "state")
BASE  = "https://api.plannrcrm.com"
ACCT  = "ccd7f98b-d6a4-46ad-bab3-242ab42826eb"
JIM   = "1a230fbc-4026-4bd1-851f-18cb55ab19f2"

plan_map = json.load(open(os.path.join(STATE, "plan_map.json")))

def lcsv(p):
    with open(p, newline="", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))

w5 = lcsv(os.path.join(IMP, "W5_plans.csv"))

sc = SecretClient(vault_url="https://kv-plannr-ingestion.vault.azure.net/",
                  credential=DefaultAzureCredential())
pat = sc.get_secret("plannr-prod-pat").value
s = requests.Session()
s.headers.update({"Authorization": f"Bearer {pat}", "Accept": "application/json",
                  "Content-Type": "application/json", "X-PLANNR-ACCOUNT-UUID": ACCT})

def get_plan(uuid):
    r = s.get(f"{BASE}/api/v1/plans/{uuid}")
    body = r.json()
    return (body.get("data") or body) if isinstance(body, dict) else body

w5_ce = [r for r in w5 if (r.get("contract_enquiry_account_uuid") or "").strip()]
src_pl  = (w5_ce[0].get("external_reference_value") or "").strip()
dest_pl = plan_map.get(src_pl, "")
print(f"Test plan dest UUID: {dest_pl}")

plan = get_plan(dest_pl)
print(f"Plan keys ({len(plan)}):")
for k, v in sorted(plan.items()):
    vstr = repr(v)[:100] if not isinstance(v, (dict, list)) else f"({type(v).__name__} {list(v.keys()) if isinstance(v,dict) else len(v)} items)"
    print(f"  {k}: {vstr}")

print("\n--- Contract/enquiry-related keys ---")
for k, v in plan.items():
    if "contract" in k.lower() or "enquiry" in k.lower() or "case" in k.lower():
        print(f"  {k}: {v!r}")

# Try PUT with the field name
print("\n--- PUT test ---")
rp = s.put(f"{BASE}/api/v1/plans/{dest_pl}",
           json={"contract_enquiry_account_uuid": JIM})
print(f"PUT HTTP {rp.status_code}")
rp_data = rp.json()
if "data" in rp_data:
    rp_data = rp_data["data"]
for k, v in rp_data.items():
    if "contract" in k.lower() or "enquiry" in k.lower():
        print(f"  response {k}: {v!r}")

plan2 = get_plan(dest_pl)
for k, v in plan2.items():
    if "contract" in k.lower() or "enquiry" in k.lower():
        print(f"  after GET {k}: {v!r}")

print("\n=== PROBE DONE ===")
