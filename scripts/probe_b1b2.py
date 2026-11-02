"""Probe B1 provider coverage across W4+W5, and B2 contract_enquiry field name."""
import csv, json, os, sys, requests
from azure.identity import DefaultAzureCredential
from azure.keyvault.secrets import SecretClient

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

IMP   = r"C:\dev\AR Migration\import_20261027"
STATE = os.path.join(IMP, "state")
BASE  = "https://api.plannrcrm.com"
ACCT  = "ccd7f98b-d6a4-46ad-bab3-242ab42826eb"
UNKNOWN = "175e2a37-200f-400f-8140-e0d7206db1b1"
JIM     = "1a230fbc-4026-4bd1-851f-18cb55ab19f2"

plan_map = json.load(open(os.path.join(STATE,"plan_map.json")))
inv_plan = {v: k for k, v in plan_map.items()}

def lcsv(p):
    with open(p, newline="", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))

w4  = lcsv(os.path.join(IMP, "W4_plans.csv"))
w5  = lcsv(os.path.join(IMP, "W5_plans.csv"))

print(f"W4 rows: {len(w4)}  W5 rows: {len(w5)}")

# Build combined src lookup
src_lookup = {}
for row in w4:
    src = (row.get("external_reference_value") or "").strip()
    if src:
        src_lookup[src] = row
for row in w5:
    src = (row.get("external_reference_value") or "").strip()
    if src:
        src_lookup[src] = row

# How many W4 rows have real provider?
w4_real = sum(1 for r in w4 if (r.get("provider_uuid") or "").strip() not in ("", UNKNOWN))
w5_real = sum(1 for r in w5 if (r.get("provider_uuid") or "").strip() not in ("", UNKNOWN))
print(f"W4 rows with real provider: {w4_real}")
print(f"W5 rows with real provider: {w5_real}")

# Sample 3 unknown-provider dest plan UUIDs
samples = [
    "8f01a57d-b573-485a-9c7e-7e422db73f6c",
    "6c920658-cd82-4706-b1f2-0512d5d886d6",
    "b23a5386-ad5b-44e7-a024-df445240b51e",
]
for du in samples:
    src = inv_plan.get(du, "")
    row = src_lookup.get(src, {})
    pu  = (row.get("provider_uuid") or "").strip()
    file_hit = "W4" if src in {(r.get("external_reference_value") or "").strip() for r in w4} else "W5"
    print(f"  dest={du[:8]} src={src[:8] if src else 'MISS'} file={file_hit} prov={pu[:8] if pu else 'blank'} type={row.get('type','?')}")

# Probe contract_enquiry field
sc = SecretClient(vault_url="https://kv-plannr-ingestion.vault.azure.net/",
                  credential=DefaultAzureCredential())
pat = sc.get_secret("plannr-prod-pat").value
s = requests.Session()
s.headers.update({"Authorization": f"Bearer {pat}", "Accept": "application/json",
                  "Content-Type": "application/json", "X-PLANNR-ACCOUNT-UUID": ACCT})

# Use the plan that was just stopped on
w5_ce = [r for r in w5 if (r.get("contract_enquiry_account_uuid") or "").strip()]
if w5_ce:
    src_pl  = (w5_ce[0].get("external_reference_value") or "").strip()
    dest_pl = plan_map.get(src_pl, "")
    print(f"\nCE probe plan: {dest_pl}")

    # GET before
    rg0 = s.get(f"{BASE}/api/v1/plans/{dest_pl}")
    b0  = rg0.json()
    ce_keys = {k: v for k, v in b0.items() if "contract" in k.lower() or "enquiry" in k.lower()}
    print(f"CE fields before PUT: {ce_keys}")

    # PUT
    rp = s.put(f"{BASE}/api/v1/plans/{dest_pl}",
               json={"contract_enquiry_account_uuid": JIM})
    print(f"PUT -> HTTP {rp.status_code}")
    if rp.status_code in (200, 201):
        pb = rp.json()
        ce_put = {k: v for k, v in pb.items() if "contract" in k.lower() or "enquiry" in k.lower()}
        print(f"CE fields in PUT response: {ce_put}")

    # GET after
    rg = s.get(f"{BASE}/api/v1/plans/{dest_pl}")
    ga = rg.json()
    ce_after = {k: v for k, v in ga.items() if "contract" in k.lower() or "enquiry" in k.lower()}
    print(f"CE fields after GET:  {ce_after}")

print("\n=== PROBE DONE ===")
