"""B1 correction: PUT provider_uuid=null for no-provider sub_types; rebuild AR list."""
import csv, json, requests, sys, time
from azure.identity import DefaultAzureCredential
from azure.keyvault.secrets import SecretClient

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

BASE = "https://api.plannrcrm.com"
ACCT = "ccd7f98b-d6a4-46ad-bab3-242ab42826eb"
DATA_ROOT = "C:/dev/AR Migration/import_20261027"
STATE = f"{DATA_ROOT}/state"

NO_PROV_SUB_TYPES = {"fixed_asset", "cash_account", "group_protection", "loan_credit"}

sc  = SecretClient(vault_url="https://kv-plannr-ingestion.vault.azure.net/",
                   credential=DefaultAzureCredential())
pat = sc.get_secret("plannr-prod-pat").value
s   = requests.Session()
s.headers.update({"Authorization": f"Bearer {pat}", "Accept": "application/json",
                  "Content-Type": "application/json", "X-PLANNR-ACCOUNT-UUID": ACCT})

def unwrap(r):
    d = r.json()
    return d.get("data", d)

# Load b1_ar_list (522 dest UUIDs)
with open(f"{STATE}/b1_ar_list.json") as f:
    b1 = json.load(f)
b1_uuids = {e["uuid"] for e in b1}
print(f"B1 list: {len(b1_uuids)} plans")

# Build sub_type lookup from original W4+W5 (external_reference_value → sub_type)
plan_sub_type_src = {}
for fn in [f"{DATA_ROOT}/W4_plans.csv", f"{DATA_ROOT}/W5_plans.csv"]:
    with open(fn, newline="", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            ext_ref = row.get("external_reference_value", "").strip()
            sub_t   = row.get("sub_type", "").strip()
            if ext_ref:
                plan_sub_type_src[ext_ref] = sub_t

with open(f"{STATE}/plan_map.json") as f:
    plan_map = json.load(f)
inv_plan = {v: k for k, v in plan_map.items()}

# Map dest uuid → sub_type via inv_plan → plan_sub_type_src
needs_null = []  # PUT provider_uuid=null
needs_prov = []  # stays on AR list
sub_type_unknown = []

for e in b1:
    dest_uuid = e["uuid"]
    src_uuid  = inv_plan.get(dest_uuid, "")
    sub_type  = plan_sub_type_src.get(src_uuid, "")
    if not sub_type:
        sub_type_unknown.append({"uuid": dest_uuid, "src": src_uuid})
    elif sub_type in NO_PROV_SUB_TYPES:
        needs_null.append({"uuid": dest_uuid, "sub_type": sub_type})
    else:
        needs_prov.append({"uuid": dest_uuid, "sub_type": sub_type, "reason": "src_no_provider"})

print(f"needs_null (PUT provider=null): {len(needs_null)}")
print(f"needs_prov (AR list):           {len(needs_prov)}")
print(f"sub_type unknown:               {len(sub_type_unknown)}")
if sub_type_unknown:
    for x in sub_type_unknown[:5]:
        print(f"  {x}")

# Test: one PUT provider_uuid=null
test = needs_null[0]
r = s.put(f"{BASE}/api/v1/plans/{test['uuid']}", json={"provider_uuid": None})
print(f"\nTest PUT null provider HTTP {r.status_code}")
if r.status_code != 200:
    print(f"!!STOP!! {r.text[:200]}"); sys.exit(1)
pl = unwrap(r)
prov = pl.get("provider")
print(f"  provider after PUT: {prov}")
if prov is not None:
    print("!!STOP!! provider not cleared"); sys.exit(1)
print("Test OK")

# GET it back
rg = s.get(f"{BASE}/api/v1/plans/{test['uuid']}")
pl2 = unwrap(rg)
prov2 = pl2.get("provider")
print(f"GET verify provider: {prov2}")
if prov2 is not None:
    print("!!STOP!! provider still set on GET"); sys.exit(1)

# Bulk PUT the rest
errors = []
for i, item in enumerate(needs_null[1:], 1):
    r = s.put(f"{BASE}/api/v1/plans/{item['uuid']}", json={"provider_uuid": None})
    if r.status_code != 200:
        errors.append({"uuid": item["uuid"], "http": r.status_code, "body": r.text[:100]})
    if (i % 50) == 0:
        print(f"  ...PUT {i+1}/{len(needs_null)} done")

print(f"\nBulk PUT complete: {len(needs_null)} total, {len(errors)} errors")
if errors:
    for e in errors[:5]:
        print(f"  {e}")

# Rebuild AR list (only those that genuinely need a provider)
ar_list = needs_prov[:]
if sub_type_unknown:
    # Keep unknowns on AR list as "sub_type unknown"
    for x in sub_type_unknown:
        ar_list.append({"uuid": x["uuid"], "sub_type": "unknown", "reason": "src_no_provider"})

with open(f"{STATE}/b1_ar_list.json", "w") as f:
    json.dump(ar_list, f, indent=2)
print(f"\nSaved b1_ar_list.json: {len(ar_list)} entries (need manual provider)")
print(f"\n=== B1 DONE ===")
