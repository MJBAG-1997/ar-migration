"""
fix_b1_noprov.py — Task 5: clear provider_uuid=null for plan sub_types
that don't need a provider (property, mortgage, physical assets).
Count before, write one first + GET back, then rest. Count after.
Rule 9: UUIDs only in output.
"""
import json, os, sys, time
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from azure.identity import DefaultAzureCredential
from azure.keyvault.secrets import SecretClient
import requests
import openpyxl

BASE_URL          = "https://api.plannrcrm.com"
DEST_ACCOUNT_UUID = "ccd7f98b-d6a4-46ad-bab3-242ab42826eb"
IMPORT_DIR        = r"C:\dev\AR Migration\import_20261027"
STATE_DIR         = os.path.join(IMPORT_DIR, "state")
B1_AR_LIST        = os.path.join(STATE_DIR, "b1_ar_list.json")
WORKBOOK_PATH     = r"C:\dev\AR Migration\hfp_AR_actions_20261029.xlsx"

# Sub-types that don't need a financial provider
NO_PROV_TYPES = {
    # Property
    "property", "property_rental", "property_investment", "property_btl",
    "property_overseas", "property_other", "property_holdings",
    "property_commercial", "property_development",
    # Mortgage / debt
    "mortgage", "mortgage_buy_to_let", "mortgage_business_buy_to_let",
    "mortgage_unknown", "mortgage_endowment", "mortgage_interest_only",
    # Physical assets
    "motor_vehicle", "land", "own_business", "collectible_art_other",
    "collectible_other", "trust",
    # Unknown physical/debt variants
    "fixed_asset_unknown", "cash_account_unknown", "loan_credit_unknown",
}

sc       = SecretClient(vault_url="https://kv-plannr-ingestion.vault.azure.net/",
                        credential=DefaultAzureCredential())
dest_pat = sc.get_secret("plannr-prod-pat").value

class API:
    def __init__(self, pat, acct):
        self._s = requests.Session()
        self._s.headers.update({
            "Authorization":         f"Bearer {pat}",
            "Accept":                "application/json",
            "X-PLANNR-ACCOUNT-UUID": acct,
        })
    def get(self, path, **kw):
        return self._req("GET", path, **kw)
    def put(self, path, **kw):
        return self._req("PUT", path, **kw)
    def _req(self, method, path, **kw):
        url    = f"{BASE_URL}{path}" if path.startswith("/") else path
        is_get = (method == "GET")
        kw.setdefault("timeout", 60 if is_get else 120)
        max_attempts = 6 if is_get else 1
        for attempt in range(max_attempts):
            try:
                r = self._s.request(method, url, **kw)
            except requests.exceptions.RequestException as e:
                if is_get and attempt < 5:
                    wait = 15 * (attempt + 1)
                    print(f"  [NET] {method} {path[:60]} retry in {wait}s", flush=True)
                    time.sleep(wait)
                    continue
                raise
            if r.status_code != 429:
                return r
            wait = int(r.headers.get("Retry-After", 30))
            print(f"  [429] {method} {path[:60]} — {wait}s", flush=True); time.sleep(wait)
        return r

def load_json(path):
    with open(path, encoding="utf-8") as f: return json.load(f)

dest = API(dest_pat, DEST_ACCOUNT_UUID)

# Load AR list
ar_list = load_json(B1_AR_LIST)
print(f"Loaded {len(ar_list)} plans from b1_ar_list.json", flush=True)

# Partition into clear vs keep
to_clear = [p for p in ar_list if p.get("sub_type") in NO_PROV_TYPES]
to_keep  = [p for p in ar_list if p.get("sub_type") not in NO_PROV_TYPES]

print(f"Plans to clear (no provider needed): {len(to_clear)}", flush=True)
print(f"Plans to keep (provider review needed): {len(to_keep)}", flush=True)
print(f"Sub-types to keep:", flush=True)
from collections import Counter
keep_dist = Counter(p.get("sub_type") for p in to_keep)
for st, n in keep_dist.most_common():
    print(f"  {n:>4}  {st}", flush=True)

if not to_clear:
    print("Nothing to clear. Exiting.", flush=True)
    sys.exit(0)

# ── One first, GET back ───────────────────────────────────────────────────────
first = to_clear[0]
print(f"\nProbe: PUT plan {first['uuid'][:8]} sub_type={first['sub_type']}", flush=True)
r = dest.put(f"/api/v1/plans/{first['uuid']}", json={"provider_uuid": None})
if r.status_code not in (200, 201):
    print(f"  REFUSED: HTTP {r.status_code} {r.text[:200]}", flush=True)
    print("STOP: first write refused", flush=True)
    sys.exit(1)
print(f"  PUT HTTP {r.status_code}", flush=True)

# GET back and confirm provider_uuid is null
rg = dest.get(f"/api/v1/plans/{first['uuid']}")
if rg.status_code != 200:
    print(f"  GET back HTTP {rg.status_code}", flush=True)
    sys.exit(1)
body = rg.json()
plan = body.get("data") or body
prov = plan.get("provider_uuid") if isinstance(plan, dict) else "?"
print(f"  GET back provider_uuid={prov} — {'OK' if prov is None else 'UNEXPECTED'}", flush=True)
if prov is not None:
    print("STOP: provider_uuid not null after PUT", flush=True)
    sys.exit(1)

# ── Count before ─────────────────────────────────────────────────────────────
# "Before" = ar_list length = 407
before = len(ar_list)
print(f"\nBefore: {before} plans in AR list", flush=True)

# ── Clear remaining ───────────────────────────────────────────────────────────
print(f"\nClearing {len(to_clear) - 1} remaining plans...", flush=True)
cleared = 1
refused = []

for plan in to_clear[1:]:
    r = dest.put(f"/api/v1/plans/{plan['uuid']}", json={"provider_uuid": None})
    if r.status_code in (200, 201):
        cleared += 1
        if cleared % 50 == 0:
            print(f"  {cleared}/{len(to_clear)} cleared", flush=True)
    else:
        refused.append({"uuid": plan["uuid"], "sub_type": plan.get("sub_type"),
                        "http": r.status_code, "body": r.text[:100]})
        print(f"  [WARN] plan {plan['uuid'][:8]} sub_type={plan.get('sub_type')} "
              f"HTTP {r.status_code}", flush=True)

print(f"\nCleared: {cleared} / {len(to_clear)}", flush=True)
if refused:
    print(f"Refused: {len(refused)}", flush=True)
    for p in refused:
        print(f"  uuid={p['uuid'][:8]} sub_type={p['sub_type']} http={p['http']} body={p['body']}", flush=True)

# ── Count after ───────────────────────────────────────────────────────────────
# remaining = to_keep + plans we tried to clear but were refused
remaining_list = list(to_keep)
refused_uuids = {r["uuid"] for r in refused}
for p in ar_list:
    if p["uuid"] in refused_uuids and p not in remaining_list:
        remaining_list.append(p)

after = len(remaining_list)
expected_after = before - cleared
print(f"\nBefore: {before}, Cleared: {cleared}, Refused: {len(refused)}, After: {after} (expected {expected_after})", flush=True)
if after != expected_after:
    print(f"STOP: count mismatch after write", flush=True)
    sys.exit(1)

# Save updated AR list
with open(B1_AR_LIST, "w", encoding="utf-8") as f:
    json.dump(remaining_list, f, indent=2)
print(f"Saved updated b1_ar_list.json ({len(remaining_list)} entries)", flush=True)

# ── Refresh workbook Unknown Provider tab ─────────────────────────────────────
print(f"\nRefreshing workbook Unknown Provider tab...", flush=True)

# Fetch updated plan details for remaining plans
wb = openpyxl.load_workbook(WORKBOOK_PATH)
if "Unknown Provider" not in wb.sheetnames:
    print(f"  [WARN] 'Unknown Provider' sheet not found in workbook", flush=True)
else:
    ws = wb["Unknown Provider"]
    # Clear existing data rows (keep header)
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.value = None

    # Fetch plan details for remaining list
    headers_written = False
    row_idx = 2
    for plan_entry in remaining_list:
        r = dest.get(f"/api/v1/plans/{plan_entry['uuid']}")
        if r.status_code != 200:
            print(f"  [WARN] GET plan {plan_entry['uuid'][:8]} HTTP {r.status_code}", flush=True)
            continue
        body = r.json()
        plan = body.get("data") or body
        if not isinstance(plan, dict):
            continue
        prov_uuid = plan.get("provider_uuid") or plan.get("provider", {}).get("uuid") or ""
        ws.cell(row=row_idx, column=1).value = plan.get("uuid", "")
        ws.cell(row=row_idx, column=2).value = plan.get("sub_type", "")
        ws.cell(row=row_idx, column=3).value = prov_uuid
        ws.cell(row=row_idx, column=4).value = plan_entry.get("reason", "")
        row_idx += 1
        if row_idx % 50 == 2:
            print(f"  {row_idx - 2} rows written...", flush=True)

    # Trim blank rows at end
    max_row = ws.max_row
    for row in range(max_row, 1, -1):
        if ws.cell(row=row, column=1).value:
            break
        for cell in ws[row]:
            cell.value = None

    wb.save(WORKBOOK_PATH)
    print(f"  Unknown Provider tab: {row_idx - 2} rows remaining", flush=True)
    print(f"  Workbook saved.", flush=True)

print("\n=== DONE ===", flush=True)
