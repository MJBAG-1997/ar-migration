"""
probe_dest_folders.py — find root folder UUID on dest accounts/plans
"""
import json, os, sys, time
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from azure.identity import DefaultAzureCredential
from azure.keyvault.secrets import SecretClient
import requests

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

BASE_URL          = "https://api.plannrcrm.com"
DEST_ACCOUNT_UUID = "ccd7f98b-d6a4-46ad-bab3-242ab42826eb"
STATE_DIR         = r"C:\dev\AR Migration\import_20261027\state"

sc = SecretClient(vault_url="https://kv-plannr-ingestion.vault.azure.net/",
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
        url = f"{BASE_URL}{path}" if path.startswith("/") else path
        for _ in range(3):
            r = self._s.request("GET", url, **kw)
            if r.status_code != 429: return r
            time.sleep(int(r.headers.get("Retry-After", 30)))
        return r
    def post(self, path, **kw):
        url = f"{BASE_URL}{path}" if path.startswith("/") else path
        for _ in range(3):
            r = self._s.request("POST", url, **kw)
            if r.status_code != 429: return r
            time.sleep(int(r.headers.get("Retry-After", 30)))
        return r

dest = API(dest_pat, DEST_ACCOUNT_UUID)

client_map = json.load(open(os.path.join(STATE_DIR, "client_map.json")))
plan_map   = json.load(open(os.path.join(STATE_DIR, "plan_map.json")))

# Try GET /api/v1/files/account/{dest_uuid} on first few dest accounts — find any with files
print("=== Probing dest account files for folder UUIDs ===")
dest_uuids = list(client_map.values())
for du in dest_uuids[:10]:
    r = dest.get(f"/api/v1/files/account/{du}", params={"per_page": 5, "include": "folder"})
    if r.status_code != 200:
        print(f"  {du[:8]}: HTTP {r.status_code}")
        continue
    body = r.json()
    data = body.get("data", body) if isinstance(body, dict) else body
    if isinstance(data, list) and data:
        f0 = data[0]
        folder = f0.get("folder") or {}
        print(f"  {du[:8]}: {len(data)} files — folder={json.dumps(folder)[:120]}")
        # Get this folder's parents to find root
        fuid = folder.get("uuid") if isinstance(folder, dict) else None
        if fuid:
            rg = dest.get(f"/api/v1/folder/{fuid}")
            if rg.status_code == 200:
                fg = (rg.json().get("data") or rg.json()) or {}
                print(f"    folder {fuid[:8]}: name={fg.get('name')!r} "
                      f"is_root={fg.get('is_root_folder')} "
                      f"doc_type={fg.get('documentable_type')} "
                      f"doc_uuid={str(fg.get('documentable_uuid',''))[:8]}")
                parents = fg.get("parents") or fg.get("ancestors") or []
                print(f"    parents: {[p.get('uuid','')[:8]+':'+p.get('name','') for p in parents]}")
        break
    else:
        print(f"  {du[:8]}: 0 files")

# Also try a plan with files
print()
print("=== Probing dest plan files ===")
dest_plan_uuids = list(plan_map.values())
for dp in dest_plan_uuids[:20]:
    r = dest.get(f"/api/v1/files/plan/{dp}", params={"per_page": 5})
    if r.status_code != 200:
        continue
    body = r.json()
    data = body.get("data", body) if isinstance(body, dict) else body
    if isinstance(data, list) and data:
        f0 = data[0]
        folder = f0.get("folder") or {}
        fuid = folder.get("uuid") if isinstance(folder, dict) else None
        print(f"  plan {dp[:8]}: {len(data)} files — folder_uuid={str(fuid or '')[:8]}")
        if fuid:
            rg = dest.get(f"/api/v1/folder/{fuid}")
            if rg.status_code == 200:
                fg = (rg.json().get("data") or rg.json()) or {}
                print(f"    name={fg.get('name')!r} is_root={fg.get('is_root_folder')} "
                      f"doc_type={fg.get('documentable_type')}")
        break

# Try creating a test folder on a dest account using spec-correct fields
print()
print("=== Test: POST /api/v1/folder with spec fields ===")
test_du = dest_uuids[0]
# First GET existing files to find a folder UUID to use as parent
r2 = dest.get(f"/api/v1/files/account/{test_du}", params={"per_page": 5})
if r2.status_code == 200:
    data2 = (r2.json().get("data") or r2.json()) or []
    if isinstance(data2, list) and data2:
        f0 = data2[0]
        folder_info = f0.get("folder") or {}
        root_uuid = folder_info.get("uuid") if isinstance(folder_info, dict) else None
        if root_uuid:
            # Navigate to root
            rg2 = dest.get(f"/api/v1/folder/{root_uuid}")
            if rg2.status_code == 200:
                fg2 = (rg2.json().get("data") or rg2.json()) or {}
                parents = fg2.get("parents") or fg2.get("ancestors") or []
                if parents:
                    root_uuid = parents[0].get("uuid") or root_uuid
                elif fg2.get("is_root_folder"):
                    root_uuid = fg2.get("uuid")
            print(f"  Using parent={root_uuid[:8]} to create test folder")
            rp = dest.post("/api/v1/folder", json={"name": "PROBE_DELETE_ME", "parent": root_uuid})
            print(f"  POST /api/v1/folder (name+parent): HTTP {rp.status_code} {rp.text[:200]}")
        else:
            print("  No folder UUID from existing files")
    else:
        print(f"  Account {test_du[:8]} has no files on dest")
        # Try without parent
        rp = dest.post("/api/v1/folder", json={"name": "PROBE_DELETE_ME"})
        print(f"  POST /api/v1/folder (name only): HTTP {rp.status_code} {rp.text[:200]}")
else:
    print(f"  GET files HTTP {r2.status_code}")
