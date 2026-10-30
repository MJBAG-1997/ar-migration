"""
morning_c_check.py — Tasks 2, 3, 4-analysis:
  Part A: 17 failed files — UUID, client, ext, size, failure type, HTTP status
  Part B: duplicate check — dest "Previous Plannr" count vs src count per client
  Part C: date mismatch count — src original_created_at vs created_at
Rule 9: UUIDs only in output.
"""
import json, os, sys, time
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from azure.identity import DefaultAzureCredential
from azure.keyvault.secrets import SecretClient
import requests

BASE_URL          = "https://api.plannrcrm.com"
DEST_ACCOUNT_UUID = "ccd7f98b-d6a4-46ad-bab3-242ab42826eb"
SRC_ACCOUNT_UUID  = "db0ab7f5-8e52-4c89-a915-29fefeeca0d4"
IMPORT_DIR        = r"C:\dev\AR Migration\import_20261027"
STATE_DIR         = os.path.join(IMPORT_DIR, "state")

sc       = SecretClient(vault_url="https://kv-plannr-ingestion.vault.azure.net/",
                        credential=DefaultAzureCredential())
dest_pat = sc.get_secret("plannr-prod-pat").value
src_pat  = sc.get_secret("plannr-hfp-pat").value

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
        kw.setdefault("timeout", 60)
        for attempt in range(6):
            try:
                r = self._s.request("GET", url, **kw)
            except requests.exceptions.RequestException as e:
                if attempt < 5:
                    wait = 15 * (attempt + 1)
                    print(f"  [NET] GET {path[:60]} retry in {wait}s", flush=True)
                    time.sleep(wait)
                    continue
                raise
            if r.status_code != 429:
                return r
            wait = int(r.headers.get("Retry-After", 30))
            print(f"  [429] GET {path[:60]} — {wait}s", flush=True); time.sleep(wait)
        return r

def get_paged(api, path, params=None, max_pages=100):
    items, page = [], 1
    while page <= max_pages:
        pr = dict(params or {}); pr["page"] = page; pr["per_page"] = 200
        r = api.get(path, params=pr)
        if r.status_code != 200:
            print(f"  [WARN] GET {path} HTTP {r.status_code}", flush=True); break
        body = r.json()
        if isinstance(body, list): items.extend(body); break
        data = body.get("data", [])
        if isinstance(data, list):   items.extend(data)
        elif isinstance(data, dict): items.append(data)
        meta = body.get("meta", {}) if isinstance(body, dict) else {}
        if page >= (meta.get("last_page") or 1): break
        page += 1
    return items

def load_json(path):
    if not os.path.exists(path): return {}
    with open(path, encoding="utf-8") as f: return json.load(f)

def save_json(path, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)

dest = API(dest_pat, DEST_ACCOUNT_UUID)
src  = API(src_pat,  SRC_ACCOUNT_UUID)

client_map   = load_json(os.path.join(STATE_DIR, "client_map.json"))
file_state   = load_json(os.path.join(STATE_DIR, "file_state.json"))
folder_state = load_json(os.path.join(STATE_DIR, "folder_state.json"))
src_uuids    = list(client_map.keys())

# Known failures from taskC_run7.log
FAILED_PREFIXES = {
    "98cd40df": ("upload", 403),
    "8a54f8ba": ("upload", 403),
    "5a50fc29": ("folder", None),
    "beb8e1a6": ("upload", 403),
    "99c02522": ("upload", 403),
    "863857fb": ("upload", 403),
    "23f6bdc2": ("upload", 403),
    "7168b608": ("upload", 403),
    "32a87f1a": ("upload", 403),
    "99caacd6": ("upload", 403),
    "0a04d057": ("upload", 403),
    "7f5ca74a": ("upload", 403),
    "762504a1": ("download", 404),
    "8f8373a4": ("upload", 403),
    "8b076ff6": ("upload", 403),
    "22bd524d": ("download", 403),
    "e4fd07bf": ("download", 403),
}

# ─── Fetch all src account files (used by all three parts) ───────────────────
print(f"Fetching src files for {len(src_uuids)} clients...", flush=True)
src_file_map  = {}   # uuid → file obj
src_count_map = {}   # src_client_uuid → count

for i, su in enumerate(src_uuids, 1):
    files = get_paged(src, f"/api/v1/files/account/{su}")
    src_count_map[su] = len(files)
    for f in files:
        fuid = f.get("uuid", "")
        if fuid:
            f["_src_account"] = su   # track which account this file belongs to
            src_file_map[fuid] = f
    if i % 50 == 0:
        print(f"  {i}/{len(src_uuids)} clients — {len(src_file_map)} files", flush=True)

print(f"  Total src files fetched: {len(src_file_map)}", flush=True)


# ─── Part A: 17 failed files ─────────────────────────────────────────────────
print("\n=== PART A: FAILED FILES ===", flush=True)

failed_full = {}   # full_uuid → details
for full_uuid, f in src_file_map.items():
    pfx = full_uuid[:8]
    if pfx in FAILED_PREFIXES:
        failed_full[full_uuid] = f

# Check for any not matched
matched_pfx = {v[:8] for v in failed_full.keys()}
unmatched = set(FAILED_PREFIXES.keys()) - matched_pfx
if unmatched:
    print(f"  [WARN] {len(unmatched)} prefixes not found in src files: {unmatched}", flush=True)

for full_uuid, f in sorted(failed_full.items()):
    pfx = full_uuid[:8]
    fail_type, http_status = FAILED_PREFIXES[pfx]
    su   = f.get("_src_account", "unknown")
    du   = client_map.get(su, "unknown")
    name = f.get("name") or f.get("file_name") or ""
    ext  = os.path.splitext(name)[1].lstrip(".").lower() or "?"
    try: size = int(f.get("size") or f.get("file_size") or 0)
    except (ValueError, TypeError): size = 0
    status_str = str(http_status) if http_status else "n/a"
    print(f"  uuid={full_uuid} client_dest={du[:8]} ext={ext} size={size} "
          f"failure={fail_type} http={status_str}", flush=True)

print(f"  Total identified: {len(failed_full)}", flush=True)


# ─── Part B: duplicate check ─────────────────────────────────────────────────
print("\n=== PART B: DUPLICATE CHECK ===", flush=True)

def list_folder_files_dest(folder_uuid):
    r = dest.get(f"/api/v1/folder/{folder_uuid}/files")
    if r.status_code != 200: return []
    body  = r.json()
    items = body.get("data", body) if isinstance(body, dict) else body
    if not isinstance(items, list): items = [items] if isinstance(items, dict) else []
    return [fi for fi in items if isinstance(fi, dict) and fi.get("uuid")]

dup_clients = 0
checked = 0
for su, du in client_map.items():
    src_n = src_count_map.get(su, 0)
    if src_n == 0:
        continue
    pp_key = f"{du}|Previous Plannr"
    pp_uuid = folder_state.get(pp_key)
    if not pp_uuid:
        continue
    dest_files = list_folder_files_dest(pp_uuid)
    dest_n = len(dest_files)
    checked += 1
    if dest_n > src_n:
        dup_clients += 1
        # Sort dest files by created_at desc; the extras are the last (dest_n - src_n)
        dest_sorted = sorted(dest_files,
                             key=lambda x: x.get("created_at") or "", reverse=True)
        extra = dest_sorted[src_n:]
        extra_uuids = [x["uuid"] for x in extra]
        print(f"  DUPLICATE dest={du[:8]} src={src_n} dest_pp={dest_n} "
              f"extra={[u[:8] for u in extra_uuids]}", flush=True)

print(f"  Clients checked: {checked}, duplicates: {dup_clients}", flush=True)


# ─── Part C: date mismatch analysis ──────────────────────────────────────────
print("\n=== PART C: DATE MISMATCH ANALYSIS ===", flush=True)

mismatches = {}   # src_uuid → {dest_uuid, src_original_created_at, src_created_at}
for src_uuid, dest_uuid in file_state.items():
    f = src_file_map.get(src_uuid)
    if not f:
        continue
    orig = (f.get("original_created_at") or "").strip()
    crt  = (f.get("created_at") or "").strip()
    # Only flag if orig is non-empty and differs from created_at
    if orig and orig != crt:
        mismatches[src_uuid] = {
            "dest_uuid":              dest_uuid,
            "src_original_created_at": orig,
            "src_created_at":          crt,
        }

save_json(os.path.join(STATE_DIR, "date_mismatch.json"), mismatches)
print(f"  Date mismatches: {len(mismatches)} (written to state/date_mismatch.json)", flush=True)
if mismatches:
    sample = list(mismatches.items())[:3]
    for s, d in sample:
        print(f"    src={s[:8]} dest={d['dest_uuid'][:8]} "
              f"orig={d['src_original_created_at'][:10]} created={d['src_created_at'][:10]}", flush=True)

print("\n=== DONE ===", flush=True)
