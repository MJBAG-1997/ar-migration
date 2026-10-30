"""
taskC_20261029.py — Task C: copy account files to dest (accounts only).
Case/circle files are already in the client folder tree on Plannr; skip them.
Rule 9: counts and UUIDs only in output.
One script talks to Plannr at a time.
"""
import io, json, os, sys, time, datetime
from collections import defaultdict

# Write directly to log file so no shell redirect is needed
_LOG_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                         "taskC_20261029.log")
_log_fh = open(_LOG_PATH, "w", encoding="utf-8", buffering=1)

class _Tee:
    def __init__(self, *streams): self._s = streams
    def write(self, data):
        for s in self._s:
            try: s.write(data)
            except Exception: pass
    def flush(self):
        for s in self._s:
            try: s.flush()
            except Exception: pass
    def reconfigure(self, **kw): pass

sys.stdout = _Tee(sys.stdout, _log_fh)
sys.stderr = _Tee(sys.stderr, _log_fh)

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from azure.identity import DefaultAzureCredential
from azure.keyvault.secrets import SecretClient
import requests

BASE_URL          = "https://api.plannrcrm.com"
DEST_ACCOUNT_UUID = "ccd7f98b-d6a4-46ad-bab3-242ab42826eb"
SRC_ACCOUNT_UUID  = "db0ab7f5-8e52-4c89-a915-29fefeeca0d4"
IMPORT_DIR        = r"C:\dev\AR Migration\import_20261027"
STATE_DIR         = os.path.join(IMPORT_DIR, "state")
FILE_STATE_PATH   = os.path.join(STATE_DIR, "file_state.json")

MAX_BYTES       = 20 * 1024 ** 3
MAX_CONSEC_FAIL = 5
PREV_FOLDER     = "Previous Plannr"


def p(*a): print(*a, flush=True)
def stop(msg): p(f"\n!!STOP!! {msg}"); sys.exit(1)

def load_json(path):
    if not os.path.exists(path): return {}
    with open(path, encoding="utf-8") as f: return json.load(f)

def save_json(path, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)


# ── API ───────────────────────────────────────────────────────────────────────
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
    def get(self, path, stream=False, **kw):
        return self._req("GET", path, stream=stream, **kw)
    def post(self, path, **kw):
        return self._req("POST", path, **kw)
    def _req(self, method, path, **kw):
        url   = f"{BASE_URL}{path}" if path.startswith("/") else path
        is_get = (method == "GET")
        kw.setdefault("timeout", 60 if is_get else 600)
        # GETs: retry up to 6× on network errors. POSTs: one attempt only.
        max_attempts = 6 if is_get else 1
        for attempt in range(max_attempts):
            try:
                r = self._s.request(method, url, **kw)
            except requests.exceptions.RequestException as e:
                if is_get and attempt < 5:
                    wait = 15 * (attempt + 1)
                    p(f"  [NET] GET {path[:60]} — {e.__class__.__name__} retry in {wait}s")
                    time.sleep(wait)
                    continue
                raise  # propagate: POSTs and exhausted GETs
            if r.status_code != 429: return r
            wait = int(r.headers.get("Retry-After", 30))
            p(f"  [429] {method} {path[:60]} — {wait}s"); time.sleep(wait)
        return r

def get_paged(api, path, params=None, max_pages=200):
    items, page = [], 1
    while page <= max_pages:
        pr = dict(params or {}); pr["page"] = page; pr["per_page"] = 200
        r = api.get(path, params=pr)
        if r.status_code != 200:
            p(f"  [WARN] GET {path} HTTP {r.status_code}"); break
        body = r.json()
        if isinstance(body, list): items.extend(body); break
        data = body.get("data", [])
        if isinstance(data, list):   items.extend(data)
        elif isinstance(data, dict): items.append(data)
        meta = body.get("meta", {}) if isinstance(body, dict) else {}
        if page >= (meta.get("last_page") or 1): break
        page += 1
    return items

dest = API(dest_pat, DEST_ACCOUNT_UUID)
src  = API(src_pat,  SRC_ACCOUNT_UUID)


def list_folder_files(folder_uuid):
    """Return list of {uuid, name, size} for files in a dest folder."""
    r = dest.get(f"/api/v1/folder/{folder_uuid}/files")
    if r.status_code != 200: return []
    body  = r.json()
    items = body.get("data", body) if isinstance(body, dict) else body
    if not isinstance(items, list): items = [items] if isinstance(items, dict) else []
    out = []
    for fi in items:
        if not isinstance(fi, dict): continue
        try: sz = int(fi.get("size") or 0)
        except (ValueError, TypeError): sz = 0
        out.append({"uuid": fi.get("uuid",""), "name": fi.get("name",""), "size": sz})
    return out


def find_subfolder(parent_uuid, name):
    """Return UUID of an immediate child folder named `name`, or None."""
    r = dest.get(f"/api/v1/folder/{parent_uuid}/tree")
    if r.status_code != 200: return None
    body  = r.json()
    items = body.get("data", body) if isinstance(body, dict) else body
    if isinstance(items, dict): items = [items]
    if not isinstance(items, list): return None
    for item in items:
        if isinstance(item, dict) and item.get("name") == name:
            return item.get("uuid") or None
    return None


client_map = load_json(os.path.join(STATE_DIR, "client_map.json"))
src_uuids  = list(client_map.keys())


# ─────────────────────────────────────────────────────────────────────────────
# C1: INVENTORY — account files only
# ─────────────────────────────────────────────────────────────────────────────
p("=== C1: INVENTORY (account files only) ===")

def file_size(f):
    sz = f.get("size") or f.get("file_size") or 0
    try: return int(sz)
    except (ValueError, TypeError): return 0

def file_folder_path(f):
    folder = f.get("folder") or {}
    if isinstance(folder, dict):
        return folder.get("path") or folder.get("full_path") or folder.get("name") or ""
    return str(folder) if folder else ""

file_inventory = {}
p(f"  Fetching account files for {len(src_uuids)} clients…")
for i, su in enumerate(src_uuids, 1):
    files = get_paged(src, f"/api/v1/files/account/{su}", max_pages=100)
    for f in files:
        fuid = f.get("uuid", "")
        if fuid and fuid not in file_inventory:
            file_inventory[fuid] = {
                "uuid":         fuid,
                "name":         f.get("name") or f.get("file_name") or "unknown",
                "size":         file_size(f),
                "src_account":  su,
                "dest_account": client_map.get(su, ""),
                "folder_path":  file_folder_path(f),
                "created_at":   f.get("created_at") or f.get("original_created_at") or "",
                "download_url": f.get("url") or f.get("download_url") or "",
            }
    if i % 50 == 0:
        p(f"  {i}/{len(src_uuids)} clients — {len(file_inventory)} files")

total_bytes = sum(fi["size"] for fi in file_inventory.values())
total_gb    = total_bytes / 1024 ** 3
p(f"\n  INVENTORY TOTALS:")
p(f"    Total files: {len(file_inventory)}")
p(f"    Total size:  {total_bytes:,} bytes = {total_gb:.2f} GB")
if total_gb > 20:
    stop(f"Total {total_gb:.2f} GB exceeds 20 GB limit")
p(f"  Under 20 GB — proceeding")

# Sort newest first
sorted_items   = sorted(file_inventory.items(),
                        key=lambda x: x[1].get("created_at") or "",
                        reverse=True)
file_inventory = dict(sorted_items)
newest = next(iter(file_inventory.values())).get("created_at", "?")[:10] if file_inventory else "?"
p(f"  Sorted newest-first (newest: {newest})")

est_secs = len(file_inventory) * 3
eta_dt   = datetime.datetime.now() + datetime.timedelta(seconds=est_secs)
p(f"  Estimated finish: {eta_dt.strftime('%H:%M')} (~{est_secs//60} min at 3 s/file)")

save_json(os.path.join(STATE_DIR, "file_inventory.json"),
          {"total_files": len(file_inventory), "total_bytes": total_bytes,
           "total_gb": round(total_gb, 3)})


# ─────────────────────────────────────────────────────────────────────────────
# FOLDER HELPERS
# ─────────────────────────────────────────────────────────────────────────────
folder_state = load_json(os.path.join(STATE_DIR, "folder_state.json"))

def _fkey(acct_uuid, path):
    return f"{acct_uuid}|{path}"

def _root_folder_uuid(acct_uuid):
    key = _fkey(acct_uuid, "")
    if key in folder_state:
        return folder_state[key]
    r = dest.get(f"/api/v1/account/{acct_uuid}/documents")
    if r.status_code != 200:
        p(f"  [WARN] GET root folder {acct_uuid[:8]}: HTTP {r.status_code} {r.text[:80]}")
        return None
    body = r.json()
    root = body.get("data") or body
    if isinstance(root, list): root = root[0] if root else {}
    ruid = (root.get("uuid", "") if isinstance(root, dict) else "")
    if not ruid:
        p(f"  [WARN] root folder {acct_uuid[:8]}: no UUID in response")
        return None
    folder_state[key] = ruid
    save_json(os.path.join(STATE_DIR, "folder_state.json"), folder_state)
    return ruid

def get_or_create_folder(acct_uuid, path):
    """Return dest folder UUID for path under acct_uuid, creating it if needed.
    Permissions clients_read/write set to False; GET confirms before returning.
    """
    key = _fkey(acct_uuid, path)
    if key in folder_state:
        return folder_state[key]

    root = _root_folder_uuid(acct_uuid)
    if not root:
        return None

    parts         = [pt.strip() for pt in path.replace("\\", "/").split("/") if pt.strip()]
    current_parent = root
    built          = ""
    for part in parts:
        built       = (built + "/" + part).lstrip("/")
        partial_key = _fkey(acct_uuid, built)
        if partial_key in folder_state:
            current_parent = folder_state[partial_key]
            continue
        try:
            r = dest.post("/api/v1/folder", json={"name": part, "parent": current_parent})
            folder_ok = r.status_code in (200, 201)
        except requests.exceptions.RequestException as e:
            p(f"  [NET] create folder {part!r}: {e.__class__.__name__} — checking parent")
            folder_ok = False
            r = None
        if not folder_ok:
            if r is not None:
                p(f"  [WARN] create folder {part!r}: HTTP {r.status_code} {r.text[:80]}")
            # Check if it was created anyway
            existing = find_subfolder(current_parent, part)
            if existing:
                p(f"  Folder {existing[:8]} ({part!r}) found in parent — using it")
                folder_uuid = existing
            else:
                # Retry once
                try:
                    r2 = dest.post("/api/v1/folder",
                                   json={"name": part, "parent": current_parent})
                    if r2.status_code not in (200, 201):
                        p(f"  [WARN] folder retry {part!r}: HTTP {r2.status_code} {r2.text[:80]}")
                        return None
                    folder_uuid = ((r2.json().get("data") or r2.json()) or {}).get("uuid","")
                except requests.exceptions.RequestException as e2:
                    p(f"  [WARN] folder retry {part!r}: {e2.__class__.__name__}")
                    return None
            if not folder_uuid:
                p(f"  [WARN] create folder {part!r}: no UUID")
                return None
        else:
            folder_uuid = ((r.json().get("data") or r.json()) or {}).get("uuid", "")
            if not folder_uuid:
                p(f"  [WARN] create folder {part!r}: no UUID in response {r.text[:80]}")
                return None
        # Set permissions
        rp = dest.post(f"/api/v1/folder/{folder_uuid}/permissions",
                       json={"clients_read": False, "clients_write": False})
        if rp.status_code not in (200, 201):
            p(f"  [WARN] set perms on folder {folder_uuid[:8]}: HTTP {rp.status_code}")
        # GET to confirm before use
        rg = dest.get(f"/api/v1/folder/{folder_uuid}")
        if rg.status_code == 200:
            fg = (rg.json().get("data") or rg.json()) or {}
            p(f"  Folder {folder_uuid[:8]} ({part!r}) confirmed: "
              f"clients_read={fg.get('clients_read','?')} "
              f"clients_write={fg.get('clients_write','?')}")
        else:
            p(f"  [WARN] GET confirm folder {folder_uuid[:8]}: HTTP {rg.status_code}")
        folder_state[partial_key] = folder_uuid
        save_json(os.path.join(STATE_DIR, "folder_state.json"), folder_state)
        current_parent = folder_uuid

    return folder_state.get(key)


# ─────────────────────────────────────────────────────────────────────────────
# C3: COPY FILES
# ─────────────────────────────────────────────────────────────────────────────
p("\n=== C3: COPY FILES ===")

file_state = load_json(FILE_STATE_PATH)
p(f"  State: {len(file_state)} files already copied")

def upload_one(fi):
    src_fuid = fi["uuid"]
    fname    = fi.get("name", "file.dat")
    created  = fi.get("created_at", "")
    du       = fi.get("dest_account", "")

    if not du:
        p(f"  [WARN] {src_fuid[:8]}: no dest account")
        return None

    # Download via direct API endpoint — no pre-signed URL expiry issues
    try:
        dl_r = src.get(f"/api/v1/file/{src_fuid}/download", stream=True)
        if dl_r.status_code != 200:
            p(f"  [WARN] download {src_fuid[:8]}: HTTP {dl_r.status_code}")
            return None
        file_bytes = dl_r.content
    except Exception as e:
        p(f"  [WARN] download {src_fuid[:8]}: {e}")
        return None

    # Determine dest folder
    src_folder  = fi.get("folder_path") or ""
    dest_path   = PREV_FOLDER + ("/" + src_folder.lstrip("/") if src_folder else "")
    folder_uuid = get_or_create_folder(du, dest_path)
    if not folder_uuid:
        p(f"  [WARN] {src_fuid[:8]}: could not get/create folder")
        return None

    # Upload — POST with 600s timeout (set via _req); no auto-retry in _req
    upload_data = {"name": fname, "folder_uuid": folder_uuid}
    if created:
        upload_data["original_created_at"] = created

    def _do_upload():
        return dest.post("/api/v1/file/direct-upload",
                         data=upload_data,
                         files={"file": (fname, io.BytesIO(file_bytes))})

    def _check_existing():
        """Look for a file with the same name and size already in the folder."""
        for fi in list_folder_files(folder_uuid):
            if fi["name"] == fname and fi["size"] == len(file_bytes):
                return fi["uuid"]
        return None

    try:
        up_r = _do_upload()
        net_err = None
    except requests.exceptions.RequestException as e:
        net_err = e
        up_r = None

    if net_err or (up_r is not None and up_r.status_code not in (200, 201)):
        if net_err:
            p(f"  [NET] upload {src_fuid[:8]}: {net_err.__class__.__name__} — checking folder")
        else:
            p(f"  [WARN] upload {src_fuid[:8]}: HTTP {up_r.status_code} — checking folder")
        existing = _check_existing()
        if existing:
            p(f"  Upload confirmed via folder check: dest={existing[:8]}")
            return existing
        # Not found — retry once
        p(f"  Not found in folder — retrying upload once")
        try:
            up_r = _do_upload()
        except requests.exceptions.RequestException as e2:
            p(f"  [WARN] upload retry {src_fuid[:8]}: {e2.__class__.__name__}")
            return None
        if up_r.status_code not in (200, 201):
            p(f"  [WARN] upload retry {src_fuid[:8]}: HTTP {up_r.status_code} {up_r.text[:80]}")
            return None

    dest_fuid = ((up_r.json().get("data") or up_r.json()) or {}).get("uuid", "")
    return dest_fuid or None


# ── One-client test first ──────────────────────────────────────────────────
test_su    = src_uuids[0]
test_du    = client_map.get(test_su, "")
test_files = [fi for fi in file_inventory.values() if fi["src_account"] == test_su]
p(f"\n  Test client ({test_su[:8]}…): {len(test_files)} files")

test_copied = 0
for fi in test_files:
    fuid = fi["uuid"]
    if fuid in file_state:
        test_copied += 1; continue
    dfuid = upload_one(fi)
    if dfuid:
        file_state[fuid] = dfuid
        save_json(FILE_STATE_PATH, file_state)
        test_copied += 1
    else:
        p(f"  [WARN] test file {fuid[:8]} failed")

# Verify on dest
if test_du:
    dest_files = get_paged(dest, f"/api/v1/files/account/{test_du}", max_pages=10)
    p(f"  Test verify: src={len(test_files)} files  dest={len(dest_files)} files  copied={test_copied}")
    if len(test_files) > 0 and test_copied == 0:
        stop("Test client upload: 0 files copied — check warnings above")
p("  Test client done")

# ── Copy remaining ─────────────────────────────────────────────────────────
remaining = [(fuid, fi) for fuid, fi in file_inventory.items() if fuid not in file_state]
p(f"\n  Copying {len(remaining)} remaining files…")
consec_fail   = 0
copied_this_run = 0

for i, (fuid, fi) in enumerate(remaining, 1):
    dfuid = upload_one(fi)
    if dfuid:
        file_state[fuid] = dfuid
        save_json(FILE_STATE_PATH, file_state)
        copied_this_run += 1
        consec_fail = 0
        if copied_this_run % 50 == 0:
            p(f"  {copied_this_run} copied ({len(file_state)}/{len(file_inventory)} total)")
    else:
        consec_fail += 1
        p(f"  [FAIL #{consec_fail}] {fuid[:8]}")
        if consec_fail >= MAX_CONSEC_FAIL:
            stop(f"{MAX_CONSEC_FAIL} consecutive failures — state saved: "
                 f"{len(file_state)}/{len(file_inventory)}")

p(f"\n  Done: {len(file_state)}/{len(file_inventory)} files in state")
if len(file_state) < len(file_inventory):
    p(f"  [WARN] {len(file_inventory) - len(file_state)} files missing — check warnings")
else:
    p("  All files copied")
p("\n=== TASK C COMPLETE ===")
