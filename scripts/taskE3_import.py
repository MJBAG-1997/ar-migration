"""
taskE3_import.py — import 8 task notes using corrected multipart upload
Uses already-written import_20261027/E3_task_notes.csv
"""
import io, json, os, sys, time
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

from azure.identity import DefaultAzureCredential
from azure.keyvault.secrets import SecretClient
import requests

BASE_URL          = "https://api.plannrcrm.com"
DEST_ACCOUNT_UUID = "ccd7f98b-d6a4-46ad-bab3-242ab42826eb"
IMPORT_DIR        = r"C:\dev\AR Migration\import_20261027"
STATE_DIR         = os.path.join(IMPORT_DIR, "state")

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
        return self._req("GET", path, **kw)

    def post(self, path, **kw):
        return self._req("POST", path, **kw)

    def _req(self, method, path, **kw):
        url = f"{BASE_URL}{path}" if path.startswith("/") else path
        for _ in range(4):
            r = self._s.request(method, url, **kw)
            if r.status_code != 429:
                return r
            wait = int(r.headers.get("Retry-After", 30))
            print(f"  [429] {method} {path} — {wait}s", flush=True)
            time.sleep(wait)
        return r

    def upload(self, csv_bytes, filename, type_str, mode):
        # Suppress session Content-Type so requests sets multipart boundary
        return self._s.post(
            f"{BASE_URL}/api/v1/import",
            headers={"Content-Type": None},
            files={"file": (filename, io.BytesIO(csv_bytes), "text/csv")},
            data={"type": type_str, "mode": mode},
        )


def p(*a):
    print(*a, flush=True)


def stop(msg):
    p(f"\n!!STOP!! {msg}")
    sys.exit(1)


def get_paged(api, path, max_pages=200, params=None):
    items, page = [], 1
    while page <= max_pages:
        pr = dict(params or {})
        pr["page"] = page
        pr["per_page"] = 200
        r = api.get(path, params=pr)
        if r.status_code != 200:
            p(f"  [WARN] {path} → HTTP {r.status_code}")
            break
        body = r.json()
        if isinstance(body, list):
            items.extend(body)
            break
        data = body.get("data", [])
        if isinstance(data, list):
            items.extend(data)
        elif isinstance(data, dict):
            items.append(data)
        meta = body.get("meta", {}) if isinstance(body, dict) else {}
        if page >= (meta.get("last_page") or 1):
            break
        page += 1
    return items


dest = API(dest_pat, DEST_ACCOUNT_UUID)

csv_path = os.path.join(IMPORT_DIR, "E3_task_notes.csv")
if not os.path.exists(csv_path):
    stop(f"E3_task_notes.csv not found at {csv_path}")

with open(csv_path, "rb") as f:
    csv_bytes = f.read()

# Count rows
row_count = csv_bytes.decode("utf-8", errors="replace").count("\n") - 1
p(f"=== E3: TASK NOTES IMPORT ===")
p(f"  CSV rows: {row_count}")

# Count notes before
notes_before = get_paged(dest, "/api/v1/notes")
p(f"  Notes on dest before: {len(notes_before)}")

# Upload
p("  Uploading E3_task_notes.csv…")
r = dest.upload(csv_bytes, "E3_task_notes.csv", "note", "validate")
p(f"  Upload HTTP {r.status_code}")
if r.status_code not in (200, 201):
    stop(f"E3 upload failed: {r.text[:300]}")

body = r.json()
imp_uuid = (body.get("data") or body or {}).get("uuid", "")
p(f"  Import UUID: {imp_uuid}")
if not imp_uuid:
    stop(f"No import UUID returned: {body}")

# Map headers
p("  Mapping headers…")
map_r = dest.post(f"/api/v1/import/{imp_uuid}/map")
p(f"  Map HTTP {map_r.status_code}")

# Start import
p("  Starting import…")
start_r = dest.post(f"/api/v1/import/{imp_uuid}/start")
p(f"  Start HTTP {start_r.status_code}")
if start_r.status_code not in (200, 201):
    stop(f"Start failed: {start_r.text[:200]}")

# Poll until complete (max 5 min)
p("  Polling…")
deadline = time.time() + 300
last_progress = -1
while time.time() < deadline:
    poll_r = dest.get(f"/api/v1/import/{imp_uuid}")
    if poll_r.status_code != 200:
        p(f"  [WARN] poll HTTP {poll_r.status_code}")
        time.sleep(10)
        continue
    pb = (poll_r.json().get("data") or poll_r.json()) or {}
    status   = pb.get("status", "")
    progress = pb.get("progress", 0)
    total    = pb.get("total_jobs", 0)
    failed   = pb.get("failed_jobs", 0)
    if progress != last_progress:
        p(f"  status={status} progress={progress}/{total} failed={failed}")
        last_progress = progress
    if status in ("complete", "completed", "failed", "error"):
        break
    time.sleep(5)
else:
    stop(f"Import timed out after 5 min. Last status={status} progress={progress}/{total}")

if failed:
    p(f"  [WARN] {failed} failed jobs")

# Count after
notes_after = get_paged(dest, "/api/v1/notes")
delta = len(notes_after) - len(notes_before)
p(f"  Notes after: {len(notes_after)}  (+{delta} from {len(notes_before)})")
if delta < row_count:
    p(f"  [WARN] expected +{row_count}, got +{delta}")
else:
    p(f"  ✓ E3 task notes imported: {delta} notes added")

p("=== E3 DONE ===")
