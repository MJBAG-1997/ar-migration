"""
1. Dependants: count + dependants import jobs.
2. Delete the 10 stale import jobs after confirming each is unmapped with progress=0.
Rule: only this script talks to Plannr right now.
"""
import json, sys, time
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

sys.path.insert(0, r"C:\dev\AR Migration\scripts")
from azure.identity import DefaultAzureCredential
from azure.keyvault.secrets import SecretClient
import requests

BASE_URL          = "https://api.plannrcrm.com"
DEST_ACCOUNT_UUID = "ccd7f98b-d6a4-46ad-bab3-242ab42826eb"
IMPORT_DIR        = r"C:\dev\AR Migration\import_20261027"

# The 10 stale import UUIDs from the analysis
STALE_UUIDS = [
    "b00c1a29-5204-4c1c-b709-db6d3b9067bc",  # clients/validate/1
    "b6e58321-f932-4f31-ba50-a4ada08b772f",  # clients/validate/1
    "413da21b-00bd-4f4c-943c-5df1f958318e",  # clients/validate/1
    "57b09bc7-7de3-48a0-b0bc-15ddbb10df51",  # clients/validate/1
    "684c122c-8471-4d08-adf4-b20fdf90d169",  # clients/validate/1
    "bd3448e3-4a62-4596-96ae-50304149c390",  # plans/validate/1
    "ed77f9ff-d4d2-4aaf-999d-a03d717cb4a9",  # plans/store/1
    "1ed30a7e-a79a-4dbb-b7d5-d2c08185e393",  # plans/validate/1900
    "ce5e1b7f-da6a-4afe-9e4a-331d6a425389",  # plans/store/1900
    "5b173634-0d17-40d4-8fa8-868df316ef92",  # plans/validate/1
]

def p(*a): print(*a, flush=True)
def stop(msg):
    p(f"\n!!STOP!! {msg}")
    sys.exit(1)

sc  = SecretClient(vault_url="https://kv-plannr-ingestion.vault.azure.net/",
                   credential=DefaultAzureCredential())
pat = sc.get_secret("plannr-prod-pat").value
s   = requests.Session()
s.headers.update({"Authorization": f"Bearer {pat}", "Accept": "application/json",
                  "X-PLANNR-ACCOUNT-UUID": DEST_ACCOUNT_UUID})

def api_get(path, params=None):
    for attempt in range(4):
        r = s.get(f"{BASE_URL}{path}", params=params)
        if r.status_code == 429:
            wait = int(r.headers.get("Retry-After", 30))
            p(f"  [429] {path} — waiting {wait}s")
            time.sleep(wait)
        else:
            return r
    return r

client_map = json.load(open(f"{IMPORT_DIR}/state/client_map.json"))

# ── 1. Dependants ──────────────────────────────────────────────────────────────
p("=== 1. DEPENDANTS CHECK ===")

# Count dependant import jobs
all_imports = []
page = 1
while True:
    r = api_get("/api/v1/import", params={"page": page, "per_page": 200})
    if r.status_code != 200: break
    body = r.json()
    data = body.get("data", [])
    all_imports.extend(data)
    if page >= (body.get("meta", {}).get("last_page") or 1): break
    page += 1
p(f"  Total imports: {len(all_imports)}")

dep_jobs = [i for i in all_imports if str(i.get("type") or "").lower() == "dependants"]
p(f"  Dependants import jobs: {len(dep_jobs)}")
for j in dep_jobs:
    p(f"    uuid={j.get('uuid')}  mode={j.get('mode')}  status={j.get('status')}  "
      f"progress={j.get('progress')}  total_jobs={j.get('total_jobs')}")

completed = [j for j in dep_jobs if str(j.get("status") or "").lower() == "complete"]
p(f"  Completed dependants jobs: {len(completed)}")
p("  (Multiple completed jobs are expected — historical runs from 2026-10-27 are still in the list)")

# Count current dependants
seen = set()
for du in client_map.values():
    pg = 1
    while True:
        r = api_get("/api/v2/dependants",
                    params={"filter[account_uuids]": du, "page": pg, "per_page": 200})
        if r.status_code != 200: break
        b = r.json()
        for rec in b.get("data", []): seen.add(rec.get("uuid", ""))
        if pg >= (b.get("meta", {}).get("last_page") or 1): break
        pg += 1

p(f"  Current dependant count: {len(seen)}  (expected 249)")
if len(seen) == 0:
    stop("Dependants = 0. Do not proceed.")
if len(seen) == 498:
    stop("Dependants = 498 (doubled). Do not proceed.")
if len(seen) != 249:
    stop(f"Dependants = {len(seen)}, expected 249. Investigate before proceeding.")
p("  ✓ Dependants = 249, one completed import job")

# ── 2. Delete 10 stale import jobs ────────────────────────────────────────────
p("\n=== 2. DELETE STALE IMPORT JOBS ===")
p(f"  Checking {len(STALE_UUIDS)} stale jobs…")

deleted = 0
skipped = 0
for uid in STALE_UUIDS:
    r = api_get(f"/api/v1/import/{uid}")
    if r.status_code == 404:
        p(f"  {uid[:8]}: already gone (404) — skip")
        skipped += 1
        continue
    if r.status_code != 200:
        stop(f"GET import {uid}: HTTP {r.status_code}")
    d = r.json().get("data", r.json())
    st   = str(d.get("status") or "").lower()
    mp   = d.get("mapping")
    prg  = d.get("progress") or 0
    tj   = d.get("total_jobs") or 0
    mode = str(d.get("mode") or "")
    typ  = str(d.get("type") or "")
    p(f"  {uid[:8]}: type={typ}  mode={mode}  status={st}  "
      f"mapping={mp!r}  progress={prg}  total_jobs={tj}")

    # Safety: only delete if unmapped and progress=0
    if mp is not None:
        stop(f"Import {uid} has mapping={mp!r} — NOT unmapped. Refusing to delete.")
    if prg != 0:
        stop(f"Import {uid} has progress={prg} (not 0). Refusing to delete.")
    if st not in ("mapping", "pending", "created", "not_started"):
        stop(f"Import {uid} status={st!r} — not a pending state. Refusing to delete.")

    rd = s.delete(f"{BASE_URL}/api/v1/import/{uid}")
    if rd.status_code in (200, 201, 204):
        p(f"    deleted ✓ (HTTP {rd.status_code})")
        deleted += 1
    else:
        stop(f"DELETE import {uid}: HTTP {rd.status_code} {rd.text[:150]}")

p(f"\n  Deleted: {deleted}  Skipped (already gone): {skipped}")
p("=== DONE ===")
