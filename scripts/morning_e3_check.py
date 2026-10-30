"""
morning_e3_check.py — Task 1: verify 8 task notes landed on dest clients.
Rule 9: UUIDs only, no client names or note contents in output.
"""
import csv, json, os, sys, time
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from azure.identity import DefaultAzureCredential
from azure.keyvault.secrets import SecretClient
import requests

BASE_URL          = "https://api.plannrcrm.com"
DEST_ACCOUNT_UUID = "ccd7f98b-d6a4-46ad-bab3-242ab42826eb"
IMPORT_DIR        = r"C:\dev\AR Migration\import_20261027"
E3_CSV            = os.path.join(IMPORT_DIR, "E3_task_notes.csv")

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
        url = f"{BASE_URL}{path}" if path.startswith("/") else path
        kw.setdefault("timeout", 60)
        for attempt in range(6):
            try:
                r = self._s.request("GET", url, **kw)
            except requests.exceptions.RequestException as e:
                if attempt < 5:
                    wait = 15 * (attempt + 1)
                    print(f"  [NET] GET {path[:60]} retry in {wait}s ({e.__class__.__name__})", flush=True)
                    time.sleep(wait)
                    continue
                raise
            if r.status_code != 429:
                return r
            wait = int(r.headers.get("Retry-After", 30))
            print(f"  [429] GET {path[:60]} — {wait}s", flush=True); time.sleep(wait)
        return r

def get_paged(api, path, params=None, max_pages=50):
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

dest = API(dest_pat, DEST_ACCOUNT_UUID)

# Load 8 rows from E3 CSV
rows = []
with open(E3_CSV, encoding="utf-8-sig") as f:
    for row in csv.DictReader(f):
        rows.append(row)
print(f"Loaded {len(rows)} rows from E3_task_notes.csv", flush=True)

# Group by account_uuid — 3 clients
from collections import defaultdict
by_acct = defaultdict(list)
for row in rows:
    by_acct[row["account_uuid"].strip()].append(row)

print(f"Distinct dest accounts: {len(by_acct)}", flush=True)

found = 0
missing = 0

for acct_uuid, acct_rows in by_acct.items():
    notes = get_paged(dest, f"/api/v1/client/{acct_uuid}/all-notes")
    print(f"  account={acct_uuid[:8]} — {len(notes)} notes on dest", flush=True)

    # Index dest notes by external_reference_value
    dest_by_extref = {}
    for n in notes:
        erv = (n.get("external_reference_value") or "").strip()
        if erv:
            dest_by_extref[erv] = n

    # Build date → list of dest notes index for this account
    from collections import defaultdict as _dd
    notes_by_date = _dd(list)
    for n in notes:
        nd = (n.get("created_at") or "")[:10]
        if nd:
            notes_by_date[nd].append(n)

    used_uuids = {}  # ext_ref → note_uuid matched

    for row in acct_rows:
        ext_ref = (row.get("external_reference_value") or "").strip()
        created = (row.get("created_at") or "").strip()

        if ext_ref in dest_by_extref:
            note = dest_by_extref[ext_ref]
            used_uuids[ext_ref] = note.get("uuid","?")
            print(f"    ext_ref={ext_ref[:8]} created={created} → FOUND (ext_ref match) note_uuid={note.get('uuid','?')[:8]}", flush=True)
            found += 1
        else:
            # Fallback: match by created_at date — convert DD/MM/YYYY → YYYY-MM-DD if needed
            created_date = created[:10] if created else ""
            if len(created_date) == 10 and created_date[2] == "/":
                parts = created_date.split("/")
                if len(parts) == 3:
                    created_date = f"{parts[2]}-{parts[1]}-{parts[0]}"
            day_notes = notes_by_date.get(created_date, [])
            # Check if enough notes exist for this date to cover all ext_refs with this date
            same_date_expected = sum(
                1 for r in acct_rows
                if (r.get("created_at") or "")[:10].replace("/", "-") == created_date or
                   ("{}-{}-{}".format(*(r.get("created_at","")[:10].split("/"))[::-1]) == created_date)
            )
            if len(day_notes) >= same_date_expected:
                # Enough distinct notes on this date
                used_uuids[ext_ref] = day_notes[0].get("uuid","?")
                print(f"    ext_ref={ext_ref[:8]} created={created} → FOUND (date match, {len(day_notes)} notes on {created_date}) note_uuid={day_notes[0].get('uuid','?')[:8]}", flush=True)
                found += 1
            elif day_notes:
                # Notes exist but fewer than expected → suspicious
                used_uuids[ext_ref] = day_notes[0].get("uuid","?")
                print(f"    ext_ref={ext_ref[:8]} created={created} → SUSPECT ({len(day_notes)} notes on {created_date}, expected {same_date_expected}) note_uuid={day_notes[0].get('uuid','?')[:8]}", flush=True)
                missing += 1
            else:
                print(f"    ext_ref={ext_ref[:8]} created={created} → MISSING (0 notes on {created_date})", flush=True)
                missing += 1

print(f"\nSummary: Found {found} / {len(rows)}, Missing {missing}", flush=True)
