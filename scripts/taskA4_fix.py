"""
taskA4_fix.py — Re-run A4 with correct date parsing.
v2 CSV uses DD/MM/YYYY; dest notes use YYYY-MM-DD (ISO first 10 chars).
"""
import csv, difflib, html, json, os, re, sys, time
from collections import defaultdict

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from azure.identity import DefaultAzureCredential
from azure.keyvault.secrets import SecretClient
import requests

BASE_URL          = "https://api.plannrcrm.com"
DEST_ACCOUNT_UUID = "ccd7f98b-d6a4-46ad-bab3-242ab42826eb"
IMPORT_DIR        = r"C:\dev\AR Migration\import_20261027"
STATE_DIR         = os.path.join(IMPORT_DIR, "state")
V2_CSV            = os.path.join(IMPORT_DIR, "W8_notes_final_v2.csv")
SIM_THRESHOLD     = 0.90


def p(*a): print(*a, flush=True)
def stop(msg): p(f"\n!!STOP!! {msg}"); sys.exit(1)

def load_csv(path):
    with open(path, newline="", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))

def load_json(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)

def to_ymd(s):
    """Convert any date string to YYYY-MM-DD."""
    s = str(s or "").strip()
    m = re.match(r'^(\d{2})/(\d{2})/(\d{4})', s)
    if m: return f"{m.group(3)}-{m.group(2)}-{m.group(1)}"
    return s[:10]

def norm_html(text):
    text = re.sub(r'<[^>]+>', ' ', str(text or ''))
    text = html.unescape(text)
    return re.sub(r'\s+', ' ', text).strip().lower()

def strip_bracket_header(text):
    return re.sub(r'^\s*\[[^\]]{0,80}\]\s*', '', text).strip()

def sim(a, b):
    if not a and not b: return 1.0
    if not a or not b: return 0.0
    return difflib.SequenceMatcher(None, a, b).ratio()


sc       = SecretClient(vault_url="https://kv-plannr-ingestion.vault.azure.net/",
                        credential=DefaultAzureCredential())
dest_pat = sc.get_secret("plannr-prod-pat").value

class API:
    def __init__(self, pat, acct):
        self._s = requests.Session()
        self._s.headers.update({"Authorization": f"Bearer {pat}",
                                 "Accept": "application/json",
                                 "X-PLANNR-ACCOUNT-UUID": acct})
    def get(self, path, **kw):
        url = f"{BASE_URL}{path}" if path.startswith("/") else path
        for _ in range(4):
            r = self._s.request("GET", url, **kw)
            if r.status_code != 429: return r
            time.sleep(int(r.headers.get("Retry-After", 30)))
        return r

def get_paged(api, path, max_pages=500, params=None):
    items, page = [], 1
    while page <= max_pages:
        pr = dict(params or {}); pr["page"] = page; pr["per_page"] = 200
        r = api.get(path, params=pr)
        if r.status_code != 200:
            p(f"  [WARN] {path} → HTTP {r.status_code}"); break
        body = r.json()
        if isinstance(body, list): items.extend(body); break
        data = body.get("data", [])
        if isinstance(data, list): items.extend(data)
        elif isinstance(data, dict): items.append(data)
        meta = body.get("meta", {}) if isinstance(body, dict) else {}
        if page >= (meta.get("last_page") or 1): break
        page += 1
    return items

dest = API(dest_pat, DEST_ACCOUNT_UUID)
client_map = load_json(os.path.join(STATE_DIR, "client_map.json"))
dest_uuids = list(client_map.values())

p("=== A4: UNMATCHED NOTES ANALYSIS (fixed date parsing) ===")

# Load v2
v2_rows = load_csv(V2_CSV)
p(f"  v2 rows: {len(v2_rows)}")

# Build v2 lookup: (dest_uuid, YYYY-MM-DD) → set of norms
v2_by_client_date = defaultdict(list)
for row in v2_rows:
    src_acct  = (row.get("account_uuid") or "").strip()
    dest_acct = client_map.get(src_acct, "")
    if not dest_acct: continue
    dt = to_ymd(row.get("created_at", ""))   # DD/MM/YYYY → YYYY-MM-DD
    if not dt or dt == "": continue
    nfull  = norm_html(row.get("contents", ""))
    nstrip = strip_bracket_header(nfull)
    v2_by_client_date[(dest_acct, dt)].append((nfull, nstrip))

v2_exact = defaultdict(set)
for (da, dt), norms in v2_by_client_date.items():
    for nfull, _ in norms:
        if nfull:
            v2_exact[(da, dt)].add(nfull)

# Fetch dest notes
p(f"  Fetching dest notes for {len(dest_uuids)} clients…")
dest_note_map = {}
for i, du in enumerate(dest_uuids, 1):
    notes = get_paged(dest, f"/api/v1/client/{du}/all-notes")
    for note in notes:
        nuid = str(note.get("uuid", ""))
        if nuid and nuid not in dest_note_map:
            dest_note_map[nuid] = {"note": note, "client_uuid": du}
    if i % 50 == 0:
        p(f"    {i}/{len(dest_uuids)} — {len(dest_note_map)} notes")
p(f"  Total dest notes: {len(dest_note_map)}")

# Match
matched_once = 0
matched_multi = 0
unmatched_notes = {}

for nuid, info in dest_note_map.items():
    note = info["note"]
    du   = info["client_uuid"]
    dt   = str(note.get("created_at") or "")[:10]   # dest: YYYY-MM-DD already
    norm = norm_html(note.get("contents", ""))
    exact_set = v2_exact.get((du, dt), set())
    if norm in exact_set:
        matched_once += 1
    else:
        cnt = sum(1 for n in exact_set if n == norm)
        if cnt > 1:
            matched_multi += 1
        elif cnt == 1:
            matched_once += 1
        else:
            unmatched_notes[nuid] = info

p(f"  Matched once:  {matched_once}")
p(f"  Matched multi: {matched_multi}")
p(f"  Unmatched:     {len(unmatched_notes)}")

# Near-dup check
near_dup_count = 0
remainder = []

for nuid, info in unmatched_notes.items():
    note = info["note"]
    du   = info["client_uuid"]
    dt   = str(note.get("created_at") or "")[:10]
    note_norm  = norm_html(note.get("contents", ""))
    note_strip = strip_bracket_header(note_norm)

    candidates = v2_by_client_date.get((du, dt), [])
    best_sim = 0.0
    for (v2full, v2strip) in candidates:
        s = max(sim(note_norm, v2full), sim(note_norm, v2strip),
                sim(note_strip, v2full), sim(note_strip, v2strip))
        if s > best_sim: best_sim = s
        if best_sim >= SIM_THRESHOLD: break

    if best_sim >= SIM_THRESHOLD:
        near_dup_count += 1
    else:
        notable_type = str(note.get("notable_type") or note.get("type") or "").lower()
        author_obj   = note.get("created_by") or note.get("author") or {}
        if isinstance(author_obj, dict):
            author = author_obj.get("name","") or "unknown"
        else:
            author = str(author_obj)
        remainder.append({
            "uuid":        nuid,
            "client_uuid": du,
            "date":        dt,
            "type":        notable_type,
            "author":      author,
        })

p(f"\n  Near-duplicates (≥{int(SIM_THRESHOLD*100)}% similar): {near_dup_count}")
p(f"  Remaining (not near-dup):                              {len(remainder)}")

# Group remainder by date+type, list 3 UUIDs per group
groups = defaultdict(list)
for item in remainder:
    groups[(item["date"], item["type"], item["author"])].append(item)

p(f"\n  Remainder grouped by date+type+author ({len(groups)} groups):")
for (dt, typ, auth), items in sorted(groups.items()):
    p(f"\n    date={dt}  type={typ!r}  author={auth!r}  count={len(items)}")
    for item in items[:3]:
        p(f"      note={item['uuid']}  client={item['client_uuid']}")

p("\n=== A4 FIX COMPLETE ===")
