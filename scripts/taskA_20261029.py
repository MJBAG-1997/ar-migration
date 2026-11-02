"""
taskA_20261029.py — Task A small fixes (2026-10-29 evening)
A1: Premiums  — list plans where dest has more premiums than the file
A2: Expenditure — categorise 2,233 custom/overwritten rows
A3: Delete 2 empty circles (confirm 0 members; count before/after)
A4: 103 unmatched notes near-duplicate analysis
Rule 9: counts and UUIDs only — no client names or note contents.
One script talks to Plannr at a time.
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
SRC_ACCOUNT_UUID  = "db0ab7f5-8e52-4c89-a915-29fefeeca0d4"
IMPORT_DIR        = r"C:\dev\AR Migration\import_20261027"
EXPORTS_DIR       = r"C:\dev\AR Migration\exports_20261026"
STATE_DIR         = os.path.join(IMPORT_DIR, "state")

PREM_CSV       = os.path.join(IMPORT_DIR, "W6_premiums.csv")
EXP_STATE_FILE = os.path.join(STATE_DIR, "exp_custom_state.json")
V2_CSV         = os.path.join(IMPORT_DIR, "W8_notes_final_v2.csv")

# find the expenditure export — name may have spaces
_exp_candidates = [f for f in os.listdir(EXPORTS_DIR)
                   if "expenditure" in f.lower() and f.endswith(".csv")]
EXP_CSV = os.path.join(EXPORTS_DIR, _exp_candidates[0]) if _exp_candidates else None

EMPTY_CIRCLES = [
    "5ec4636c-bfcd-4a61-9a12-e7707d8508f7",
    "fbadcb0c-fd70-4322-8e07-28bf26b5cdf9",
]

SIM_THRESHOLD = 0.90


def p(*a): print(*a, flush=True)
def stop(msg): p(f"\n!!STOP!! {msg}"); sys.exit(1)

def load_csv(path):
    with open(path, newline="", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))

def load_json(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


# ── API ───────────────────────────────────────────────────────────────────────
sc       = SecretClient(vault_url="https://kv-plannr-ingestion.vault.azure.net/",
                        credential=DefaultAzureCredential())
dest_pat = sc.get_secret("plannr-prod-pat").value
src_pat  = sc.get_secret("plannr-hfp-pat").value

class API:
    def __init__(self, pat, acct):
        self._s = requests.Session()
        self._s.headers.update({
            "Authorization":       f"Bearer {pat}",
            "Accept":              "application/json",
            "X-PLANNR-ACCOUNT-UUID": acct,
        })
    def get(self, path, **kw):    return self._req("GET",    path, **kw)
    def delete(self, path, **kw): return self._req("DELETE", path, **kw)
    def _req(self, method, path, **kw):
        url = f"{BASE_URL}{path}" if path.startswith("/") else path
        for _ in range(4):
            r = self._s.request(method, url, **kw)
            if r.status_code != 429: return r
            wait = int(r.headers.get("Retry-After", 30))
            p(f"  [429] {method} {path} — {wait}s"); time.sleep(wait)
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
src  = API(src_pat,  SRC_ACCOUNT_UUID)

client_map = load_json(os.path.join(STATE_DIR, "client_map.json"))
plan_map   = load_json(os.path.join(STATE_DIR, "plan_map.json"))
dest_uuids = list(client_map.values())


# ─────────────────────────────────────────────────────────────────────────────
# A1: PREMIUMS
# ─────────────────────────────────────────────────────────────────────────────
p("=== A1: PREMIUMS ===")

prem_rows = load_csv(PREM_CSV)
# file_counts: dest_plan_uuid → number of CSV rows for that plan
file_counts = defaultdict(int)
for row in prem_rows:
    sp = (row.get("plan_uuid") or "").strip()
    dp = plan_map.get(sp)
    if dp:
        file_counts[dp] += 1
p(f"  CSV rows: {len(prem_rows)}  dest plan UUIDs: {len(file_counts)}")

anomalies = {}  # dest_plan_uuid → {file_rows, dest_count, premiums}
for dp, fc in sorted(file_counts.items()):
    r = dest.get(f"/api/v1/plans/{dp}/premiums")
    if r.status_code != 200:
        p(f"  [WARN] GET premiums {dp[:8]}: HTTP {r.status_code}"); continue
    prems = r.json().get("data", [])
    if len(prems) > fc:
        anomalies[dp] = {"file_rows": fc, "dest_count": len(prems), "premiums": prems}

if not anomalies:
    p("  No plans with extra premiums ✓")
else:
    p(f"\n  Plans where dest count > CSV rows ({len(anomalies)} plans):")
    for dp, info in anomalies.items():
        p(f"\n  plan={dp}  file_rows={info['file_rows']}  dest_count={info['dest_count']}")
        for pr in info["premiums"]:
            amt     = pr.get("amount") or pr.get("premium_amount") or {}
            if isinstance(amt, dict):
                amt_fmt = amt.get("formatted", str(amt.get("amount", "?")))
            else:
                amt_fmt = str(amt)
            start   = pr.get("start_date") or pr.get("commencement_date") or "?"
            created = pr.get("created_at") or "?"
            p(f"    prem={pr.get('uuid','?')[:8]}…  amount={amt_fmt}  "
              f"start_date={start}  created_at={created}")


# ─────────────────────────────────────────────────────────────────────────────
# A2: EXPENDITURE CATEGORISATION
# ─────────────────────────────────────────────────────────────────────────────
p("\n=== A2: EXPENDITURE CATEGORISATION ===")

if not EXP_CSV:
    p("  [WARN] Expenditure export CSV not found — skipping A2")
else:
    p(f"  Expenditure CSV: {os.path.basename(EXP_CSV)}")
    exp_rows = load_csv(EXP_CSV)
    p(f"  Total rows: {len(exp_rows)}")

    # Load state: keys are row indices (strings)
    if os.path.exists(EXP_STATE_FILE):
        exp_state = load_json(EXP_STATE_FILE)
        p(f"  State file: {len(exp_state)} entries (rows created)")
    else:
        exp_state = {}
        p("  [WARN] State file not found — will count as 0 created")

    # Build field_map from expenditure structure (standard fields)
    field_map = {}
    sections = get_paged(dest, "/api/v1/expenditure-structure/section", max_pages=5)
    for sec in sections:
        sec_name = str(sec.get("name", "") or "")
        fields_r = dest.get(f"/api/v1/expenditure-structure/section/{sec.get('uuid')}/fields",
                             params={"per_page": 200})
        if fields_r.status_code == 200:
            for fld in (fields_r.json().get("data") or fields_r.json()):
                fld_name = str(fld.get("name", "") or "")
                fld_uuid = fld.get("uuid", "")
                if sec_name and fld_name and fld_uuid:
                    field_map[(sec_name, fld_name)] = fld_uuid
    p(f"  Field map: {len(field_map)} standard fields")

    # Categorise each row
    cats = defaultdict(int)
    last_writer = {}
    row_cat = {}

    for i, row in enumerate(exp_rows):
        src_uuid  = (row.get("client_uuid") or "").strip()
        dest_uuid = client_map.get(src_uuid)
        if not dest_uuid:
            row_cat[i] = "no_dest"; cats["no_dest"] += 1; continue
        sec = str(row.get("section") or "").strip()
        fld = str(row.get("type")    or "").strip()
        try:
            amt = float(row.get("amount") or "0")
        except (ValueError, TypeError):
            row_cat[i] = "amount_lte_0"; cats["amount_lte_0"] += 1; continue
        if amt <= 0:
            row_cat[i] = "amount_lte_0"; cats["amount_lte_0"] += 1; continue
        fld_uuid = field_map.get((sec, fld))
        if not fld_uuid:
            row_cat[i] = "custom"; cats["custom"] += 1; continue
        key = (dest_uuid, fld_uuid)
        if key in last_writer:
            old = last_writer[key]
            row_cat[old] = "overwritten"; cats["standard"] -= 1; cats["overwritten"] += 1
        last_writer[key] = i
        row_cat[i] = "standard"; cats["standard"] += 1

    custom_and_overwritten = {i for i, c in row_cat.items() if c in ("custom", "overwritten")}
    in_state_idxs = {int(k) for k in exp_state}

    created  = in_state_idxs & custom_and_overwritten
    zero_amt = {i for i, c in row_cat.items()
                if c == "amount_lte_0" and field_map.get(
                    (str(exp_rows[i].get("section","")).strip(),
                     str(exp_rows[i].get("type","")).strip())) is None}
    other    = custom_and_overwritten - created

    p(f"\n  Row breakdown (total {len(exp_rows)}):")
    p(f"    no_dest:      {cats['no_dest']}")
    p(f"    amount_lte_0: {cats['amount_lte_0']}  "
      f"(zero-amount non-standard subset: {len(zero_amt)})")
    p(f"    custom:       {cats['custom']}")
    p(f"    overwritten:  {cats['overwritten']}")
    p(f"    standard:     {cats['standard']}")
    p(f"\n  Of the {len(custom_and_overwritten)} custom+overwritten rows:")
    p(f"    created (in state):   {len(created)}")
    p(f"    zero-amount (skipped):{len(zero_amt)}")
    p(f"    other (not created, not zero): {len(other)}")
    if other:
        sample = list(other)[:5]
        for i in sample:
            r2 = exp_rows[i]
            p(f"      row {i}: sec={r2.get('section')!r} type={r2.get('type')!r} "
              f"amt={r2.get('amount')!r} cat={row_cat.get(i)!r}")

    # 3 created line samples: source pounds vs posted pence
    p(f"\n  3 created line amount samples (source £ vs posted pence):")
    sample_idxs = sorted(created)[:3]
    for si in sample_idxs:
        row2    = exp_rows[si]
        src_amt = row2.get("amount", "?")
        state_e = exp_state.get(str(si), {})
        pence   = state_e.get("amount", "?")
        sec2    = str(row2.get("section", "")).strip()
        name2   = str(row2.get("type", "")).strip()
        try:
            ratio = float(pence) / float(src_amt) if float(src_amt) != 0 else 0
        except Exception:
            ratio = 0
        p(f"    row {si}: sec={sec2[:20]!r} name={name2[:20]!r}  "
          f"src=£{src_amt}  posted={pence}p  ratio={ratio:.1f}×")


# ─────────────────────────────────────────────────────────────────────────────
# A3: DELETE 2 EMPTY CIRCLES
# ─────────────────────────────────────────────────────────────────────────────
p("\n=== A3: DELETE EMPTY CIRCLES ===")

def count_circles():
    circles = get_paged(dest, "/api/v1/circles", max_pages=20)
    return len(circles)

circles_before = count_circles()
p(f"  Circles before: {circles_before}  (expected 89)")

deleted_count = 0
for cuuid in EMPTY_CIRCLES:
    r = dest.get(f"/api/v1/circles/{cuuid}")
    if r.status_code == 404:
        p(f"  {cuuid[:8]}: 404 — already gone, skipping")
        continue
    if r.status_code != 200:
        stop(f"GET circle {cuuid}: HTTP {r.status_code} {r.text[:100]}")
    data = r.json()
    if isinstance(data, dict) and "data" in data:
        data = data["data"]
    members = data.get("accounts") or []
    member_count = len(members) if isinstance(members, list) else 0
    p(f"  {cuuid[:8]}: members={member_count}")
    if member_count != 0:
        stop(f"Circle {cuuid} has {member_count} member(s) — not deleting")
    rd = dest.delete(f"/api/v1/circles/{cuuid}")
    if rd.status_code not in (200, 201, 204):
        stop(f"DELETE circle {cuuid}: HTTP {rd.status_code} {rd.text[:100]}")
    p(f"    deleted ✓ (HTTP {rd.status_code})")
    deleted_count += 1

circles_after = count_circles()
p(f"  Circles after:  {circles_after}  (expected {circles_before - deleted_count})")
expected_after = circles_before - deleted_count
if circles_after != expected_after:
    stop(f"Circle count mismatch: expected {expected_after}, got {circles_after}")
p(f"  ✓ circles: {circles_before} → {circles_after} (−{deleted_count})")


# ─────────────────────────────────────────────────────────────────────────────
# A4: UNMATCHED NOTES NEAR-DUPLICATE ANALYSIS
# ─────────────────────────────────────────────────────────────────────────────
p("\n=== A4: UNMATCHED NOTES ANALYSIS ===")

def norm_html(text):
    text = re.sub(r'<[^>]+>', ' ', str(text or ''))
    text = html.unescape(text)
    return re.sub(r'\s+', ' ', text).strip().lower()

def strip_bracket_header(text):
    """Remove a leading [bracketed] line."""
    return re.sub(r'^\s*\[[^\]]{0,80}\]\s*', '', text).strip()

def sim(a, b):
    if not a and not b: return 1.0
    if not a or not b: return 0.0
    return difflib.SequenceMatcher(None, a, b).ratio()

# Load v2
v2_rows = load_csv(V2_CSV)
p(f"  v2 rows: {len(v2_rows)}")

# Build v2 lookup: (dest_uuid, date) → list of (norm_full, norm_stripped)
v2_by_client_date = defaultdict(list)
for row in v2_rows:
    src_acct  = (row.get("account_uuid") or "").strip()
    dest_acct = client_map.get(src_acct, "")
    if not dest_acct: continue
    created_raw = (row.get("created_at") or "")
    dt = created_raw[:10] if created_raw else ""
    if not dt: continue
    nfull = norm_html(row.get("contents", ""))
    nstrip = strip_bracket_header(nfull)
    v2_by_client_date[(dest_acct, dt)].append((nfull, nstrip))

# Build v2 exact-match lookup for fast matching pass
v2_exact = defaultdict(set)
for (dest_acct, dt), norms in v2_by_client_date.items():
    for nfull, _ in norms:
        if nfull:
            v2_exact[(dest_acct, dt)].add(nfull)

# Fetch all dest notes (one client at a time, deduplicate by note UUID)
p(f"  Fetching dest notes for {len(dest_uuids)} clients…")
dest_note_map = {}  # note_uuid → {note, client_uuid}
for i, du in enumerate(dest_uuids, 1):
    notes = get_paged(dest, f"/api/v1/client/{du}/all-notes")
    for note in notes:
        nuid = str(note.get("uuid", ""))
        if nuid and nuid not in dest_note_map:
            dest_note_map[nuid] = {"note": note, "client_uuid": du}
    if i % 50 == 0:
        p(f"    {i}/{len(dest_uuids)} — {len(dest_note_map)} notes so far")
p(f"  Total dest notes: {len(dest_note_map)}")

# Match dest notes against v2
matched_once = 0
matched_multi = 0
unmatched_notes = {}  # note_uuid → {note, client_uuid}

for nuid, info in dest_note_map.items():
    note = info["note"]
    du   = info["client_uuid"]
    dt   = str(note.get("created_at") or "")[:10]
    norm = norm_html(note.get("contents", ""))
    exact_set = v2_exact.get((du, dt), set())
    if norm in exact_set:
        matched_once += 1
    else:
        # count matching v2 rows (any normalised match in the day's set)
        cnt = sum(1 for n in exact_set if n == norm)
        if cnt > 1:
            matched_multi += 1
        elif cnt == 1:
            matched_once += 1
        else:
            unmatched_notes[nuid] = info

p(f"  Matched (once):   {matched_once}")
p(f"  Matched (multi):  {matched_multi}")
p(f"  Unmatched:        {len(unmatched_notes)}")

# Near-duplicate check for unmatched notes
near_dup_count = 0
remainder = []

for nuid, info in unmatched_notes.items():
    note = info["note"]
    du   = info["client_uuid"]
    dt   = str(note.get("created_at") or "")[:10]
    note_norm   = norm_html(note.get("contents", ""))
    note_strip  = strip_bracket_header(note_norm)

    candidates = v2_by_client_date.get((du, dt), [])
    best_sim = 0.0
    for (v2full, v2strip) in candidates:
        s1 = sim(note_norm,  v2full)
        s2 = sim(note_norm,  v2strip)
        s3 = sim(note_strip, v2full)
        s4 = sim(note_strip, v2strip)
        best_sim = max(best_sim, s1, s2, s3, s4)
        if best_sim >= SIM_THRESHOLD:
            break

    if best_sim >= SIM_THRESHOLD:
        near_dup_count += 1
    else:
        notable_type = str(note.get("notable_type") or note.get("type") or "").lower()
        author       = str(note.get("created_by") or note.get("author") or "IT SYSTEM")
        remainder.append({
            "uuid":         nuid,
            "client_uuid":  du,
            "date":         dt,
            "type":         notable_type,
            "author":       author,
        })

p(f"\n  Near-duplicates (≥{int(SIM_THRESHOLD*100)}% similar to a v2 row): {near_dup_count}")
p(f"  Remaining (not near-dup):                                      {len(remainder)}")

# Group remainder by date + type, list 3 UUIDs per group
groups = defaultdict(list)
for item in remainder:
    key = (item["date"], item["type"])
    groups[key].append(item)

p(f"\n  Remainder grouped by date+type ({len(groups)} groups):")
for (dt, typ), items in sorted(groups.items()):
    p(f"\n    date={dt}  type={typ!r}  count={len(items)}")
    for item in items[:3]:
        p(f"      note={item['uuid']}  client={item['client_uuid']}")

p("\n=== TASK A COMPLETE ===")
