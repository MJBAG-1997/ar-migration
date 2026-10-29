"""
Read-only reconciliation — HFP migration (2026-10-29).

1. Notes: dest all-notes for all 356 clients. Unique source refs vs 6757;
   duplicate source refs; notes with no data_migration ref. Explain C1 and C6.
2. Note dates: match to v2 CSV by source ref; compare created_at as real dates.
   Categories: same / day-month swapped / set to Oct 27-29 / other.
   Earliest and latest on both sites.
3. Duplicates by type: circles, relationships, dependants, income, valuations,
   charges — dest count vs expected, identical-twin count for any over.
4. Dependants pairs: list every duplicate pair (UUID, created_at), mark later.
5. Expenditure: rebuild field map, simulate original run against source CSV.
   Count: custom, amount<=0, overwritten, written (last), failed.

Rule 9: no client names, no note contents. UUIDs, counts, types only.
No destination writes.
"""

import csv, json, os, sys, time
from collections import defaultdict
from datetime import date, datetime, timezone

import requests

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

BASE_URL          = "https://api.plannrcrm.com"
KV_NAME           = "kv-plannr-ingestion"
DEST_PAT_SECRET   = "plannr-prod-pat"
DEST_ACCOUNT_UUID = "ccd7f98b-d6a4-46ad-bab3-242ab42826eb"
MIGRATION_CO      = "data_migration"
IMPORT_DIR        = r"C:\dev\AR Migration\import_20261027"
EXPORTS_DIR       = r"C:\dev\AR Migration\exports_20261026"

EXPECTED = {
    "circles":       87,
    "relationships": 147,
    "dependants":    249,
    "income":        316,
    "valuations":    1526,
    "charges":       542,
}

C1_UUID = "587bbc91-6c99-4114-a6c0-8416c3a68cb6"
C6_UUID = "26697c8a-e94b-4eeb-b021-d338793e63f7"

EXP_EXPORT = os.path.join(
    EXPORTS_DIR, "Client Expenditure Export HFP Migration 2026-10-27.csv")

# ── Helpers ───────────────────────────────────────────────────────────────────
def p(*a): print(*a, flush=True)

def resolve_pat():
    from azure.identity import DefaultAzureCredential
    from azure.keyvault.secrets import SecretClient
    sc = SecretClient(
        vault_url=f"https://{KV_NAME}.vault.azure.net/",
        credential=DefaultAzureCredential(),
    )
    return sc.get_secret(DEST_PAT_SECRET).value

def load_json(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)

def load_csv(path):
    with open(path, newline="", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


class Dest:
    def __init__(self, pat):
        self._s = requests.Session()
        self._s.headers.update({
            "Authorization":          f"Bearer {pat}",
            "Accept":                 "application/json",
            "X-PLANNR-ACCOUNT-UUID":  DEST_ACCOUNT_UUID,
        })

    def get(self, path, **kw):
        return self._s.get(f"{BASE_URL}{path}", **kw)

    def post(self, path, **kw):
        return self._s.post(f"{BASE_URL}{path}", **kw)


def get_paged(dest, path, max_pages=200, params=None, _warn_on_error=False):
    items, page = [], 1
    while page <= max_pages:
        pr = dict(params or {}); pr.update({"page": page, "per_page": 200})
        r = dest.get(path, params=pr)
        if r.status_code != 200:
            if _warn_on_error or page == 1:
                p(f"    [WARN] GET {path} → HTTP {r.status_code}")
            break
        body = r.json()
        # Handle flat list response (e.g. all-notes returns a plain array)
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


# ── Date helpers ──────────────────────────────────────────────────────────────
OCT_IMPORT_DATES = {date(2026, 10, 27), date(2026, 10, 28), date(2026, 10, 29)}

def parse_csv_date(s):
    """Parse DD/MM/YYYY → date.  Returns None on failure."""
    s = (s or "").strip()
    for fmt in ("%d/%m/%Y", "%d/%m/%Y %H:%M:%S", "%Y-%m-%dT%H:%M:%S%z",
                "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d"):
        try:
            if "T" in s or (len(s) == 10 and s[4] == "-"):
                dt = datetime.fromisoformat(s.replace("Z", "+00:00"))
                return dt.date()
            dt = datetime.strptime(s, fmt)
            return dt.date()
        except (ValueError, TypeError):
            continue
    return None

def parse_iso_date(s):
    """Parse ISO 8601 → date.  Returns None on failure."""
    s = (s or "").strip()
    if not s:
        return None
    try:
        dt = datetime.fromisoformat(s.replace("Z", "+00:00"))
        return dt.date()
    except (ValueError, TypeError):
        pass
    try:
        return date.fromisoformat(s[:10])
    except (ValueError, TypeError):
        return None

def extract_source_ref(note):
    """Return (data_migration source UUID or '') from a note object."""
    # Try direct fields first
    if note.get("external_reference_company") == MIGRATION_CO:
        return str(note.get("external_reference_value") or "")
    # Try nested external_references array
    for ref in (note.get("external_references") or []):
        if ref.get("third_party") == MIGRATION_CO or ref.get("company") == MIGRATION_CO:
            return str(ref.get("reference") or ref.get("value") or "")
    # Try ext_ref shorthand
    if str(note.get("ext_ref_company") or "") == MIGRATION_CO:
        return str(note.get("ext_ref_value") or "")
    return ""


# ── PART 1 + 2: Notes reconciliation ─────────────────────────────────────────
def reconcile_notes(dest, client_map, v2_rows):
    p("\n=== PART 1+2: NOTES RECONCILIATION ===")

    # Build expected source-ref set from v2
    v2_ref_to_csv_date = {}      # src_note_uuid → csv created_at string
    for row in v2_rows:
        ref = (row.get("external_reference_value") or "").strip()
        if ref:
            v2_ref_to_csv_date[ref] = row.get("created_at", "")
    p(f"  v2 file: {len(v2_rows)} rows, {len(v2_ref_to_csv_date)} unique source refs")

    # Probe one client to discover response structure and field names
    dest_uuids = list(client_map.values())
    probe_uuid = dest_uuids[0]
    # First try without extra params
    probe_r = dest.get(f"/api/v1/client/{probe_uuid}/all-notes",
                       params={"page": 1, "per_page": 5})
    p(f"  Probe HTTP {probe_r.status_code} — content-type: "
      f"{probe_r.headers.get('content-type','?')[:60]}")
    if probe_r.status_code == 200:
        pb = probe_r.json()
        if isinstance(pb, list):
            p(f"  Response is a plain list ({len(pb)} items on first fetch)")
            if pb:
                p(f"  Note object keys: {sorted(pb[0].keys())}")
        elif isinstance(pb, dict):
            p(f"  Response keys: {sorted(pb.keys())}")
            data = pb.get("data", [])
            if isinstance(data, list) and data:
                p(f"  Note object keys: {sorted(data[0].keys())}")
            elif isinstance(data, dict):
                p(f"  data is a dict, keys: {sorted(data.keys())}")

    # Fetch all notes from all 356 dest clients (no extra params)
    all_dest_notes = []       # [{uuid, created_at, src_ref, client_uuid}, ...]
    n_clients = len(dest_uuids)
    for i, du in enumerate(dest_uuids, 1):
        notes = get_paged(dest, f"/api/v1/client/{du}/all-notes",
                          max_pages=500)
        for note in notes:
            all_dest_notes.append({
                "uuid":        str(note.get("uuid") or ""),
                "created_at":  str(note.get("created_at") or ""),
                "src_ref":     extract_source_ref(note),
                "client_uuid": du,
            })
        if i % 50 == 0 or i == n_clients:
            p(f"  {i}/{n_clients} clients — {len(all_dest_notes)} notes so far")

    p(f"\n  Total dest notes: {len(all_dest_notes)}")
    p(f"  Expected (v2 file): 6757")

    # Unique source refs
    ref_to_dest_notes = defaultdict(list)   # src_ref → [dest_note_uuids]
    no_ref_notes = []
    for n in all_dest_notes:
        if n["src_ref"]:
            ref_to_dest_notes[n["src_ref"]].append(n)
        else:
            no_ref_notes.append(n)

    p(f"  Unique source refs: {len(ref_to_dest_notes)}")
    p(f"  Notes with no data_migration ref: {len(no_ref_notes)}")

    # Duplicate source refs (same src ref on multiple dest notes)
    dup_refs = {ref: notes for ref, notes in ref_to_dest_notes.items() if len(notes) > 1}
    p(f"  Source refs appearing more than once: {len(dup_refs)}")
    if dup_refs:
        dup_counts = sorted(((ref, len(ns)) for ref, ns in dup_refs.items()),
                            key=lambda x: -x[1])
        p(f"  Worst duplicates (top 5):")
        for ref, cnt in dup_counts[:5]:
            p(f"    src_ref={ref[:8]}…  occurrences={cnt}")

    # Cross-check: refs in v2 that have no dest note
    v2_refs_set     = set(v2_ref_to_csv_date.keys())
    dest_refs_set   = set(ref_to_dest_notes.keys())
    in_v2_not_dest  = v2_refs_set - dest_refs_set
    in_dest_not_v2  = dest_refs_set - v2_refs_set
    p(f"\n  In v2 but no dest note:  {len(in_v2_not_dest)}")
    p(f"  In dest but not in v2:   {len(in_dest_not_v2)}")
    if in_dest_not_v2:
        p(f"  Sample src refs in dest but not v2 (first 8):")
        for r in sorted(in_dest_not_v2)[:8]:
            p(f"    {r}")

    # Explain C1 and C6
    for label, cuuid, expected_count in [("C1", C1_UUID, 9), ("C6", C6_UUID, 147)]:
        c_notes = [n for n in all_dest_notes if n["client_uuid"] == cuuid]
        c_refs  = [n["src_ref"] for n in c_notes if n["src_ref"]]
        in_v2   = [n for n in c_notes if n["src_ref"] in v2_refs_set]
        not_in_v2 = [n for n in c_notes if n["src_ref"] and n["src_ref"] not in v2_refs_set]
        no_ref_c  = [n for n in c_notes if not n["src_ref"]]
        dup_c     = [n for n in c_notes if n["src_ref"] in dup_refs]
        p(f"\n  {label} ({cuuid[:8]}…):")
        p(f"    Total dest notes: {len(c_notes)}  expected: {expected_count}")
        p(f"    Notes matching v2 ref:  {len(in_v2)}")
        p(f"    Notes not in v2 (extra from earlier runs): {len(not_in_v2)}")
        p(f"    Notes with no data_migration ref:          {len(no_ref_c)}")
        p(f"    Notes that are duplicates of a v2 ref:     {len(dup_c)}")
        for n in not_in_v2[:5]:
            p(f"      extra: dest_uuid={n['uuid'][:8]}… src_ref={n['src_ref'][:8]}… created={n['created_at'][:10]}")

    # ── PART 2: Note dates ────────────────────────────────────────────────────
    p("\n=== PART 2: NOTE DATES ===")

    same_count       = 0
    swapped_count    = 0
    import_date_count = 0
    other_count      = 0
    no_match_count   = 0   # dest note ref not in v2

    dest_dates_all   = []
    src_dates_all    = []

    for ref, csv_date_str in v2_ref_to_csv_date.items():
        src_date = parse_csv_date(csv_date_str)
        if src_date:
            src_dates_all.append(src_date)

        dest_ns = ref_to_dest_notes.get(ref, [])
        if not dest_ns:
            no_match_count += 1
            continue

        # Use first occurrence
        dest_n    = dest_ns[0]
        dest_date = parse_iso_date(dest_n["created_at"])
        if dest_date:
            dest_dates_all.append(dest_date)

        if src_date is None or dest_date is None:
            other_count += 1
            continue

        if dest_date == src_date:
            same_count += 1
        elif dest_date in OCT_IMPORT_DATES:
            import_date_count += 1
        elif dest_date.day == src_date.month and dest_date.month == src_date.day:
            # day and month transposed
            swapped_count += 1
        else:
            other_count += 1

    # Notes in dest not in v2 also contribute to dest dates
    for ref in in_dest_not_v2:
        for n in ref_to_dest_notes[ref]:
            d = parse_iso_date(n["created_at"])
            if d:
                dest_dates_all.append(d)

    p(f"  Date comparison (v2 source dates vs dest dates):")
    p(f"    Same date:                {same_count}")
    p(f"    Day and month swapped:    {swapped_count}")
    p(f"    Set to import date (Oct 27-29 2026): {import_date_count}")
    p(f"    Other difference:         {other_count}")
    p(f"    No dest note found:       {no_match_count}")

    if src_dates_all:
        p(f"\n  Source (v2 CSV) dates:")
        p(f"    Earliest: {min(src_dates_all)}")
        p(f"    Latest:   {max(src_dates_all)}")
    if dest_dates_all:
        p(f"  Destination note dates:")
        p(f"    Earliest: {min(dest_dates_all)}")
        p(f"    Latest:   {max(dest_dates_all)}")

    return all_dest_notes, ref_to_dest_notes, no_ref_notes


# ── PART 3: Duplicates by type ────────────────────────────────────────────────
def ext_ref_key(rec):
    """Extract data_migration external_reference_value, or '' if absent."""
    if str(rec.get("external_reference_company") or "") == MIGRATION_CO:
        return str(rec.get("external_reference_value") or "")
    for ref in (rec.get("external_references") or []):
        if ref.get("third_party") == MIGRATION_CO or ref.get("company") == MIGRATION_CO:
            return str(ref.get("reference") or ref.get("value") or "")
    return ""


def twin_report(records, key_fn, label):
    """Count twins (records with identical key on same [client/plan]).  Print result."""
    groups = defaultdict(list)
    no_key = 0
    for rec in records:
        k = key_fn(rec)
        if k:
            groups[k].append(rec)
        else:
            no_key += 1
    dup_groups  = {k: v for k, v in groups.items() if len(v) > 1}
    twin_count  = sum(len(v) - 1 for v in dup_groups.values())
    p(f"    {label}: records with key: {len(records)-no_key}  without: {no_key}")
    p(f"    Duplicate groups: {len(dup_groups)}  extra records (twins): {twin_count}")
    return dup_groups


def count_all_duplicates(dest, client_map, plan_map):
    p("\n=== PART 3: DUPLICATES BY TYPE ===")

    dest_uuids = list(client_map.values())
    dest_plan_uuids = list(plan_map.values())

    # ── Circles (global, has ext_ref) ──
    circles_recs = get_paged(dest, "/api/v1/circles", max_pages=200)
    p(f"\n  circles:")
    p(f"    Dest total: {len(circles_recs)}  Expected: {EXPECTED['circles']}")
    if len(circles_recs) > EXPECTED["circles"]:
        twin_report(circles_recs, ext_ref_key, "by ext_ref")

    # ── Relationships (global then per-client; no ext_ref → key by account+related+type) ──
    rel_recs = get_paged(dest, "/api/v1/relationships", max_pages=500)
    if not rel_recs:
        p(f"\n  relationships: global returned 0 — fetching per client…")
        rel_recs = []
        for i, du in enumerate(dest_uuids, 1):
            recs = get_paged(dest, f"/api/v1/client/{du}/relationships", max_pages=50)
            for r in recs:
                r["_client_uuid"] = du
            rel_recs.extend(recs)
            if i % 50 == 0:
                p(f"    {i}/{len(dest_uuids)} — {len(rel_recs)} so far")
    p(f"\n  relationships:")
    p(f"    Dest total: {len(rel_recs)}  Expected: {EXPECTED['relationships']}")
    if len(rel_recs) > EXPECTED["relationships"]:
        def rel_key(r):
            a = str(r.get("account_uuid") or r.get("_client_uuid") or "")
            b = str(r.get("related_account_uuid") or "")
            t = str(r.get("relationship") or r.get("type") or "")
            return f"{min(a,b)}|{max(a,b)}|{t}" if a and b else ""
        twin_report(rel_recs, rel_key, "by account+related+type")

    # ── Dependants (global v2, has ext_ref) ──
    dep_recs = get_paged(dest, "/api/v2/dependants", max_pages=500)
    p(f"\n  dependants:")
    p(f"    Dest total: {len(dep_recs)}  Expected: {EXPECTED['dependants']}")
    if len(dep_recs) > EXPECTED["dependants"]:
        twin_report(dep_recs, ext_ref_key, "by ext_ref")

    # ── Income (per-client, has ext_ref) ──
    inc_recs = []
    p(f"\n  income: fetching per client…")
    for i, du in enumerate(dest_uuids, 1):
        recs = get_paged(dest, f"/api/v1/client/{du}/income-and-employment",
                          max_pages=20)
        for r in recs:
            r["_client_uuid"] = du
        inc_recs.extend(recs)
        if i % 100 == 0:
            p(f"    {i}/{len(dest_uuids)} — {len(inc_recs)} so far")
    p(f"    Dest total: {len(inc_recs)}  Expected: {EXPECTED['income']}")
    if len(inc_recs) > EXPECTED["income"]:
        twin_report(inc_recs, ext_ref_key, "by ext_ref")

    # ── Valuations (global, fallback to per-plan; no ext_ref → key by plan+date+value) ──
    val_recs = get_paged(dest, "/api/v1/valuations", max_pages=1000)
    if not val_recs:
        p(f"\n  valuations: global returned 0 — fetching per plan ({len(dest_plan_uuids)} plans)…")
        val_recs = []
        for i, pu in enumerate(dest_plan_uuids, 1):
            recs = get_paged(dest, f"/api/v1/plans/{pu}/valuations", max_pages=20)
            for r in recs:
                r["_plan_uuid"] = pu
            val_recs.extend(recs)
            if i % 200 == 0:
                p(f"    {i}/{len(dest_plan_uuids)} — {len(val_recs)} so far")
    p(f"\n  valuations:")
    p(f"    Dest total: {len(val_recs)}  Expected: {EXPECTED['valuations']}")
    if len(val_recs) > EXPECTED["valuations"]:
        def val_key(r):
            pu = str(r.get("plan_uuid") or r.get("_plan_uuid") or "")
            d  = str(r.get("valuation_date") or "")[:10]
            v  = str(r.get("value") or r.get("amount") or "")
            return f"{pu}|{d}|{v}" if pu and d else ""
        twin_report(val_recs, val_key, "by plan+date+value")

    # ── Charges (global, fallback to per-plan; no ext_ref → key by plan+type+start+%) ──
    chg_recs = get_paged(dest, "/api/v1/charges", max_pages=500)
    if not chg_recs:
        p(f"\n  charges: global returned 0 — fetching per plan…")
        chg_recs = []
        for i, pu in enumerate(dest_plan_uuids, 1):
            recs = get_paged(dest, f"/api/v1/plans/{pu}/charges", max_pages=20)
            for r in recs:
                r["_plan_uuid"] = pu
            chg_recs.extend(recs)
            if i % 200 == 0:
                p(f"    {i}/{len(dest_plan_uuids)} — {len(chg_recs)} so far")
    p(f"\n  charges:")
    p(f"    Dest total: {len(chg_recs)}  Expected: {EXPECTED['charges']}")
    if len(chg_recs) > EXPECTED["charges"]:
        def chg_key(r):
            pu = str(r.get("plan_uuid") or r.get("_plan_uuid") or "")
            t  = str(r.get("type") or "")
            s  = str(r.get("start_date") or "")[:10]
            pc = str(r.get("percentage") or r.get("amount") or "")
            return f"{pu}|{t}|{s}|{pc}" if pu else ""
        twin_report(chg_recs, chg_key, "by plan+type+start+pct")

    return dep_recs


# ── PART 4: Dependant duplicate pairs ────────────────────────────────────────
def list_dependant_pairs(dep_recs):
    p("\n=== PART 4: DEPENDANT DUPLICATE PAIRS ===")

    # Group by external_reference_value
    ref_map = defaultdict(list)
    no_ref  = 0
    for rec in dep_recs:
        ref_val = ""
        for ref in (rec.get("external_references") or []):
            if ref.get("third_party") == MIGRATION_CO or ref.get("company") == MIGRATION_CO:
                ref_val = str(ref.get("reference") or ref.get("value") or "")
                break
        if not ref_val:
            ref_val = str(rec.get("external_reference_value") or "")
        if ref_val:
            ref_map[ref_val].append(rec)
        else:
            no_ref += 1

    dup_groups = {k: v for k, v in ref_map.items() if len(v) > 1}
    p(f"  Duplicate groups (same source ref): {len(dup_groups)}")
    p(f"  Records with no ref: {no_ref}")

    pair_count = 0
    for src_ref, recs in sorted(dup_groups.items()):
        # Sort by created_at ascending
        def _key(r):
            return str(r.get("created_at") or "")
        recs_sorted = sorted(recs, key=_key)
        for j, rec in enumerate(recs_sorted):
            marker = "  [LATER — DELETE]" if j > 0 else "  [keep]"
            p(f"  src_ref={src_ref[:8]}…  uuid={rec.get('uuid','?')[:8]}…  "
              f"created={str(rec.get('created_at','?'))[:10]}  {marker}")
        pair_count += 1
        if pair_count >= 5:
            remaining = len(dup_groups) - 5
            if remaining > 0:
                p(f"  … {remaining} more duplicate groups (full list in handover notes)")
            break

    p(f"\n  Total pairs to remove (later of each): {sum(len(v)-1 for v in dup_groups.values())}")


# ── PART 5: Expenditure analysis ──────────────────────────────────────────────
def build_exp_field_map(dest):
    field_map = {}
    sections  = get_paged(dest, "/api/v1/expenditure-structure/section", max_pages=5)
    for sec in sections:
        sec_name = str(sec.get("name", "") or "")
        fields   = sec.get("fields", [])
        if not fields:
            sec_uuid = sec.get("uuid", "")
            if sec_uuid:
                rf = dest.get(f"/api/v1/expenditure-structure/section/{sec_uuid}",
                              params={"include": "fields"})
                if rf.status_code == 200:
                    fields = rf.json().get("data", {}).get("fields", [])
        for fld in fields:
            fld_name = str(fld.get("name", "") or "")
            fld_uuid = fld.get("uuid", "")
            if sec_name and fld_name and fld_uuid:
                field_map[(sec_name, fld_name)] = fld_uuid
    return field_map


def analyze_expenditure(dest, client_map):
    p("\n=== PART 5: EXPENDITURE ANALYSIS ===")

    field_map = build_exp_field_map(dest)
    p(f"  Dest expenditure field map: {len(field_map)} fields")
    if not field_map:
        p("  ERROR: could not build field map")
        return

    exp_rows = load_csv(EXP_EXPORT)
    p(f"  Source rows: {len(exp_rows)}")

    # Simulate run in source-row order; track last writer per (dest_uuid, fld_uuid)
    last_writer = {}        # (dest_uuid, fld_uuid) → row index
    row_category = {}       # row index → category string

    no_dest_uuid   = 0
    amount_lte_0   = 0
    custom_count   = 0      # (section, type) not in field_map
    mapped_count   = 0      # mapped and has amount > 0

    for i, row in enumerate(exp_rows):
        src_uuid  = (row.get("client_uuid") or "").strip()
        dest_uuid = client_map.get(src_uuid)
        if not dest_uuid:
            row_category[i] = "no_dest_uuid"
            no_dest_uuid += 1
            continue

        sec_name = str(row.get("section") or "").strip()
        fld_name = str(row.get("type") or "").strip()
        fld_uuid = field_map.get((sec_name, fld_name))
        if not fld_uuid:
            row_category[i] = "custom"
            custom_count += 1
            continue

        try:
            amount = float(row.get("amount") or "0")
        except ValueError:
            row_category[i] = "amount_lte_0"
            amount_lte_0 += 1
            continue

        if amount <= 0:
            row_category[i] = "amount_lte_0"
            amount_lte_0 += 1
            continue

        # This row would be posted — check if it overwrites an earlier row
        key = (dest_uuid, fld_uuid)
        if key in last_writer:
            # Previous row for this key is overwritten
            prev_i = last_writer[key]
            if row_category.get(prev_i) != "overwritten":
                row_category[prev_i] = "overwritten"
        last_writer[key] = i
        row_category[i] = "written"   # tentative; may be overwritten later
        mapped_count += 1

    # Final counts after full pass
    n_written    = sum(1 for c in row_category.values() if c == "written")
    n_overwritten = sum(1 for c in row_category.values() if c == "overwritten")
    n_custom      = custom_count
    n_lte0        = amount_lte_0
    n_no_dest     = no_dest_uuid

    # "Failed": rows expected to be written (last writer) but where dest field is 0/blank
    # Determine by checking dest vs simulation
    # Fetch dest values for each (dest_uuid, fld_uuid) in last_writer
    failed_count = 0
    clients_to_check = {dest_uuid for (dest_uuid, _) in last_writer.keys()}
    dest_field_values = {}  # (dest_uuid, fld_uuid) → amount_pence
    ls_map = {}  # dest_uuid → life_stage_uuid

    p(f"  Fetching dest expenditure for {len(clients_to_check)} clients to count failures…")
    n_checked = 0
    for dest_uuid in clients_to_check:
        r = dest.get(f"/api/v1/client/{dest_uuid}/expenditure")
        if r.status_code == 200:
            body      = r.json()
            exp_data  = body.get("data", body) if isinstance(body, dict) else body
            if isinstance(exp_data, dict):
                ls_uuid = body.get("current_life_stage_uuid") or exp_data.get("current_life_stage_uuid")
                ls_map[dest_uuid] = ls_uuid
                for sec in (exp_data.get("sections") or []):
                    sec_name = str(sec.get("name") or "")
                    for fld in (sec.get("fields") or []):
                        fld_uuid = fld.get("uuid") or fld.get("field_uuid") or ""
                        amt_raw  = fld.get("amount")
                        try:
                            amt_p = int(amt_raw) if amt_raw is not None else 0
                        except (TypeError, ValueError):
                            try:
                                amt_p = int(float(amt_raw) * 100)
                            except Exception:
                                amt_p = 0
                        if fld_uuid:
                            dest_field_values[(dest_uuid, fld_uuid)] = amt_p
        n_checked += 1
        if n_checked % 25 == 0:
            p(f"  …checked {n_checked}/{len(clients_to_check)}")

    # Count written rows where dest shows 0 (failed API call at time of import)
    for (dest_uuid, fld_uuid), row_i in last_writer.items():
        dest_val = dest_field_values.get((dest_uuid, fld_uuid), None)
        if dest_val == 0 or dest_val is None:
            if row_category.get(row_i) == "written":
                failed_count += 1
                row_category[row_i] = "failed"

    n_written_final = sum(1 for c in row_category.values() if c == "written")

    p(f"\n  Source rows total: {len(exp_rows)}")
    p(f"  Custom (section+type not in dest form):  {n_custom}")
    p(f"  Amount <= 0 (skipped):                   {n_lte0}")
    p(f"  No dest UUID (client not in client_map): {n_no_dest}")
    p(f"  Overwritten (later row same client+field): {n_overwritten}")
    p(f"  Written (last write, API succeeded):     {n_written_final}")
    p(f"  Failed (last write but dest shows 0):    {failed_count}")
    p(f"  Total accounted:  {n_custom + n_lte0 + n_no_dest + n_overwritten + n_written_final + failed_count}")

    # Show what the custom fields were
    custom_fields = set()
    for i, row in enumerate(exp_rows):
        if row_category.get(i) == "custom":
            sec = str(row.get("section") or "").strip()
            fld = str(row.get("type") or "").strip()
            custom_fields.add((sec, fld))
    p(f"\n  Custom field names ({len(custom_fields)}):")
    for sf in sorted(custom_fields):
        cnt = sum(1 for i, row in enumerate(exp_rows)
                  if row_category.get(i) == "custom"
                  and str(row.get("section") or "").strip() == sf[0]
                  and str(row.get("type") or "").strip() == sf[1])
        p(f"    {sf[0]} / {sf[1]}: {cnt} rows")


# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    p("=== RECONCILIATION 2026-10-29 ===\n")
    pat  = resolve_pat()
    dest = Dest(pat)

    client_map = load_json(os.path.join(IMPORT_DIR, "state", "client_map.json"))
    plan_map   = load_json(os.path.join(IMPORT_DIR, "state", "plan_map.json"))
    v2_rows    = load_csv(os.path.join(IMPORT_DIR, "W8_notes_final_v2.csv"))
    p(f"  client_map: {len(client_map)}, plan_map: {len(plan_map)}, v2: {len(v2_rows)} rows")

    all_dest_notes, ref_to_dest, no_ref_notes = reconcile_notes(dest, client_map, v2_rows)
    dep_recs = count_all_duplicates(dest, client_map, plan_map)
    list_dependant_pairs(dep_recs)
    analyze_expenditure(dest, client_map)

    p("\n" + "=" * 60)
    p("RECONCILIATION COMPLETE")
    p("=" * 60)


if __name__ == "__main__":
    main()
