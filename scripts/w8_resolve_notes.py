"""
W8 notes resolution — re-home 4,160 staff-account notes to migrated clients.

Steps:
  1. Setup check: branch state, file presence, Key Vault.
  2. Pre-flight: employee guard; count dest clients (356) and plans (1901);
     count dest notes — must be 0 or STOP.
  3. For every migrated source client (W1+W2), GET /api/v1/client/{uuid}/all-notes
     on SOURCE (paginated).  Build note_uuid → [src_client_uuid, ...].
  4. For each of the 4,160 unresolved rows, resolve or drop.
  5. Cross-check: count resolved rows where note map disagrees with current account_uuid.
  6. Write W8_notes_final_v2.csv.
  7. File checks and final report.

Rule 9: counts and UUIDs only in output — no client names, no note contents.
Source is GET-only.  No destination writes.
"""

import csv, json, os, sys, time
import requests

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# ── Constants ──────────────────────────────────────────────────────────────────

BASE_URL          = "https://api.plannrcrm.com"
KV_NAME           = "kv-plannr-ingestion"
SRC_PAT_SECRET    = "plannr-hfp-pat"
DEST_PAT_SECRET   = "plannr-prod-pat"
DEST_ACCOUNT_UUID = "ccd7f98b-d6a4-46ad-bab3-242ab42826eb"
MIGRATION_CO      = "data_migration"
IMPORT_DIR        = r"C:\dev\AR Migration\import_20261027"

EXPECTED_EMPLOYEES = {"jim hoban", "katie hoban", "tommie hoban", "valerie hoban"}

STAFF_UUIDS = {
    "db0ab7f5-8e52-4c89-a915-29fefeeca0d4",
    "6dbac8ff-90d2-4dc8-b60d-26567773a68e",
    "2153bb3c-7b76-4643-86a3-a9c53acab0b0",
    "5a8bc24f-8ff9-4671-ba49-4a5c8bfc5b2b",
    "bc00b121-9dcc-4283-bcb2-a75697fa249d",
    "2ce06c3f-6929-49f9-bfcf-946debc81209",
}

# ── Helpers ────────────────────────────────────────────────────────────────────

def p(*a):
    print(*a, flush=True)


def resolve_pats():
    from azure.identity import DefaultAzureCredential
    from azure.keyvault.secrets import SecretClient
    sc = SecretClient(
        vault_url=f"https://{KV_NAME}.vault.azure.net/",
        credential=DefaultAzureCredential(),
    )
    return sc.get_secret(SRC_PAT_SECRET).value, sc.get_secret(DEST_PAT_SECRET).value


# ── API clients ────────────────────────────────────────────────────────────────

class SourceClient:
    """GET-only; sets account UUID from /api/v1/me on init."""
    def __init__(self, pat):
        self._s = requests.Session()
        self._s.headers.update({
            "Authorization": f"Bearer {pat}",
            "Accept":        "application/json",
        })
        r     = self._s.get(f"{BASE_URL}/api/v1/me")
        inner = r.json().get("data", r.json())
        acct  = (inner.get("current_account") or {})
        uuid  = acct.get("uuid") or inner.get("uuid")
        self._s.headers["X-PLANNR-ACCOUNT-UUID"] = uuid
        p(f"  Source account UUID: {uuid}")

    def get(self, path, **kw):
        return self._s.get(f"{BASE_URL}{path}", **kw)


class DestClient:
    """GET-only destination client."""
    def __init__(self, pat):
        self._s = requests.Session()
        self._s.headers.update({
            "Authorization":          f"Bearer {pat}",
            "Accept":                 "application/json",
            "X-PLANNR-ACCOUNT-UUID":  DEST_ACCOUNT_UUID,
        })

    def get(self, path, **kw):
        return self._s.get(f"{BASE_URL}{path}", **kw)


# ── Pagination ─────────────────────────────────────────────────────────────────

def get_paged(client, path, max_pages=500, params=None):
    items, page = [], 1
    while page <= max_pages:
        pr = dict(params or {})
        pr.update({"page": page, "per_page": 200})
        r = client.get(path, params=pr)
        if r.status_code != 200:
            break
        body = r.json()
        data = body.get("data", [])
        if isinstance(data, dict):
            data = [data]
        items.extend(data)
        meta = body.get("meta", {}) if isinstance(body, dict) else {}
        if page >= (meta.get("last_page") or 1):
            break
        page += 1
    return items


# ── Data loaders ───────────────────────────────────────────────────────────────

def load_client_map():
    path = os.path.join(IMPORT_DIR, "state", "client_map.json")
    with open(path, encoding="utf-8") as f:
        return json.load(f)  # {src_uuid: dest_uuid}


def load_individual_and_joint_info():
    """Returns (individual_src: set, joint_to_owners: {joint_src: [owner_src, ...]})."""
    individual_src  = set()
    joint_to_owners = {}

    w1_path = os.path.join(IMPORT_DIR, "W1_clients.csv")
    with open(w1_path, newline="", encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            src = (r.get("external_reference_value") or "").strip()
            if src:
                individual_src.add(src)

    w2_path = os.path.join(IMPORT_DIR, "W2_clients.csv")
    with open(w2_path, newline="", encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            src  = (r.get("external_reference_value") or "").strip()
            kind = (r.get("type") or "").strip().lower()
            if not src:
                continue
            if kind == "individual":
                individual_src.add(src)
            elif kind == "joint":
                owners_raw = (r.get("owners") or "").strip()
                owners     = [o.strip() for o in owners_raw.split(",") if o.strip()]
                joint_to_owners[src] = owners

    return individual_src, joint_to_owners


# ── Resolution logic ───────────────────────────────────────────────────────────

def best_account_uuid(src_clients, individual_src, joint_to_owners, client_map):
    """Given a list of source client UUIDs, return the source UUID to store in
    account_uuid (the W8 file uses source UUIDs; the W8 runner rewrites to dest).

    Priority:
      1. First individual client in the list.
      2. First individual *owner* of the first joint client in the list.
      3. Fallback: first client in the list (trust/business).

    Returns (src_uuid_to_use, chosen_was_joint: bool).
    """
    # Pass 1: first individual
    for sc in src_clients:
        if sc in individual_src:
            return sc, False

    # Pass 2: first joint → first owner in client_map
    for sc in src_clients:
        if sc in joint_to_owners:
            for owner in joint_to_owners[sc]:
                if owner in client_map:
                    return owner, True

    # Pass 3: fallback (trust/business/unknown)
    first = src_clients[0]
    if first in joint_to_owners:
        for owner in joint_to_owners[first]:
            if owner in client_map:
                return owner, True
    return first, False


# ── Main ───────────────────────────────────────────────────────────────────────

def main():

    # ── Step 1: setup check ────────────────────────────────────────────────────
    p("\n=== SETUP CHECK ===")
    required_files = [
        "W8_notes_final.csv",
        "state/wave_state.json",
        "state/client_map.json",
        "state/plan_map.json",
        "state/w1_client_uuid_map.json",
    ]
    all_present = True
    for fn in required_files:
        present = os.path.exists(os.path.join(IMPORT_DIR, fn))
        p(f"  {'OK' if present else 'MISSING'}: {fn}")
        if not present:
            all_present = False
    if not all_present:
        p("ABORT: required files missing")
        return

    p("\n  Resolving Key Vault PATs...")
    try:
        src_pat, dest_pat = resolve_pats()
        p(f"  KV: OK — src len={len(src_pat)}, dest len={len(dest_pat)}")
    except Exception as exc:
        p(f"  KV FAILED: {exc}")
        return

    # ── Step 2: pre-flight ─────────────────────────────────────────────────────
    p("\n=== PRE-FLIGHT ===")
    dest = DestClient(dest_pat)

    # Employee guard — confirm we are on the right destination
    r0 = dest.get("/api/v1/employee", params={"per_page": 50})
    if r0.status_code != 200:
        p(f"ABORT: GET /api/v1/employee HTTP {r0.status_code}")
        return
    emps  = r0.json().get("data", [])
    names = {
        ((e.get("first_name") or "") + " " + (e.get("last_name") or "")).strip().lower()
        for e in emps
    }
    if not EXPECTED_EMPLOYEES.issubset(names):
        p(f"ABORT: employee guard failed — found {names}")
        return
    p(f"  Employee guard: PASS ({len(emps)} employees)")

    # Count clients (all paginated; only count data_migration ext-ref)
    p("  Counting dest clients...")
    all_clients = get_paged(dest, "/api/v1/client", max_pages=10,
                            params={"include": "external_references"})
    migrated_clients = [
        c for c in all_clients
        if any(ref.get("third_party") == MIGRATION_CO
               for ref in (c.get("external_references") or []))
    ]
    p(f"  Migrated clients: {len(migrated_clients)} (expected 356)")
    if len(migrated_clients) != 356:
        p(f"STOP — client count mismatch")
        return

    # Count plans
    p("  Counting dest plans...")
    all_plans = get_paged(dest, "/api/v1/plans", max_pages=20,
                          params={"include": "external_references"})
    migrated_plans = [
        pl for pl in all_plans
        if any(ref.get("third_party") == MIGRATION_CO
               for ref in (pl.get("external_references") or []))
    ]
    p(f"  Migrated plans: {len(migrated_plans)} (expected 1901)")
    if len(migrated_plans) != 1901:
        p(f"STOP — plan count mismatch")
        return

    # Count destination notes — must be 0
    p("  Counting dest notes...")
    notes_r = dest.get("/api/v1/notes", params={"per_page": 1, "page": 1})
    if notes_r.status_code == 200:
        body        = notes_r.json()
        meta        = body.get("meta", {}) if isinstance(body, dict) else {}
        total_notes = meta.get("total", 0)
        if total_notes == 0:
            # Also check data length in case meta is missing
            data_len = len(body.get("data", body if isinstance(body, list) else []))
            total_notes = max(total_notes, data_len)
        p(f"  Dest notes (via /api/v1/notes): {total_notes}")
    else:
        # Fallback: check two spot-check clients for any notes
        p(f"  GET /api/v1/notes HTTP {notes_r.status_code} — falling back to spot checks")
        spot_uuids = [
            "28cfd18c-89ca-421e-bfff-47e4db249198",
            "a598e1a7-151a-4cb6-a774-caaf918ef7b5",
        ]
        total_notes = 0
        for su in spot_uuids:
            sr = dest.get(f"/api/v1/client/{su}/all-notes", params={"per_page": 1})
            if sr.status_code == 200:
                bd = sr.json()
                data = bd.get("data", bd if isinstance(bd, list) else [])
                total_notes += len(data) if isinstance(data, list) else 0
        p(f"  Dest notes (spot-check {len(spot_uuids)} clients): {total_notes}")

    if total_notes != 0:
        p(f"STOP — destination has {total_notes} notes; cannot proceed until it is 0")
        return
    p("  Notes check: PASS (0)")

    # ── Step 3: load maps ──────────────────────────────────────────────────────
    p("\n=== LOADING MAPS ===")
    client_map = load_client_map()
    p(f"  client_map: {len(client_map)} entries")
    individual_src, joint_to_owners = load_individual_and_joint_info()
    p(f"  individual source clients: {len(individual_src)}")
    p(f"  joint source clients:      {len(joint_to_owners)}")

    # ── Step 3 (cont.): fetch all-notes from source ────────────────────────────
    p("\n=== FETCHING SOURCE ALL-NOTES ===")
    src            = SourceClient(src_pat)
    # {note_uuid: [src_client_uuid, ...]} — use sets to deduplicate per client
    note_to_clients_sets = {}  # {note_uuid: set of src_uuid}

    src_uuids    = list(client_map.keys())
    n_clients    = len(src_uuids)
    appearances  = 0

    for i, src_uuid in enumerate(src_uuids, 1):
        notes = get_paged(src, f"/api/v1/client/{src_uuid}/all-notes", max_pages=500)
        for note in notes:
            nuid = str(note.get("uuid") or "").strip()
            if nuid:
                note_to_clients_sets.setdefault(nuid, set()).add(src_uuid)
        appearances += len(notes)
        if i % 50 == 0 or i == n_clients:
            p(f"  {i}/{n_clients} clients — {appearances} note appearances, "
              f"{len(note_to_clients_sets)} unique note UUIDs so far")

    # Convert sets to lists (preserve insertion order via sorted for determinism)
    note_to_clients = {k: sorted(v, key=lambda u: src_uuids.index(u))
                       for k, v in note_to_clients_sets.items()}

    multi_note_count = sum(1 for v in note_to_clients_sets.values() if len(v) > 1)
    p(f"\n  Total appearances: {appearances}")
    p(f"  Unique note UUIDs: {len(note_to_clients)}")
    p(f"  Notes under >1 client: {multi_note_count}")

    # ── Step 4: process the 4,160 unresolved rows ──────────────────────────────
    p("\n=== PROCESSING UNRESOLVED ROWS ===")
    notes_path = os.path.join(IMPORT_DIR, "W8_notes_final.csv")
    with open(notes_path, newline="", encoding="utf-8-sig") as f:
        reader     = csv.DictReader(f)
        fieldnames = list(reader.fieldnames)
        all_rows   = list(reader)

    out_rows       = []
    rehomed        = 0
    multi_client   = 0
    dropped        = 0
    dropped_by_mt  = {}
    dropped_by_staff = {}

    for row in all_rows:
        acct = (row.get("account_uuid") or "").strip()

        if acct not in STAFF_UUIDS:
            out_rows.append(row)
            continue

        note_uuid   = (row.get("external_reference_value") or "").strip()
        src_clients = note_to_clients.get(note_uuid, [])

        if not src_clients:
            # Under no migrated client — drop
            dropped += 1
            mt = (row.get("model_type") or "").strip()
            dropped_by_mt[mt]    = dropped_by_mt.get(mt, 0) + 1
            dropped_by_staff[acct] = dropped_by_staff.get(acct, 0) + 1
            continue

        is_multi = len(src_clients) > 1
        if is_multi:
            multi_client += 1

        chosen_src, _ = best_account_uuid(src_clients, individual_src, joint_to_owners, client_map)

        new_row = dict(row)
        new_row["account_uuid"] = chosen_src
        out_rows.append(new_row)
        rehomed += 1

    p(f"  Re-homed:    {rehomed}")
    p(f"  Multi-client (used first individual): {multi_client}")
    p(f"  Dropped:     {dropped}")

    # ── Step 5: cross-check ────────────────────────────────────────────────────
    p("\n=== CROSS-CHECK (resolved rows only) ===")
    disagree = 0
    for row in all_rows:
        acct = (row.get("account_uuid") or "").strip()
        if acct in STAFF_UUIDS:
            continue
        note_uuid   = (row.get("external_reference_value") or "").strip()
        src_clients = note_to_clients.get(note_uuid, [])
        if src_clients and acct not in src_clients:
            disagree += 1
    p(f"  Disagreements: {disagree}")

    # ── Step 6: write W8_notes_final_v2.csv ───────────────────────────────────
    p("\n=== WRITING W8_notes_final_v2.csv ===")
    v2_path = os.path.join(IMPORT_DIR, "W8_notes_final_v2.csv")
    with open(v2_path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
        w.writeheader()
        w.writerows(out_rows)
    p(f"  Written: {v2_path}  ({len(out_rows)} rows)")

    # ── Step 7: file checks ────────────────────────────────────────────────────
    p("\n=== V2 FILE CHECKS ===")
    expected_total = 6795 - dropped

    unresolved_v2 = [
        r for r in out_rows
        if (r.get("account_uuid") or "").strip() in STAFF_UUIDS
    ]
    joint_remaining = [
        r for r in out_rows
        if (r.get("account_uuid") or "").strip() in joint_to_owners
    ]
    visible_yes = [
        r for r in out_rows
        if (r.get("visible_to_clients") or "").strip().lower() == "yes"
    ]

    total_ok   = len(out_rows) == expected_total
    unresolv_ok = len(unresolved_v2) == 0
    joint_ok   = len(joint_remaining) == 0
    vis_ok     = len(visible_yes) == 0

    p(f"  Total rows:              {len(out_rows)} "
      f"(expected 6795 - {dropped} = {expected_total}) "
      f"{'OK' if total_ok else 'FAIL'}")
    p(f"  Unresolved UUIDs:        {len(unresolved_v2)} "
      f"{'OK' if unresolv_ok else 'FAIL'}")
    p(f"  Rows on joint account:   {len(joint_remaining)} "
      f"{'OK' if joint_ok else 'FAIL'}")
    p(f"  Rows visible_to_clients: {len(visible_yes)} "
      f"{'OK' if vis_ok else 'FAIL'}")

    # ── Final report ───────────────────────────────────────────────────────────
    p("\n" + "=" * 60)
    p("W8 NOTES RESOLUTION REPORT")
    p("=" * 60)
    p(f"\n  Re-homed:                  {rehomed}")
    p(f"  Multi-client (1st indiv):  {multi_client}")
    p(f"  Dropped:                   {dropped}")

    if dropped_by_mt:
        p(f"\n  Dropped by model_type:")
        for mt, cnt in sorted(dropped_by_mt.items(), key=lambda x: -x[1]):
            p(f"    {mt}: {cnt}")

    if dropped_by_staff:
        p(f"\n  Dropped by staff account UUID:")
        for su, cnt in sorted(dropped_by_staff.items(), key=lambda x: -x[1]):
            p(f"    {su}: {cnt}")

    p(f"\n  Cross-check disagreements: {disagree}")

    p(f"\n  File checks (W8_notes_final_v2.csv):")
    p(f"    Total rows:        {len(out_rows)} / {expected_total}  {'OK' if total_ok else 'FAIL'}")
    p(f"    Unresolved:        {len(unresolved_v2)}  {'OK' if unresolv_ok else 'FAIL'}")
    p(f"    Joint accounts:    {len(joint_remaining)}  {'OK' if joint_ok else 'FAIL'}")
    p(f"    Visible to client: {len(visible_yes)}  {'OK' if vis_ok else 'FAIL'}")

    p(f"\n  Spot-check client UUIDs (destination):")
    spot_clients = [
        "28cfd18c-89ca-421e-bfff-47e4db249198",
        "06550252-391b-40c0-834f-f6f9bca8f65a",
        "26697c8a-e94b-4eeb-b021-d338793e63f7",
        "a598e1a7-151a-4cb6-a774-caaf918ef7b5",
        "a20b5b1c-dfb9-4cf3-bd06-eb93c45e1341",
    ]
    for dest_uuid in spot_clients:
        p(f"    {dest_uuid}")

    all_ok = total_ok and unresolv_ok and joint_ok and vis_ok
    p(f"\n  Overall: {'ALL CHECKS PASS' if all_ok else 'CHECKS FAILED — see above'}")
    p("=" * 60)


if __name__ == "__main__":
    main()
