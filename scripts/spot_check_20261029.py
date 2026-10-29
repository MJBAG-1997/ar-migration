"""
Spot-check: expected (import CSVs) vs actual (destination API) for 6 clients.
Rule 9: UUIDs, types, providers, statuses, values and counts only — no client names,
plan names or policy numbers in terminal output.

Six clients:
  C1  active W1 individual, ongoing charge on agency plan
  C2  individual from couple linked by a W2 circle (partner also shown)
  C3  Add-to-Plannr plan (prov- ext_ref) with W7 charge
  C4  client with income, expenditure and dependants
  C5  client with most notes in W8_notes_final_v2.csv
  C6  Rathbones plan + W7 charge

Explain section:
  X1  28cfd18c-89ca-421e-bfff-47e4db249198 — joint record: plans and expected charges
  X2  a20b5b1c-dfb9-4cf3-bd06-eb93c45e1341 — Add-to-Plannr plan, W7 charge status

READ-ONLY: no destination writes.
"""

import csv, json, os, sys, time

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

import requests

BASE_URL          = "https://api.plannrcrm.com"
KV_NAME           = "kv-plannr-ingestion"
DEST_PAT_SECRET   = "plannr-prod-pat"
DEST_ACCOUNT_UUID = "ccd7f98b-d6a4-46ad-bab3-242ab42826eb"
MIGRATION_CO      = "data_migration"
IMPORT_DIR        = r"C:\dev\AR Migration\import_20261027"
EXPORTS_DIR       = r"C:\dev\AR Migration\exports_20261026"

# ── Six clients (destination UUIDs) ───────────────────────────────────────────
CLIENTS = {
    "C1": "587bbc91-6c99-4114-a6c0-8416c3a68cb6",   # W1 indiv, agency plan + ongoing charge
    "C2": "ed17a935-feb4-4011-b7db-3b9575932acc",   # couple circle (partner f3890327)
    "C3": "64d5e606-e46b-4880-9ce5-2f47da8335bf",   # Add-to-Plannr plan with W7 charge
    "C4": "c80da0e1-f368-4d66-b97a-6d88bd27e736",   # income + expenditure + dependants
    "C5": "c5052bc5-1a6c-4857-bf81-bce8603bb8a9",   # most notes in W8 v2 (174)
    "C6": "26697c8a-e94b-4eeb-b021-d338793e63f7",   # Rathbones plan + 0.5% charge
}
C2_PARTNER = "f3890327-a428-4a29-9f08-60408e199371"
C2_CIRCLE  = "53b93c18-2f3d-4933-aa62-bc7dae07929d"

# Explain UUIDs
X1 = "28cfd18c-89ca-421e-bfff-47e4db249198"   # joint record
X2 = "a20b5b1c-dfb9-4cf3-bd06-eb93c45e1341"   # Add-to-Plannr client
W4_SRC_PLAN = "0341093d-9821-44e5-9f25-146adbbefa70"   # W4 Transact GIA plan

def p(*a): print(*a, flush=True)
def flag(ok, msg=""):
    return ("    " if ok else "  !! ") + msg


# ── API client ─────────────────────────────────────────────────────────────────

def resolve_dest_pat():
    from azure.identity import DefaultAzureCredential
    from azure.keyvault.secrets import SecretClient
    sc = SecretClient(
        vault_url=f"https://{KV_NAME}.vault.azure.net/",
        credential=DefaultAzureCredential(),
    )
    return sc.get_secret(DEST_PAT_SECRET).value


class Dest:
    def __init__(self, pat):
        self._s = requests.Session()
        self._s.headers.update({
            "Authorization":         f"Bearer {pat}",
            "Accept":                "application/json",
            "X-PLANNR-ACCOUNT-UUID": DEST_ACCOUNT_UUID,
        })

    def get(self, path, **kw):
        return self._s.get(f"{BASE_URL}{path}", **kw)


def get_paged(dest, path, max_pages=100, params=None):
    items, page = [], 1
    while page <= max_pages:
        pr = dict(params or {}); pr.update({"page": page, "per_page": 100})
        r = dest.get(path, params=pr)
        if r.status_code != 200:
            break
        body = r.json()
        data = body.get("data", [])
        if isinstance(data, dict): data = [data]
        items.extend(data)
        meta = body.get("meta", {}) if isinstance(body, dict) else {}
        if page >= (meta.get("last_page") or 1):
            break
        page += 1
    return items


def get_one(dest, path):
    r = dest.get(path)
    if r.status_code != 200:
        return None
    body = r.json()
    return body.get("data", body) if isinstance(body, dict) else body


# ── Load source data ────────────────────────────────────────────────────────────

def load_json(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)

def load_csv(path):
    with open(path, newline="", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def load_all_source():
    cm    = load_json(os.path.join(IMPORT_DIR, "state", "client_map.json"))   # src→dest
    pm    = load_json(os.path.join(IMPORT_DIR, "state", "plan_map.json"))      # src→dest

    # Reverse maps
    dest_to_src_client = {v: k for k, v in cm.items()}
    dest_to_src_plan   = {v: k for k, v in pm.items()}

    w1    = load_csv(os.path.join(IMPORT_DIR, "W1_clients.csv"))
    w2    = load_csv(os.path.join(IMPORT_DIR, "W2_clients.csv"))
    w4    = load_csv(os.path.join(IMPORT_DIR, "W4_plans_rewritten.csv"))
    w5    = load_csv(os.path.join(IMPORT_DIR, "W5_plans_rewritten_fixed.csv"))
    w7    = load_csv(os.path.join(IMPORT_DIR, "W7_charges_rewritten.csv"))
    w6v   = load_csv(os.path.join(IMPORT_DIR, "W6_valuations_rewritten.csv"))
    w6p   = load_csv(os.path.join(IMPORT_DIR, "W6_premiums.csv"))
    w3r   = load_csv(os.path.join(IMPORT_DIR, "W3_relationships_rewritten.csv"))
    w3d   = load_csv(os.path.join(IMPORT_DIR, "W3_dependants_rewritten.csv"))
    w3i   = load_csv(os.path.join(IMPORT_DIR, "W3_income_rewritten.csv"))
    exp   = load_csv(os.path.join(EXPORTS_DIR, "Client Expenditure Export HFP Migration 2026-10-27.csv"))
    v2    = load_csv(os.path.join(IMPORT_DIR, "W8_notes_final_v2.csv"))

    # Index by dest client UUID or plan UUID
    def first_client(plan_row):
        return (plan_row.get("client_uuids") or "").split(",")[0].strip()

    # Plans per dest client UUID (W4 + W5) — index by ALL clients in client_uuids
    plans_by_dest_client = {}
    for plan in (w4 + w5):
        all_clients = [(c.strip()) for c in (plan.get("client_uuids") or "").split(",") if c.strip()]
        for dc in all_clients:
            plans_by_dest_client.setdefault(dc, []).append(plan)

    # W7 charges by dest plan UUID
    charges_by_dest_plan = {}
    for ch in w7:
        charges_by_dest_plan.setdefault(ch["plan_uuid"], []).append(ch)

    # W6 valuations by dest plan UUID
    vals_by_dest_plan = {}
    for v in w6v:
        vals_by_dest_plan.setdefault(v["plan_uuid"], []).append(v)

    # W6 premiums by dest plan UUID
    prems_by_dest_plan = {}
    for pr in w6p:
        dp = pm.get(pr.get("plan_uuid") or "")
        if dp:
            prems_by_dest_plan.setdefault(dp, []).append(pr)

    # W3 relationships by dest client UUID (rewritten → dest UUIDs)
    rels_by_dest_client = {}
    for r in w3r:
        dc = r.get("account_uuid") or ""
        rels_by_dest_client.setdefault(dc, []).append(r)

    # W3 dependants by dest client UUID (rewritten → dest UUIDs)
    deps_by_dest_client = {}
    for r in w3d:
        for uid in (r.get("account_uuids") or "").split(","):
            uid = uid.strip()
            if uid:
                deps_by_dest_client.setdefault(uid, []).append(r)

    # W3 income by dest client UUID (rewritten → dest UUIDs)
    income_by_dest_client = {}
    for r in w3i:
        dc = r.get("client_uuid") or ""
        income_by_dest_client.setdefault(dc, []).append(r)

    # Expenditure by source client UUID (not rewritten)
    exp_by_src_client = {}
    for r in exp:
        sc = r.get("client_uuid") or ""
        exp_by_src_client.setdefault(sc, []).append(r)

    # Notes in W8 v2 by source account_uuid
    notes_v2_by_src = {}
    for r in v2:
        acct = (r.get("account_uuid") or "").strip()
        notes_v2_by_src[acct] = notes_v2_by_src.get(acct, 0) + 1

    # Source client lookup by ext_ref (for W1/W2 rows)
    w1_by_src = {r["external_reference_value"]: r for r in w1 if r.get("external_reference_value")}
    w2_by_src = {r["external_reference_value"]: r for r in w2 if r.get("external_reference_value")}

    return {
        "cm": cm, "pm": pm,
        "dest_to_src_client": dest_to_src_client,
        "dest_to_src_plan":   dest_to_src_plan,
        "plans_by_dc":        plans_by_dest_client,
        "charges_by_dp":      charges_by_dest_plan,
        "vals_by_dp":         vals_by_dest_plan,
        "prems_by_dp":        prems_by_dest_plan,
        "rels_by_dc":         rels_by_dest_client,
        "deps_by_dc":         deps_by_dest_client,
        "income_by_dc":       income_by_dest_client,
        "exp_by_sc":          exp_by_src_client,
        "notes_v2_by_src":    notes_v2_by_src,
        "w1_by_src":          w1_by_src,
        "w2_by_src":          w2_by_src,
    }


# ── Fetch destination data per client ─────────────────────────────────────────

def fetch_client_data(dest, dest_uuid):
    """Returns dict with all destination data for one client."""
    d = {}
    client = get_one(dest, f"/api/v1/client/{dest_uuid}")
    d["client"] = client

    plans = get_paged(dest, f"/api/v1/client/{dest_uuid}/plans")
    d["plans"] = plans

    plan_details = {}
    for pl in plans:
        puuid = pl.get("uuid")
        if not puuid:
            continue
        detail = get_one(dest, f"/api/v1/plans/{puuid}")
        charges = get_paged(dest, f"/api/v1/plans/{puuid}/charges")
        vals = get_paged(dest, f"/api/v1/plans/{puuid}/valuations")
        plan_details[puuid] = {"detail": detail, "charges": charges, "valuations": vals}
    d["plan_details"] = plan_details

    # Count endpoints
    rels  = get_paged(dest, f"/api/v1/client/{dest_uuid}/linked-clients")
    deps  = get_paged(dest, f"/api/v1/client/{dest_uuid}/dependants")
    inc   = get_paged(dest, f"/api/v1/client/{dest_uuid}/income")

    # Expenditure: try API
    exp_r = dest.get(f"/api/v1/client/{dest_uuid}/expenditure")
    if exp_r.status_code == 200:
        exp_data = exp_r.json()
        # expenditure returns nested structure; count line items
        exp_items = exp_data.get("data", exp_data) if isinstance(exp_data, dict) else exp_data
        if isinstance(exp_items, dict):
            # May be {sections: [{fields: [{amount...}]}]}
            exp_count = sum(
                1 for sec in (exp_items.get("sections") or [])
                for fld in (sec.get("fields") or [])
                if fld.get("amount") and fld["amount"] != 0
            )
        else:
            exp_count = len(exp_items) if isinstance(exp_items, list) else 0
    else:
        exp_count = -1

    # Premiums: via plans
    prem_count = 0
    for puuid, pd in plan_details.items():
        pr = get_paged(dest, f"/api/v1/plans/{puuid}/premiums")
        prem_count += len(pr)

    d["rels"]        = rels
    d["deps"]        = deps
    d["income"]      = inc
    d["exp_count"]   = exp_count
    d["prem_count"]  = prem_count
    return d


# ── Format helpers ─────────────────────────────────────────────────────────────

def fmt_val(v):
    if v is None: return "None"
    try: return f"£{float(v):,.2f}"
    except: return str(v)

def latest_val(vals):
    """Latest valuation from list [{valuation_date, value}]."""
    if not vals: return None, None
    sv = sorted(vals, key=lambda x: str(x.get("valuation_date") or ""), reverse=True)
    return sv[0].get("value"), sv[0].get("valuation_date")

def src_latest_val(src_vals):
    """Latest from W6 CSV rows."""
    if not src_vals: return None, None
    sv = sorted(src_vals, key=lambda x: str(x.get("valuation_date") or ""), reverse=True)
    return sv[0].get("value"), sv[0].get("valuation_date")

def match(a, b, label):
    if str(a or "").strip().lower() == str(b or "").strip().lower():
        return f"    {label}: {a} ✓"
    return f"  !! {label}: expected={a} actual={b}"

def match_count(exp, act, label):
    if exp == act:
        return f"    {label}: {act} ✓"
    return f"  !! {label}: expected={exp} actual={act}"


# ── Print one client block ─────────────────────────────────────────────────────

def print_client_block(label, dest_uuid, src, dest_data):
    p(f"\n{'─'*60}")
    p(f"{label}: {dest_uuid}")
    src_uuid = src.get("dest_to_src_client", {}).get(dest_uuid, "?")
    p(f"  source UUID: {src_uuid}")

    # Status and adviser
    client = dest_data.get("client") or {}
    src_row = (src["w1_by_src"].get(src_uuid) or src["w2_by_src"].get(src_uuid) or {})

    exp_status_uuid = src_row.get("status_uuid") or "ebbdb871-c154-49a0-a212-b8dbe1874c66"
    act_status      = client.get("status") or {}
    act_status_uuid = act_status.get("uuid") or act_status.get("id") or str(act_status)

    exp_adviser = src_row.get("assigned_adviser_uuid") or ""
    act_adviser = (client.get("assigned_adviser") or {}).get("uuid") or ""
    if not act_adviser:
        act_adviser = client.get("assigned_adviser_uuid") or ""

    p(f"\n  STATUS & ADVISER:")
    p(match(exp_status_uuid[:8], act_status_uuid[:8] if act_status_uuid else "?",
            f"status_uuid (first 8)"))
    p(match(exp_adviser[:8] if exp_adviser else "", act_adviser[:8] if act_adviser else "",
            "adviser_uuid (first 8)"))

    # Plans
    exp_plans_raw = src.get("plans_by_dc", {}).get(dest_uuid, [])
    act_plans = dest_data.get("plans", [])
    plan_details = dest_data.get("plan_details", {})

    # Deduplicate exp_plans by external_reference_value (plan listed once per client)
    seen_plan_srcs = set()
    exp_plans = []
    for ep in exp_plans_raw:
        k = ep.get("external_reference_value") or id(ep)
        if k not in seen_plan_srcs:
            seen_plan_srcs.add(k); exp_plans.append(ep)

    p(f"\n  PLANS: expected={len(exp_plans)} actual={len(act_plans)}")
    if len(exp_plans) != len(act_plans):
        p(f"  !! plan count mismatch")

    for ep in exp_plans:
        ep_src_uuid  = ep.get("external_reference_value") or ""
        ep_dest_uuid = src["pm"].get(ep_src_uuid) or ep.get("uuid") or ""
        ep_type      = ep.get("type") or ""
        ep_sub       = ep.get("sub_type") or ""
        ep_prov      = ep.get("provider") or ""
        ep_status    = ep.get("status") or "active"
        ep_under_adv = ep.get("under_advice") or ""
        ep_agency    = ep.get("agency_in_your_name") or ""

        # Latest expected valuation
        ep_vals      = src["vals_by_dp"].get(ep_dest_uuid, [])
        ep_val, ep_vdate = src_latest_val(ep_vals)

        # Expected charges
        ep_charges   = src["charges_by_dp"].get(ep_dest_uuid, [])

        # Actual plan
        pd = plan_details.get(ep_dest_uuid, {})
        ap = pd.get("detail") or {}
        act_type   = ap.get("type") or ""
        act_sub    = ap.get("sub_type") or ""
        prov_raw   = ap.get("provider") or ""
        act_prov   = prov_raw.get("name") if isinstance(prov_raw, dict) else str(prov_raw)
        stat_raw   = ap.get("status") or ""
        act_status = (stat_raw.get("name") if isinstance(stat_raw, dict) else str(stat_raw)).lower()
        act_ua     = str(ap.get("under_advice") or "")
        act_agency = str(ap.get("agency_in_your_name") or "")

        # Latest actual valuation
        act_vals = pd.get("valuations", [])
        av, avdate = latest_val(act_vals)

        # Actual charges
        act_charges = pd.get("charges", [])

        p(f"\n    plan {ep_dest_uuid[:8]}… (src {ep_src_uuid[:8]}…)")
        p(match(ep_type,   act_type,   "type"))
        p(match(ep_sub,    act_sub,    "sub_type"))
        p(match(ep_prov,   act_prov,   "provider"))
        p(match(ep_status, (act_status or "").lower(), "status"))
        p(match(str(ep_under_adv).lower(), str(act_ua).lower(), "under_advice"))
        p(match(str(ep_agency).lower(), str(act_agency).lower(), "agency_in_your_name"))

        # Valuation
        if ep_val:
            p(match(str(ep_val), str(av or ""), f"latest_value ({ep_vdate})"))
        else:
            p(f"    latest_value: act={fmt_val(av)} on {avdate} (no W6 import expected)")

        # Charges
        if ep_charges:
            for ec in ep_charges:
                pct  = ec.get("percentage") or ""
                amt  = ec.get("amount") or ""
                sd   = ec.get("start_date") or ""
                rate_str = f"{pct}%" if pct else f"£{amt}" if amt else "?"
                p(f"    charge expected: {rate_str} from {sd}")
                # Find matching actual charge
                found = [ac for ac in act_charges
                         if (str(ac.get("percentage") or "") == str(pct)
                             or str(ac.get("amount") or "") == str(amt))]
                if found:
                    ac = found[0]
                    a_pct = ac.get("percentage") or ""; a_amt = ac.get("amount") or ""
                    a_sd  = ac.get("start_date") or ""
                    a_rate = f"{a_pct}%" if a_pct else f"£{a_amt}" if a_amt else "?"
                    ok = (str(pct) == str(a_pct) and str(sd) == str(a_sd))
                    p(f"    {'✓' if ok else '!!'} charge actual:   {a_rate} from {a_sd}")
                else:
                    p(f"    !! charge actual: NOT FOUND (actual charges={len(act_charges)})")
        else:
            if act_charges:
                p(f"    !! no charge expected but actual has {len(act_charges)}")
            else:
                p(f"    charge: none expected, none actual ✓")

    # Counts
    exp_rels  = len(src["rels_by_dc"].get(dest_uuid, []))
    act_rels  = len(dest_data.get("rels", []))
    exp_deps  = len(src["deps_by_dc"].get(dest_uuid, []))
    act_deps  = len(dest_data.get("deps", []))
    exp_inc   = len(src["income_by_dc"].get(dest_uuid, []))
    act_inc   = len(dest_data.get("income", []))
    exp_prems = sum(len(src["prems_by_dp"].get(ep_dest_uuid, []))
                    for ep in exp_plans
                    for ep_dest_uuid in [src["pm"].get(ep.get("external_reference_value") or "") or ""])
    act_prems = dest_data.get("prem_count", 0)

    # Expenditure: count from source export by source UUID
    exp_exp_lines = len(src["exp_by_sc"].get(src_uuid, []))
    act_exp_count = dest_data.get("exp_count", -1)

    # Notes in W8 v2 (by source UUID)
    notes_exp = src["notes_v2_by_src"].get(src_uuid, 0)

    p(f"\n  COUNTS:")
    p(match_count(exp_rels,  act_rels,  "relationships"))
    p(match_count(exp_deps,  act_deps,  "dependants"))
    p(match_count(exp_inc,   act_inc,   "income_lines"))
    p(match_count(exp_prems, act_prems, "premiums"))
    if act_exp_count < 0:
        p(f"    expenditure_lines: exp={exp_exp_lines} act=API_ERROR")
    else:
        p(match_count(exp_exp_lines, act_exp_count, "expenditure_lines"))
    p(f"    notes_in_W8_v2: {notes_exp} (dest shows 0 until W8 runs)")


# ── Explain X1 and X2 ─────────────────────────────────────────────────────────

def print_explain(x_label, dest_uuid, src, dest_data, extra=None):
    p(f"\n{'─'*60}")
    p(f"EXPLAIN {x_label}: {dest_uuid}")
    src_uuid = src["dest_to_src_client"].get(dest_uuid, "?")
    p(f"  source UUID: {src_uuid}")

    exp_plans = src["plans_by_dc"].get(dest_uuid, [])
    act_plans = dest_data.get("plans", [])
    plan_details = dest_data.get("plan_details", {})

    p(f"\n  Plans owned (expected {len(exp_plans)}, actual {len(act_plans)}):")
    for ep in exp_plans:
        ep_src  = ep.get("external_reference_value") or ""
        ep_dest = src["pm"].get(ep_src) or ""
        ptype   = ep.get("type") or ""
        psub    = ep.get("sub_type") or ""
        pprov   = ep.get("provider") or ""
        pstat   = ep.get("status") or "active"
        pagency = ep.get("agency_in_your_name") or ""

        charges = src["charges_by_dp"].get(ep_dest, [])
        ch_str  = f"{len(charges)} W7 charge(s)" if charges else "no W7 charge"

        is_atp = ep_src.startswith("prov-") or ep_src == W4_SRC_PLAN
        atp_mark = " [Add-to-Plannr]" if is_atp else ""

        p(f"    plan_src={ep_src[:8]}… dest={ep_dest[:8] if ep_dest else '?'}…"
          f"  type={ptype}/{psub}  provider={pprov}  agency={pagency}"
          f"  {ch_str}{atp_mark}")

    if extra:
        p(f"\n  {extra}")


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    p("=== SPOT-CHECK 2026-10-29 ===\n")
    p("Loading source data...")
    src = load_all_source()
    p(f"  client_map:  {len(src['cm'])} entries")
    p(f"  plan_map:    {len(src['pm'])} entries")

    p("\nConnecting to destination...")
    pat  = resolve_dest_pat()
    dest = Dest(pat)

    # Fetch all client data
    all_dest_data = {}
    all_uuids = list(CLIENTS.values()) + [C2_PARTNER, X1, X2]
    p(f"Fetching destination data for {len(all_uuids)} clients...")
    for i, du in enumerate(all_uuids, 1):
        p(f"  {i}/{len(all_uuids)}: {du[:8]}…")
        all_dest_data[du] = fetch_client_data(dest, du)

    # ── Print 6 client blocks ──────────────────────────────────────────────────
    p("\n" + "=" * 60)
    p("SPOT-CHECK RESULTS")
    p("=" * 60)

    for label, dest_uuid in CLIENTS.items():
        print_client_block(label, dest_uuid, src, all_dest_data[dest_uuid])

    # C2 partner also shown
    p(f"\n{'─'*60}")
    p(f"C2 PARTNER: {C2_PARTNER}")
    p(f"  circle ext_ref: {C2_CIRCLE}")
    c2p_src = src["dest_to_src_client"].get(C2_PARTNER, "?")
    p(f"  partner source UUID: {c2p_src}")
    c2p_plans = src["plans_by_dc"].get(C2_PARTNER, [])
    p(f"  partner plans in W4+W5: {len(c2p_plans)}")

    # ── Explain X1 and X2 ──────────────────────────────────────────────────────
    p("\n" + "=" * 60)
    p("EXPLAIN SECTION")
    p("=" * 60)

    # X1: joint record
    x1_extra = ("Joint record (type=joint). Plans listed above are all plans "
                "where client_uuids includes this dest UUID. W7 charges that "
                "had client_uuid=joint were rewritten to the joint's dest UUID; "
                "Plannr may attach them to the joint record.")
    print_explain("X1 (joint)", X1, src, all_dest_data[X1], extra=x1_extra)

    # X2: Add-to-Plannr client
    x2_atp_plans = [ep for ep in src["plans_by_dc"].get(X2, [])
                    if (ep.get("external_reference_value") or "").startswith("prov-")
                    or (ep.get("external_reference_value") or "") == W4_SRC_PLAN]
    x2_w4_plan = next((ep for ep in src["plans_by_dc"].get(X2, [])
                       if (ep.get("external_reference_value") or "") == W4_SRC_PLAN), None)
    if x2_w4_plan:
        w4_dest = src["pm"].get(W4_SRC_PLAN)
        w4_charges = src["charges_by_dp"].get(w4_dest or "", [])
        if w4_charges:
            extra_msg = f"W4 plan ({W4_SRC_PLAN[:8]}…) has {len(w4_charges)} W7 charge(s)."
        else:
            extra_msg = (
                f"W4 plan ({W4_SRC_PLAN[:8]}…, Transact GIA) has NO W7 charge. "
                f"Reason: this plan was a wrap product already in Plannr (is_wrap_product=True), "
                f"not in the source system export. It appears in W4 but was omitted from "
                f"W7_charges because its rate was not in the Transact rate sheet (blank rate) "
                f"and it is not classified as an Add-to-Plannr prov- plan in the charge "
                f"build logic — it lacks a Section A/B rate entry. Any service charge "
                f"for this plan needs to be entered manually in Plannr."
            )
    else:
        extra_msg = "W4 plan not found in plans for X2 dest UUID."
    print_explain("X2 (Add-to-Plannr)", X2, src, all_dest_data[X2], extra=extra_msg)

    p("\n" + "=" * 60)
    p("SPOT-CHECK COMPLETE")
    p("=" * 60)


if __name__ == "__main__":
    main()
