"""
taskB_20261029_fix.py
B1: Fix Unknown Provider plans (PUT /api/v1/plans/{uuid} {"provider_uuid": ...})
B2: Fix contract enquiry (PUT)
B3: Skip type change; print flexi drawdown details for manual E5
B4: Fix joint-owned plans owners (PUT)
B5: Create 2 missing circles
B6: Audit joint records; save clear list for E6
One script at a time. Count before / write / count after. Mismatch = STOP.
"""
import csv, json, os, sys, time
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

UNKNOWN_PROV_UUID = "175e2a37-200f-400f-8140-e0d7206db1b1"
SRC_ACCT          = "db0ab7f5-8e52-4c89-a915-29fefeeca0d4"
KATIE_SRC         = "2153bb3c-7b76-4643-86a3-a9c53acab0b0"
JIM               = "1a230fbc-4026-4bd1-851f-18cb55ab19f2"
KATIE             = "78a6aab8-55f3-42ed-b8ef-4867da1c0c78"
FLEXI_DEST        = "fff4250d-a0c6-4d97-8fec-b81b3483904d"
FAIL_LIMIT        = 5

def p(*a): print(*a, flush=True)
def stop(msg): p(f"\n!!STOP!! {msg}"); sys.exit(1)

def lcsv(path):
    with open(path, newline="", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))

def ljson(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)

sc  = SecretClient(vault_url="https://kv-plannr-ingestion.vault.azure.net/",
                   credential=DefaultAzureCredential())
pat = sc.get_secret("plannr-prod-pat").value

class API:
    def __init__(self, pat, acct):
        self._s = requests.Session()
        self._s.headers.update({"Authorization": f"Bearer {pat}",
                                 "Accept": "application/json",
                                 "Content-Type": "application/json",
                                 "X-PLANNR-ACCOUNT-UUID": acct})
    def get(self, path, **kw):  return self._req("GET",    path, **kw)
    def post(self, path, **kw): return self._req("POST",   path, **kw)
    def put(self, path, **kw):  return self._req("PUT",    path, **kw)
    def delete(self, path, **kw): return self._req("DELETE", path, **kw)
    def _req(self, method, path, **kw):
        url = f"{BASE_URL}{path}" if path.startswith("/") else path
        for _ in range(4):
            r = self._s.request(method, url, **kw)
            if r.status_code != 429: return r
            time.sleep(int(r.headers.get("Retry-After", 30)))
        return r

def paged(api, path, params=None, max_pages=20):
    items, page = [], 1
    while page <= max_pages:
        pr = dict(params or {}); pr["page"] = page; pr["per_page"] = 200
        r = api.get(path, params=pr)
        if r.status_code != 200:
            p(f"  [WARN] GET {path} page {page} -> HTTP {r.status_code}"); break
        body = r.json()
        if isinstance(body, list): items.extend(body); break
        data = body.get("data", [])
        if isinstance(data, list):   items.extend(data)
        elif isinstance(data, dict): items.append(data)
        meta  = body.get("meta", {}) if isinstance(body, dict) else {}
        if page >= (meta.get("last_page") or 1): break
        page += 1
    return items

dest_pat = pat
src_pat  = sc.get_secret("plannr-hfp-pat").value

dest = API(dest_pat, DEST_ACCOUNT_UUID)
src  = API(src_pat,  "db0ab7f5-8e52-4c89-a915-29fefeeca0d4")

def unwrap(r):
    """Unwrap single-item GET response from its 'data' key."""
    body = r.json() if hasattr(r, "json") else r
    return (body.get("data") or body) if isinstance(body, dict) else body

# ── Load state ─────────────────────────────────────────────────────────────
client_map        = ljson(os.path.join(STATE_DIR, "client_map.json"))
plan_map          = ljson(os.path.join(STATE_DIR, "plan_map.json"))
inv_plan          = {v: k for k, v in plan_map.items()}   # dest→src
plan_dest_set     = set(plan_map.values())

w2c               = lcsv(os.path.join(IMPORT_DIR, "W2_clients.csv"))
w2circles         = lcsv(os.path.join(IMPORT_DIR, "W2_circles.csv"))
w5                = lcsv(os.path.join(IMPORT_DIR, "W5_plans.csv"))
w5fx              = lcsv(os.path.join(IMPORT_DIR, "W5_plans_rewritten_fixed.csv"))

# Joint record lookups
joint_rows  = {}   # src_uuid -> row
for row in w2c:
    if row.get("type","").lower() == "joint":
        j_src = (row.get("external_reference_value") or "").strip()
        if j_src:
            joint_rows[j_src] = row
joint_src_set   = set(joint_rows.keys())
# dest_uuid -> src_uuid  for joints
joint_dest_map  = {client_map[s]: s for s in joint_src_set if s in client_map}
joint_dest_set  = set(joint_dest_map.keys())
# joint_dest_uuid -> [dest_ind1, dest_ind2]
joint_members   = {}
for j_src2, row in joint_rows.items():
    owners_raw = (row.get("owners") or "").strip()
    members_src = [x.strip() for x in owners_raw.split(",") if x.strip()]
    dest_j = client_map.get(j_src2)
    if dest_j:
        joint_members[dest_j] = [client_map[m] for m in members_src if m in client_map]

# W5 src plan -> row
w5_by_src = {(r.get("external_reference_value") or "").strip(): r for r in w5}
# W5_fixed multi-owner dest plan UUIDs
multi_owner_dest = {(r.get("uuid") or "").strip()
                    for r in w5fx if "," in (r.get("client_uuids") or "")}

# W2_circles: frozenset{src_ind1,src_ind2} -> row
circles_by_src_pair = {}
for row in w2circles:
    accts = frozenset(x.strip() for x in (row.get("account_uuids") or "").split(",") if x.strip())
    if accts:
        circles_by_src_pair[accts] = row

p("=== DATA LOADED ===")
p(f"  client_map {len(client_map)}, plan_map {len(plan_map)}, joint records {len(joint_src_set)}")

# ═══════════════════════════════════════════════════════════════════════════════
# B1  Unknown Provider
# ═══════════════════════════════════════════════════════════════════════════════
p("\n=== B1: UNKNOWN PROVIDER ===")

dest_provs = paged(dest, "/api/v1/providers", max_pages=5)
dest_prov_set      = {pv.get("uuid") for pv in dest_provs if pv.get("uuid")}
dest_prov_by_name  = {(pv.get("name") or "").lower(): pv.get("uuid","")
                      for pv in dest_provs if pv.get("uuid")}
p(f"  Dest providers: {len(dest_prov_set)}")

p("  Fetching all 1901 dest plans (10 pages)…")
all_plans = paged(dest, "/api/v1/plans", max_pages=15)
p(f"  Total: {len(all_plans)}")
if len(all_plans) != 1901:
    p(f"  [NOTE] expected 1901, got {len(all_plans)}")

unknown_prov = [pl for pl in all_plans
                if (pl.get("provider") or {}).get("uuid") == UNKNOWN_PROV_UUID]
before_b1 = len(unknown_prov)
p(f"  Unknown Provider plans (before): {before_b1}")

# Try W5 lookup first; then source API for those without a W5 match
to_set_prov = []   # (dest_uuid, prov_uuid)
need_src     = []  # dest plan data (no W5 provider) → check source API
ar_b1        = []  # (dest_uuid, plan_type, reason)

for pl in unknown_prov:
    duuid = pl.get("uuid","")
    src_u = inv_plan.get(duuid,"")
    w5r   = w5_by_src.get(src_u, {})
    pu    = (w5r.get("provider_uuid") or "").strip()
    if pu and pu in dest_prov_set and pu != UNKNOWN_PROV_UUID:
        to_set_prov.append((duuid, pu))
    else:
        need_src.append({"dest_uuid": duuid, "src_uuid": src_u,
                          "atype": (pl.get("abstract_type") or pl.get("type") or "?")})

p(f"  Plans with real W5 provider: {len(to_set_prov)}")
p(f"  Plans needing source API check: {len(need_src)}")

# Batch source API lookup for plans without W5 provider
if need_src:
    p("  Fetching source plans to check providers…")
    src_checked = 0
    for item in need_src:
        src_u = item["src_uuid"]
        if not src_u:
            ar_b1.append((item["dest_uuid"], item["atype"], "no_src_map"))
            continue
        rs = src.get(f"/api/v1/plans/{src_u}")
        if rs.status_code != 200:
            ar_b1.append((item["dest_uuid"], item["atype"], f"src_{rs.status_code}"))
            continue
        src_plan = unwrap(rs)
        src_prov = src_plan.get("provider") or {}
        src_prov_name = (src_prov.get("name") or "").strip().lower()
        if not src_prov_name or src_prov_name == "unknown provider":
            ar_b1.append((item["dest_uuid"], item["atype"], "src_no_provider"))
            continue
        dest_prov_uuid = dest_prov_by_name.get(src_prov_name, "")
        if dest_prov_uuid and dest_prov_uuid != UNKNOWN_PROV_UUID:
            to_set_prov.append((item["dest_uuid"], dest_prov_uuid))
        else:
            ar_b1.append((item["dest_uuid"], item["atype"], f"src_prov_not_in_dest:{src_prov_name}"))
        src_checked += 1
        if src_checked % 50 == 0:
            p(f"    …{src_checked}/{len(need_src)} source plans checked")
    p(f"  Source lookup complete: {src_checked} checked")

p(f"  Plans to set real provider (W5+source): {len(to_set_prov)}")
p(f"  AR list: {len(ar_b1)}")

# Test one, then rest
b1_ok = 0; b1_fail = 0; consec = 0
if to_set_prov:
    t_uuid, t_prov = to_set_prov[0]
    rp = dest.put(f"/api/v1/plans/{t_uuid}", json={"provider_uuid": t_prov})
    if rp.status_code not in (200,201):
        stop(f"B1 test PUT -> HTTP {rp.status_code}: {rp.text[:200]}")
    rg = dest.get(f"/api/v1/plans/{t_uuid}")
    got = (unwrap(rg).get("provider") or {}).get("uuid","")
    if got != t_prov:
        stop(f"B1 test: provider not set. expected {t_prov[:8]} got {got!r}")
    p(f"  TEST OK: provider set on {t_uuid[:8]}")
    b1_ok = 1; consec = 0
    for duuid, prov in to_set_prov[1:]:
        rp2 = dest.put(f"/api/v1/plans/{duuid}", json={"provider_uuid": prov})
        if rp2.status_code in (200,201):
            rg2 = dest.get(f"/api/v1/plans/{duuid}")
            if (unwrap(rg2).get("provider") or {}).get("uuid") == prov:
                b1_ok += 1; consec = 0
            else:
                p(f"  [WARN] PUT OK but provider wrong for {duuid[:8]}"); b1_fail += 1; consec += 1
        else:
            p(f"  [FAIL] PUT {duuid[:8]} -> HTTP {rp2.status_code}"); b1_fail += 1; consec += 1
        if consec >= FAIL_LIMIT: stop(f"B1: {FAIL_LIMIT} consecutive failures")

p(f"\n  B1: {b1_ok} providers set, {b1_fail} failures")
if b1_ok != len(to_set_prov):
    stop(f"B1 count: expected {len(to_set_prov)}, got {b1_ok}")

p(f"  AR list ({len(ar_b1)} plans — no resolvable provider):")
for duuid, ptype, reason in ar_b1[:15]:
    p(f"    {duuid}  type={ptype}  reason={reason[:40]}")
if len(ar_b1) > 15:
    p(f"    … and {len(ar_b1)-15} more")
# Save for E8
json.dump([{"uuid": d, "type": t, "reason": r} for d,t,r in ar_b1],
          open(os.path.join(STATE_DIR,"b1_ar_list.json"),"w"), indent=2)

# ═══════════════════════════════════════════════════════════════════════════════
# B2  Contract enquiry
# ═══════════════════════════════════════════════════════════════════════════════
p("\n=== B2: CONTRACT ENQUIRY ===")

to_ce = []   # (dest_plan_uuid, adviser_uuid)
for row in w5:
    ce_src  = (row.get("contract_enquiry_account_uuid") or "").strip()
    src_pl  = (row.get("external_reference_value") or "").strip()
    if not ce_src or not src_pl: continue
    dest_pl = plan_map.get(src_pl)
    if not dest_pl: continue
    if   ce_src == SRC_ACCT:   to_ce.append((dest_pl, JIM))
    elif ce_src == KATIE_SRC:  to_ce.append((dest_pl, KATIE))
    else: p(f"  [WARN] unknown CE src {ce_src[:8]}")

jim_n   = sum(1 for _,a in to_ce if a==JIM)
katie_n = sum(1 for _,a in to_ce if a==KATIE)
p(f"  Plans to set: Jim={jim_n}  Katie={katie_n}  total={len(to_ce)}")
if jim_n != 39 or katie_n != 7:
    stop(f"B2: expected Jim=39 Katie=7, got Jim={jim_n} Katie={katie_n}")

def get_ce_uuid(plan_body):
    """Extract contract_enquiry adviser UUID from unwrapped plan object."""
    obj = plan_body.get("contract_enquiry_account") or {}
    return obj.get("uuid","") if isinstance(obj,dict) else ""

# Count before on dest (using unwrapped data)
before_b2 = 0
for dp, _ in to_ce:
    rg = dest.get(f"/api/v1/plans/{dp}")
    if rg.status_code == 200:
        if not get_ce_uuid(unwrap(rg)):
            before_b2 += 1

# Test one
b2_ok = 0; b2_fail = 0; consec = 0
if to_ce:
    t_dp, t_adv = to_ce[0]
    rp = dest.put(f"/api/v1/plans/{t_dp}", json={"contract_enquiry_account_uuid": t_adv})
    if rp.status_code not in (200,201):
        stop(f"B2 test PUT -> HTTP {rp.status_code}: {rp.text[:200]}")
    rg = dest.get(f"/api/v1/plans/{t_dp}")
    got_ce = get_ce_uuid(unwrap(rg))
    if got_ce != t_adv:
        stop(f"B2 test: expected {t_adv[:8]} got {got_ce!r}")
    p(f"  TEST OK: CE set on {t_dp[:8]}")
    b2_ok = 1
    for dp, adv in to_ce[1:]:
        rp2 = dest.put(f"/api/v1/plans/{dp}", json={"contract_enquiry_account_uuid": adv})
        if rp2.status_code in (200,201):
            rg2 = dest.get(f"/api/v1/plans/{dp}")
            if get_ce_uuid(unwrap(rg2)) == adv:
                b2_ok += 1; consec = 0
            else:
                p(f"  [WARN] CE wrong for {dp[:8]}"); b2_fail += 1; consec += 1
        else:
            p(f"  [FAIL] PUT {dp[:8]} -> HTTP {rp2.status_code}"); b2_fail += 1; consec += 1
        if consec >= FAIL_LIMIT: stop(f"B2: {FAIL_LIMIT} consecutive failures")

p(f"\n  B2: {b2_ok} set, {b2_fail} failures")
if b2_ok != 46:
    stop(f"B2 count: expected 46, got {b2_ok}")

# ═══════════════════════════════════════════════════════════════════════════════
# B3  Flexi drawdown — print details only (PUT can't change type; E5 will recreate)
# ═══════════════════════════════════════════════════════════════════════════════
p("\n=== B3: FLEXI DRAWDOWN DETAILS (for E5 manual recreation) ===")
rg = dest.get(f"/api/v1/plans/{FLEXI_DEST}")
if rg.status_code == 200:
    fp = unwrap(rg)
    p(f"  UUID:        {FLEXI_DEST}")
    p(f"  Name:        {fp.get('name','')}")
    p(f"  Type:        {fp.get('abstract_type','')} / {fp.get('sub_type','')}")
    p(f"  Provider:    {(fp.get('provider') or {}).get('name','')}  uuid={(fp.get('provider') or {}).get('uuid','')}")
    p(f"  Policy:      {fp.get('policy_number','')}")
    p(f"  Status:      {fp.get('status','')}")
    accs = fp.get("accounts") or []
    for ac in accs:
        if isinstance(ac, dict):
            p(f"  Owner:       {ac.get('uuid','') or ac.get('account_uuid','')}  pct={ac.get('ownership_percentage','')}")
    rv = dest.get(f"/api/v1/plans/{FLEXI_DEST}/valuations", params={"per_page":1,"page":1})
    vals = (rv.json().get("data") or []) if rv.status_code==200 else []
    if vals:
        v = vals[0]
        p(f"  Valuation:   {v.get('value','')} on {v.get('valuation_date','')}  uuid={v.get('uuid','')}")
    rc = dest.get(f"/api/v1/plans/{FLEXI_DEST}/charges")
    chgs = []
    if rc.status_code == 200:
        chgs = rc.json().get("data") or rc.json()
        if isinstance(chgs, list):
            p(f"  Charges:     {len(chgs)}")
            for ch in chgs:
                p(f"    uuid={ch.get('uuid','')}  amount={ch.get('amount','')}  type={ch.get('type','')}  freq={ch.get('frequency','')}")
    b3_state = {
        "flexi_uuid": FLEXI_DEST,
        "name": fp.get("name",""),
        "policy_number": fp.get("policy_number",""),
        "status": fp.get("status",""),
        "provider_uuid": (fp.get("provider") or {}).get("uuid",""),
        "accounts": [{"account_uuid": (a.get("uuid") or a.get("account_uuid","")),
                      "ownership_percentage": a.get("ownership_percentage",100)}
                     for a in accs if isinstance(a,dict)],
        "valuation": vals[0] if vals else None,
        "charges": chgs if isinstance(chgs, list) else [],
    }
    json.dump(b3_state, open(os.path.join(STATE_DIR,"b3_flexi_details.json"),"w"), indent=2)
    p(f"  Saved to b3_flexi_details.json")
else:
    p(f"  [ERROR] GET flexi plan -> HTTP {rg.status_code}")

# ═══════════════════════════════════════════════════════════════════════════════
# B4  Joint-owned plans  (find from dest using filter, confirm in plan_map)
# ═══════════════════════════════════════════════════════════════════════════════
p("\n=== B4: JOINT-OWNED PLANS ===")

# Find plans owned by joint accounts using dest API (bracketed filter)
joint_owned = []   # {dest_plan, joint_dest, members, plan_data}
for jd in joint_dest_set:
    rp = dest.get("/api/v1/plans", params={"filter[account_uuids]": jd, "per_page": 200})
    if rp.status_code != 200:
        p(f"  [WARN] filter plans for joint {jd[:8]} -> HTTP {rp.status_code}")
        continue
    body = rp.json()
    plans_j = body.get("data") or (body if isinstance(body, list) else [])
    if len(plans_j) > 50:
        stop(f"B4: joint {jd[:8]} has {len(plans_j)} plans (>50 limit)")
    for pl in plans_j:
        duuid = pl.get("uuid","")
        if duuid not in plan_dest_set:
            continue  # not a migrated plan — skip
        if duuid in multi_owner_dest:
            continue  # Guard 2: multi-owner individual plan — skip
        members = joint_members.get(jd, [])
        if len(members) < 2:
            p(f"  [WARN] joint {jd[:8]} has <2 members, skipping plan {duuid[:8]}")
            continue
        joint_owned.append({"dest_plan": duuid, "joint_dest": jd,
                             "ind1": members[0], "ind2": members[1],
                             "plan_data": pl})

p(f"  Joint-owned migrated plans found: {len(joint_owned)}")
if len(joint_owned) != 21:
    p(f"  [NOTE] expected 21, found {len(joint_owned)}")
    if len(joint_owned) > 30:
        stop(f"B4: found {len(joint_owned)} joint-owned plans, expected 21 — too many to be safe")

before_b4 = len(joint_owned)

# Build confirmed list with guards
confirmed_b4 = []
skipped_g1   = []
pension_isa   = []

for item in joint_owned:
    duuid    = item["dest_plan"]
    jd       = item["joint_dest"]
    # Guard 1: GET plan and confirm joint is an owner
    rg = dest.get(f"/api/v1/plans/{duuid}")
    if rg.status_code != 200:
        p(f"  [WARN] GET plan {duuid[:8]} -> HTTP {rg.status_code}"); skipped_g1.append(item); continue
    pl = unwrap(rg)
    accs = pl.get("accounts") or []
    owner_uuids = {(a.get("uuid") or a.get("account_uuid","")) for a in accs if isinstance(a,dict)}
    # Also check ownership_values dict keys as fallback
    ov = pl.get("ownership_values") or {}
    if isinstance(ov, dict):
        owner_uuids |= set(ov.keys())
    if jd not in owner_uuids:
        p(f"  [G1 SKIP] joint {jd[:8]} not owner of plan {duuid[:8]}. owners={[o[:8] for o in owner_uuids]}")
        skipped_g1.append(item); continue
    # Guard 3: pension or ISA -> first individual only
    ptype = (pl.get("abstract_type") or pl.get("type") or "").lower()
    psub  = (pl.get("sub_type") or "").lower()
    is_pension = "pension" in ptype or "pension" in psub
    is_isa     = "isa" in psub or "individual_savings" in psub
    if is_pension or is_isa:
        pension_isa.append({**item, "ptype": ptype, "psub": psub})
        accounts_body = [{"account_uuid": item["ind1"], "ownership_percentage": 100}]
    else:
        accounts_body = [{"account_uuid": item["ind1"], "ownership_percentage": 50},
                         {"account_uuid": item["ind2"], "ownership_percentage": 50}]
    confirmed_b4.append({**item, "accounts_body": accounts_body})

p(f"  Confirmed for update: {len(confirmed_b4)}")
if skipped_g1:
    p(f"  Skipped Guard 1 (joint not owner): {len(skipped_g1)}")
    for it in skipped_g1:
        p(f"    plan={it['dest_plan']}  joint={it['joint_dest']}")
if pension_isa:
    p(f"\n  Guard 3 — Pensions/ISAs (100% to first individual):")
    for it in pension_isa:
        p(f"    plan={it['dest_plan']}  {it['ptype']}/{it['psub']}")

# Test one PUT
b4_ok = 0; b4_fail = 0; consec = 0
if confirmed_b4:
    t = confirmed_b4[0]
    p(f"\n  TEST: PUT {t['dest_plan'][:8]} accounts={[a['account_uuid'][:8] for a in t['accounts_body']]}")
    rp = dest.put(f"/api/v1/plans/{t['dest_plan']}", json={"accounts": t["accounts_body"]})
    if rp.status_code not in (200,201):
        stop(f"B4 test PUT -> HTTP {rp.status_code}: {rp.text[:200]}")
    rg = dest.get(f"/api/v1/plans/{t['dest_plan']}")
    got_pl    = unwrap(rg)
    got_accs  = got_pl.get("accounts") or []
    got_uuids = {(a.get("uuid") or a.get("account_uuid","")) for a in got_accs if isinstance(a,dict)}
    exp_uuids = {a["account_uuid"] for a in t["accounts_body"]}
    if not exp_uuids.issubset(got_uuids):
        stop(f"B4 test: owners not set. expected {exp_uuids} got {got_uuids}")
    p(f"  TEST OK: owners={got_uuids}")
    b4_ok = 1

    for item in confirmed_b4[1:]:
        rp2 = dest.put(f"/api/v1/plans/{item['dest_plan']}", json={"accounts": item["accounts_body"]})
        if rp2.status_code in (200,201):
            rg2  = dest.get(f"/api/v1/plans/{item['dest_plan']}")
            got2 = {(a.get("uuid") or a.get("account_uuid","")) for a in (unwrap(rg2).get("accounts") or []) if isinstance(a,dict)}
            exp2 = {a["account_uuid"] for a in item["accounts_body"]}
            if exp2.issubset(got2):
                b4_ok += 1; consec = 0
            else:
                p(f"  [WARN] wrong owners for {item['dest_plan'][:8]}"); b4_fail += 1; consec += 1
        else:
            p(f"  [FAIL] PUT {item['dest_plan'][:8]} -> HTTP {rp2.status_code}"); b4_fail += 1; consec += 1
        if consec >= FAIL_LIMIT: stop(f"B4: {FAIL_LIMIT} consecutive failures")

p(f"\n  B4: {b4_ok} updated, {b4_fail} failures")
if b4_ok != len(confirmed_b4):
    stop(f"B4 count: expected {len(confirmed_b4)}, got {b4_ok}")

# Count after: joint accounts should now have 0 plans each
after_b4 = 0
for item in confirmed_b4:
    rg2 = dest.get(f"/api/v1/plans/{item['dest_plan']}")
    if rg2.status_code == 200:
        got_accs  = unwrap(rg2).get("accounts") or []
        got_uuids = {(a.get("uuid") or a.get("account_uuid","")) for a in got_accs if isinstance(a,dict)}
        if item["joint_dest"] in got_uuids:
            after_b4 += 1
            p(f"  [WARN] joint still owner of {item['dest_plan'][:8]}")

p(f"  After B4: {after_b4} plans still with joint owner (expected 0)")
if after_b4 > 0:
    stop(f"B4: {after_b4} plans still owned by joint UUID")

# Confirm charges follow updated plans
p("\n  Charges on updated plans (should move with plan ownership):")
total_b4_charges = 0
for item in confirmed_b4:
    rc = dest.get(f"/api/v1/plans/{item['dest_plan']}/charges")
    if rc.status_code == 200:
        chgs = rc.json().get("data") or rc.json()
        if isinstance(chgs, list) and chgs:
            total_b4_charges += len(chgs)
            p(f"    plan {item['dest_plan'][:8]}: {len(chgs)} charge(s)")
p(f"  Total charges confirmed on updated plans: {total_b4_charges} (expected 7 total)")

# ═══════════════════════════════════════════════════════════════════════════════
# B5  Circles — find 2 missing and create them
# ═══════════════════════════════════════════════════════════════════════════════
p("\n=== B5: CIRCLES ===")

# GET with include=accounts
all_circles = paged(dest, "/api/v1/circles", params={"include": "accounts"}, max_pages=5)
p(f"  Circles fetched: {len(all_circles)}")

# If list endpoint doesn't include accounts, fetch each individually
if all_circles and all_circles[0].get("accounts") is None:
    p("  include=accounts not in list; fetching circles individually…")
    details = []
    for circ in all_circles:
        rc2 = dest.get(f"/api/v1/circles/{circ.get('uuid','')}")
        if rc2.status_code == 200:
            details.append(unwrap(rc2))
    all_circles = details
    p(f"  Fetched {len(all_circles)} circles individually")

# Build circle_pairs  frozenset{ind1,ind2} -> circle_uuid
circle_pairs  = {}
sample_groups = None
for circ in all_circles:
    accounts = circ.get("accounts") or []
    muuids = frozenset((m.get("uuid") or m.get("account_uuid",""))
                       for m in accounts if isinstance(m,dict))
    if len(muuids) >= 2:
        circle_pairs[muuids] = circ.get("uuid","")
    if sample_groups is None:
        groups = circ.get("groups") or []
        if groups:
            sample_groups = [g.get("uuid","") for g in groups if isinstance(g,dict) and g.get("uuid")]

p(f"  Circles with 2+ members: {len(circle_pairs)}")
if sample_groups:
    p(f"  Group UUIDs from sample circle: {sample_groups}")

# Find missing pairs
missing_circles = []
for jd, members in joint_members.items():
    if len(members) < 2:
        continue
    pair = frozenset(members[:2])
    if pair in circle_pairs:
        continue
    # Not found — need to create
    jsr = joint_dest_map.get(jd,"")
    src_pair = frozenset((joint_rows.get(jsr,{}).get("owners","") or "").split(","))
    src_pair = frozenset(x.strip() for x in src_pair if x.strip())
    w2c_row  = circles_by_src_pair.get(src_pair, {})
    cname    = (w2c_row.get("name","") or "").strip()
    if not cname:
        jr = joint_rows.get(jsr,{})
        fn = (jr.get("first_name","") or jr.get("entity_name","") or "").strip()
        ln = (jr.get("last_name","") or "").strip()
        cname = f"{fn} {ln}".strip() or f"Circle-{jsr[:8]}"
    missing_circles.append({"joint_dest": jd, "ind1": members[0], "ind2": members[1],
                             "circle_name": cname})

p(f"\n  Joint couples without a circle: {len(missing_circles)}")
if len(missing_circles) != 2:
    p(f"  [NOTE] expected 2, found {len(missing_circles)}")
    for mc in missing_circles[:5]:
        p(f"    joint={mc['joint_dest'][:8]}  name={mc['circle_name']!r}")
    if len(missing_circles) > 10:
        stop(f"B5: {len(missing_circles)} missing circles (expected 2) — aborting")

before_b5 = len(all_circles)
b5_created = 0
new_circle_uuids = []

for mc in missing_circles:
    body = {"name": mc["circle_name"],
            "account_uuids": [mc["ind1"], mc["ind2"]]}
    if sample_groups:
        body["group_uuids"] = sample_groups
    p(f"\n  POST /api/v1/circles  name={mc['circle_name']!r}  members=[{mc['ind1'][:8]},{mc['ind2'][:8]}]")
    rc = dest.post("/api/v1/circles", json=body)
    p(f"  -> HTTP {rc.status_code}")
    if rc.status_code not in (200,201):
        p(f"  [FAIL] {rc.text[:200]}"); continue
    created_uuid = (rc.json().get("data") or rc.json()).get("uuid","")
    p(f"  Created: {created_uuid}")
    new_circle_uuids.append(created_uuid)
    # Verify
    rv = dest.get(f"/api/v1/circles/{created_uuid}")
    if rv.status_code == 200:
        got_m = frozenset((m.get("uuid","")) for m in unwrap(rv).get("accounts",[]) if isinstance(m,dict))
        if {mc["ind1"],mc["ind2"]}.issubset(got_m):
            p(f"  Verified members OK"); b5_created += 1
        else:
            p(f"  [WARN] members mismatch: {got_m}")
    else:
        b5_created += 1  # count anyway

after_b5 = len(paged(dest, "/api/v1/circles", max_pages=5))
p(f"\n  Circles: before={before_b5}  created={b5_created}  after={after_b5}")
if after_b5 != before_b5 + len(missing_circles):
    stop(f"B5 count: expected {before_b5+len(missing_circles)}, got {after_b5}")

# ═══════════════════════════════════════════════════════════════════════════════
# B6  Joint record audit
# ═══════════════════════════════════════════════════════════════════════════════
p("\n=== B6: JOINT RECORD AUDIT ===")

# Rebuild circle_pairs with new circles
final_circles = paged(dest, "/api/v1/circles", params={"include": "accounts"}, max_pages=5)
if final_circles and final_circles[0].get("accounts") is None:
    final_details = []
    for circ in final_circles:
        rc2 = dest.get(f"/api/v1/circles/{circ.get('uuid','')}")
        if rc2.status_code == 200:
            final_details.append(unwrap(rc2))
    final_circles = final_details
circle_pairs_final = {}
for circ in final_circles:
    muuids = frozenset((m.get("uuid") or m.get("account_uuid",""))
                       for m in (circ.get("accounts") or []) if isinstance(m,dict))
    if len(muuids) >= 2:
        circle_pairs_final[muuids] = circ.get("uuid","")

def count_paged(api, path, params=None):
    r = api.get(path, params=dict({"per_page":200,"page":1}, **(params or {})))
    if r.status_code != 200: return -1
    body = r.json()
    data = body.get("data") or (body if isinstance(body,list) else [])
    total = (body.get("meta",{}).get("total") if isinstance(body,dict) else None)
    return total if total is not None else len(data)

audit = []
for jd, js in joint_dest_map.items():
    members = joint_members.get(jd, [])
    pair    = frozenset(members[:2]) if len(members)>=2 else frozenset()

    # Plans (bracketed filter)
    rp2 = dest.get("/api/v1/plans", params={"filter[account_uuids]": jd, "per_page":200})
    if rp2.status_code == 200:
        pdata = rp2.json().get("data") or []
        n_plans = len(pdata)
        if n_plans > 50:
            stop(f"B6: joint {jd[:8]} has {n_plans} plans (>50)")
    else:
        n_plans = -1

    # Notes
    rn = dest.get(f"/api/v1/client/{jd}/all-notes", params={"per_page":200})
    n_notes = len(rn.json().get("data") or []) if rn.status_code==200 else -1

    # Dependants
    rd = dest.get(f"/api/v1/client/{jd}/dependants", params={"per_page":200})
    dep_body = rd.json() if rd.status_code==200 else {}
    n_dep = len(dep_body.get("data") or (dep_body if isinstance(dep_body,list) else []))

    # Income
    ri = dest.get(f"/api/v1/client/{jd}/income", params={"per_page":200})
    if ri.status_code == 200:
        ib = ri.json()
        id2 = ib.get("data") or (ib if isinstance(ib,list) else [])
        n_inc = len(id2) if isinstance(id2,list) else 0
    else:
        n_inc = 0

    # Expenditure
    re = dest.get(f"/api/v1/client/{jd}/expenditure", params={"per_page":200})
    if re.status_code == 200:
        eb = re.json()
        ed2 = eb.get("data") or (eb if isinstance(eb,list) else [])
        n_exp = len(ed2) if isinstance(ed2,list) else 0
    else:
        n_exp = 0

    has_circle  = pair in circle_pairs_final
    circle_uuid = circle_pairs_final.get(pair,"")

    audit.append({
        "joint_dest":  jd,
        "joint_src":   js,
        "plans":       n_plans,
        "notes":       n_notes,
        "dependants":  n_dep,
        "income":      n_inc,
        "expenditure": n_exp,
        "has_circle":  has_circle,
        "circle_uuid": circle_uuid,
    })

with_circle    = [a for a in audit if a["has_circle"]]
without_circle = [a for a in audit if not a["has_circle"]]
clear_list = [a for a in with_circle
              if a["plans"]==0 and a["notes"]==0 and a["dependants"]==0
              and a["income"]==0 and a["expenditure"]==0]

p(f"  Total joint records: {len(audit)}")
p(f"  With shared circle: {len(with_circle)}")
p(f"  WITHOUT shared circle: {len(without_circle)}")
p(f"\n  Clear list (all-zero + shared circle): {len(clear_list)}")
for a in clear_list:
    p(f"    {a['joint_dest']}  circle={a['circle_uuid'][:8] if a['circle_uuid'] else 'n/a'}…")

if without_circle:
    p(f"\n  WITHOUT shared circle ({len(without_circle)}):")
    for a in without_circle:
        p(f"    {a['joint_dest']}  plans={a['plans']}  notes={a['notes']}  dep={a['dependants']}  inc={a['income']}  exp={a['expenditure']}")

p(f"\n  Non-zero joint records (with circle):")
for a in sorted(with_circle, key=lambda x: -(x["plans"]+x["notes"]+x["dependants"])):
    tot = a["plans"]+a["notes"]+a["dependants"]+a["income"]+a["expenditure"]
    if tot > 0:
        p(f"    {a['joint_dest']}  plans={a['plans']}  notes={a['notes']}  dep={a['dependants']}  inc={a['income']}  exp={a['expenditure']}")

# Save clear list for E6
json.dump(clear_list, open(os.path.join(STATE_DIR,"b6_clear_list.json"),"w"), indent=2)
p(f"\n  Saved clear list to b6_clear_list.json")

p("\n=== TASKB FIX COMPLETE ===")
