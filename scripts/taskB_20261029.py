"""
taskB_20261029.py — Task B plans clean-up (2026-10-29 evening)
B1: Unknown Provider — set from source or clear (if type doesn't need one)
B2: Contract enquiry — assign 39 → Jim, 7 → Katie
B3: Flexi drawdown stub — set type=pension sub_type=personal_pension_plan
B4: Joint-owned plans — owners → two individuals 50/50; move 7 charges
B5: Circles — create circles for 2 joint couples that have none
B6: Joint record audit — list joints with 0 data and shared circle vs not
Rule 9: counts and UUIDs only — no client names or note contents.
One script talks to Plannr at a time.
"""
import csv, io, json, os, re, sys, time
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
STATE_DIR         = os.path.join(IMPORT_DIR, "state")

W5_CSV   = os.path.join(IMPORT_DIR, "W5_plans.csv")
W2C_CSV  = os.path.join(IMPORT_DIR, "W2_clients.csv")
W2CI_CSV = os.path.join(IMPORT_DIR, "W2_circles.csv")
W7_CSV   = os.path.join(IMPORT_DIR, "W7_charges.csv")

JIM_UUID   = "1a230fbc-4026-4bd1-851f-18cb55ab19f2"
KATIE_UUID = "78a6aab8-55f3-42ed-b8ef-4867da1c0c78"

FLEXI_POLICY = "IH01853021"  # policy number suffix for the flexi drawdown stub

# types needing a provider
TYPES_NEED_PROVIDER = {
    "annuity", "general_insurance", "investment",
    "mortgage", "pension", "protection_policy",
}


def p(*a): print(*a, flush=True)
def stop(msg): p(f"\n!!STOP!! {msg}"); sys.exit(1)

def load_csv(path):
    with open(path, newline="", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))

def load_json(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)

def csv_bytes(rows, fieldnames):
    buf = io.StringIO()
    w = csv.DictWriter(buf, fieldnames=fieldnames, extrasaction="ignore")
    w.writeheader()
    for row in rows:
        w.writerow({k: row.get(k, "") for k in fieldnames})
    return buf.getvalue().encode("utf-8")


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
    def get(self, path, **kw):    return self._req("GET",    path, **kw)
    def post(self, path, **kw):   return self._req("POST",   path, **kw)
    def put(self, path, **kw):    return self._req("PUT",    path, **kw)
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

# reverse maps
dest_to_src_client = {v: k for k, v in client_map.items()}
dest_to_src_plan   = {v: k for k, v in plan_map.items()}


# ── Import runner ─────────────────────────────────────────────────────────────
def run_import(imp_type, mode, csv_data_bytes, label=""):
    label = label or f"{imp_type}/{mode}"
    p(f"\n  [{label}] Creating import…")
    r = dest.post("/api/v1/import", data={
        "type": imp_type, "mode": mode,
    }, files={"file": ("import.csv", csv_data_bytes, "text/csv")})
    if r.status_code not in (200, 201):
        stop(f"Create import {label}: HTTP {r.status_code} {r.text[:200]}")
    imp_uuid = (r.json().get("data") or r.json()).get("uuid")
    p(f"  [{label}] UUID={imp_uuid}")
    time.sleep(8)

    r2 = dest.get(f"/api/v1/import/{imp_uuid}")
    if r2.status_code != 200:
        stop(f"GET import {imp_uuid}: HTTP {r2.status_code}")
    hdrs = (r2.json().get("data") or r2.json()).get("headers") or []
    if not hdrs:
        stop(f"Import {imp_uuid} has no headers — cannot map")
    mapping = {h: h for h in hdrs}
    p(f"  [{label}] Mapping {len(hdrs)} headers")

    r3 = dest.post(f"/api/v1/import/map/{imp_uuid}", json={"headers": mapping})
    if r3.status_code not in (200, 201):
        stop(f"Map import {imp_uuid}: HTTP {r3.status_code} {r3.text[:200]}")

    r4 = dest.post(f"/api/v1/import/start/{imp_uuid}")
    if r4.status_code not in (200, 201):
        stop(f"Start import {imp_uuid}: HTTP {r4.status_code} {r4.text[:200]}")

    p(f"  [{label}] Polling…")
    deadline = time.time() + 300
    last_prog = -1
    while time.time() < deadline:
        time.sleep(8)
        rp = dest.get(f"/api/v1/import/{imp_uuid}")
        if rp.status_code != 200:
            p(f"  [{label}] poll HTTP {rp.status_code}"); continue
        d = (rp.json().get("data") or rp.json())
        st   = str(d.get("status") or "").lower()
        prog = d.get("progress") or 0
        tj   = d.get("total_jobs") or 0
        fj   = d.get("failed_jobs") or 0
        if prog != last_prog:
            p(f"  [{label}] status={st}  {prog}/{tj}  failed={fj}")
            last_prog = prog
        if st == "complete":
            p(f"  [{label}] ✓ complete  total={tj}  failed={fj}")
            return d
        if st in ("failed", "error"):
            stop(f"Import {imp_uuid} status={st}")
    # timeout
    rp = dest.get(f"/api/v1/import/{imp_uuid}")
    d = (rp.json().get("data") or rp.json())
    stop(f"Import {imp_uuid} timed out (5 min): "
         f"status={d.get('status')}  progress={d.get('progress')}/{d.get('total_jobs')}")


# ═════════════════════════════════════════════════════════════════════════════
# ANALYSIS PHASE
# ═════════════════════════════════════════════════════════════════════════════

# ── Load source plan data ─────────────────────────────────────────────────────
p("Loading source plan data…")
w5_rows = load_csv(W5_CSV)
w5_by_src_uuid = {}
for row in w5_rows:
    su = (row.get("external_reference_value") or "").strip()
    if su:
        w5_by_src_uuid[su] = row
p(f"  W5_plans.csv: {len(w5_rows)} rows, {len(w5_by_src_uuid)} unique src plan UUIDs")

w2c_rows = load_csv(W2C_CSV)
joint_records = {}   # joint_src_uuid → {dest_uuid, owners: [src_ind1, src_ind2]}
for row in w2c_rows:
    if str(row.get("type", "")).lower() == "joint":
        src_uuid = (row.get("external_reference_value") or "").strip()
        dest_uuid = client_map.get(src_uuid)
        owners_raw = (row.get("owners") or "").strip()
        owners = [u.strip() for u in owners_raw.split(",") if u.strip()]
        if src_uuid and dest_uuid:
            joint_records[src_uuid] = {
                "src_uuid":  src_uuid,
                "dest_uuid": dest_uuid,
                "owners":    owners,
            }
p(f"  W2_clients.csv: {len(w2c_rows)} total, {len(joint_records)} joint records")

joint_dest_uuids = {v["dest_uuid"] for v in joint_records.values()}

# ── B2: Contract enquiry analysis ────────────────────────────────────────────
p("\n=== B2: CONTRACT ENQUIRY ANALYSIS ===")
# Plans in source with a contract_enquiry set
src_enquiry_plans = {}  # dest_plan_uuid → dest_enquiry_uuid
for su, row in w5_by_src_uuid.items():
    ce = (row.get("contract_enquiry_account_uuid") or "").strip()
    dp = plan_map.get(su)
    if ce and dp:
        src_enquiry_plans[dp] = ce

p(f"  Plans with contract_enquiry in source CSV: {len(src_enquiry_plans)}")
jim_plans   = [dp for dp, ce in src_enquiry_plans.items() if ce == JIM_UUID]
katie_plans = [dp for dp, ce in src_enquiry_plans.items() if ce == KATIE_UUID]
other_ce    = [dp for dp, ce in src_enquiry_plans.items()
               if ce not in (JIM_UUID, KATIE_UUID)]
p(f"    → Jim ({JIM_UUID[:8]}…):   {len(jim_plans)}")
p(f"    → Katie ({KATIE_UUID[:8]}…): {len(katie_plans)}")
p(f"    → Other adviser:           {len(other_ce)}")
if other_ce:
    p(f"  [WARN] {len(other_ce)} plans have a different adviser UUID — inspect before assigning")
    for dp in other_ce[:5]:
        su2 = dest_to_src_plan.get(dp, "?")
        ce2 = src_enquiry_plans[dp]
        p(f"    plan_dest={dp[:8]}  ce_uuid={ce2}")

# Now check dest: which of those are already set vs blank
p(f"  Checking dest for {len(src_enquiry_plans)} plans…")
blank_ce_plans = {}  # dest_plan_uuid → desired_ce_uuid (Jim or Katie)
for dp, ce in src_enquiry_plans.items():
    r = dest.get(f"/api/v1/plans/{dp}")
    if r.status_code != 200:
        p(f"  [WARN] GET plan {dp[:8]}: HTTP {r.status_code}"); continue
    body = r.json()
    plan_data = body.get("data") or body
    dest_ce = (plan_data.get("contract_enquiry_account_uuid") or
               ((plan_data.get("contract_enquiry") or {}).get("uuid") or "")).strip()
    if not dest_ce:
        blank_ce_plans[dp] = ce

p(f"  Plans with blank contract_enquiry on dest: {len(blank_ce_plans)}")
p(f"    → Jim:   {sum(1 for v in blank_ce_plans.values() if v == JIM_UUID)}")
p(f"    → Katie: {sum(1 for v in blank_ce_plans.values() if v == KATIE_UUID)}")


# ── B3: Flexi drawdown stub ───────────────────────────────────────────────────
p("\n=== B3: FLEXI DRAWDOWN STUB ===")
flexi_src_uuid = None
flexi_dest_uuid = None
for su, row in w5_by_src_uuid.items():
    pn = str(row.get("policy_number") or "")
    if FLEXI_POLICY in pn:
        flexi_src_uuid = su
        flexi_dest_uuid = plan_map.get(su)
        p(f"  Found: src={su[:8]}…  dest={flexi_dest_uuid}  policy={pn!r}")
        p(f"  Source type={row.get('type')!r}  sub_type={row.get('sub_type')!r}")
        break

if not flexi_dest_uuid:
    stop(f"Could not find plan with policy number containing {FLEXI_POLICY!r}")

r = dest.get(f"/api/v1/plans/{flexi_dest_uuid}")
if r.status_code != 200:
    stop(f"GET flexi plan {flexi_dest_uuid}: HTTP {r.status_code}")
fp = (r.json().get("data") or r.json())
p(f"  Dest currently: type={fp.get('type')!r}  sub_type={fp.get('sub_type')!r}"
  f"  provider_uuid={str(fp.get('provider_uuid') or '')[:8]}…")


# ── B4: Joint ownership analysis ─────────────────────────────────────────────
p("\n=== B4: JOINT OWNERSHIP ANALYSIS ===")
# Plans where source client_uuids = a joint record
joint_owned_plans = {}  # src_plan_uuid → {dest_plan_uuid, joint_src, ind1_src, ind2_src}
for su, row in w5_by_src_uuid.items():
    cu_raw = (row.get("client_uuids") or "").strip()
    cu_parts = [u.strip() for u in cu_raw.split(",") if u.strip()]
    cu = cu_parts[0] if cu_parts else ""
    if cu in joint_records:
        jr = joint_records[cu]
        dp = plan_map.get(su)
        if dp:
            joint_owned_plans[su] = {
                "src_plan_uuid":  su,
                "dest_plan_uuid": dp,
                "joint_src_uuid": cu,
                "joint_dest_uuid": jr["dest_uuid"],
                "ind1_src": jr["owners"][0] if len(jr["owners"]) > 0 else "",
                "ind2_src": jr["owners"][1] if len(jr["owners"]) > 1 else "",
                "ind1_dest": client_map.get(jr["owners"][0], "") if len(jr["owners"]) > 0 else "",
                "ind2_dest": client_map.get(jr["owners"][1], "") if len(jr["owners"]) > 1 else "",
            }

p(f"  Joint-owned plans in source CSV: {len(joint_owned_plans)}")
missing_owners = [v for v in joint_owned_plans.values()
                  if not v["ind1_dest"] or not v["ind2_dest"]]
if missing_owners:
    p(f"  [WARN] {len(missing_owners)} joint plans have unmapped individual owners")
    for v in missing_owners[:3]:
        p(f"    plan={v['dest_plan_uuid'][:8]}  ind1_src={v['ind1_src'][:8]}  ind2_src={v['ind2_src'][:8]}")

# Find charges on joint records in dest
p(f"  Fetching charges for {len(joint_dest_uuids)} joint dest UUIDs…")
joint_charges = {}  # charge_dest_uuid → {plan_uuid, joint_dest_uuid, first_ind_dest}
for jdu in joint_dest_uuids:
    charges = get_paged(dest, "/api/v2/charges",
                        params={"filter[account_uuids]": jdu})
    for ch in charges:
        cuuid = ch.get("uuid", "")
        plan_obj = ch.get("plan") or {}
        plan_du  = plan_obj.get("uuid", "") if isinstance(plan_obj, dict) else ""
        # find first individual owner of this plan
        src_pu = dest_to_src_plan.get(plan_du, "")
        jo_info = joint_owned_plans.get(src_pu)
        first_ind = jo_info["ind1_dest"] if jo_info else ""
        if not first_ind:
            # find from W5 directly
            w5r = w5_by_src_uuid.get(src_pu, {})
            cu_raw = (w5r.get("client_uuids") or "").strip()
            cu_parts = [u.strip() for u in cu_raw.split(",") if u.strip()]
            first_ind = client_map.get(cu_parts[0], "") if cu_parts else ""
        if cuuid:
            joint_charges[cuuid] = {
                "charge_uuid":     cuuid,
                "plan_dest_uuid":  plan_du,
                "joint_dest_uuid": jdu,
                "first_ind_dest":  first_ind,
            }

p(f"  Charges on joint dest records: {len(joint_charges)}")
if joint_charges:
    for cv in list(joint_charges.values())[:5]:
        p(f"    charge={cv['charge_uuid'][:8]}  plan={cv['plan_dest_uuid'][:8]}  "
          f"first_ind={cv['first_ind_dest'][:8]}")


# ── B1: Unknown Provider analysis ────────────────────────────────────────────
p("\n=== B1: UNKNOWN PROVIDER ANALYSIS ===")
# Get dest providers → find Unknown Provider UUID
dest_providers = get_paged(dest, "/api/v1/providers", max_pages=20)
dest_prov_by_name = {str(p2.get("name","")).strip(): p2.get("uuid","")
                     for p2 in dest_providers}
unknown_prov_uuid = dest_prov_by_name.get("Unknown Provider", "")
p(f"  Dest providers: {len(dest_providers)}  Unknown Provider UUID: {unknown_prov_uuid[:8] if unknown_prov_uuid else 'NOT FOUND'}…")

if not unknown_prov_uuid:
    # Try alternative spellings
    for name, uuid in dest_prov_by_name.items():
        if "unknown" in name.lower():
            unknown_prov_uuid = uuid
            p(f"  Found as {name!r}: {uuid[:8]}…")
            break

# Get source providers
src_providers = get_paged(src, "/api/v1/providers", max_pages=20)
src_prov_by_uuid = {p2.get("uuid",""): str(p2.get("name","")).strip()
                    for p2 in src_providers}

# Fetch all dest plans per client, find Unknown Provider ones
p(f"  Fetching plans for {len(dest_uuids)} clients to find Unknown Provider…")
unknown_prov_plans = {}  # dest_plan_uuid → plan_data
for i, du in enumerate(dest_uuids, 1):
    plans = get_paged(dest, "/api/v1/plans", params={"filter[client_uuid]": du})
    for plan in plans:
        puid = plan.get("uuid", "")
        prov_obj = plan.get("provider") or {}
        prov_uuid = (plan.get("provider_uuid") or
                     (prov_obj.get("uuid") if isinstance(prov_obj, dict) else ""))
        prov_name = (str(plan.get("provider_name") or "") or
                     (prov_obj.get("name","") if isinstance(prov_obj, dict) else ""))
        if (prov_uuid and prov_uuid == unknown_prov_uuid) or \
           "unknown" in prov_name.lower():
            if puid not in unknown_prov_plans:
                plan["_dest_uuid"] = puid
                plan["_prov_uuid_found"] = prov_uuid
                unknown_prov_plans[puid] = plan
    if i % 50 == 0:
        p(f"    {i}/{len(dest_uuids)} — {len(unknown_prov_plans)} Unknown Provider so far")

p(f"  Total Unknown Provider plans: {len(unknown_prov_plans)}")

# For each Unknown Provider plan: get src provider name
p(f"  Looking up source providers for {len(unknown_prov_plans)} plans…")
prov_set_plans  = {}  # dest_plan_uuid → dest_prov_uuid to set
prov_clear_plans= []  # dest_plan_uuid (type doesn't need provider)
prov_ar_plans   = []  # dest_plan_uuid (type needs provider, no match)
prov_no_src_plan= []  # dest_plan_uuid (no src plan found)

for dp, plan in unknown_prov_plans.items():
    ptype = str(plan.get("type") or "").lower()
    sp = dest_to_src_plan.get(dp, "")
    if not sp:
        prov_no_src_plan.append(dp); continue
    # GET src plan
    rs = src.get(f"/api/v1/plans/{sp}")
    if rs.status_code != 200:
        prov_no_src_plan.append(dp); continue
    src_plan = (rs.json().get("data") or rs.json())
    src_prov_uuid = str(src_plan.get("provider_uuid") or "").strip()
    src_prov_name = src_prov_by_uuid.get(src_prov_uuid, "")

    if src_prov_name and src_prov_name != "Unknown Provider":
        dest_match_uuid = dest_prov_by_name.get(src_prov_name, "")
        if dest_match_uuid:
            prov_set_plans[dp] = dest_match_uuid
            continue

    # No matching dest provider — check if type needs provider
    if ptype in TYPES_NEED_PROVIDER:
        prov_ar_plans.append(dp)
    else:
        prov_clear_plans.append(dp)

p(f"\n  Provider action breakdown:")
p(f"    Set to source provider:         {len(prov_set_plans)}")
p(f"    Clear (type doesn't need one):  {len(prov_clear_plans)}")
p(f"    AR list (needs but no match):   {len(prov_ar_plans)}")
p(f"    No src plan found:              {len(prov_no_src_plan)}")
p(f"    Total:                          "
  f"{len(prov_set_plans)+len(prov_clear_plans)+len(prov_ar_plans)+len(prov_no_src_plan)}")

if prov_ar_plans:
    p(f"\n  AR list — plans to review ({len(prov_ar_plans)}):")
    for dp in prov_ar_plans[:20]:
        plan = unknown_prov_plans[dp]
        p(f"    plan={dp}  type={plan.get('type')!r}")
    if len(prov_ar_plans) > 20:
        p(f"    … {len(prov_ar_plans)-20} more")


# ── B5: Circles analysis ─────────────────────────────────────────────────────
p("\n=== B5: CIRCLES ANALYSIS ===")
all_circles = get_paged(dest, "/api/v1/circles", max_pages=20)
p(f"  Circles on dest: {len(all_circles)}")
# Build set of (frozenset{member_uuids}) for quick lookup
circle_pairs = {}  # frozenset{ind1,ind2} → circle_uuid
for circ in all_circles:
    members = circ.get("accounts") or []
    member_uuids = frozenset(
        m.get("uuid","") for m in members if isinstance(m, dict) and m.get("uuid"))
    circ_uuid = circ.get("uuid","")
    if len(member_uuids) >= 2:
        circle_pairs[member_uuids] = circ_uuid

joints_need_circle = []  # {joint_src_uuid, ind1_dest, ind2_dest}
for js, jr in joint_records.items():
    if len(jr["owners"]) < 2: continue
    d1 = client_map.get(jr["owners"][0], "")
    d2 = client_map.get(jr["owners"][1], "")
    if not d1 or not d2: continue
    pair = frozenset({d1, d2})
    has_circle = any(pair.issubset(k) for k in circle_pairs)
    if not has_circle:
        joints_need_circle.append({
            "joint_src_uuid": js,
            "joint_dest_uuid": jr["dest_uuid"],
            "ind1_dest": d1,
            "ind2_dest": d2,
        })

p(f"  Joint couples without shared circle: {len(joints_need_circle)}")
for jnc in joints_need_circle:
    p(f"    joint_dest={jnc['joint_dest_uuid'][:8]}…  "
      f"ind1={jnc['ind1_dest'][:8]}…  ind2={jnc['ind2_dest'][:8]}…")


# ── B6: Joint record audit ────────────────────────────────────────────────────
p("\n=== B6: JOINT RECORD AUDIT ===")

def count_for_joint(du):
    """Return dict of counts for the 6 data types on a joint dest UUID."""
    counts = {}
    # plans
    plans = get_paged(dest, "/api/v1/plans", params={"filter[client_uuid]": du})
    counts["plans"] = len(plans)
    # charges
    charges = get_paged(dest, "/api/v2/charges", params={"filter[account_uuids]": du})
    counts["charges"] = len(charges)
    # dependants
    deps = get_paged(dest, "/api/v2/dependants", params={"filter[account_uuids]": du})
    counts["dependants"] = len(deps)
    # income
    incs = get_paged(dest, "/api/v2/income", params={"filter[account_uuids]": du})
    counts["income"] = len(incs)
    # expenditure
    exp_r = dest.get(f"/api/v1/client/{du}/expenditure")
    if exp_r.status_code == 200:
        exp_data = exp_r.json().get("data", [])
        counts["expenditure"] = len(exp_data) if isinstance(exp_data, list) else 0
    else:
        counts["expenditure"] = -1
    # notes
    notes = get_paged(dest, f"/api/v1/client/{du}/all-notes")
    counts["notes"] = len(notes)
    return counts

# Check each joint record (may take a while)
p(f"  Auditing {len(joint_records)} joint records…")
b6_zero_shared  = []
b6_zero_no_circ = []
b6_has_data     = []

for i, (js, jr) in enumerate(joint_records.items(), 1):
    du = jr["dest_uuid"]
    if i % 10 == 0:
        p(f"  {i}/{len(joint_records)}…")
    counts = count_for_joint(du)
    all_zero = all(v == 0 for v in counts.values())
    # check if owners share a circle
    d1 = client_map.get(jr["owners"][0], "") if len(jr["owners"]) > 0 else ""
    d2 = client_map.get(jr["owners"][1], "") if len(jr["owners"]) > 1 else ""
    pair = frozenset({d1, d2}) if d1 and d2 else frozenset()
    has_circle = any(pair.issubset(k) for k in circle_pairs) if pair else False
    record = {"joint_dest_uuid": du, "counts": counts,
              "has_circle": has_circle, "ind1": d1, "ind2": d2}
    if all_zero and has_circle:
        b6_zero_shared.append(record)
    elif all_zero and not has_circle:
        b6_zero_no_circ.append(record)
    else:
        b6_has_data.append(record)

p(f"\n  Joints with 0 data + shared circle (potentially removable): {len(b6_zero_shared)}")
for rec in b6_zero_shared[:10]:
    p(f"    joint={rec['joint_dest_uuid']}  "
      f"ind1={rec['ind1'][:8]}…  ind2={rec['ind2'][:8]}…")
if len(b6_zero_shared) > 10:
    p(f"    … {len(b6_zero_shared)-10} more")

p(f"\n  Joints with 0 data but NO shared circle: {len(b6_zero_no_circ)}")
for rec in b6_zero_no_circ[:10]:
    p(f"    joint={rec['joint_dest_uuid']}  "
      f"ind1={rec['ind1'][:8]}…  ind2={rec['ind2'][:8]}…")

p(f"\n  Joints with data (leave as-is): {len(b6_has_data)}")
for rec in b6_has_data[:5]:
    c = rec["counts"]
    p(f"    joint={rec['joint_dest_uuid'][:8]}…  "
      f"plans={c['plans']}  charges={c['charges']}  deps={c['dependants']}  "
      f"income={c['income']}  exp={c['expenditure']}  notes={c['notes']}")


# ═════════════════════════════════════════════════════════════════════════════
# COUNT BEFORE WRITES
# ═════════════════════════════════════════════════════════════════════════════
p("\n=== COUNTS BEFORE ===")

def count_unknown_provider():
    n = 0
    for du in dest_uuids:
        plans = get_paged(dest, "/api/v1/plans", params={"filter[client_uuid]": du})
        for plan in plans:
            prov_obj = plan.get("provider") or {}
            prov_name = (str(plan.get("provider_name") or "") or
                         (prov_obj.get("name","") if isinstance(prov_obj, dict) else ""))
            if "unknown" in prov_name.lower():
                n += 1
    return n

def count_blank_ce():
    n = 0
    for dp in list(blank_ce_plans.keys()):
        r = dest.get(f"/api/v1/plans/{dp}")
        if r.status_code == 200:
            pd2 = (r.json().get("data") or r.json())
            ce = (pd2.get("contract_enquiry_account_uuid") or
                  ((pd2.get("contract_enquiry") or {}).get("uuid") or ""))
            if not ce: n += 1
    return n

unknown_before = len(unknown_prov_plans)
blank_ce_before = len(blank_ce_plans)
joint_plans_before = len(joint_owned_plans)
charges_before = len(joint_charges)
circles_before_b5 = len(all_circles)

p(f"  Unknown Provider plans:   {unknown_before}")
p(f"  Blank contract_enquiry:   {blank_ce_before}")
p(f"  Joint-owned plans:        {joint_plans_before}")
p(f"  Charges on joint records: {charges_before}")
p(f"  Circles (before B5):      {circles_before_b5}")


# ═════════════════════════════════════════════════════════════════════════════
# TEST: Clear one Unknown Provider plan via update import (blank provider)
# ═════════════════════════════════════════════════════════════════════════════
p("\n=== PROBE: TEST CLEAR PROVIDER ===")

if not prov_clear_plans:
    p("  No plans to clear — skipping probe")
    clear_via_null = False
else:
    probe_plan = prov_clear_plans[0]
    p(f"  Probing plan {probe_plan} with blank provider_uuid in update import…")

    probe_csv = csv_bytes(
        [{"uuid": probe_plan, "provider_uuid": ""}],
        ["uuid", "provider_uuid"],
    )
    run_import("plans", "update", probe_csv, label="probe-clear-provider")

    rp = dest.get(f"/api/v1/plans/{probe_plan}")
    if rp.status_code != 200:
        stop(f"GET probe plan {probe_plan}: HTTP {rp.status_code}")
    probe_result = (rp.json().get("data") or rp.json())
    pr_prov_uuid = str(probe_result.get("provider_uuid") or "")
    pr_prov_obj  = probe_result.get("provider") or {}
    pr_prov_name = str(probe_result.get("provider_name") or
                       (pr_prov_obj.get("name","") if isinstance(pr_prov_obj, dict) else ""))
    still_unknown = ("unknown" in pr_prov_name.lower() or
                     pr_prov_uuid == unknown_prov_uuid)
    p(f"  After probe: provider_uuid={pr_prov_uuid[:8] if pr_prov_uuid else 'blank'}…  "
      f"provider_name={pr_prov_name!r}")

    if still_unknown:
        p("  Blank didn't clear — will use PUT provider_uuid: null for clear plans")
        clear_via_null = True
    else:
        p("  ✓ Blank clears provider — using update import for all clear plans")
        clear_via_null = False


# ═════════════════════════════════════════════════════════════════════════════
# WRITE: PLANS UPDATE IMPORT (combined B1+B2+B3+B4)
# ═════════════════════════════════════════════════════════════════════════════
p("\n=== WRITE: PLANS UPDATE IMPORT ===")

plan_update_rows = {}  # dest_plan_uuid → update dict

# B1 set-provider rows
for dp, prov_uuid in prov_set_plans.items():
    plan_update_rows.setdefault(dp, {})["uuid"] = dp
    plan_update_rows[dp]["provider_uuid"] = prov_uuid

# B1 clear-provider rows (only if blank works; if null needed, handled separately)
if not clear_via_null:
    for dp in prov_clear_plans:
        plan_update_rows.setdefault(dp, {})["uuid"] = dp
        plan_update_rows[dp]["provider_uuid"] = ""

# B2 contract enquiry
for dp, ce_uuid in blank_ce_plans.items():
    plan_update_rows.setdefault(dp, {})["uuid"] = dp
    plan_update_rows[dp]["contract_enquiry_account_uuid"] = ce_uuid

# B3 flexi drawdown type fix
if flexi_dest_uuid:
    plan_update_rows.setdefault(flexi_dest_uuid, {})["uuid"] = flexi_dest_uuid
    plan_update_rows[flexi_dest_uuid]["type"]     = "pension"
    plan_update_rows[flexi_dest_uuid]["sub_type"] = "personal_pension_plan"

# B4 joint ownership
for su, info in joint_owned_plans.items():
    dp = info["dest_plan_uuid"]
    d1 = info["ind1_dest"]
    d2 = info["ind2_dest"]
    if not d1 or not d2:
        p(f"  [WARN] joint plan {dp[:8]} missing individual dest UUIDs — skipping"); continue
    plan_update_rows.setdefault(dp, {})["uuid"] = dp
    plan_update_rows[dp]["client_uuids"]                = f"{d1},{d2}"
    plan_update_rows[dp]["client_ownership_percentage"] = "50,50"

total_plan_rows = len(plan_update_rows)
p(f"  Plan update rows: {total_plan_rows}")
if total_plan_rows == 0:
    p("  Nothing to update — skipping plans import")
else:
    # Determine all columns needed
    all_cols = set()
    for row in plan_update_rows.values():
        all_cols.update(row.keys())
    all_cols = ["uuid"] + sorted(c for c in all_cols if c != "uuid")

    plans_import_csv = csv_bytes(list(plan_update_rows.values()), all_cols)
    run_import("plans", "update", plans_import_csv, label="plans-update")

# If blank didn't clear, use PUT for clear-provider plans
if clear_via_null and prov_clear_plans:
    p(f"\n  Using PUT provider_uuid:null for {len(prov_clear_plans)} clear-provider plans…")
    cleared_ok = 0
    for dp in prov_clear_plans:
        rp = dest.put(f"/api/v1/plans/{dp}", json={"provider_uuid": None})
        if rp.status_code in (200, 201):
            cleared_ok += 1
        else:
            p(f"  [WARN] PUT {dp[:8]}: HTTP {rp.status_code} {rp.text[:80]}")
    p(f"  PUT cleared: {cleared_ok}/{len(prov_clear_plans)}")


# ═════════════════════════════════════════════════════════════════════════════
# WRITE: CHARGES UPDATE IMPORT (B4 — move charges from joints to first individual)
# ═════════════════════════════════════════════════════════════════════════════
p("\n=== WRITE: CHARGES UPDATE IMPORT ===")

if not joint_charges:
    p("  No charges to move — skipping")
else:
    charge_rows = []
    for cuuid, info in joint_charges.items():
        if not info["first_ind_dest"]:
            p(f"  [WARN] charge {cuuid[:8]} — no first individual dest UUID; skipping")
            continue
        charge_rows.append({
            "uuid":        cuuid,
            "client_uuid": info["first_ind_dest"],
        })
    p(f"  Charge rows to update: {len(charge_rows)}")
    if charge_rows:
        charges_csv = csv_bytes(charge_rows, ["uuid", "client_uuid"])
        run_import("charges", "update", charges_csv, label="charges-update")


# ═════════════════════════════════════════════════════════════════════════════
# WRITE: CIRCLES STORE IMPORT (B5)
# ═════════════════════════════════════════════════════════════════════════════
p("\n=== WRITE: CIRCLES CREATE IMPORT ===")

if not joints_need_circle:
    p("  No circles to create — skipping")
else:
    circle_import_rows = []
    for jnc in joints_need_circle:
        # Get entity_name for the joint record (from w2c)
        js_uuid = jnc["joint_src_uuid"]
        w2r = next((r for r in w2c_rows
                    if (r.get("external_reference_value") or "").strip() == js_uuid), {})
        # Rule 9: don't log the name; just use it in the import
        circle_import_rows.append({
            "name":          w2r.get("entity_name", f"Joint {js_uuid[:8]}"),
            "account_uuids": f"{jnc['ind1_dest']},{jnc['ind2_dest']}",
        })

    p(f"  Creating {len(circle_import_rows)} circles…")
    circles_csv = csv_bytes(circle_import_rows, ["name", "account_uuids"])
    circles_before_write = len(get_paged(dest, "/api/v1/circles", max_pages=20))
    run_import("circles", "store", circles_csv, label="circles-new")
    circles_after_write = len(get_paged(dest, "/api/v1/circles", max_pages=20))
    p(f"  Circles: {circles_before_write} → {circles_after_write} "
      f"(+{circles_after_write - circles_before_write})")
    if circles_after_write - circles_before_write != len(circle_import_rows):
        stop(f"Expected +{len(circle_import_rows)} circles, got "
             f"+{circles_after_write - circles_before_write}")
    p(f"  ✓ circles created")


# ═════════════════════════════════════════════════════════════════════════════
# COUNT AFTER
# ═════════════════════════════════════════════════════════════════════════════
p("\n=== COUNTS AFTER ===")

# B1 Unknown Provider
p("  Re-counting Unknown Provider plans…")
after_unknown = 0
for du in dest_uuids:
    plans2 = get_paged(dest, "/api/v1/plans", params={"filter[client_uuid]": du})
    for plan2 in plans2:
        prov_obj2 = plan2.get("provider") or {}
        pn2 = (str(plan2.get("provider_name") or "") or
               (prov_obj2.get("name","") if isinstance(prov_obj2, dict) else ""))
        if "unknown" in pn2.lower():
            after_unknown += 1
expected_unknown = len(prov_ar_plans) + len(prov_no_src_plan)
p(f"  Unknown Provider: {unknown_before} → {after_unknown}  (expected ≤{expected_unknown})")
if after_unknown > expected_unknown:
    stop(f"Still {after_unknown} Unknown Provider plans; expected ≤{expected_unknown}")

# B2 Contract enquiry
p("  Re-checking blank contract enquiry…")
after_blank_ce = 0
for dp in blank_ce_plans:
    r2 = dest.get(f"/api/v1/plans/{dp}")
    if r2.status_code == 200:
        pd2 = (r2.json().get("data") or r2.json())
        ce2 = (pd2.get("contract_enquiry_account_uuid") or
               ((pd2.get("contract_enquiry") or {}).get("uuid") or ""))
        if not ce2: after_blank_ce += 1
p(f"  Blank contract_enquiry: {blank_ce_before} → {after_blank_ce}  (expected 0)")
if after_blank_ce != 0:
    stop(f"Still {after_blank_ce} blank contract_enquiry plans")

# B3 Flexi drawdown
if flexi_dest_uuid:
    r3 = dest.get(f"/api/v1/plans/{flexi_dest_uuid}")
    if r3.status_code == 200:
        fp3 = (r3.json().get("data") or r3.json())
        p(f"  Flexi drawdown: type={fp3.get('type')!r}  sub_type={fp3.get('sub_type')!r}")
        if fp3.get("type") != "pension":
            stop(f"Flexi plan type={fp3.get('type')!r}; expected 'pension'")
        if "personal_pension" not in str(fp3.get("sub_type") or "").lower():
            stop(f"Flexi plan sub_type={fp3.get('sub_type')!r}; expected personal_pension*")
        p("  ✓ flexi drawdown type fixed")

# B4 Joint ownership
p("  Re-checking joint-owned plans…")
after_joint_plans = 0
for du in joint_dest_uuids:
    plans_joint = get_paged(dest, "/api/v1/plans", params={"filter[client_uuid]": du})
    after_joint_plans += len(plans_joint)
p(f"  Plans still on joint dest accounts: {after_joint_plans}  (expected 0)")
if after_joint_plans != 0:
    stop(f"Still {after_joint_plans} plans on joint records; expected 0")
p("  ✓ joint ownership transferred")

p("\n=== TASK B COMPLETE ===")
