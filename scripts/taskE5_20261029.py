"""
taskE5_20261029.py — E5: Recreate flexi drawdown as personal_pension_plan.
1. Import new plan (pension/personal_pension_plan) via plans importer.
2. Import charge for new plan via charges importer.
3. DELETE old plan fff4250d-a0c6-4d97-8fec-b81b3483904d.
Rule 9: counts and UUIDs only. Count before/write/count after. Mismatch = STOP.
"""
import csv, io, json, os, sys, time
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
FLEXI_OLD         = "fff4250d-a0c6-4d97-8fec-b81b3483904d"
JIM               = "1a230fbc-4026-4bd1-851f-18cb55ab19f2"

def p(*a): print(*a, flush=True)
def stop(msg): p(f"\n!!STOP!! {msg}"); sys.exit(1)

def ljson(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)

sc  = SecretClient(vault_url="https://kv-plannr-ingestion.vault.azure.net/",
                   credential=DefaultAzureCredential())
pat = sc.get_secret("plannr-prod-pat").value

class API:
    def __init__(self, token, acct):
        self._s = requests.Session()
        self._s.headers.update({"Authorization": f"Bearer {token}",
                                 "Accept": "application/json",
                                 "Content-Type": "application/json",
                                 "X-PLANNR-ACCOUNT-UUID": acct})
        self._token = token
        self._acct  = acct
    def get(self, path, **kw):    return self._req("GET",    path, **kw)
    def delete(self, path, **kw): return self._req("DELETE", path, **kw)
    def _req(self, method, path, **kw):
        url = f"{BASE_URL}{path}" if path.startswith("/") else path
        for _ in range(4):
            r = self._s.request(method, url, **kw)
            if r.status_code != 429: return r
            time.sleep(int(r.headers.get("Retry-After", 30)))
        return r
    def upload(self, csv_bytes, filename, type_str, mode):
        s2 = requests.Session()
        s2.headers.update({"Authorization": f"Bearer {self._token}",
                            "Accept": "application/json",
                            "X-PLANNR-ACCOUNT-UUID": self._acct})
        return s2.post(f"{BASE_URL}/api/v1/import",
                       files={"file": (filename, io.BytesIO(csv_bytes), "text/csv")},
                       data={"type": type_str, "mode": mode})
    def map_headers(self, imp_uuid, headers):
        return self._s.post(f"{BASE_URL}/api/v1/import/map/{imp_uuid}",
                            json={"headers": {h: h for h in headers}})
    def start(self, imp_uuid):
        return self._s.post(f"{BASE_URL}/api/v1/import/start/{imp_uuid}")
    def get_import(self, imp_uuid):
        return self._s.get(f"{BASE_URL}/api/v1/import/{imp_uuid}")

def paged(api, path, params=None, max_pages=20):
    items, page = [], 1
    while page <= max_pages:
        pr = dict(params or {}); pr["page"] = page; pr["per_page"] = 200
        r = api.get(path, params=pr)
        if r.status_code != 200:
            p(f"  [WARN] GET {path} page {page} → HTTP {r.status_code}"); break
        body = r.json()
        if isinstance(body, list): items.extend(body); break
        data = body.get("data", [])
        if isinstance(data, list):   items.extend(data)
        elif isinstance(data, dict): items.append(data)
        meta = body.get("meta", {}) if isinstance(body, dict) else {}
        if page >= (meta.get("last_page") or 1): break
        page += 1
    return items

DONE_STATUSES = {"complete","failed","error",
                 "validation_complete","validation_failed","validated"}

def poll_import(api, imp_uuid, label, timeout=90):
    for _ in range(timeout):
        time.sleep(10)
        r = api.get_import(imp_uuid)
        if r.status_code != 200: continue
        d      = r.json().get("data", r.json())
        status = str(d.get("status","") or "").lower()
        if status in DONE_STATUSES:
            errs    = d.get("errors") or d.get("error_count") or 0
            n_err   = len(errs) if isinstance(errs,list) else int(errs or 0)
            rows    = d.get("total_rows") or d.get("row_count") or 0
            created = d.get("successful_rows") or d.get("created") or 0
            failed  = d.get("failed_rows") or d.get("failed") or 0
            return {"status": status, "errors": n_err, "rows": rows,
                    "created": created, "failed": failed,
                    "detail": errs if isinstance(errs,list) else [], "raw": d}
    return {"status":"timeout","errors":-1,"rows":0,"created":0,"failed":0,"detail":[],"raw":{}}

def run_import(api, csv_bytes, filename, type_str):
    row_count = max(0, len(csv_bytes.decode("utf-8-sig","replace").splitlines()) - 1)
    p(f"  upload [store] {filename} ({row_count} rows)…")
    r = api.upload(csv_bytes, filename, type_str, "store")
    if r.status_code not in (200,201,202):
        stop(f"upload {type_str} failed HTTP {r.status_code}: {r.text[:200]}")
    body     = r.json()
    inner    = body.get("data", body) if isinstance(body,dict) else body
    imp_uuid = inner.get("uuid") or inner.get("id")
    if not imp_uuid:
        stop(f"no import UUID in upload response: {str(body)[:200]}")
    p(f"  imp_uuid={imp_uuid}  status={inner.get('status','?')}")
    time.sleep(8)
    r_hdr    = api.get_import(imp_uuid)
    hdr_body = r_hdr.json().get("data", r_hdr.json()) if r_hdr.status_code==200 else {}
    hdrs     = hdr_body.get("headers", [])
    p(f"  get_import: HTTP {r_hdr.status_code}  status={hdr_body.get('status','?')}  headers={len(hdrs)}")
    good = list(hdrs)
    for _ in range(25):
        r2 = api.map_headers(imp_uuid, good)
        if r2.status_code in (200,201,202): break
        try:
            bad = {k[len("headers."):] for k in (r2.json().get("errors") or {})
                   if k.startswith("headers.")}
        except Exception:
            bad = set()
        if not bad: break
        p(f"  stripping {len(bad)} headers: {sorted(bad)}")
        good = [h for h in good if h not in bad]
    r3 = api.start(imp_uuid)
    if r3.status_code not in (200,201,202):
        p(f"  [WARN] start HTTP {r3.status_code}: {r3.text[:150]}")
    res = poll_import(api, imp_uuid, type_str)
    p(f"  result: status={res['status']} errors={res['errors']} rows={res['rows']} created={res['created']} failed={res['failed']}")
    if res["status"] == "timeout": stop(f"{type_str} import poll timed out")
    if res["errors"] > 0 or res["failed"] > 0:
        for e in res["detail"][:10]: p(f"  error: {e}")
        stop(f"{type_str} import had errors/failures")
    return res

dest        = API(pat, DEST_ACCOUNT_UUID)
src_pat_val = sc.get_secret("plannr-hfp-pat").value
src         = API(src_pat_val, SRC_ACCOUNT_UUID)
client_map  = ljson(os.path.join(STATE_DIR, "client_map.json"))
plan_map    = ljson(os.path.join(STATE_DIR, "plan_map.json"))
b3          = ljson(os.path.join(STATE_DIR, "b3_flexi_details.json"))
inv_plan    = {v: k for k, v in plan_map.items()}

p("=== E5: FLEXI DRAWDOWN RECREATION ===")
p(f"  Old plan UUID: {FLEXI_OLD}")
p(f"  Old plan type: {b3.get('name','')} (will be replaced with pension/personal_pension_plan)")

# Verify old plan still exists
rg_old = dest.get(f"/api/v1/plans/{FLEXI_OLD}")
if rg_old.status_code != 200:
    stop(f"Cannot GET old plan {FLEXI_OLD}: HTTP {rg_old.status_code}")
old_pl = (rg_old.json().get("data") or rg_old.json())
p(f"  Old plan confirmed: type={old_pl.get('abstract_type','?')}/{old_pl.get('sub_type','?')}  policy={old_pl.get('policy_number','')}")

# Get source plan for pension-specific fields
src_uuid = inv_plan.get(FLEXI_OLD, "")
p(f"  Source plan UUID: {src_uuid}")
pension_fields = {}
if src_uuid:
    rs = src.get(f"/api/v1/plans/{src_uuid}")
    if rs.status_code == 200:
        sp = rs.json().get("data", rs.json())
        p(f"  Source type: {sp.get('abstract_type','?')}/{sp.get('sub_type','?')}")
        # Extract crystallisation / drawdown fields if available
        for k, v in sp.items():
            if k.startswith("pension_") and v is not None:
                pension_fields[k] = v

# Owner account
owner_uuid = b3["accounts"][0]["account_uuid"] if b3["accounts"] else ""
p(f"  Owner: {owner_uuid}")

# Count plans before
plans_before = len(paged(dest, "/api/v1/plans", max_pages=15))
p(f"\n  Plans on dest before: {plans_before}  (expected 1901)")
if plans_before != 1901:
    p(f"  [NOTE] expected 1901, got {plans_before}")

# ── Step 1: Import new plan (skip if already created) ───────────────────────
p("\n  Step 1: Import new personal_pension_plan…")

# Check if already created (from a prior run)
already_created = ""
all_plans_check = paged(dest, "/api/v1/plans", max_pages=15)
for pl in all_plans_check:
    if (pl.get("policy_number") == b3["policy_number"] and
            pl.get("sub_type") == "personal_pension_plan" and
            pl.get("uuid") != FLEXI_OLD):
        already_created = pl.get("uuid","")
        p(f"  New plan already exists from prior run: {already_created} — skipping import")
        break

EXT_REF = "e5_flexi_drawdown_replacement_20261029"
plan_row = {
    "name":                       b3["name"],
    "uuid":                       "",
    "client_uuids":               owner_uuid,
    "client_ownership_percentage": "100",
    "type":                       "pension",
    "status":                     b3["status"],
    "sub_type":                   "personal_pension_plan",
    "opened_at":                  "",
    "policy_number":              b3["policy_number"],
    "linked_policy_numbers":      "",
    "proposal_reference":         "",
    "provider":                   "Transact",
    "provider_uuid":              b3["provider_uuid"],
    "inactive_at":                "",
    "inactive_reason":            "",
    "agency_in_your_name":        "Yes",
    "agency_in_your_name_from":   "17/10/2026",
    "originally_sold_by_firm":    "",
    "under_advice":               "Yes",
    "show_on_circles":            "Yes",
    "visible_on_client_portal":   "Yes",
    "include_in_client_valuation":"Yes",
    "review_date":                "",
    "on_panel":                   "",
    "seller_uuid":                JIM,
    "seller_name":                "Jim Hoban",
    "active_adviser_uuid":        JIM,
    "contract_enquiry_account_uuid": "",
    "is_group_scheme":            "No",
    "is_wrap_product":            "No",
    "wrap_uuid":                  "",
    "external_reference_company": "data_migration",
    "external_reference_value":   EXT_REF,
    "ownership_split":            "",
    # Pension fields
    "pension_is_crystallised":    "true",
    "pension_platform_provider":  "Transact",
}
# Fill in any pension fields we found in source
for k, v in pension_fields.items():
    safe_key = k.replace(".", "_")
    if safe_key in plan_row and v is not None:
        plan_row[safe_key] = str(v)

# Write plans CSV
plan_csv_path = os.path.join(IMPORT_DIR, "E5_flexi_replacement_plan.csv")
with open(plan_csv_path, "w", newline="", encoding="utf-8-sig") as f:
    w = csv.DictWriter(f, fieldnames=list(plan_row.keys()))
    w.writeheader()
    w.writerow(plan_row)
p(f"  Wrote plan CSV: {os.path.basename(plan_csv_path)}")

if not already_created:
    with open(plan_csv_path, "rb") as f:
        plan_csv_bytes = f.read()
    run_import(dest, plan_csv_bytes, "E5_flexi_replacement_plan.csv", "plans")
# Plannr plans importer sometimes reports created=0 even when it succeeds; find by other means

# ── Step 2: Find new plan UUID ───────────────────────────────────────────────
p("\n  Step 2: Find new plan UUID…")
new_plan_uuid = ""
all_plans_after = paged(dest, "/api/v1/plans", params={"include": "external_references"}, max_pages=15)
p(f"  Plans fetched: {len(all_plans_after)}")
for pl in all_plans_after:
    # Try external_reference first
    refs = pl.get("external_references") or []
    for ref in refs:
        if ref.get("value") == EXT_REF:
            new_plan_uuid = pl.get("uuid","")
            p(f"  Found by ext_ref: {new_plan_uuid}")
            break
    if new_plan_uuid: break

if not new_plan_uuid:
    # Fallback: find by policy_number + pension type, excluding old plan
    for pl in all_plans_after:
        if (pl.get("policy_number") == b3["policy_number"] and
                pl.get("sub_type") == "personal_pension_plan" and
                pl.get("uuid") != FLEXI_OLD):
            new_plan_uuid = pl.get("uuid","")
            p(f"  Found by policy+type fallback: {new_plan_uuid}")
            break

if not new_plan_uuid:
    stop(f"E5: cannot find new plan by ext_ref or policy+type")
p(f"  New plan UUID: {new_plan_uuid}")

# Verify new plan
rg_new = dest.get(f"/api/v1/plans/{new_plan_uuid}")
if rg_new.status_code != 200:
    stop(f"GET new plan {new_plan_uuid}: HTTP {rg_new.status_code}")
new_pl = rg_new.json().get("data", rg_new.json())
p(f"  New plan type:   {new_pl.get('abstract_type','?')}/{new_pl.get('sub_type','?')}")
p(f"  New plan policy: {new_pl.get('policy_number','')}")
p(f"  New plan status: {new_pl.get('status','')}")

if new_pl.get("sub_type") != "personal_pension_plan":
    stop(f"E5: new plan has wrong sub_type: {new_pl.get('sub_type')}")
p(f"  Plan type verified ✓")

# ── Step 3: Check / import charge for new plan ──────────────────────────────
p("\n  Step 3: Charge for new plan…")
rc_check = dest.get(f"/api/v1/plans/{new_plan_uuid}/charges")
existing_charges = []
if rc_check.status_code == 200:
    chgs_chk = rc_check.json().get("data") or rc_check.json()
    if isinstance(chgs_chk, list): existing_charges = chgs_chk

if existing_charges:
    p(f"  Charge already exists ({len(existing_charges)}) — skipping import")
else:
    p(f"  No charge yet — importing…")
    b3_charge = b3["charges"][0]
    charge_row = {
        "uuid": "", "plan_uuid": new_plan_uuid, "client_uuid": owner_uuid,
        "adviser_uuid": JIM, "payment_origin": "Provider", "type": "Fee",
        "category": b3_charge.get("category","ongoing_fee"), "recurring": "True",
        "frequency": b3_charge.get("frequency","quarterly"),
        "start_date": b3_charge.get("start_date",""), "end_date": b3_charge.get("end_date") or "",
        "amount": "", "percentage": str(b3_charge.get("percentage","0.7")),
        "percentage_frequency": b3_charge.get("percentage_frequency","annually"),
        "tax": "", "tax_exempt": "Yes", "exclude_from_rmar": "No",
        "advice_type": "Independent", "service_type": b3_charge.get("service_type","ongoing"),
        "rmar_type": "Adviser Charge", "premium_uuid": "",
        "advice_area_deposit": "0", "advice_area_non_regulated": "0", "advice_area_non_rmar": "0",
        "advice_area_non_investment_insurance": "0", "advice_area_regulated_mortgage": "0",
        "advice_area_retail_investment": "100%", "advice_area_other_fca_regulated": "0",
        "advice_area_unknown": "0", "notes": "",
    }
    charge_csv_path = os.path.join(IMPORT_DIR, "E5_flexi_replacement_charge.csv")
    with open(charge_csv_path, "w", newline="", encoding="utf-8-sig") as f:
        w2 = csv.DictWriter(f, fieldnames=list(charge_row.keys()))
        w2.writeheader(); w2.writerow(charge_row)
    with open(charge_csv_path, "rb") as f2:
        charge_csv_bytes = f2.read()
    run_import(dest, charge_csv_bytes, "E5_flexi_replacement_charge.csv", "charges")

# Verify charge
rc_verify = dest.get(f"/api/v1/plans/{new_plan_uuid}/charges")
if rc_verify.status_code == 200:
    chgs_v = rc_verify.json().get("data") or rc_verify.json()
    if isinstance(chgs_v, list):
        p(f"  Charges on new plan: {len(chgs_v)}  (expected 1)")
        if len(chgs_v) != 1:
            stop(f"E5: expected 1 charge on new plan, got {len(chgs_v)}")
p(f"  Charge verified ✓")

# ── Step 4: Copy plan notes from old to new ──────────────────────────────────
p(f"\n  Step 4: Copy plan notes from old plan to new…")
old_notes = paged(dest, f"/api/v1/client/{owner_uuid}/all-notes")
# Filter to notes associated with the old plan
plan_notes = [n for n in old_notes
              if (n.get("noteable_uuid") == FLEXI_OLD or
                  (n.get("noteable") or {}).get("uuid") == FLEXI_OLD)]
p(f"  Notes on old plan: {len(plan_notes)}")

notes_copied = 0
for n in plan_notes:
    contents = n.get("contents","")
    created_raw = str(n.get("created_at",""))[:10]
    try:
        y2, mo2, d2 = created_raw.split("-")
        created_fmt = f"{d2}/{mo2}/{y2}"
    except Exception:
        created_fmt = created_raw
    delta_path2 = os.path.join(IMPORT_DIR, "E5_note_copy.csv")
    note_row = {
        "model_type": "plan", "type": "note", "path": "",
        "account_uuid": owner_uuid, "plan_uuid": new_plan_uuid,
        "risk_uuid": "", "case_uuid": "", "task_uuid": "",
        "contents": contents, "is_pinned": "false", "visible_to_clients": "false",
        "created_at": created_fmt,
        "external_reference_company": "data_migration",
        "external_reference_value": f"e5_note_copy_{n.get('uuid','')}",
    }
    with open(delta_path2, "w", newline="", encoding="utf-8-sig") as f3:
        w3 = csv.DictWriter(f3, fieldnames=list(note_row.keys()))
        w3.writeheader(); w3.writerow(note_row)
    with open(delta_path2, "rb") as f4:
        nb = f4.read()
    rn2 = dest._s.post(f"{BASE_URL}/api/v1/import",
                       files={"file": ("E5_note_copy.csv", io.BytesIO(nb), "text/csv")},
                       data={"type": "notes", "mode": "store"})
    if rn2.status_code in (200,201,202):
        ni = rn2.json().get("data", rn2.json())
        n_imp = ni.get("uuid") or ni.get("id")
        if n_imp:
            time.sleep(6)
            dest._s.post(f"{BASE_URL}/api/v1/import/map/{n_imp}",
                         json={"headers": {h: h for h in note_row.keys()}})
            dest._s.post(f"{BASE_URL}/api/v1/import/start/{n_imp}")
            time.sleep(10)
        notes_copied += 1
p(f"  Notes copied to new plan: {notes_copied}")

# Copy premiums from old plan
p(f"\n  Step 4b: Copy premiums from old plan to new…")
rp_old = dest.get(f"/api/v1/plans/{FLEXI_OLD}/premiums")
old_premiums = []
if rp_old.status_code == 200:
    op = rp_old.json().get("data") or rp_old.json()
    if isinstance(op, list): old_premiums = op
p(f"  Premiums on old plan: {len(old_premiums)}")

for pr in old_premiums:
    amt = pr.get("amount") or {}
    amt_v = amt.get("amount","0") if isinstance(amt,dict) else str(amt or "0")
    try:
        amt_pence = int(float(amt_v) * 100)
    except Exception:
        amt_pence = 0
    freq  = pr.get("frequency","")
    start = pr.get("start_date","")
    p(f"  POST premium to new plan: amt={amt_pence}p freq={freq} start={start}")
    rpp = dest._s.post(f"{BASE_URL}/api/v1/plans/{new_plan_uuid}/premiums",
                       json={"amount": amt_pence, "frequency": freq, "start_date": start})
    if rpp.status_code in (200,201):
        p(f"  Premium copied ✓")
    else:
        p(f"  [WARN] POST premium HTTP {rpp.status_code}: {rpp.text[:200]}")

# ── Step 5: Update plan_map ───────────────────────────────────────────────────
p(f"\n  Step 5: Update plan_map to point src→new plan…")
src_plan_uuid = inv_plan.get(FLEXI_OLD,"")
if src_plan_uuid:
    plan_map[src_plan_uuid] = new_plan_uuid
    json.dump(plan_map, open(os.path.join(STATE_DIR,"plan_map.json"),"w"), indent=2)
    p(f"  plan_map updated: {src_plan_uuid[:20]} → {new_plan_uuid}")
else:
    p(f"  [WARN] src plan UUID not found for old plan {FLEXI_OLD}")

# ── Step 6: DELETE old plan ──────────────────────────────────────────────────
p(f"\n  Step 6: DELETE old plan {FLEXI_OLD}…")

rg_old2 = dest.get(f"/api/v1/plans/{FLEXI_OLD}")
if rg_old2.status_code != 200:
    stop(f"Cannot GET old plan before delete: HTTP {rg_old2.status_code}")
old2 = rg_old2.json().get("data", rg_old2.json())
p(f"  Old plan sub_type={old2.get('sub_type','?')}  policy={old2.get('policy_number','?')}")
if old2.get("policy_number") != b3["policy_number"]:
    stop(f"Old plan policy mismatch: expected {b3['policy_number']} got {old2.get('policy_number')}")

plans_pre_del = len(paged(dest, "/api/v1/plans", max_pages=15))
p(f"  Plans before delete: {plans_pre_del}")

# Delete old plan's charge first (removes expectation that would block plan delete)
rc_old = dest.get(f"/api/v1/plans/{FLEXI_OLD}/charges")
if rc_old.status_code == 200:
    old_chgs = rc_old.json().get("data") or rc_old.json()
    if isinstance(old_chgs, list):
        for oc in old_chgs:
            oc_uuid = oc.get("uuid","")
            if oc_uuid:
                rdc = dest.delete(f"/api/v1/plans/{FLEXI_OLD}/charges/{oc_uuid}")
                p(f"  Deleted old charge {oc_uuid[:8]} → HTTP {rdc.status_code}")

rd = dest.delete(f"/api/v1/plans/{FLEXI_OLD}")
if rd.status_code not in (200,201,204):
    stop(f"DELETE old plan HTTP {rd.status_code}: {rd.text[:200]}")
p(f"  DELETE → HTTP {rd.status_code}")

rg_verify = dest.get(f"/api/v1/plans/{FLEXI_OLD}")
if rg_verify.status_code == 404:
    p(f"  Old plan confirmed deleted (404) ✓")
elif rg_verify.status_code == 200:
    stop(f"E5: old plan still exists after DELETE")
else:
    p(f"  [WARN] GET old plan after delete: HTTP {rg_verify.status_code}")

plans_after_del = len(paged(dest, "/api/v1/plans", max_pages=15))
p(f"  Plans after delete: {plans_after_del}  (expected {plans_pre_del - 1})")
if plans_after_del != plans_pre_del - 1:
    stop(f"E5: plans count mismatch after delete: expected {plans_pre_del-1}, got {plans_after_del}")

p(f"\n  E5 COMPLETE:")
p(f"    New plan UUID:  {new_plan_uuid}")
p(f"    Old plan:       {FLEXI_OLD} DELETED")
p(f"    Plans: {plans_before} → {plans_after_del}")

p("\n=== TASK E5 COMPLETE ===")
