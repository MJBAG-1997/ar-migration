"""
ZZ end-to-end test — HFP AR migration (27 Oct 2026).

Two fake ZZ households through every importer in store mode.
Inter-wave UUID resolution: after W1, GET dest client UUIDs by name;
after W5, GET dest plan UUIDs by policy_number.  All cross-entity
fields use actual destination UUIDs.

Includes one 'Add-to-Plannr' style plan (external_reference_company=
'data_migration_provider') plus a statement charge — proves those 32
charges can go in W7 rather than by hand.

Security constraints:
  - Destination writes: store-mode imports + one expenditure POST only.
  - Confirm destination by employee list, never /api/v1/firm.
  - ZZ identities are fake — OK to display in terminal (Rule 9 = PII).

STOP: after listing ZZ UUIDs for deletion.
"""

import csv, io, json, os, sys, time
import requests

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# ─── Constants ────────────────────────────────────────────────────────────────

BASE_URL          = "https://api.plannrcrm.com"
KV_NAME           = "kv-plannr-ingestion"
DEST_PAT_SECRET   = "plannr-prod-pat"
DEST_ACCOUNT_UUID = "ccd7f98b-d6a4-46ad-bab3-242ab42826eb"
MIGRATION_CO      = "data_migration"
# Add-to-Plannr plans also use "data_migration"; ext_ref_value is prefixed "prov-"
# so it doesn't collide with regular plan source UUIDs.

ADVISER_UUID = "1a230fbc-4026-4bd1-851f-18cb55ab19f2"   # Jim Hoban dest
STATUS_UUID  = "ebbdb871-c154-49a0-a212-b8dbe1874c66"   # live

EXP_SECTION_UUID = "ba6fe581-5986-49fa-a581-7e2ef54e2959"   # Essential
EXP_FIELD_UUID   = "a40acf58-f09e-423d-87e7-61a70bae939b"   # Utilities

EXPECTED_EMPLOYEES = {"jim hoban", "katie hoban", "tommie hoban", "valerie hoban"}

# ─── ZZ fake source entity UUIDs ─────────────────────────────────────────────

ZZ_CL_A1  = "00000000-0000-4000-8000-000000000001"
ZZ_CL_A2  = "00000000-0000-4000-8000-000000000002"
ZZ_CL_B1  = "00000000-0000-4000-8000-000000000003"
ZZ_CIR_1  = "00000000-0000-4000-8000-000000001001"
ZZ_DEP_1  = "00000000-0000-4000-8000-000000002001"
ZZ_PL_001 = "00000000-0000-4000-8000-000000003001"   # Joint ISA
ZZ_PL_002 = "00000000-0000-4000-8000-000000003002"   # Protection
ZZ_PL_003 = "00000000-0000-4000-8000-000000003003"   # Rathbones investment
ZZ_PORT_1 = "ZZ-PORT-001"                            # Add-to-Plannr portfolio number
ZZ_NT_001 = "00000000-0000-4000-8000-000000004001"
ZZ_NT_002 = "00000000-0000-4000-8000-000000004002"
ZZ_NT_003 = "00000000-0000-4000-8000-000000004003"

# Map: policy_number → source UUID (for plan UUID resolution after W5)
POLICY_TO_SRC = {
    "ZZ-POL-001": ZZ_PL_001,
    "ZZ-POL-002": ZZ_PL_002,
    "ZZ-POL-003": ZZ_PL_003,
    ZZ_PORT_1:    ZZ_PORT_1,
}

# Map: last_name → source UUID (for client UUID resolution after W1)
LAST_NAME_TO_SRC = {
    "HsldA-One": ZZ_CL_A1,
    "HsldA-Two": ZZ_CL_A2,
    "HsldB-One": ZZ_CL_B1,
}

def p(*a): print(*a, flush=True)


# ─── PAT resolution ───────────────────────────────────────────────────────────

def resolve_pat():
    from azure.identity import DefaultAzureCredential
    from azure.keyvault.secrets import SecretClient
    sc = SecretClient(
        vault_url=f"https://{KV_NAME}.vault.azure.net/",
        credential=DefaultAzureCredential(),
    )
    return sc.get_secret(DEST_PAT_SECRET).value


# ─── Destination client ────────────────────────────────────────────────────────

class Dest:
    def __init__(self, pat):
        self._s = requests.Session()
        self._s.headers.update({
            "Authorization": f"Bearer {pat}",
            "Accept": "application/json",
            "X-PLANNR-ACCOUNT-UUID": DEST_ACCOUNT_UUID,
        })

    def get(self, path, **kw):
        return self._s.get(f"{BASE_URL}{path}", **kw)

    def post(self, path, **kw):
        return self._s.post(f"{BASE_URL}{path}", **kw)

    def post_import(self, csv_bytes, filename, type_str):
        return self._s.post(
            f"{BASE_URL}/api/v1/import",
            files={"file": (filename, io.BytesIO(csv_bytes), "text/csv")},
            data={"type": type_str, "mode": "store"},
        )

    def post_map(self, imp_uuid, headers_map):
        return self._s.post(f"{BASE_URL}/api/v1/import/map/{imp_uuid}",
                            json={"headers": headers_map})

    def post_start(self, imp_uuid):
        return self._s.post(f"{BASE_URL}/api/v1/import/start/{imp_uuid}")

    def get_import(self, imp_uuid):
        return self._s.get(f"{BASE_URL}/api/v1/import/{imp_uuid}")


# ─── Helpers ──────────────────────────────────────────────────────────────────

def make_csv(headers, rows):
    buf = io.StringIO()
    w = csv.DictWriter(buf, fieldnames=headers, extrasaction="ignore")
    w.writeheader()
    for row in rows:
        w.writerow({h: str(row.get(h, "") or "") for h in headers})
    return buf.getvalue().encode("utf-8")


def run_import(dest, csv_bytes, filename, type_str, label=""):
    p(f"\n  → importing {label or filename}…")
    r = dest.post_import(csv_bytes, filename, type_str)
    if r.status_code not in (200, 201, 202):
        p(f"    ERROR upload: HTTP {r.status_code}: {r.text[:300]}")
        return {"status": "error", "errors": r.text[:200], "uuid": None}
    body     = r.json()
    inner    = body.get("data", body) if isinstance(body, dict) else body
    imp_uuid = inner.get("uuid") or inner.get("id")

    time.sleep(8)
    r_hdr = dest.get_import(imp_uuid)
    hdrs  = ((r_hdr.json().get("data", r_hdr.json()) if r_hdr.status_code == 200 else {})
             .get("headers", []))
    r2 = dest.post_map(imp_uuid, {h: h for h in hdrs})
    if r2.status_code not in (200, 201, 202):
        p(f"    WARN map: HTTP {r2.status_code}")

    r3 = dest.post_start(imp_uuid)
    if r3.status_code not in (200, 201, 202):
        p(f"    WARN start: HTTP {r3.status_code}")

    for _ in range(30):
        time.sleep(10)
        r4 = dest.get_import(imp_uuid)
        if r4.status_code != 200:
            continue
        d4     = r4.json().get("data", r4.json())
        status = str(d4.get("status", "") or "").lower()
        if status in ("complete", "failed", "error", "validation_complete",
                      "validation_failed", "validated"):
            errs  = d4.get("errors") or d4.get("error_count") or 0
            n_err = len(errs) if isinstance(errs, list) else int(errs or 0)
            detail = errs[:3] if isinstance(errs, list) and errs else ""
            p(f"    {status}, errors={n_err}" + (f": {detail}" if detail else ""))
            return {"status": status, "errors": n_err, "uuid": imp_uuid, "detail": errs}
    p("    WARN: timed out polling")
    return {"status": "timeout", "errors": -1, "uuid": imp_uuid}


def get_paged(dest, path, max_pages=10, params=None):
    items, page = [], 1
    while page <= max_pages:
        pr = dict(params or {})
        pr.update({"page": page, "per_page": 200})
        r = dest.get(path, params=pr)
        if r.status_code != 200:
            break
        body = r.json()
        data = body.get("data", [])
        if isinstance(data, dict):
            data = [data]
        items.extend(data)
        meta      = body.get("meta", {})
        last_page = meta.get("last_page", 1)
        if page >= last_page:
            break
        page += 1
    return items


def sub_src(val, uuid_map):
    """Replace source UUIDs with dest UUIDs; handles comma-separated lists."""
    parts  = [x.strip() for x in str(val or "").split(",") if x.strip()]
    mapped = [uuid_map.get(x, x) for x in parts]
    return ",".join(mapped)


def resolve_client_uuids(dest):
    """
    GET all clients whose first_name starts with 'ZZ-', map by last_name.
    If W1 was blocked by duplicate-client check, the pre-existing clients are used.
    If both old and new exist (6 total), takes first match per last_name.
    Returns {source_uuid: dest_uuid}.
    """
    all_cl = get_paged(dest, "/api/v1/client",
                       params={"filter[name]": "ZZ-", "per_page": 200})
    result = {}
    for c in all_cl:
        last = str(c.get("last_name", "") or "")
        src  = LAST_NAME_TO_SRC.get(last)
        if src and src not in result:   # first match wins
            result[src] = c["uuid"]
    if len(result) < 3:
        p(f"  WARN: only {len(result)} ZZ clients resolved — "
          f"last_names found: {[c.get('last_name') for c in all_cl]}")
    return result


def resolve_plan_uuids(dest):
    """
    GET all plans, filter by policy_number matching ZZ plans.
    Returns {source_uuid: dest_uuid}.
    """
    all_pl = get_paged(dest, "/api/v1/plans", max_pages=20)
    result = {}
    for pl in all_pl:
        pol = str(pl.get("policy_number", "") or "")
        src = POLICY_TO_SRC.get(pol)
        if src:
            result[src] = pl["uuid"]
    if len(result) < 4:
        pols = [pl.get("policy_number","") for pl in all_pl if str(pl.get("policy_number","")).startswith("ZZ")]
        p(f"  WARN: only {len(result)} ZZ plans resolved — found ZZ policy_numbers: {pols}")
    return result


# ─── ZZ CSV data ──────────────────────────────────────────────────────────────

CLIENT_KEYS = [
    "uuid","type","status_uuid","first_name","last_name","date_of_birth",
    "phone","email","assigned_adviser_uuid","groups","send_welcome_email",
    "external_reference_company","external_reference_value",
]

ZZ_CLIENTS = [
    {"type": "individual", "status_uuid": STATUS_UUID,
     "first_name": "ZZ-Alpha", "last_name": "HsldA-One",
     "date_of_birth": "1975-03-15", "phone": "+441234567891",
     "email": "zz.alpha1@test.invalid",
     "assigned_adviser_uuid": ADVISER_UUID, "groups": "",
     "send_welcome_email": "No",
     "external_reference_company": MIGRATION_CO,
     "external_reference_value": ZZ_CL_A1},
    {"type": "individual", "status_uuid": STATUS_UUID,
     "first_name": "ZZ-Alpha", "last_name": "HsldA-Two",
     "date_of_birth": "1977-08-22", "phone": "+441234567892",
     "email": "zz.alpha2@test.invalid",
     "assigned_adviser_uuid": ADVISER_UUID, "groups": "",
     "send_welcome_email": "No",
     "external_reference_company": MIGRATION_CO,
     "external_reference_value": ZZ_CL_A2},
    {"type": "individual", "status_uuid": STATUS_UUID,
     "first_name": "ZZ-Beta", "last_name": "HsldB-One",
     "date_of_birth": "1965-11-05", "phone": "",
     "email": "zz.beta1@test.invalid",
     "assigned_adviser_uuid": ADVISER_UUID, "groups": "",
     "send_welcome_email": "No",
     "external_reference_company": MIGRATION_CO,
     "external_reference_value": ZZ_CL_B1},
]

CIRCLE_KEYS = ["uuid","name","account_uuids","group_uuids",
               "external_reference_company","external_reference_value"]

# account_uuids use source UUIDs here; rebuilt with dest UUIDs before import
ZZ_CIRCLES = [
    {"name": "ZZ Test Circle",
     "account_uuids": f"{ZZ_CL_A1},{ZZ_CL_A2}",
     "group_uuids": "",
     "external_reference_company": MIGRATION_CO,
     "external_reference_value": ZZ_CIR_1},
]

REL_KEYS = ["account_uuid","related_account_uuid","relationship","inverse_relationship"]

# account_uuids use source UUIDs here; rebuilt with dest UUIDs before import
ZZ_RELS = [
    {"account_uuid": ZZ_CL_A1, "related_account_uuid": ZZ_CL_A2,
     "relationship": "spouse", "inverse_relationship": "spouse"},
]

DEP_KEYS = ["uuid","first_name","last_name","relationship","date_of_birth",
            "is_financially_dependant","notes","account_uuids",
            "external_reference_company","external_reference_value"]

# account_uuids use source UUIDs here; rebuilt with dest UUIDs before import
ZZ_DEPS = [
    {"first_name": "ZZ-Dep", "last_name": "Child-One",
     "relationship": "child", "date_of_birth": "2015-06-20",
     "is_financially_dependant": "True", "notes": "",
     "account_uuids": f"{ZZ_CL_A1},{ZZ_CL_B1}",
     "external_reference_company": MIGRATION_CO,
     "external_reference_value": ZZ_DEP_1},
]

PLAN_KEYS = [
    "name","uuid","client_uuids","client_ownership_percentage","type","status",
    "sub_type","opened_at","policy_number","provider","under_advice",
    "agency_in_your_name","agency_in_your_name_from","on_panel",
    "active_adviser_uuid","is_wrap_product",
    "external_reference_company","external_reference_value",
]

# client_uuids use source UUIDs here; rebuilt with dest UUIDs before import
ZZ_PLANS = [
    # Joint ISA — on_panel blank (key test), two owners
    {"name": "ZZ Joint ISA",
     "client_uuids": f"{ZZ_CL_A1},{ZZ_CL_A2}",
     "client_ownership_percentage": "50,50",
     "type": "investment", "sub_type": "isa_stocks_and_shares",
     "status": "active", "policy_number": "ZZ-POL-001",
     "provider": "Transact",
     "under_advice": "Yes", "agency_in_your_name": "Yes",
     "agency_in_your_name_from": "01/01/2020",
     "on_panel": "",
     "active_adviser_uuid": ADVISER_UUID,
     "is_wrap_product": "No",
     "external_reference_company": MIGRATION_CO,
     "external_reference_value": ZZ_PL_001},
    # Protection policy — monetary commission charge
    {"name": "ZZ Protection Life",
     "client_uuids": ZZ_CL_A1,
     "client_ownership_percentage": "",
     "type": "protection_policy", "sub_type": "whole_of_life",
     "status": "active", "policy_number": "ZZ-POL-002",
     "provider": "AIG Life",
     "under_advice": "Yes", "agency_in_your_name": "",
     "agency_in_your_name_from": "",
     "on_panel": "",
     "active_adviser_uuid": ADVISER_UUID,
     "is_wrap_product": "No",
     "external_reference_company": MIGRATION_CO,
     "external_reference_value": ZZ_PL_002},
    # Rathbones investment — quarterly % charge
    {"name": "ZZ Rathbones Portfolio",
     "client_uuids": ZZ_CL_A2,
     "client_ownership_percentage": "",
     "type": "investment", "sub_type": "general_investment_account",
     "status": "active", "policy_number": "ZZ-POL-003",
     "provider": "Rathbone Investment Management Ltd",
     "under_advice": "Yes", "agency_in_your_name": "Yes",
     "agency_in_your_name_from": "01/01/2020",
     "on_panel": "",
     "active_adviser_uuid": ADVISER_UUID,
     "is_wrap_product": "No",
     "external_reference_company": MIGRATION_CO,
     "external_reference_value": ZZ_PL_003},
    # Add-to-Plannr plan — ext_ref_company=data_migration, value prefixed "prov-"
    {"name": "ZZ Transact Portfolio",
     "client_uuids": ZZ_CL_A1,
     "client_ownership_percentage": "",
     "type": "investment", "sub_type": "general_investment_account",
     "status": "active", "policy_number": ZZ_PORT_1,
     "provider": "Transact",
     "under_advice": "Yes", "agency_in_your_name": "Yes",
     "agency_in_your_name_from": "01/01/2020",
     "on_panel": "",
     "active_adviser_uuid": ADVISER_UUID,
     "is_wrap_product": "No",
     "external_reference_company": MIGRATION_CO,
     "external_reference_value": f"prov-{ZZ_PORT_1}"},
]

VAL_KEYS = ["plan_uuid","sub_account_uuid","policy_number","valuation_date","value"]

# plan_uuid uses source UUIDs here; rebuilt with dest UUIDs before import
ZZ_VALS = [
    {"plan_uuid": ZZ_PL_001, "valuation_date": "2026-10-01", "value": "250000.00"},
    {"plan_uuid": ZZ_PL_002, "valuation_date": "2026-10-01", "value": "50000.00"},
    {"plan_uuid": ZZ_PL_003, "valuation_date": "2026-10-01", "value": "75000.00"},
    {"plan_uuid": ZZ_PORT_1, "valuation_date": "2026-10-01", "value": "100000.00"},
]

CHARGE_KEYS = [
    "uuid","plan_uuid","client_uuid","adviser_uuid","payment_origin","type","category",
    "recurring","frequency","start_date","end_date","amount","percentage",
    "percentage_frequency","tax","tax_exempt","exclude_from_rmar","advice_type",
    "service_type","rmar_type","premium_uuid",
    "advice_area_deposit","advice_area_non_regulated","advice_area_non_rmar",
    "advice_area_non_investment_insurance","advice_area_regulated_mortgage",
    "advice_area_retail_investment","advice_area_other_fca_regulated","advice_area_unknown",
    "notes",
]

def _base_charge(plan_uuid, client_uuid):
    return {
        "uuid": "", "plan_uuid": plan_uuid, "client_uuid": client_uuid,
        "adviser_uuid": ADVISER_UUID,
        "category": "ongoing_fee", "recurring": "True",
        "end_date": "", "amount": "", "tax": "",
        "tax_exempt": "Yes", "exclude_from_rmar": "No",
        "advice_type": "Independent", "service_type": "ongoing",
        "rmar_type": "Adviser Charge", "premium_uuid": "",
        "advice_area_deposit": "", "advice_area_non_regulated": "",
        "advice_area_non_rmar": "", "advice_area_non_investment_insurance": "",
        "advice_area_regulated_mortgage": "", "advice_area_retail_investment": "100%",
        "advice_area_other_fca_regulated": "", "advice_area_unknown": "", "notes": "",
    }

# plan_uuid/client_uuid use source UUIDs; rebuilt with dest UUIDs before import
ZZ_CHARGES = [
    # Sec A: provider statement annual % charge on joint ISA
    {**_base_charge(ZZ_PL_001, ZZ_CL_A1),
     "payment_origin": "Provider", "type": "Fee",
     "frequency": "quarterly", "start_date": "2024-01-01",
     "percentage": "0.5000", "percentage_frequency": "annually"},
    # Rathbones quarterly % charge
    {**_base_charge(ZZ_PL_003, ZZ_CL_A2),
     "payment_origin": "Provider", "type": "Fee",
     "frequency": "quarterly", "start_date": "2024-04-01",
     "percentage": "0.6000", "percentage_frequency": "quarterly"},
    # Protection monetary commission
    {**_base_charge(ZZ_PL_002, ZZ_CL_A1),
     "payment_origin": "Provider", "type": "Commission",
     "frequency": "monthly", "start_date": "2022-06-01",
     "amount": "100.00", "percentage": "", "percentage_frequency": ""},
    # Add-to-Plannr charge
    {**_base_charge(ZZ_PORT_1, ZZ_CL_A1),
     "payment_origin": "Provider", "type": "Fee",
     "frequency": "quarterly", "start_date": "2024-01-01",
     "percentage": "0.6000", "percentage_frequency": "annually"},
]

NOTE_KEYS = ["model_type","type","path","account_uuid","plan_uuid","risk_uuid","case_uuid",
             "task_uuid","circle_uuid","contents","is_pinned","visible_to_clients","created_at",
             "external_reference_company","external_reference_value"]

# account_uuid/plan_uuid use source UUIDs; rebuilt with dest UUIDs before import
ZZ_NOTES = [
    {"model_type": "account", "type": "note",
     "account_uuid": ZZ_CL_A1, "plan_uuid": "",
     "contents": "[task note by ZZ Test on 2026-10-27]\nZZ test task note body.",
     "is_pinned": "False", "visible_to_clients": "False",
     "created_at": "2026-10-01T09:00:00Z",
     "external_reference_company": MIGRATION_CO,
     "external_reference_value": ZZ_NT_001},
    {"model_type": "account", "type": "note",
     "account_uuid": ZZ_CL_A2, "plan_uuid": "",
     "contents": "[case note by ZZ Test on 2026-10-27]\nZZ test case note body.",
     "is_pinned": "False", "visible_to_clients": "False",
     "created_at": "2026-10-01T09:00:00Z",
     "external_reference_company": MIGRATION_CO,
     "external_reference_value": ZZ_NT_002},
    {"model_type": "plan", "type": "note",
     "account_uuid": ZZ_CL_A1, "plan_uuid": ZZ_PL_001,
     "contents": "ZZ test plan note body.",
     "is_pinned": "False", "visible_to_clients": "False",
     "created_at": "2026-10-01T09:00:00Z",
     "external_reference_company": MIGRATION_CO,
     "external_reference_value": ZZ_NT_003},
]


# ─── Main ─────────────────────────────────────────────────────────────────────

def main():
    p("=" * 68)
    p("ZZ end-to-end test — HFP AR migration 27 Oct 2026")
    p("=" * 68)

    p("\n[1] Resolving DEST PAT from Key Vault…")
    pat  = resolve_pat()
    dest = Dest(pat)
    p(f"  PAT len={len(pat)}")

    p("\n[2] Confirming destination by employee list…")
    r_emp = dest.get("/api/v1/employee")
    if r_emp.status_code != 200:
        p(f"  ERROR: GET /api/v1/employee → HTTP {r_emp.status_code}")
        sys.exit(1)
    emps = r_emp.json().get("data", [])
    emp_names = {f"{e.get('first_name','')} {e.get('last_name','')}".strip().lower()
                 for e in emps}
    missing = EXPECTED_EMPLOYEES - emp_names
    if missing:
        p(f"  ERROR: expected employees not found: {missing}")
        p(f"  Found: {sorted(emp_names)}")
        sys.exit(1)
    p(f"  OK — {len(emps)} employees, all 4 Hobans present")

    # ── W1: clients ───────────────────────────────────────────────────────────
    p("\n[3] W1 — clients…")
    results = {}
    w1_csv = make_csv(CLIENT_KEYS, ZZ_CLIENTS)
    results["W1_clients"] = run_import(dest, w1_csv, "ZZ_W1_clients.csv", "clients",
                                       "W1 clients (3)")

    # Resolve dest client UUIDs (works even if W1 was blocked by duplicate check)
    p("  Resolving dest client UUIDs…")
    cl_map = resolve_client_uuids(dest)
    p(f"  Client map: {cl_map}")
    if len(cl_map) < 3:
        p("  ERROR: could not resolve all 3 ZZ client UUIDs — aborting")
        sys.exit(1)

    # ── W2: circles ───────────────────────────────────────────────────────────
    p("\n[4] W2 — circles…")
    circles_built = [
        {**r, "account_uuids": sub_src(r["account_uuids"], cl_map)}
        for r in ZZ_CIRCLES
    ]
    results["W2_circles"] = run_import(dest,
        make_csv(CIRCLE_KEYS, circles_built),
        "ZZ_W2_circles.csv", "circles", "W2 circles (1)")

    # ── W3: relationships ─────────────────────────────────────────────────────
    p("\n[5] W3 — relationships…")
    rels_built = [
        {**r,
         "account_uuid":         cl_map.get(r["account_uuid"],         r["account_uuid"]),
         "related_account_uuid": cl_map.get(r["related_account_uuid"], r["related_account_uuid"])}
        for r in ZZ_RELS
    ]
    results["W3_rels"] = run_import(dest,
        make_csv(REL_KEYS, rels_built),
        "ZZ_W3_relationships.csv", "relationships", "W3 relationships (1)")

    # ── W3: dependants ────────────────────────────────────────────────────────
    p("\n[6] W3 — dependants…")
    deps_built = [
        {**r, "account_uuids": sub_src(r["account_uuids"], cl_map)}
        for r in ZZ_DEPS
    ]
    results["W3_deps"] = run_import(dest,
        make_csv(DEP_KEYS, deps_built),
        "ZZ_W3_dependants.csv", "dependants", "W3 dependants (1)")

    # ── W5: plans ─────────────────────────────────────────────────────────────
    p("\n[7] W5 — plans (inter-wave UUID resolution will follow)…")
    plans_built = [
        {**r, "client_uuids": sub_src(r["client_uuids"], cl_map)}
        for r in ZZ_PLANS
    ]
    results["W5_plans"] = run_import(dest,
        make_csv(PLAN_KEYS, plans_built),
        "ZZ_W5_plans.csv", "plans", "W5 plans (4)")

    # Resolve dest plan UUIDs
    p("  Resolving dest plan UUIDs…")
    pl_map = resolve_plan_uuids(dest)
    p(f"  Plan map: {pl_map}")
    if len(pl_map) < 4:
        p("  WARN: could not resolve all 4 ZZ plan UUIDs — W6-W8 may have errors")

    # ── W6: valuations ────────────────────────────────────────────────────────
    p("\n[8] W6 — valuations…")
    vals_built = [
        {**r, "plan_uuid": pl_map.get(r["plan_uuid"], r["plan_uuid"])}
        for r in ZZ_VALS
    ]
    results["W6_vals"] = run_import(dest,
        make_csv(VAL_KEYS, vals_built),
        "ZZ_W6_valuations.csv", "valuations", "W6 valuations (4)")

    # ── W7: charges ───────────────────────────────────────────────────────────
    p("\n[9] W7 — charges…")
    charges_built = [
        {**r,
         "plan_uuid":   pl_map.get(r["plan_uuid"],   r["plan_uuid"]),
         "client_uuid": cl_map.get(r["client_uuid"], r["client_uuid"])}
        for r in ZZ_CHARGES
    ]
    results["W7_charges"] = run_import(dest,
        make_csv(CHARGE_KEYS, charges_built),
        "ZZ_W7_charges.csv", "charges", "W7 charges (4)")

    # ── W8: notes ─────────────────────────────────────────────────────────────
    p("\n[10] W8 — notes…")
    notes_built = [
        {**r,
         "account_uuid": cl_map.get(r["account_uuid"], r["account_uuid"]),
         "plan_uuid":    pl_map.get(r.get("plan_uuid",""), r.get("plan_uuid",""))}
        for r in ZZ_NOTES
    ]
    results["W8_notes"] = run_import(dest,
        make_csv(NOTE_KEYS, notes_built),
        "ZZ_W8_notes.csv", "notes", "W8 notes (3)")

    # ── Expenditure via API ───────────────────────────────────────────────────
    p("\n[11] Expenditure write via API…")
    a1_dest = cl_map.get(ZZ_CL_A1)
    if not a1_dest:
        p("  WARN: no dest UUID for ZZ_CL_A1 — skipping expenditure write")
    else:
        # Get client's current life stage UUID (required for the POST path)
        r_exp_get = dest.get(f"/api/v1/client/{a1_dest}/expenditure")
        ls_uuid = (r_exp_get.json().get("current_life_stage_uuid")
                   if r_exp_get.status_code == 200 else None)
        if not ls_uuid:
            p(f"  WARN: could not get current_life_stage_uuid (HTTP {r_exp_get.status_code}) — skipping")
        else:
            exp_path = f"/api/v1/client/{a1_dest}/expenditure/{ls_uuid}/{EXP_FIELD_UUID}"
            r_exp = dest.post(exp_path, json={"amount": 1500})
            p(f"  POST {exp_path}")
            p(f"  → HTTP {r_exp.status_code}: {r_exp.text[:200]}")

    # ── GET ZZ records ────────────────────────────────────────────────────────
    p("\n[12] Getting ZZ records for verification…")

    # Full client details
    def get_client_detail(src_uuid):
        d_uuid = cl_map.get(src_uuid)
        if not d_uuid:
            return None
        r = dest.get(f"/api/v1/client/{d_uuid}")
        return r.json().get("data", r.json()) if r.status_code == 200 else None

    a1_detail = get_client_detail(ZZ_CL_A1)
    a2_detail = get_client_detail(ZZ_CL_A2)

    # Full plan details
    def get_plan_detail(src_uuid):
        d_uuid = pl_map.get(src_uuid)
        if not d_uuid:
            return None
        r = dest.get(f"/api/v1/plans/{d_uuid}")
        return r.json().get("data", r.json()) if r.status_code == 200 else None

    pl001_detail  = get_plan_detail(ZZ_PL_001)
    pl002_detail  = get_plan_detail(ZZ_PL_002)
    pl003_detail  = get_plan_detail(ZZ_PL_003)
    plport_detail = get_plan_detail(ZZ_PORT_1)

    # Charges per plan
    def get_plan_charges(plan_detail):
        if not plan_detail:
            return []
        uid = plan_detail.get("uuid")
        r   = dest.get(f"/api/v1/plans/{uid}/charges")
        if r.status_code != 200:
            return []
        d = r.json().get("data", r.json())
        return d if isinstance(d, list) else [d]

    ch_pl001  = get_plan_charges(pl001_detail)
    ch_pl002  = get_plan_charges(pl002_detail)
    ch_pl003  = get_plan_charges(pl003_detail)
    ch_plport = get_plan_charges(plport_detail)

    # Notes per client (account-level)
    def get_client_notes(src_uuid):
        d_uuid = cl_map.get(src_uuid)
        if not d_uuid:
            return []
        r = dest.get(f"/api/v1/client/{d_uuid}/all-notes")
        if r.status_code != 200:
            return []
        d = r.json().get("data", r.json())
        return d if isinstance(d, list) else [d]

    notes_a1 = get_client_notes(ZZ_CL_A1)
    notes_a2 = get_client_notes(ZZ_CL_A2)

    # ── Report ────────────────────────────────────────────────────────────────
    p("\n" + "=" * 68)
    p("ZZ TEST REPORT")
    p("=" * 68)

    p("\n── Import results ────────────────────────────────────────────────────────")
    all_ok = True
    for label, res in results.items():
        status = res.get("status", "?")
        errs   = res.get("errors", "?")
        mark   = "✓" if status == "complete" and errs == 0 else "✗"
        if mark == "✗":
            all_ok = False
        p(f"  {mark} {label}: {status}, errors={errs}")

    add_to_pl_worked = (pl_map.get(ZZ_PORT_1) is not None)
    p(f"\n  Add-to-Plannr plan found in dest: {add_to_pl_worked}")
    p(f"  Add-to-Plannr charge attached:    {len(ch_plport)} charge(s)")

    p("\n── Relationships ─────────────────────────────────────────────────────────")
    for src_uuid, label in [(ZZ_CL_A1, "ZZ_CL_A1"), (ZZ_CL_A2, "ZZ_CL_A2")]:
        d_uuid = cl_map.get(src_uuid)
        if not d_uuid:
            p(f"  {label}: NOT FOUND")
            continue
        r_lc = dest.get(f"/api/v1/client/{d_uuid}/linked-clients")
        if r_lc.status_code != 200:
            p(f"  {label}: GET /linked-clients HTTP {r_lc.status_code}")
            continue
        linked = r_lc.json() if isinstance(r_lc.json(), list) else r_lc.json().get("data", [])
        if not linked:
            p(f"  {label}: no linked clients")
        for lc in linked:
            p(f"  {label} ({d_uuid}): linked to {lc.get('uuid','')} "
              f"'{lc.get('name','')}' as {lc.get('relationship','')}")

    p("\n── on_panel display ──────────────────────────────────────────────────────")
    for label, detail in [("ZZ_PL_001 joint ISA",       pl001_detail),
                           ("ZZ_PL_002 protection",      pl002_detail),
                           ("ZZ_PL_003 Rathbones",       pl003_detail),
                           ("ZZ_PORT_1 Transact",        plport_detail)]:
        if not detail:
            p(f"  {label}: NOT FOUND")
            continue
        p(f"  {label}: on_panel={detail.get('on_panel')!r}")

    p("\n── Groups ────────────────────────────────────────────────────────────────")
    for label, detail in [("ZZ_CL_A1", a1_detail), ("ZZ_CL_A2", a2_detail)]:
        if not detail:
            p(f"  {label}: NOT FOUND")
            continue
        p(f"  {label}: groups={detail.get('groups')!r}")

    p("\n── Charges (start_date / % / amount / freq) ──────────────────────────────")
    for plan_label, charges in [("ZZ_PL_001 joint ISA",  ch_pl001),
                                  ("ZZ_PL_002 protection", ch_pl002),
                                  ("ZZ_PL_003 Rathbones",  ch_pl003),
                                  ("ZZ_PORT_1 Transact",   ch_plport)]:
        if not charges:
            p(f"  {plan_label}: no charges")
            continue
        for ch in charges:
            p(f"  {plan_label}: type={ch.get('type')} start={ch.get('start_date')} "
              f"pct={ch.get('percentage')} pct_freq={ch.get('percentage_frequency')} "
              f"amt={ch.get('amount')} freq={ch.get('frequency')}")

    p("\n── Notes (first 120 chars) ───────────────────────────────────────────────")
    for label, notes in [("ZZ_CL_A1", notes_a1), ("ZZ_CL_A2", notes_a2)]:
        zz_notes = [n for n in notes if "ZZ test" in str(n.get("contents",""))
                    or str(n.get("external_reference_value","")).startswith("00000000-0000-4000-8000-0000000040")]
        if not zz_notes:
            p(f"  {label}: no ZZ notes found")
        for n in zz_notes[:3]:
            c    = str(n.get("contents") or "")
            mtype = n.get("notable_type") or n.get("model_type") or n.get("type") or "?"
            p(f"  {label} [{mtype}]: {c[:120]!r}")

    # Plan note may be on the plan directly, also check account notes
    pl_note = next((n for n in notes_a1
                    if n.get("model_type") == "plan"
                    and "ZZ test plan note" in str(n.get("contents",""))), None)
    pn_body = pl_note.get("contents","")[:80] if pl_note else "not found (may be on plan directly)"
    p(f"  ZZ_PL_001 plan note in a1 notes: {pn_body!r}")

    p("\n── Dates and phone numbers ───────────────────────────────────────────────")
    for label, detail in [("ZZ_CL_A1", a1_detail), ("ZZ_CL_A2", a2_detail)]:
        if not detail:
            p(f"  {label}: NOT FOUND")
            continue
        p(f"  {label}: DOB={detail.get('date_of_birth')} phone={detail.get('phone')!r}")

    # ── ZZ UUIDs to delete ────────────────────────────────────────────────────
    p("\n" + "=" * 68)
    p("ZZ UUIDs TO DELETE")
    p("=" * 68)

    # Clients — all ZZ clients in dest (may include both old and new if both were created)
    p("\n  [clients]")
    all_zz_cl = get_paged(dest, "/api/v1/client",
                          params={"filter[name]": "ZZ-", "per_page": 200})
    for c in all_zz_cl:
        last  = str(c.get("last_name","") or "")
        first = str(c.get("first_name","") or "")
        p(f"  client  {first} {last}: {c['uuid']}")

    # Plans
    if pl_map:
        p("\n  [plans]")
        for src, d_uuid in sorted(pl_map.items()):
            label = {ZZ_PL_001: "001 joint ISA",
                     ZZ_PL_002: "002 protection",
                     ZZ_PL_003: "003 Rathbones",
                     ZZ_PORT_1: "PORT Transact"}.get(src, src)
            p(f"  plan    {label}: {d_uuid}")

    # Dependants
    all_deps = get_paged(dest, "/api/v2/dependants", params={"include": "accounts"})
    zz_deps = [d for d in all_deps
               if str(d.get("external_reference_value","")).startswith("00000000-0000-4000-8000-0000000020")]
    for d in zz_deps:
        p(f"  dep     ZZ_DEP_1: {d['uuid']}")

    # Circles
    if a1_detail:
        for circ in (a1_detail.get("circles") or []):
            if "ZZ" in str(circ.get("name","")):
                p(f"  circle  {circ.get('name')}: {circ.get('uuid')}")

    # Charges
    for plan_label, charges in [("plan 001", ch_pl001), ("plan 002", ch_pl002),
                                  ("plan 003", ch_pl003), ("plan PORT", ch_plport)]:
        for ch in charges:
            p(f"  charge  on {plan_label}: {ch.get('uuid')}")

    # Notes
    for label, notes in [("a1", notes_a1), ("a2", notes_a2)]:
        for n in notes:
            if "ZZ test" in str(n.get("contents","")) or \
               str(n.get("external_reference_value","")).startswith("00000000-0000-4000-8000-0000000040"):
                p(f"  note    {label}: {n.get('uuid')}")

    # Expenditure
    if a1_dest:
        r_exp_list = dest.get(f"/api/v1/client/{a1_dest}/expenditure")
        if r_exp_list.status_code == 200:
            for e in (r_exp_list.json().get("data", []) or []):
                # amount stored in pence: 1500 pence = £15 (test amount)
                if str(e.get("amount", {}).get("amount", "") if isinstance(e.get("amount"), dict) else e.get("amount", "")) == "1500":
                    p(f"  expenditure a1 utilities (1500p = £15): {e.get('uuid')}")

    p("\n" + "=" * 68)
    p("STOP — ZZ test complete. Delete ZZ records before running W1–W8.")
    p("=" * 68)


if __name__ == "__main__":
    main()
