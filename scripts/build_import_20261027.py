"""
HFP AR Migration — Build import files (27 Oct 2026).

Pre-build checks 1 & 2, build W1-W8 CSVs, validate W1.

Rule 9: aggregates only in terminal; PII to DATA_ROOT Excel only.
D19: expenditure via API — no W3_expenses.csv produced here.
"""
import csv, json, os, re, sys, time
from collections import defaultdict

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# ─── Paths ────────────────────────────────────────────────────────────────────
DATA_ROOT    = r"C:\dev\AR Migration"
EXPORTS_26   = os.path.join(DATA_ROOT, "exports_20261026")
EXPORTS_14   = os.path.join(DATA_ROOT, "exports_20261014")
SNAP_DIR     = os.path.join(DATA_ROOT, "snapshot_20261012T090906Z")
RETURNED_DIR = os.path.join(DATA_ROOT, "returned_20261026")
IMPORT_DIR   = os.path.join(DATA_ROOT, "import_20261027")
LISTS_XLSX   = os.path.join(DATA_ROOT, "hfp_build_lists_20261027.xlsx")
CHANGES_XLSX = os.path.join(DATA_ROOT, "hfp_changes_20261026.xlsx")

SENT_CLIENTS_XLSX = os.path.join(DATA_ROOT, "Plannr_Clients.xlsx")
SENT_PLANS_XLSX   = os.path.join(DATA_ROOT, "Plannr_Plans.xlsx")
SENT_CHARGES_XLSX = os.path.join(DATA_ROOT, "Plannr_Charges.xlsx")
RETURNED_CHARGES  = os.path.join(RETURNED_DIR, "Plannr Charges.xlsx")
RETURNED_CLIENTS  = os.path.join(RETURNED_DIR, "Clients with no Adviser.xlsx")
SIGNOFF_V6_XLSX   = os.path.join(DATA_ROOT, "hfp_signoff_pack_v6.xlsx")
TRANSACT_RATES_XLSX = os.path.join(DATA_ROOT,
    "Hoban_Financial_Planning_Ltd_260922_Payment_Rates_113727875.xlsx")

BASE_URL          = "https://api.plannrcrm.com"
KV_NAME           = "kv-plannr-ingestion"
SRC_PAT_SECRET    = "plannr-hfp-pat"
DEST_PAT_SECRET   = "plannr-prod-pat"
DEST_ACCOUNT_UUID = "ccd7f98b-d6a4-46ad-bab3-242ab42826eb"
MIGRATION_CO      = "data_migration"

# ─── Mappings ────────────────────────────────────────────────────────────────
ADVISER_MAP = {
    "db0ab7f5-8e52-4c89-a915-29fefeeca0d4": "1a230fbc-4026-4bd1-851f-18cb55ab19f2",
    "2153bb3c-7b76-4643-86a3-a9c53acab0b0": "78a6aab8-55f3-42ed-b8ef-4867da1c0c78",
    "5a8bc24f-8ff9-4671-ba49-4a5c8bfc5b2b": "49b93c74-dffa-4abe-8aa7-3e58b62f3375",
    "6dbac8ff-90d2-4dc8-b60d-26567773a68e": "380dd547-558e-4c99-8382-c0534e6b871c",
}
ADVISER_NAME_DEST = {
    "jim":           "1a230fbc-4026-4bd1-851f-18cb55ab19f2",
    "jim hoban":     "1a230fbc-4026-4bd1-851f-18cb55ab19f2",
    "katie":         "78a6aab8-55f3-42ed-b8ef-4867da1c0c78",
    "katie hoban":   "78a6aab8-55f3-42ed-b8ef-4867da1c0c78",
    "tommie":        "49b93c74-dffa-4abe-8aa7-3e58b62f3375",
    "tommie hoban":  "49b93c74-dffa-4abe-8aa7-3e58b62f3375",
    "tomme":         "49b93c74-dffa-4abe-8aa7-3e58b62f3375",
    "valerie":       "380dd547-558e-4c99-8382-c0534e6b871c",
    "valerie hoban": "380dd547-558e-4c99-8382-c0534e6b871c",
}
STATUS_MAP = {
    "605a9298-486b-4726-ba6e-8b822fdc0cb7": "ebbdb871-c154-49a0-a212-b8dbe1874c66",
    "c12cdf7e-0b2f-433f-b258-b9d31e51f823": "e631424e-56fc-4fa6-81a1-ddd029659bba",
    "1684bd3d-c040-4d93-96b9-90a8d63f0280": "e631424e-56fc-4fa6-81a1-ddd029659bba",
    "0fdbefb1-0772-490d-bde0-4b441a2b214f": "e631424e-56fc-4fa6-81a1-ddd029659bba",
    "03c963a9-e01f-4bea-a4b0-1b198c765e60": "e0ac230a-e68d-4f4e-beaa-cb2db82bb86a",
    "28b8adb6-8b22-4c47-ad63-4d6e958f622e": "e0ac230a-e68d-4f4e-beaa-cb2db82bb86a",
    "67546b27-ccb1-4952-bf84-9100db9c142b": "e0ac230a-e68d-4f4e-beaa-cb2db82bb86a",
    "5df2253b-7cb1-4c66-ac77-bae887eb93f6": "ebbdb871-c154-49a0-a212-b8dbe1874c66",
}
DEFAULT_STATUS_UUID = "ebbdb871-c154-49a0-a212-b8dbe1874c66"
DEFAULT_GROUP_UUID  = "b5749882-0a12-4be3-a022-7cf98d4b56b9"

# ─── Helpers ─────────────────────────────────────────────────────────────────

def p(*a): print(*a, flush=True)

def v(val):
    """Normalise a cell value to a comparable string."""
    s = str(val).strip() if val is not None else ""
    if s.lower() in ("none", "nan"):
        return ""
    # normalise date DD/MM/YYYY → YYYY-MM-DD
    m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})$", s)
    if m:
        return f"{m.group(3)}-{int(m.group(2)):02d}-{int(m.group(1)):02d}"
    # normalise YYYY-MM-DDT... → YYYY-MM-DD
    m = re.match(r"^(\d{4}-\d{2}-\d{2})(?:[ T].+)?$", s)
    if m:
        return m.group(1)
    # normalise numbers
    try:
        f = float(s.replace(",", ""))
        return str(int(f)) if f == int(f) else str(f)
    except (ValueError, OverflowError):
        pass
    return s.lower()

def dd_to_iso(s):
    """DD/MM/YYYY → YYYY-MM-DD, or passthrough if already ISO."""
    if not s:
        return ""
    m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})$", str(s).strip())
    if m:
        return f"{m.group(3)}-{int(m.group(2)):02d}-{int(m.group(1)):02d}"
    m = re.match(r"^(\d{4}-\d{2}-\d{2})", str(s).strip())
    if m:
        return m.group(1)
    return str(s).strip()

def load_csv(path):
    with open(path, encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))

def load_xlsx_rows(path, sheet):
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    if not rows:
        return [], []
    hdrs = [str(c) if c is not None else "" for c in rows[0]]
    data = [dict(zip(hdrs, row)) for row in rows[1:] if any(c is not None for c in row)]
    return hdrs, data

def load_json(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)

def write_csv(path, rows, keys):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=keys, extrasaction="ignore")
        w.writeheader()
        for row in rows:
            w.writerow({k: ("" if row.get(k) is None else row.get(k, "")) for k in keys})

def map_adviser(src):
    return ADVISER_MAP.get(str(src or "").strip(), "")

def map_status(src):
    return STATUS_MAP.get(str(src or "").strip(), DEFAULT_STATUS_UUID)

# ─── API clients ─────────────────────────────────────────────────────────────

def resolve_pats():
    from azure.identity import DefaultAzureCredential
    from azure.keyvault.secrets import SecretClient
    sc = SecretClient(
        vault_url=f"https://{KV_NAME}.vault.azure.net",
        credential=DefaultAzureCredential(),
    )
    return sc.get_secret(SRC_PAT_SECRET).value, sc.get_secret(DEST_PAT_SECRET).value


class SourceClient:
    """GET-only guard on source."""
    def __init__(self, pat):
        import requests
        self._s = requests.Session()
        self._s.headers.update({
            "Authorization": f"Bearer {pat}",
            "Accept": "application/json",
            "User-Agent": "HFP-AR-Migration/1.0 build_import_20261027",
        })
        r = self._s.get(f"{BASE_URL}/api/v1/me")
        inner = r.json().get("data", r.json())
        acct = inner.get("current_account", {}) or {}
        acct_uuid = acct.get("uuid") or inner.get("uuid")
        self._s.headers["X-PLANNR-ACCOUNT-UUID"] = acct_uuid
        p(f"  Source account: {acct_uuid}")

    def get(self, path, **kw):
        return self._s.get(f"{BASE_URL}{path}", **kw)


class DestClient:
    """POST /api/v1/import only on destination."""
    def __init__(self, pat):
        import requests
        self._s = requests.Session()
        self._s.headers.update({
            "Authorization": f"Bearer {pat}",
            "Accept": "application/json",
            "User-Agent": "HFP-AR-Migration/1.0 build_import_20261027",
            "X-PLANNR-ACCOUNT-UUID": DEST_ACCOUNT_UUID,
        })

    def post_import(self, files, data):
        return self._s.post(f"{BASE_URL}/api/v1/import", files=files, data=data)

    def post_map(self, imp_uuid, headers_map=None):
        body = {"headers": headers_map} if headers_map else {}
        return self._s.post(f"{BASE_URL}/api/v1/import/map/{imp_uuid}", json=body)

    def post_start(self, imp_uuid):
        return self._s.post(f"{BASE_URL}/api/v1/import/start/{imp_uuid}")

    def get_import(self, imp_uuid):
        return self._s.get(f"{BASE_URL}/api/v1/import/{imp_uuid}")


# ─── Pre-build check 1: missing clients/plans ────────────────────────────────

def pre_check_1(src):
    p("\n── Pre-build check 1 ─────────────────────────────────────────────────────")

    cl_26 = load_csv(os.path.join(EXPORTS_26, "Client Export HFP Migration 2026-10-27.csv"))
    cl_14 = load_csv(os.path.join(EXPORTS_14, "Client Export HFP Migration 2026-10-14.csv"))
    pl_26 = load_csv(os.path.join(EXPORTS_26, "Plan Export HFP Migration 2026-10-27.csv"))
    pl_14 = load_csv(os.path.join(EXPORTS_14, "Plan Export HFP Migration 2026-10-14.csv"))

    ids_26  = {r["uuid"] for r in cl_26 if r.get("uuid")}
    ids_14  = {r["uuid"] for r in cl_14 if r.get("uuid")}
    pids_26 = {r["uuid"] for r in pl_26 if r.get("uuid")}
    pids_14 = {r["uuid"] for r in pl_14 if r.get("uuid")}

    missing_cl = ids_14 - ids_26
    missing_pl = pids_14 - pids_26
    p(f"  Missing from 26 Oct: {len(missing_cl)} clients, {len(missing_pl)} plans")

    cl_14_map = {r["uuid"]: r for r in cl_14}
    pl_14_map = {r["uuid"]: r for r in pl_14}

    alive_cl, dead_cl, alive_pl, dead_pl = [], [], [], []
    for uid in sorted(missing_cl):
        ok = src.get(f"/api/v1/client/{uid}").status_code == 200
        (alive_cl if ok else dead_cl).append(uid)
    for uid in sorted(missing_pl):
        ok = src.get(f"/api/v1/plans/{uid}").status_code == 200
        (alive_pl if ok else dead_pl).append(uid)

    # Supplement 26 Oct with still-alive 14 Oct records
    for uid in alive_cl:
        cl_26.append(cl_14_map[uid])
    for uid in alive_pl:
        pl_26.append(pl_14_map[uid])

    p(f"  Clients: {len(alive_cl)} alive (added from 14 Oct), {len(dead_cl)} deleted")
    p(f"  Plans:   {len(alive_pl)} alive (added from 14 Oct), {len(dead_pl)} deleted")

    # Cross-reference deleted clients against "No status, no plans" Exceptions tab
    _, exc_rows = load_xlsx_rows(SIGNOFF_V6_XLSX, "Exceptions")
    no_status_no_plans = {
        str(r.get("source_uuid", "") or "").strip()
        for r in exc_rows
        if str(r.get("category", "") or "").strip().lower() == "no status, no plans"
        and r.get("source_uuid")
    }
    dead_cl_no_status = [u for u in dead_cl if u in no_status_no_plans]
    p(f"  Deleted clients that were 'No status, no plans': {len(dead_cl_no_status)}")

    return {
        "cl_26": cl_26, "pl_26": pl_26,
        "alive_cl": alive_cl, "dead_cl": dead_cl,
        "alive_pl": alive_pl, "dead_pl": dead_pl,
        "dead_cl_no_status": dead_cl_no_status,
    }


# ─── Pre-build check 2: re-count AR edits ────────────────────────────────────

def pre_check_2():
    p("\n── Pre-build check 2: AR edits recount ──────────────────────────────────")

    SKIP = {"Review", "basis", "indicative_annual_income"}

    def changed_cells(sent_rows, ret_rows, comp_cols, id_col=None):
        """Count cells where norm(sent) != norm(ret), both blank→filled and filled→blank."""
        n = 0
        if id_col:
            sm = {str(r.get(id_col, "") or "").strip(): r for r in sent_rows if r.get(id_col)}
            rm = {str(r.get(id_col, "") or "").strip(): r for r in ret_rows  if r.get(id_col)}
            for uid, sr in sm.items():
                rr = rm.get(uid)
                if rr is None:
                    continue
                for col in comp_cols:
                    sv = v(sr.get(col))
                    rv = v(rr.get(col))
                    if sv != rv:
                        n += 1
        else:
            for sr, rr in zip(sent_rows, ret_rows):
                for col in comp_cols:
                    sv = v(sr.get(col))
                    rv = v(rr.get(col))
                    if sv != rv:
                        n += 1
        return n

    def comp_cols_for(hdrs):
        return [h for h in hdrs if not h.startswith("INFO:") and
                not h.startswith("ORIGINAL_") and h not in SKIP]

    results = {}

    # Clients
    sh, sr = load_xlsx_rows(SENT_CLIENTS_XLSX, "Clients")
    _, rr  = load_xlsx_rows(RETURNED_CLIENTS, "Clients")
    cc = comp_cols_for(sh)
    n = changed_cells(sr, rr, cc, id_col="uuid")
    results["Clients"] = n
    p(f"  Clients: {n} changed cells")

    # Plans — merge all 3 returned Plan files
    sh, sr = load_xlsx_rows(SENT_PLANS_XLSX, "Plans")
    cc_pl = comp_cols_for(sh)
    sm_pl = {str(r.get("uuid", "") or "").strip(): dict(r) for r in sr if r.get("uuid")}
    merged_pl = {uid: dict(row) for uid, row in sm_pl.items()}
    for fname in sorted(os.listdir(RETURNED_DIR)):
        if fname.startswith("Plans") and fname.endswith(".xlsx"):
            _, ret_pl = load_xlsx_rows(os.path.join(RETURNED_DIR, fname), "Plans")
            for rr in ret_pl:
                uid = str(rr.get("uuid", "") or "").strip()
                if uid and uid in merged_pl:
                    for col in cc_pl:
                        rv = v(rr.get(col))
                        if rv and rv != v(sm_pl[uid].get(col)):
                            merged_pl[uid][col] = rr.get(col)
    n_pl = sum(
        1 for uid, mr in merged_pl.items()
        for col in cc_pl
        if v(mr.get(col)) != v(sm_pl[uid].get(col))
    )
    results["Plans"] = n_pl
    p(f"  Plans: {n_pl} changed cells")

    # Provider plans — merge all 3 returned files by (Portfolio number, Wrapper ref)
    sh_prov, sr_prov = load_xlsx_rows(SENT_PLANS_XLSX, "Provider_plans_not_in_Plannr")
    cc_prov = comp_cols_for(sh_prov)

    def pkey(r):
        return (str(r.get("Portfolio number", "") or "").strip(),
                str(r.get("Wrapper ref", "") or "").strip())

    sm_prov = {pkey(r): dict(r) for r in sr_prov}
    merged_prov = {k: dict(v2) for k, v2 in sm_prov.items()}
    for fname in sorted(os.listdir(RETURNED_DIR)):
        if fname.startswith("Plans") and fname.endswith(".xlsx"):
            _, ret_prov = load_xlsx_rows(os.path.join(RETURNED_DIR, fname),
                                         "Provider_plans_not_in_Plannr")
            for rr in ret_prov:
                k = pkey(rr)
                if k in merged_prov:
                    for col in cc_prov:
                        rv = v(rr.get(col))
                        if rv and rv != v(sm_prov[k].get(col)):
                            merged_prov[k][col] = rr.get(col)

    n_prov = sum(
        1 for k, mr in merged_prov.items()
        for col in cc_prov
        if v(mr.get(col)) != v(sm_prov[k].get(col))
    )
    results["Provider_plans_not_in_Plannr"] = n_prov
    p(f"  Provider plans: {n_prov} changed cells")

    # Charges — match by external_reference_value (returned workbook may be reordered)
    sh_ch, sr_ch = load_xlsx_rows(SENT_CHARGES_XLSX, "Charges")
    _, rr_ch     = load_xlsx_rows(RETURNED_CHARGES, "Charges")
    cc_ch = comp_cols_for(sh_ch)
    n_ch = changed_cells(sr_ch, rr_ch, cc_ch, id_col="external_reference_value")
    results["Charges"] = n_ch
    p(f"  Charges: {n_ch} changed cells")

    return results, merged_prov


# ─── Build W1 + W2 clients ───────────────────────────────────────────────────

CLIENT_KEYS = [
    "uuid","type","client_regulatory_classification","status_uuid","title",
    "first_name","middle_name","last_name","maiden_name","entity_name","email",
    "assigned_adviser_uuid","date_of_birth","gender","ni_number","nickname",
    "target_retirement_age","smoker_status","cigarettes_daily","smoker_since",
    "residential_status","country_of_residence","marital_status",
    "is_vulnerable","is_uk_resident","is_uk_domicile",
    "expected_retirement_age","state_retirement_age","nationalities",
    "vulnerability_details","vulnerability_sub_type","vulnerability_last_assessment_date",
    "vulnerability_duration","procedure_adaptation","adaption_details_notes","client_preferences",
    "current_state_of_health","state_of_health_explanation","long_term_care_needed","long_term_care_notes",
    "has_will","will_notes","will_last_reviewed","has_power_of_attorney","power_of_attorney_notes",
    "first_contact_date","units_of_alcohol_weekly","height","weight","health_notes",
    "has_iht_planning","iht_planning_notes","has_gifted_family","gift_notes",
    "is_anticipating_inheritance","anticipating_inheritance_notes",
    "is_will_accurate","will_executors","will_location","will_bequests_detail",
    "is_poa_registered_epa_or_lpa","are_poa_appointed_for_health_or_wealth",
    "poa_attorney_appointed_at","are_poa_documents_registered",
    "is_in_receipt_of_your_state_pension","has_obtained_a_forecast","date_of_forecast",
    "projected_state_pension","years_of_national_insurance_contributions","pension_notes",
    "phone","is_prospect","prospect_converted_at","value",
    "service_level","next_review_date","previous_review_date","anniversary_review_date",
    "terms_of_business_at","client_agreement_at",
    "address_1","address_2","address_3","address_4","town_city","county","postcode","country",
    "owners","send_welcome_email","date_of_marriage","is_deceased","deceased_at",
    "is_married_to_partner","checked_sanctions_at","lead_source","write_to_whom","passport_number",
    "do_not_contact_via_fax","do_not_contact_via_mail","do_not_contact_via_sms",
    "do_not_contact_via_email","do_not_contact_via_telephone","do_not_visit","do_not_market",
    "groups","external_reference_company","external_reference_value",
    "archived_reason","archived_notes",
]

_ADVISER_NAMES = {"jim hoban","jim","katie","katie hoban","tommie","tommie hoban","tomme","valerie","valerie hoban"}

def load_review_overrides():
    """Load client review entries: return (deleted_uuids, adviser_overrides)."""
    _, rows = load_xlsx_rows(CHANGES_XLSX, "Review entries")
    deleted = set()
    adviser = {}
    for r in rows:
        if str(r.get("sheet", "") or "") != "Clients":
            continue
        uid = str(r.get("uuid", "") or "").strip()
        new = str(r.get("new_value", "") or "").strip().lower()
        if "client has been deleted" in new:
            deleted.add(uid)
        elif uid and new and new not in ("", "no adviser", "client has been deleted"):
            adviser[uid] = new
    return deleted, adviser

def classify_review_entries():
    """Return (applied_list, unapplied_list) from Review entries tab."""
    _, rows = load_xlsx_rows(CHANGES_XLSX, "Review entries")
    applied, unapplied = [], []
    for r in rows:
        sheet = str(r.get("sheet", "") or "").strip()
        uid   = str(r.get("uuid", "") or "").strip()
        old   = str(r.get("old_value", "") or "").strip()
        new   = str(r.get("new_value", "") or "").strip()
        new_l = new.lower()
        if sheet == "Clients":
            if "deleted" in new_l:
                applied.append({"sheet": sheet, "uuid": uid, "type": "client_deleted", "new_value": new})
            elif new_l in _ADVISER_NAMES:
                applied.append({"sheet": sheet, "uuid": uid, "type": "adviser", "new_value": new})
            else:
                unapplied.append({"sheet": sheet, "uuid": uid, "old_value": old, "new_value": new})
        elif sheet == "Plans":
            if old == "No adviser" and "deleted" in new_l:
                applied.append({"sheet": sheet, "uuid": uid, "type": "plan_deleted_client", "new_value": new})
            else:
                unapplied.append({"sheet": sheet, "uuid": uid, "old_value": old, "new_value": new})
    return applied, unapplied

def build_clients(cl_26):
    p("\n── Building W1/W2 clients ────────────────────────────────────────────────")
    deleted_uuids, adviser_overrides = load_review_overrides()

    w1, w2, manual_phones = [], [], []
    for row in cl_26:
        uid = str(row.get("uuid", "") or "").strip()
        if not uid or uid in deleted_uuids:
            continue

        out = {k: str(row.get(k, "") or "").strip() for k in CLIENT_KEYS}
        out["uuid"] = ""  # store mode
        dest_status = map_status(row.get("status_uuid", ""))
        out["status_uuid"] = dest_status
        out["send_welcome_email"] = "No"
        out["groups"] = ""  # dest account has no groups configured
        out["external_reference_company"] = MIGRATION_CO
        out["external_reference_value"] = uid

        # archived_reason required when status is ARCHIVED/DECEASED/INACTIVE at dest
        ARCHIVED_DEST  = "e631424e-56fc-4fa6-81a1-ddd029659bba"
        SRC_DECEASED   = "1684bd3d-c040-4d93-96b9-90a8d63f0280"
        if dest_status == ARCHIVED_DEST and not out.get("archived_reason"):
            src_status = str(row.get("status_uuid", "") or "").strip()
            out["archived_reason"] = "Deceased" if src_status == SRC_DECEASED else "Left firm - Firm decision"

        # Strip spurious quote-wrappers from phone (e.g. '"+447889725956"' → '+447889725956')
        ph = out.get("phone", "")
        if ph.startswith('"') and ph.endswith('"'):
            ph = ph[1:-1]
        # Plannr rejects non-+44 numbers; blank them and record for manual entry after import
        if ph and not ph.startswith("+44"):
            manual_phones.append({"external_reference_value": uid, "phone": ph})
            ph = ""
        out["phone"] = ph

        # Adviser: Review override has priority
        if uid in adviser_overrides:
            dest_adv = ADVISER_NAME_DEST.get(adviser_overrides[uid], "")
        else:
            dest_adv = map_adviser(row.get("assigned_adviser_uuid", ""))
        out["assigned_adviser_uuid"] = dest_adv

        # Owners: comma-separated source UUIDs (pass through as-is for joint)
        out["owners"] = str(row.get("owners", "") or "").replace("|", ",")

        t = str(row.get("type", "") or "").lower()
        (w1 if t == "individual" else w2).append(out)

    if manual_phones:
        p(f"  Manual phones (non-+44, blanked): {len(manual_phones)}")
    p(f"  W1 individual: {len(w1)}, W2 non-individual: {len(w2)}")
    return w1, w2, manual_phones


# ─── Build W2 circles ────────────────────────────────────────────────────────

CIRCLE_KEYS = ["uuid","name","account_uuids","group_uuids",
               "external_reference_company","external_reference_value"]

def build_circles():
    p("\n── Building W2 circles ───────────────────────────────────────────────────")
    rows = load_csv(os.path.join(EXPORTS_26, "Circle Export HFP Migration 2026-10-27.csv"))
    circles = {}
    for r in rows:
        cid = str(r.get("circle_uuid", "") or "").strip()
        aid = str(r.get("account_uuid", "") or "").strip()
        if not cid:
            continue
        if cid not in circles:
            circles[cid] = {"name": str(r.get("circle_name", "") or "").strip(), "members": []}
        if aid:
            circles[cid]["members"].append(aid)

    out = []
    for cid, c in circles.items():
        out.append({
            "uuid": "",
            "name": c["name"],
            "account_uuids": ",".join(c["members"]),
            "group_uuids": "",
            "external_reference_company": MIGRATION_CO,
            "external_reference_value": cid,
        })
    p(f"  W2 circles: {len(out)}")
    return out


# ─── Build W3 relationships ──────────────────────────────────────────────────

REL_KEYS = ["account_uuid","related_account_uuid","relationship","inverse_relationship"]

INVERSE_REL = {
    "parent": "child",       "child": "parent",
    "grandparent": "grandchild", "grandchild": "grandparent",
    "aunt_uncle": "niece_nephew", "niece_nephew": "aunt_uncle",
    "spouse": "spouse",      "partner": "partner",
    "sibling": "sibling",    "cousin": "cousin",
    "civil-partner": "civil-partner", "co-habiting": "co-habiting",
    "ex-spouse": "ex-spouse", "primary-contact": "primary-contact",
    "other": "other",
}

def build_relationships():
    p("\n── Building W3 relationships ─────────────────────────────────────────────")
    acct = load_json(os.path.join(SNAP_DIR, "account_details.json"))
    cl_26 = load_csv(os.path.join(EXPORTS_26, "Client Export HFP Migration 2026-10-27.csv"))
    alive = {r["uuid"] for r in cl_26 if r.get("uuid")}

    # Collect per-pair directional relationships
    pair_rels = defaultdict(list)  # sorted_pair -> [(from_uuid, to_uuid, rel_type), ...]
    for cl_uuid, det in acct.items():
        for lc in det.get("linked_clients", []):
            rel_uuid = str(lc.get("uuid", "") or "").strip()
            rel_type = str(lc.get("relationship", "") or "").strip()
            if not rel_uuid or not rel_type:
                continue
            if cl_uuid not in alive or rel_uuid not in alive:
                continue
            pair = tuple(sorted([cl_uuid, rel_uuid]))
            pair_rels[pair].append((cl_uuid, rel_uuid, rel_type))

    # One row per pair: relationship from canonical direction + inverse
    rows = []
    for pair, rels in pair_rels.items():
        a, b = pair
        rel_ab = next((r[2] for r in rels if r[0] == a), None)
        rel_ba = next((r[2] for r in rels if r[0] == b), None)
        if rel_ab:
            inv = rel_ba if rel_ba else INVERSE_REL.get(rel_ab, rel_ab)
            rows.append({"account_uuid": a, "related_account_uuid": b,
                         "relationship": rel_ab, "inverse_relationship": inv})
        elif rel_ba:
            rows.append({"account_uuid": b, "related_account_uuid": a,
                         "relationship": rel_ba,
                         "inverse_relationship": INVERSE_REL.get(rel_ba, rel_ba)})

    p(f"  W3 relationships: {len(rows)} ({len(pair_rels)} unique pairs, {247 - len(rows)} deduped)")
    return rows


# ─── Build W3 dependants ─────────────────────────────────────────────────────

DEP_KEYS = ["uuid","first_name","last_name","known_as","relationship","date_of_birth",
            "is_financially_dependant","notes","account_uuids",
            "external_reference_company","external_reference_value"]

def build_dependants():
    p("\n── Building W3 dependants ────────────────────────────────────────────────")
    deps = load_json(os.path.join(EXPORTS_26, "dependants.json"))
    rows = []
    for d in deps:
        uid = str(d.get("uuid", "") or "").strip()
        accts = d.get("accounts", []) or []
        acct_uuids = ",".join(str(a.get("uuid", "") or "") for a in accts if a.get("uuid"))
        rows.append({
            "uuid": "",
            "first_name": d.get("first_name", "") or "",
            "last_name": d.get("last_name", "") or "",
            "known_as": d.get("known_as", "") or "",
            "relationship": d.get("relationship", "") or "",
            "date_of_birth": d.get("date_of_birth", "") or "",
            "is_financially_dependant": d.get("is_financially_dependant", "") or "",
            "notes": d.get("notes", "") or "",
            "account_uuids": acct_uuids,
            "external_reference_company": MIGRATION_CO,
            "external_reference_value": uid,
        })
    p(f"  W3 dependants: {len(rows)}")
    return rows


# ─── Build W3 income & employment ────────────────────────────────────────────

INC_KEYS = [
    "client_uuid","name","type","type_other","start_date","end_date",
    "amount","amount_type","frequency","is_taxable",
    "employment_uuid","employment_employer","employment_job_title",
    "employment_occupation_code","employment_status",
    "employment_started_at","employment_ended_at",
    "external_reference_company","external_reference_value",
    "employment_external_reference_company","employment_external_reference_value",
]

def build_income():
    p("\n── Building W3 income ────────────────────────────────────────────────────")
    acct = load_json(os.path.join(SNAP_DIR, "account_details.json"))
    rows = []
    for cl_uuid, det in acct.items():
        emp_list = det.get("employment", []) or []
        emp_by_uid = {str(e.get("uuid", "") or ""): e for e in emp_list if e.get("uuid")}
        primary_emp = emp_list[0] if emp_list else None
        for inc in det.get("income", []) or []:
            inc_uuid = str(inc.get("uuid", "") or "").strip()
            emp_uid  = str(inc.get("employment_uuid", "") or "").strip()
            emp = emp_by_uid.get(emp_uid) or (primary_emp if not emp_uid else None) or {}
            rows.append({
                "client_uuid": cl_uuid,
                "name": inc.get("name", "") or "",
                "type": inc.get("type", "") or "",
                "type_other": inc.get("type_other", "") or "",
                "start_date": inc.get("start_date", "") or "",
                "end_date": inc.get("end_date", "") or "",
                "amount": inc.get("amount", "") or "",
                "amount_type": inc.get("amount_type", "") or "",
                "frequency": inc.get("frequency", "") or "",
                "is_taxable": inc.get("is_taxable", "") or "",
                "employment_uuid": str(emp.get("uuid", "") or ""),
                "employment_employer": str(emp.get("employer", "") or ""),
                "employment_job_title": str(emp.get("job_title", "") or ""),
                "employment_occupation_code": str(emp.get("occupation_code", "") or ""),
                "employment_status": str(emp.get("status", "") or ""),
                "employment_started_at": str(emp.get("started_at", "") or ""),
                "employment_ended_at": str(emp.get("ended_at", "") or ""),
                "external_reference_company": MIGRATION_CO,
                "external_reference_value": inc_uuid,
                "employment_external_reference_company": MIGRATION_CO if emp.get("uuid") else "",
                "employment_external_reference_value": str(emp.get("uuid", "") or ""),
            })
    p(f"  W3 income: {len(rows)}")
    return rows


# ─── Build W4/W5 plans ───────────────────────────────────────────────────────

PLAN_KEYS = [
    "name","uuid","client_uuids","client_ownership_percentage","type","status","sub_type",
    "opened_at","policy_number","linked_policy_numbers","proposal_reference","provider",
    "provider_uuid","inactive_at","inactive_reason","agency_in_your_name","agency_in_your_name_from",
    "originally_sold_by_firm","under_advice","show_on_circles","visible_to_circles_uuids",
    "visible_on_client_portal","include_in_client_valuation","review_date","on_panel",
    "seller_uuid","seller_name","active_adviser_uuid","contract_enquiry_account_uuid",
    "is_group_scheme","is_wrap_product","wrap_uuid","external_reference_company","external_reference_value",
    "annuity_platform_provider","annuity_life_assured","annuity_sum_assured","annuity_term_type",
    "annuity_term_maturity_date","annuity_term_age","annuity_term_maturity_value","annuity_guaranteed_term",
    "annuity_payment_type","annuity_income_amount","annuity_income_frequency","annuity_income_from",
    "annuity_income_escalation","annuity_initial_purchase_amount",
    "cash_account_term_type","cash_account_term_maturity_date","cash_account_term_age",
    "cash_account_interest_type","cash_account_interest_rate","cash_account_platform_provider",
    "fixed_asset_address_1","fixed_asset_address_2","fixed_asset_address_3","fixed_asset_address_4",
    "fixed_asset_town_city","fixed_asset_county","fixed_asset_postcode","fixed_asset_country",
    "fixed_asset_new_property","fixed_asset_price_listed","fixed_asset_purchase_price",
    "fixed_asset_asset_description","fixed_asset_asset_image_path",
    "general_insurance_ncd","general_insurance_ncd_protected","general_insurance_ncd_years",
    "general_insurance_renewal_date","general_insurance_term_type","general_insurance_term_age",
    "general_insurance_term_maturity_date","group_protection_renewal_date","group_protection_term_type",
    "investment_offshore","investment_term_type","investment_term_maturity_date","investment_term_age",
    "investment_interest_type","investment_interest_rate","investment_platform_provider",
    "investment_in_trust","investment_trustee_details",
    "loan_credit_original_balance","loan_credit_outstanding_balance","loan_credit_interest_type",
    "loan_credit_interest_rate","loan_credit_loan_amount","loan_credit_repayment_amount",
    "loan_credit_repayment_frequency","loan_credit_term_type","loan_credit_term_maturity_date",
    "loan_credit_term_age","loan_credit_details",
    "mortgage_property_type","mortgage_payment_amount","mortgage_payment_frequency",
    "mortgage_initial_interest","mortgage_initial_period_end","mortgage_initial_payment",
    "mortgage_interest_type","mortgage_interest_rate","mortgage_equity_released",
    "mortgage_term_type","mortgage_term_maturity_date","mortgage_term_age",
    "mortgage_loan_quote","mortgage_loan_type","mortgage_mig_added","mortgage_mig_premium",
    "mortgage_repayment_type","mortgage_repayment_method","mortgage_repayment_term",
    "mortgage_other_repayment_method","mortgage_scheme_fixed_date","mortgage_base_payment",
    "mortgage_has_transfer_penalties","mortgage_is_portable_mortgage",
    "mortgage_original_balance","mortgage_outstanding_balance",
    "mortgage_transfer_consent","mortgage_transfer_penalty",
    "mortgage_contract_exchanged","mortgage_date_complete","mortgage_date_confirmed",
    "mortgage_date_exchanged","mortgage_date_submitted","mortgage_mortgage_completed",
    "mortgage_offer_confirmed","mortgage_offer_confirmed_date","mortgage_offer_date",
    "mortgage_offer_issued","mortgage_on_risk","mortgage_proposal_date","mortgage_submitted",
    "mortgage_declined_notes",
    "pension_accrual_rate","pension_year_accrued","pension_retirement_date","pension_pensionable_salary",
    "pension_employer_percentage","pension_employer_frequency","pension_employee_percentage",
    "pension_employee_frequency","pension_transfer_value","pension_date_of_transfer_valuation",
    "pension_widows_percentage","pension_contracted_out","pension_had_avc",
    "pension_in_house_avc_amount","pension_in_house_avc_percentage",
    "pension_death_benefit_amount","pension_projected_income","pension_plan_retirement_age",
    "pension_platform_provider","pension_state_retirement_age","pension_term_type",
    "pension_term_maturity_date","pension_term_age","pension_salary_sacrifice","pension_waiver",
    "pension_is_crystallised","pension_crystallisation_amount","pension_crystallisation_benefit",
    "pension_crystallisation_fund","pension_crystallisation_start_date","pension_crystallisation_end_date",
    "pension_crystallisation_frequency","pension_fund_amount","pension_uncrystallised_fund",
    "pension_income_taken_amount","pension_maximum_income_amount","pension_tax_free_amount",
    "pension_lifetime_allowance_percentage","pension_lta_applied","pension_mpaa_trigger",
    "pension_pcls_percentage","pension_reason",
    "protection_policy_in_trust","protection_policy_trustee_details",
    "protection_policy_has_pre_existing_conditions","protection_policy_pre_existing_conditions_notes",
    "protection_policy_term_type","protection_policy_term_maturity_date","protection_policy_term_age",
    "ownership_split",
]

def _norm_name(s):
    return re.sub(r"[^a-z0-9]", "", str(s).lower())

def build_plans(pl_26, cl_26, merged_prov):
    p("\n── Building W4/W5 plans ──────────────────────────────────────────────────")

    # Client name → UUID lookup for provider "Add to Plannr"
    cl_name_map = {}
    for row in cl_26:
        uid = str(row.get("uuid", "") or "").strip()
        if not uid:
            continue
        full = _norm_name(str(row.get("name", "") or f"{row.get('first_name','')} {row.get('last_name','')}"))
        if full:
            cl_name_map[full] = uid
        fn = _norm_name(str(row.get("first_name", "") or ""))
        ln = _norm_name(str(row.get("last_name", "") or ""))
        if ln:
            cl_name_map[ln + fn[:1]] = uid

    w4, w5, zero_val, prov_add, prov_already = [], [], [], [], []

    for row in pl_26:
        uid = str(row.get("uuid", "") or "").strip()
        if not uid:
            continue
        out = {k: str(row.get(k, "") or "").strip() for k in PLAN_KEYS}
        out["uuid"] = ""
        out["on_panel"] = ""
        out["active_adviser_uuid"] = map_adviser(row.get("active_adviser_uuid", ""))
        out["seller_uuid"] = map_adviser(row.get("seller_uuid", ""))
        out["client_uuids"] = str(row.get("client_uuids", "") or "").replace("|", ",")
        out["external_reference_company"] = MIGRATION_CO
        out["external_reference_value"] = uid

        cv = str(row.get("current_value", "") or "").strip()
        try:
            cv_f = float(cv.replace(",", "")) if cv else 0.0
        except ValueError:
            cv_f = 0.0
        if cv_f == 0.0:
            zero_val.append({"plan_uuid": uid, "current_value": cv,
                             "current_value_at": str(row.get("current_value_at", "") or "")})

        is_wrap = str(row.get("is_wrap_product", "") or "").lower() in ("true", "yes", "1")
        (w4 if is_wrap else w5).append(out)

    # Provider "Add to Plannr" from merged_prov
    seen_ext_refs = set()
    for pr in merged_prov.values():
        action = str(pr.get("action", "") or "").strip().lower()
        port_num = str(pr.get("Portfolio number", "") or "").strip()
        wrap_ref = str(pr.get("Wrapper ref", "") or "").strip()
        # Build unique ext-ref value: prov-{port_num}-{wrap_ref} or prov-{port_num}
        ext_ref_val = f"prov-{port_num}-{wrap_ref}" if wrap_ref else f"prov-{port_num}"
        if action == "add to plannr" and port_num and ext_ref_val not in seen_ext_refs:
            seen_ext_refs.add(ext_ref_val)
            client_name_raw = str(pr.get("Client in your Plannr", "") or "").strip()
            port_name = str(pr.get("Portfolio name", "") or "").strip()
            source = str(pr.get("source", "") or "").strip()
            amount = str(pr.get("Amount", "") or "").strip()
            # policy_number = wrapper_ref when available (unique per plan), else port_num
            policy_num = wrap_ref if wrap_ref else port_num

            # Match client
            cl_uid = cl_name_map.get(_norm_name(client_name_raw), "")
            if not cl_uid:
                parts = client_name_raw.split()
                if len(parts) >= 2:
                    cl_uid = cl_name_map.get(_norm_name(parts[-1]) + _norm_name(parts[0])[:1], "")

            plan_type = "pension" if any(w in port_name.lower() for w in ("sipp", "pension", "stps")) else "investment"
            new_row = {k: "" for k in PLAN_KEYS}
            new_row.update({
                "name": port_name,
                "uuid": "",
                "type": plan_type,
                "client_uuids": cl_uid,
                "policy_number": policy_num,
                "provider": source,
                "status": "active",
                "under_advice": "Yes",
                "agency_in_your_name": "Yes",
                "agency_in_your_name_from": "",
                "on_panel": "",
                "external_reference_company": MIGRATION_CO,
                "external_reference_value": ext_ref_val,
            })
            w5.append(new_row)
            prov_add.append({"port_name": port_name, "port_num": port_num,
                             "ext_ref_val": ext_ref_val, "policy_num": policy_num,
                             "client_name": client_name_raw, "client_uuid": cl_uid,
                             "amount": amount, "source": source})

        elif action == "already in plannr":
            prov_already.append({
                "port_name": str(pr.get("Portfolio name", "") or "").strip(),
                "port_num": port_num,
                "client_name": str(pr.get("Client in your Plannr", "") or "").strip(),
                "amount": str(pr.get("Amount", "") or "").strip(),
                "source": str(pr.get("source", "") or "").strip(),
            })

    p(f"  W4 wrappers: {len(w4)}, W5 other: {len(w5)} (inc {len(prov_add)} Add-to-Plannr)")
    p(f"  Zero/blank valuation plans: {len(zero_val)}, Provider already in Plannr: {len(prov_already)}")
    return w4, w5, zero_val, prov_add, prov_already


# ─── Build W6 premiums ───────────────────────────────────────────────────────

PREM_KEYS = ["uuid","plan_uuid","transfer_type","date_advice_given","amount","percentage",
             "frequency","source","payment_method","start_date","end_date","pensionable_salary",
             "details","transferred_from_plan_uuid","transferred_from_plan_type",
             "transferred_from_plan_sub_type"]

def build_premiums():
    p("\n── Building W6 premiums ──────────────────────────────────────────────────")
    rows = load_csv(os.path.join(EXPORTS_26, "Premium Export HFP Migration 2026-10-27.csv"))
    out = []
    for r in rows:
        row = {k: str(r.get(k, "") or "").strip() for k in PREM_KEYS}
        row["uuid"] = ""
        out.append(row)
    p(f"  W6 premiums: {len(out)}")
    return out


# ─── Build W6 valuations ─────────────────────────────────────────────────────

VAL_KEYS = ["plan_uuid","sub_account_uuid","policy_number","valuation_date","value"]

def build_valuations(pl_26):
    p("\n── Building W6 valuations ────────────────────────────────────────────────")

    # Load changed valuations from Valuations tab
    _, val_rows = load_xlsx_rows(CHANGES_XLSX, "Valuations")
    val_overrides = {}
    for r in val_rows:
        uid = str(r.get("uuid", "") or "").strip()
        new_val  = str(r.get("new_value_gbp", "") or "").strip()
        new_date = dd_to_iso(str(r.get("new_date", "") or ""))
        if uid and (new_val or new_date):
            val_overrides[uid] = (new_val, new_date)

    # plan_uuid → (current_value, current_value_at) from 26 Oct plan export
    pl_export = {r["uuid"]: r for r in pl_26 if r.get("uuid")}

    snap = load_json(os.path.join(SNAP_DIR, "plans_dedup.json"))

    out, zero_blank = [], []
    for plan_uid, plan in snap.items():
        if not isinstance(plan, dict):
            continue
        val_obj = plan.get("value") or {}
        if not isinstance(val_obj, dict):
            continue

        # Default from snapshot (pence → pounds)
        try:
            pounds = float(str(val_obj.get("amount", 0) or 0)) / 100.0
        except (TypeError, ValueError):
            pounds = 0.0
        val_date = str(val_obj.get("amount_at", "") or "").strip()[:10]

        # Fallback to 26 Oct plan export if snapshot date is blank
        if not val_date and plan_uid in pl_export:
            pl_row = pl_export[plan_uid]
            exp_val  = str(pl_row.get("current_value", "") or "").strip()
            exp_date = dd_to_iso(str(pl_row.get("current_value_at", "") or ""))
            if exp_val:
                try:
                    pounds = float(exp_val.replace(",", ""))
                except ValueError:
                    pass
            if exp_date:
                val_date = exp_date

        # Apply Valuations tab override
        if plan_uid in val_overrides:
            new_val, new_date = val_overrides[plan_uid]
            if new_val:
                try:
                    pounds = float(str(new_val).replace(",", ""))
                except ValueError:
                    pass
            if new_date:
                val_date = new_date

        if plan_uid not in pl_export:
            continue  # plan deleted between snapshot and 26 Oct export — skip

        pol_num = str((pl_export.get(plan_uid) or {}).get("policy_number", "") or "").strip()

        row = {
            "plan_uuid": plan_uid,
            "sub_account_uuid": "",
            "policy_number": pol_num,
            "valuation_date": val_date,
            "value": f"{pounds:.2f}",
        }
        if pounds == 0.0 and not val_date:
            zero_blank.append({"plan_uuid": plan_uid, "value": f"{pounds:.2f}", "valuation_date": val_date})
            continue  # no valuation data at all — skip
        out.append(row)
        if pounds == 0.0 or not val_date:
            zero_blank.append({"plan_uuid": plan_uid, "value": f"{pounds:.2f}", "valuation_date": val_date})

    p(f"  W6 valuations: {len(out)} ({len(zero_blank)} at zero or no date; {sum(1 for z in zero_blank if z['value']=='0.00' and not z['valuation_date'])} zero+no-date skipped)")

    # Item 4: compare W6 totals vs 26 Oct export totals for the same plans
    pl_type_map = {r["uuid"]: str(r.get("type", "") or "") for r in pl_26}
    w6_by_type, exp_by_type = defaultdict(float), defaultdict(float)
    for row in out:
        puid = row["plan_uuid"]
        pt   = pl_type_map.get(puid, "unknown")
        w6_by_type[pt]  += float(row["value"] or 0)
        try:
            exp_v = float(str((pl_export.get(puid) or {}).get("current_value", "") or "").replace(",", ""))
        except ValueError:
            exp_v = 0.0
        exp_by_type[pt] += exp_v
    w6_total  = sum(w6_by_type.values())
    exp_total = sum(exp_by_type.values())
    p(f"  Val check (W6 plans only): W6=£{w6_total:,.0f}  export=£{exp_total:,.0f}  gap=£{exp_total - w6_total:,.0f}")
    for pt in sorted(set(list(w6_by_type) + list(exp_by_type))):
        p(f"    {pt or 'blank'}: W6=£{w6_by_type[pt]:,.0f} | export=£{exp_by_type[pt]:,.0f}")

    return out, zero_blank


# ─── Build W7 charges ────────────────────────────────────────────────────────

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

def build_charges(prov_add=None):
    p("\n── Building W7 charges ───────────────────────────────────────────────────")

    _, sent_ch = load_xlsx_rows(SENT_CHARGES_XLSX, "Charges")
    _, ret_ch  = load_xlsx_rows(RETURNED_CHARGES,  "Charges")

    # Alive plans (to filter deleted plans)
    pl_26_raw = load_csv(os.path.join(EXPORTS_26, "Plan Export HFP Migration 2026-10-27.csv"))
    alive_plans = {r["uuid"] for r in pl_26_raw if r.get("uuid")}

    # Add-to-Plannr: use prov_add (built in build_plans) which has resolved client_uuid, ext_ref_val
    # prov_add rows: {"port_name", "port_num", "ext_ref_val", "policy_num",
    #                 "client_name", "client_uuid", "amount", "source"}
    if prov_add is None:
        prov_add = []
    add_to_pl_transact  = [r for r in prov_add
                           if str(r.get("source", "") or "").strip().lower() == "transact"]
    add_to_pl_rathbones = [r for r in prov_add
                           if str(r.get("source", "") or "").strip().lower() == "rathbones"]

    # Payment Rates lookup: portfolio_number → {cash, inv} annual rate (decimal, e.g. 0.006)
    rate_by_port = {}
    try:
        _, pay_rate_rows = load_xlsx_rows(TRANSACT_RATES_XLSX, "Payment Rates Summary")
        for r in pay_rate_rows:
            pn = str(r.get("Portfolio Number", "") or "").strip()
            if not pn:
                continue
            try:
                cash = float(r.get("Annual Payment - Cash") or 0)
            except (TypeError, ValueError):
                cash = 0.0
            try:
                inv = float(r.get("Annual Payment - Investments") or 0)
            except (TypeError, ValueError):
                inv = 0.0
            if pn not in rate_by_port:
                rate_by_port[pn] = {"cash": cash, "inv": inv}
            else:
                rate_by_port[pn]["cash"] = max(rate_by_port[pn]["cash"], cash)
                rate_by_port[pn]["inv"]  = max(rate_by_port[pn]["inv"],  inv)
    except Exception as e:
        p(f"  WARN: could not load Payment Rates: {e}")

    # Returned workbook ext-ref → row lookup (may be in different order from sent)
    ret_by_extref = {}
    for rr in ret_ch:
        ext = str(rr.get("external_reference_value", "") or "").strip()
        if ext:
            ret_by_extref[ext] = rr

    def _read_pct(raw):
        """Parse percentage raw value → corrected float, or None if zero/blank/non-numeric."""
        s = str(raw or "").strip()
        if not s:
            return None
        try:
            v = float(s.replace("%", "").replace(",", ""))
        except ValueError:
            return None
        if v == 0.0:
            return None
        if 0 < v < 0.1:
            v *= 100.0  # Excel stores 0.6% as 0.006
        return v

    def _sent_row(r, pct):
        """Build a CHARGE_KEYS row from a sent-workbook row and a resolved percentage float."""
        return {
            "uuid":        "",
            "plan_uuid":   str(r.get("plan_uuid", "") or "").strip(),
            "client_uuid": str(r.get("client_uuid", "") or "").strip(),
            "adviser_uuid": map_adviser(str(r.get("adviser_uuid", "") or "")),
            "payment_origin": str(r.get("payment_origin", "") or "Provider").strip(),
            "type":         str(r.get("type", "") or "Fee").strip(),
            "category":     ("ongoing_fee" if str(r.get("category","") or "") in ("","ongoing") else str(r.get("category","")).strip()),
            "recurring":    str(r.get("recurring", "") or "True").strip(),
            "frequency":    str(r.get("frequency", "") or "").strip(),
            "start_date":   str(r.get("start_date", "") or "").strip(),
            "end_date":     "",
            "amount":       str(r.get("amount", "") or "").strip(),
            "percentage":   f"{pct:.4f}" if pct is not None else "",
            "percentage_frequency": str(r.get("percentage_frequency", "") or "annually").strip(),
            "tax": "",
            "tax_exempt":         str(r.get("tax_exempt", "") or "Yes").strip(),
            "exclude_from_rmar":  str(r.get("exclude_from_rmar", "") or "No").strip(),
            "advice_type":        str(r.get("advice_type", "") or "Independent").strip(),
            "service_type":       ("ongoing" if str(r.get("service_type","") or "") in ("","ongoing_service") else str(r.get("service_type","")).strip()),
            "rmar_type":          str(r.get("rmar_type", "") or "Adviser Charge").strip(),
            "premium_uuid": "",
            "advice_area_deposit": "",
            "advice_area_non_regulated": "",
            "advice_area_non_rmar": "",
            "advice_area_non_investment_insurance": "",
            "advice_area_regulated_mortgage": "",
            "advice_area_retail_investment": "100%",
            "advice_area_other_fca_regulated": "",
            "advice_area_unknown": "",
            "notes": str(r.get("notes", "") or "").strip(),
        }

    # ── Sec A: Provider statement + Rathbones (nonzero rate, alive plans) ─────
    sec_a, high_rates = [], []
    for r in sent_ch:
        basis = str(r.get("basis", "") or "").strip()
        if basis not in ("Provider statement", "Rathbones"):
            continue
        plan_uid = str(r.get("plan_uuid", "") or "").strip()
        if not plan_uid or plan_uid not in alive_plans:
            continue
        pct = _read_pct(r.get("percentage"))
        if pct is None:
            continue
        if pct > 3.0:
            high_rates.append({"plan_uuid": plan_uid,
                                "client_uuid": str(r.get("client_uuid", "") or ""),
                                "percentage": f"{pct:.4f}"})
        sec_a.append(_sent_row(r, pct))

    # ── Sec B: Your Plannr ongoing charges (alive plans) ──────────────────────
    sec_b = []
    for r in sent_ch:
        basis = str(r.get("basis", "") or "").strip()
        if basis != "Your Plannr":
            continue
        plan_uid = str(r.get("plan_uuid", "") or "").strip()
        if not plan_uid or plan_uid not in alive_plans:
            continue
        pct = _read_pct(r.get("percentage"))
        row = _sent_row(r, pct)
        # If no percentage, use amount (Sec B may be fixed-amount charges)
        if not row["percentage"]:
            row["amount"] = str(r.get("amount", "") or r.get("ORIGINAL_amount", "") or "").strip()
        sec_b.append(row)

    # ── Sec C: Rate-needed — AR-filled percentages from returned workbook ──────
    rate_filled, rate_blank = [], []
    rate_entered = {"rate": 0, "zero": 0, "blank": 0, "text": 0}
    sec_c = []
    for sr in sent_ch:
        if str(sr.get("basis", "") or "").strip() != "Rate needed":
            continue
        ext_ref  = str(sr.get("external_reference_value", "") or "").strip()
        plan_uid = str(sr.get("plan_uuid", "") or "").strip()
        cl_uid   = str(sr.get("client_uuid", "") or "").strip()
        info_adv = str(sr.get("INFO: Adviser", "") or "").strip()

        rr = ret_by_extref.get(ext_ref, {})
        # Distinguish None/missing from explicit zero (int 0 or float 0.0)
        if not rr or rr.get("percentage") is None:
            rate_entered["blank"] += 1
            rate_blank.append({"plan_uuid": plan_uid, "client_uuid": cl_uid,
                                "ar_entered": "blank", "info_adviser": info_adv})
            continue
        raw_pct = rr["percentage"]
        # Excel stores blank as None (handled above); zero appears as int 0 or float 0.0
        try:
            v = float(str(raw_pct).replace("%", "").replace(",", "").strip())
        except (ValueError, AttributeError):
            s = str(raw_pct or "").strip()
            if not s:
                rate_entered["blank"] += 1
                rate_blank.append({"plan_uuid": plan_uid, "client_uuid": cl_uid,
                                    "ar_entered": "blank", "info_adviser": info_adv})
            else:
                rate_entered["text"] += 1
                rate_blank.append({"plan_uuid": plan_uid, "client_uuid": cl_uid,
                                    "ar_entered": f"text:{s[:30]}", "info_adviser": info_adv})
            continue

        if v == 0.0:
            rate_entered["zero"] += 1
            rate_blank.append({"plan_uuid": plan_uid, "client_uuid": cl_uid,
                                "ar_entered": "zero", "info_adviser": info_adv})
            continue

        rate_entered["rate"] += 1
        if 0 < v < 0.1:
            v *= 100.0
        if v > 3.0:
            high_rates.append({"plan_uuid": plan_uid, "client_uuid": cl_uid,
                                "percentage": f"{v:.4f}"})

        row = _sent_row(sr, v)
        # Sec C: use returned frequency if present
        ret_freq = str((rr or {}).get("frequency", "") or "").strip()
        if ret_freq:
            row["frequency"] = ret_freq
        sec_c.append(row)
        rate_filled.append({"plan_uuid": plan_uid, "client_uuid": cl_uid,
                            "percentage": f"{v:.4f}", "info_adviser": info_adv})

    # ── Add-to-Plannr charges — in W7 with ext_ref_val as plan_uuid placeholder ─
    # plan_uuid is set to ext_ref_val; the import runner resolves it after W5.
    add_ch_report = []
    add_ch_w7     = []
    for r in add_to_pl_transact:
        pn       = str(r.get("port_num", "") or "").strip()
        name     = str(r.get("port_name", "") or "").strip()
        cl       = str(r.get("client_uuid", "") or "").strip()
        ext_ref  = str(r.get("ext_ref_val", "") or "").strip()
        adv_uuid = map_adviser("")
        rates    = rate_by_port.get(pn, {})
        inv_rate  = rates.get("inv", 0.0)
        cash_rate = rates.get("cash", 0.0)
        rate_dec  = inv_rate if inv_rate > 0 else cash_rate
        add_ch_report.append({
            "portfolio_number": pn,
            "ext_ref_val":      ext_ref,
            "port_name": name,
            "client_uuid": cl,
            "annual_rate_dec": f"{rate_dec:.4f}" if rate_dec else "",
            "pct_for_plannr":  f"{rate_dec * 100:.4f}" if rate_dec else "",
            "source": "Transact Payment Rates",
        })
        if rate_dec > 0 and ext_ref and cl:
            ch_row = {k: "" for k in CHARGE_KEYS}
            ch_row.update({
                "plan_uuid":   ext_ref,   # placeholder; resolved by import runner after W5
                "client_uuid": cl,
                "adviser_uuid": adv_uuid,
                "payment_origin": "Provider",
                "type":     "Fee",
                "category": "ongoing_fee",
                "recurring": "True",
                "frequency": "quarterly",
                "start_date": "",
                "percentage": f"{rate_dec * 100:.4f}",
                "percentage_frequency": "annually",
                "tax_exempt": "Yes",
                "exclude_from_rmar": "No",
                "advice_type": "Independent",
                "service_type": "ongoing",
                "rmar_type": "Adviser Charge",
                "advice_area_retail_investment": "100%",
            })
            add_ch_w7.append(ch_row)
    for r in add_to_pl_rathbones:
        pn      = str(r.get("port_num", "") or "").strip()
        name    = str(r.get("port_name", "") or "").strip()
        cl      = str(r.get("client_uuid", "") or "").strip()
        ext_ref = str(r.get("ext_ref_val", "") or "").strip()
        amt     = str(r.get("amount", "") or "").strip()
        add_ch_report.append({
            "portfolio_number": pn,
            "ext_ref_val":      ext_ref,
            "port_name": name,
            "client_uuid": cl,
            "annual_rate_dec": "",
            "pct_for_plannr":  "",
            "source": f"Rathbones — amount={amt}",
        })

    all_charges = sec_a + sec_b + sec_c + add_ch_w7
    p(f"  Rate-needed 206: rate={rate_entered['rate']}, zero={rate_entered['zero']}, "
      f"blank={rate_entered['blank']}, text={rate_entered['text']}")
    p(f"  W7 charges: {len(all_charges)} ({len(sec_a)} sec A, {len(sec_b)} sec B, "
      f"{len(sec_c)} sec C, {len(add_ch_w7)} Add-to-Plannr)")
    p(f"  Add-to-Plannr report: {len(add_ch_report)} rows ({len(add_ch_w7)} with rate → in W7)")
    p(f"  High rates >3%: {len(high_rates)}")
    return all_charges, rate_filled, rate_blank, high_rates, add_ch_report


# ─── Build W8 notes ───────────────────────────────────────────────────────────

NOTE_KEYS = ["model_type","type","path","account_uuid","plan_uuid","risk_uuid","case_uuid",
             "task_uuid","circle_uuid","contents","is_pinned","visible_to_clients","created_at",
             "external_reference_company","external_reference_value"]

def build_notes():
    p("\n── Building W8 notes ─────────────────────────────────────────────────────")
    rows = load_csv(os.path.join(EXPORTS_26, "Note Export HFP Migration 2026-10-27.csv"))
    out, no_client, rehomed = [], [], 0
    mt_counts = {}

    for r in rows:
        mt = str(r.get("model_type", "") or "").strip().lower()
        if mt == "board_task":
            continue

        uid        = str(r.get("uuid", "") or "").strip()
        contents   = str(r.get("contents", "") or "").strip()
        created_at = str(r.get("created_at", "") or "").strip()

        # Resolve account_uuid: direct field first, then first from account_uuids
        acct_uid = str(r.get("account_uuid", "") or "").strip()
        if not acct_uid:
            au = str(r.get("account_uuids", "") or "").replace("|", ",").replace(";", ",")
            acct_uid = au.split(",")[0].strip()

        if not acct_uid:
            no_client.append({"note_uuid": uid, "model_type": mt})
            continue

        # Determine output model_type and entity UUID fields
        if mt in ("task", "case"):
            # AR decision: re-home task/case notes to client account record
            task_name  = str(r.get("task_name", "") or r.get("case_name", "") or "").strip()
            author     = str(r.get("author", "") or "").strip()
            date_short = dd_to_iso(created_at)[:10] if created_at else ""
            header_parts = [f"[{mt} note"]
            if author:
                header_parts.append(f"by {author}")
            if date_short:
                header_parts.append(f"on {date_short}")
            if task_name:
                header_parts.append(f'"{task_name}"')
            header = " ".join(header_parts) + "]"
            contents  = header + "\n" + contents
            out_mt    = "account"
            out_plan  = ""
            out_risk  = ""
            rehomed  += 1
        elif mt == "plan":
            out_mt   = "plan"
            out_plan = str(r.get("plan_uuid", "") or "").strip()
            out_risk = ""
        elif mt == "risk":
            # Risk profiles not imported; re-home to client account with header
            author     = str(r.get("author", "") or "").strip()
            date_short = dd_to_iso(created_at)[:10] if created_at else ""
            header_parts = ["[risk note"]
            if author:      header_parts.append(f"by {author}")
            if date_short:  header_parts.append(f"on {date_short}")
            contents  = " ".join(header_parts) + "]\n" + contents
            out_mt    = "account"
            out_plan  = ""
            out_risk  = ""
            rehomed  += 1
        elif mt in ("circle", "charge"):
            # Re-home to client account with header
            author     = str(r.get("author", "") or "").strip()
            date_short = dd_to_iso(created_at)[:10] if created_at else ""
            header_parts = [f"[{mt} note"]
            if author:      header_parts.append(f"by {author}")
            if date_short:  header_parts.append(f"on {date_short}")
            contents  = " ".join(header_parts) + "]\n" + contents
            out_mt    = "account"
            out_plan  = ""
            out_risk  = ""
            rehomed  += 1
        else:
            # account — attach directly
            out_mt   = mt
            out_plan = str(r.get("plan_uuid", "") or "").strip()
            out_risk = ""

        out.append({
            "model_type": out_mt,
            "type": str(r.get("type", "") or "").strip(),
            "path": str(r.get("path", "") or "").strip(),
            "account_uuid": acct_uid,
            "plan_uuid": out_plan,
            "risk_uuid": out_risk,
            "case_uuid": "",
            "task_uuid": "",
            "circle_uuid": "",
            "contents": contents,
            "is_pinned": str(r.get("is_pinned", "") or "False").strip(),
            "visible_to_clients": str(r.get("visible_to_clients", "") or "False").strip(),
            "created_at": created_at,
            "external_reference_company": MIGRATION_CO,
            "external_reference_value": uid,
        })
        mt_counts[mt] = mt_counts.get(mt, 0) + 1

    p(f"  W8 notes: {len(out)} ({rehomed} task/case re-homed, {len(no_client)} no-account skipped)")
    p(f"  By source type: { {k: mt_counts.get(k,0) for k in ('account','task','case','plan','risk','circle','charge')} }")
    return out, no_client, rehomed


# ─── Validate W1 ─────────────────────────────────────────────────────────────

def validate_w1(dest):
    p("\n── Validating W1 against destination ────────────────────────────────────")
    w1_path = os.path.join(IMPORT_DIR, "W1_clients.csv")
    with open(w1_path, "rb") as fh:
        r = dest.post_import(
            files={"file": ("W1_clients.csv", fh, "text/csv")},
            data={"type": "clients", "mode": "validate"},
        )
    if r.status_code not in (200, 201, 202):
        p(f"  ERROR creating import: HTTP {r.status_code}: {r.text[:300]}")
        return {"status": "error", "error_count": -1}

    body  = r.json()
    inner = body.get("data", body) if isinstance(body, dict) else body
    imp_uuid = inner.get("uuid") or inner.get("id")
    p(f"  Import UUID: {imp_uuid}")

    # Wait for server to parse headers, then map with identity mapping {col: col}
    time.sleep(8)
    r_hdr = dest.get_import(imp_uuid)
    csv_hdrs = (r_hdr.json().get("data", r_hdr.json()) if r_hdr.status_code == 200 else {}).get("headers", [])
    r2 = dest.post_map(imp_uuid, headers_map={h: h for h in csv_hdrs})
    if r2.status_code not in (200, 201, 202):
        p(f"  WARN map: HTTP {r2.status_code}: {r2.text[:200]}")

    r3 = dest.post_start(imp_uuid)
    if r3.status_code not in (200, 201, 202):
        p(f"  WARN start: HTTP {r3.status_code}: {r3.text[:100]}")

    for i in range(60):
        time.sleep(10)
        r4 = dest.get_import(imp_uuid)
        if r4.status_code != 200:
            p(f"  WARN poll {i+1}: HTTP {r4.status_code}")
            continue
        d4 = r4.json().get("data", r4.json())
        status = str(d4.get("status", "") or "").lower()
        p(f"  poll {i+1}: {status}")
        if status in ("complete", "failed", "error", "validated",
                      "validation_complete", "validation_failed"):
            errs = d4.get("errors") or d4.get("error_count") or 0
            n_err = len(errs) if isinstance(errs, list) else int(errs or 0)
            p(f"  Final: status={status}, errors={n_err}")
            return {"status": status, "error_count": n_err, "uuid": imp_uuid}

    p("  WARN: timed out polling")
    return {"status": "timeout", "error_count": -1}


# ─── Write lists Excel ───────────────────────────────────────────────────────

def write_lists(sheets):
    import openpyxl
    from openpyxl.styles import Font, PatternFill
    p("\n── Writing lists Excel ───────────────────────────────────────────────────")
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    hdr_font = Font(bold=True)
    hdr_fill = PatternFill(fill_type="solid", fgColor="F2F2F2")
    for name, (hdrs, rows) in sheets.items():
        ws = wb.create_sheet(title=name[:31])
        ws.append(hdrs)
        for cell in ws[1]:
            cell.font = hdr_font
            cell.fill = hdr_fill
        for row in rows:
            ws.append([str(row.get(h, "") or "") for h in hdrs])
    wb.save(LISTS_XLSX)
    p(f"  Saved: {LISTS_XLSX}")


# ─── Main ─────────────────────────────────────────────────────────────────────

def main():
    p("=" * 68)
    p("HFP AR Migration — Build import files 27 Oct 2026")
    p("=" * 68)
    os.makedirs(IMPORT_DIR, exist_ok=True)

    p("\n[1] Resolving PATs from Key Vault...")
    src_pat, dest_pat = resolve_pats()
    p(f"  SRC len={len(src_pat)}, DEST len={len(dest_pat)}")

    p("\n[2] Connecting source + destination...")
    src  = SourceClient(src_pat)
    dest = DestClient(dest_pat)

    # ── Pre-build checks ──────────────────────────────────────────────────────
    pc1 = pre_check_1(src)
    cl_26 = pc1["cl_26"]
    pl_26 = pc1["pl_26"]

    pc2_results, merged_prov = pre_check_2()

    # ── Build waves ───────────────────────────────────────────────────────────
    w1, w2_cl, manual_phones = build_clients(cl_26)
    w2_cir       = build_circles()
    w3_rel       = build_relationships()
    w3_dep       = build_dependants()
    w3_inc       = build_income()
    w4, w5, zero_val, prov_add, prov_already = build_plans(pl_26, cl_26, merged_prov)
    w6_prem      = build_premiums()
    w6_val, zero_blank_val = build_valuations(pl_26)
    w7, rate_filled, rate_blank, high_rates, add_ch_report = build_charges(prov_add)
    w8, notes_no_cl, rehomed = build_notes()

    # ── Write CSVs ────────────────────────────────────────────────────────────
    p("\n── Writing CSVs ──────────────────────────────────────────────────────────")
    write_csv(os.path.join(IMPORT_DIR, "W1_clients.csv"),    w1,      CLIENT_KEYS)
    write_csv(os.path.join(IMPORT_DIR, "W2_clients.csv"),    w2_cl,   CLIENT_KEYS)
    write_csv(os.path.join(IMPORT_DIR, "W2_circles.csv"),    w2_cir,  CIRCLE_KEYS)
    write_csv(os.path.join(IMPORT_DIR, "W3_relationships.csv"), w3_rel, REL_KEYS)
    write_csv(os.path.join(IMPORT_DIR, "W3_dependants.csv"), w3_dep,  DEP_KEYS)
    write_csv(os.path.join(IMPORT_DIR, "W3_income.csv"),     w3_inc,  INC_KEYS)
    write_csv(os.path.join(IMPORT_DIR, "W4_plans.csv"),      w4,      PLAN_KEYS)
    write_csv(os.path.join(IMPORT_DIR, "W5_plans.csv"),      w5,      PLAN_KEYS)
    write_csv(os.path.join(IMPORT_DIR, "W6_premiums.csv"),   w6_prem, PREM_KEYS)
    write_csv(os.path.join(IMPORT_DIR, "W6_valuations.csv"), w6_val,  VAL_KEYS)
    write_csv(os.path.join(IMPORT_DIR, "W7_charges.csv"),    w7,      CHARGE_KEYS)
    write_csv(os.path.join(IMPORT_DIR, "W8_notes.csv"),      w8,      NOTE_KEYS)
    p("  Done.")

    # ── Validate W1 ───────────────────────────────────────────────────────────
    val = validate_w1(dest)

    # ── Build lists Excel ─────────────────────────────────────────────────────
    rev_applied, rev_unapplied = classify_review_entries()
    _, rev_rows = load_xlsx_rows(CHANGES_XLSX, "Review entries")
    cl_rev = [{k: str(r.get(k,"") or "") for k in ["uuid","info_name","column","old_value","new_value"]}
              for r in rev_rows if str(r.get("sheet","") or "") == "Clients"]
    pl_rev = [{k: str(r.get(k,"") or "") for k in ["uuid","info_name","column","old_value","new_value"]}
              for r in rev_rows if str(r.get("sheet","") or "") == "Plans"]

    write_lists({
        "PC1 deleted clients":      (["uuid"],
            [{"uuid": u} for u in pc1["dead_cl"]]),
        "PC1 alive clients":        (["uuid"],
            [{"uuid": u} for u in pc1["alive_cl"]]),
        "PC1 deleted plans":        (["uuid"],
            [{"uuid": u} for u in pc1["dead_pl"]]),
        "PC1 alive plans":          (["uuid"],
            [{"uuid": u} for u in pc1["alive_pl"]]),
        "PC2 AR edits":             (["sheet", "changed_cells"],
            [{"sheet": k, "changed_cells": v2} for k, v2 in pc2_results.items()]),
        "Plans zero valuation":     (["plan_uuid", "current_value", "current_value_at"],
            zero_val),
        "Val zero or no date":      (["plan_uuid", "value", "valuation_date"],
            zero_blank_val),
        "Provider add to Plannr":   (["port_name", "port_num", "client_name", "client_uuid", "amount", "source"],
            prov_add),
        "Provider already in Pl":   (["port_name", "port_num", "client_name", "amount", "source"],
            prov_already),
        "Charges sec C filled":     (["plan_uuid", "client_uuid", "percentage", "info_adviser"],
            rate_filled),
        "Charges sec C omitted":    (["plan_uuid", "client_uuid", "ar_entered", "info_adviser"],
            rate_blank),
        "Charges high rate >3pct":  (["plan_uuid", "client_uuid", "percentage"],
            high_rates),
        "W7 add-to-Plannr":         (["portfolio_number", "ext_ref_val", "port_name",
                                       "client_uuid", "annual_rate_dec", "pct_for_plannr", "source"],
            add_ch_report),
        "Review entries clients":   (["uuid", "info_name", "column", "old_value", "new_value"],
            cl_rev),
        "Review entries plans":     (["uuid", "info_name", "column", "old_value", "new_value"],
            pl_rev),
        "Review applied":           (["sheet", "uuid", "type", "new_value"],
            rev_applied),
        "Review unapplied":         (["sheet", "uuid", "old_value", "new_value"],
            rev_unapplied),
        "Notes no account":         (["note_uuid", "model_type"],
            notes_no_cl),
        "Manual phones":            (["external_reference_value", "phone"],
            manual_phones),
    })

    # ── Report ────────────────────────────────────────────────────────────────
    p("\n" + "=" * 68)
    p("REPORT")
    p("=" * 68)
    p(f"  Pre-check 1 clients:     {len(pc1['alive_cl'])} alive (added), {len(pc1['dead_cl'])} deleted"
      f" ({len(pc1['dead_cl_no_status'])} were 'No status, no plans')")
    p(f"  Pre-check 1 plans:       {len(pc1['alive_pl'])} alive (added), {len(pc1['dead_pl'])} deleted")
    p(f"  Pre-check 2 AR edits:")
    for sheet, n in pc2_results.items():
        p(f"    {sheet}: {n} changed cells")
    p(f"  W1  individual clients:  {len(w1)}")
    p(f"  W2  non-indiv clients:   {len(w2_cl)}")
    p(f"  W2  circles:             {len(w2_cir)}")
    p(f"  W3  relationships:       {len(w3_rel)}")
    p(f"  W3  dependants:          {len(w3_dep)}")
    p(f"  W3  income:              {len(w3_inc)}")
    p(f"  W3  expenditure:         via API (D19) — not built")
    p(f"  W4  wrapper plans:       {len(w4)}")
    p(f"  W5  other plans:         {len(w5)} (inc {len(prov_add)} Add-to-Plannr)")
    p(f"  W6  premiums:            {len(w6_prem)}")
    p(f"  W6  valuations:          {len(w6_val)} ({len(zero_blank_val)} at zero/no date)")
    add_tp_in_w7 = [r for r in add_ch_report if r.get("annual_rate_dec")]
    p(f"  W7  charges:             {len(w7)} ({len(w7)-len(rate_filled)-len(add_tp_in_w7)} sec A/B, "
      f"{len(rate_filled)} sec C, {len(add_tp_in_w7)} Add-to-Plannr)")
    p(f"       Add-to-Plannr total: {len(add_ch_report)} ({len(add_ch_report)-len(add_tp_in_w7)} no-rate/Rathbones not in W7)")
    p(f"       high rates >3%:     {len(high_rates)}")
    p(f"       rate-needed loaded: {len(rate_filled)}, omitted: {len(rate_blank)}")
    p(f"  W8  notes:               {len(w8)} ({rehomed} re-homed, {len(notes_no_cl)} skipped no account)")
    p(f"  Review entries:          {len(rev_applied)} applied, {len(rev_unapplied)} unapplied (see lists)")
    p(f"  Validate W1:             status={val['status']}, errors={val['error_count']}")
    p("=" * 68)


if __name__ == "__main__":
    main()
