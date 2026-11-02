"""
Build 'Ed Green - Plannr Import Pack (Gap Review)' from a ceding-CRM export.

Reads:  Ed_Green_Client_Data_-_Draft_v1.xlsx  (ceding CRM export)
        Plannr_Importer_Templates_08042026.xlsx (Plannr importer contract)
Writes: Ed_Green_Plannr_Import_Pack_Gap_Review_v1.xlsx

Tabs use exact Plannr template column order plus x_ helper columns at the right.
Cell colours: RED = mandatory, missing, Ed to complete. AMBER = present but
needs confirmation / inferred. GREY columns = Beyond-side, Ed ignores.
BLUE columns = source references and the externalreference migration key; do not edit.
Dropdown lists and sub_type validity are taken from Plannr_Reference_Data*.xlsx / the
importer template when present, so they match the destination site.

Re-run against Draft v2 by changing SOURCE below.
"""
import re
import os
import sys
import glob
import difflib
import subprocess

# Self-bootstrap so it runs under any Python (e.g. plain PowerShell), not only one that
# already has the data stack installed.
try:
    import pandas as pd
    from openpyxl import Workbook
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "--quiet", "pandas", "openpyxl"])
    import pandas as pd
    from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.comments import Comment


def _read_book(path):
    """Read all sheets of an xlsx; prefer calamine, fall back to openpyxl if it is not installed."""
    try:
        return pd.read_excel(path, sheet_name=None, engine="calamine")
    except Exception:
        return pd.read_excel(path, sheet_name=None, engine="openpyxl")

# SOURCE can be a FOLDER of the AR's CSV exports, or a combined .xlsx. Pass it as the first
# command-line argument. A folder is preferred: no 31-char tab-name truncation, no combine step.
if len(sys.argv) > 1:
    SOURCE = sys.argv[1]
elif os.path.isdir("SDA"):
    SOURCE = "SDA"
elif os.path.isfile("SDA.xlsx"):
    SOURCE = "SDA.xlsx"
else:
    SOURCE = "SDA"   # errors helpfully below if missing
TEMPLATES = "Plannr_Importer_Templates_08042026.xlsx"
GO_LIVE = "2026-11-02"          # charge start date / valuation date
DEST_ADVISER_UUID = "8dff5edc-2830-42d2-8a23-38ea8d6006bb"
MATCHES_FILE = r"C:\Users\markj\OneDrive - Beyond Advice\Technology - Documents\workspace\AR Migration\Higgs\stage1\higgs_matches_v4.xlsx"
_USE_MATCHES_EARLY = "--no-matches" not in sys.argv
OUTPUT_DIR = (r"C:\Users\markj\OneDrive - Beyond Advice\Technology - Documents\workspace\AR Migration\Higgs\stage2_build"
              if _USE_MATCHES_EARLY else
              r"C:\Users\markj\OneDrive - Beyond Advice\Technology - Documents\workspace\AR Migration\Higgs\stage2_regression")
os.makedirs(OUTPUT_DIR, exist_ok=True)
START_FALLBACK = "2026-01-01"   # plan start where no date_of_sale or date_of_transfer
REFDATA = "Plannr_Reference_Data*.xlsx"   # Plannr export of provider directory etc. (glob)


def _read_csv(path):
    try:
        return pd.read_csv(path, dtype=str, keep_default_na=False, encoding="utf-8-sig")
    except UnicodeDecodeError:
        try:
            return pd.read_csv(path, dtype=str, keep_default_na=False, encoding="latin-1")
        except pd.errors.EmptyDataError:
            return pd.DataFrame()
    except pd.errors.EmptyDataError:
        return pd.DataFrame()


if os.path.isdir(SOURCE):
    _raw = {}
    for _fn in sorted(os.listdir(SOURCE)):
        if _fn.lower().endswith(".csv") and not _fn.startswith("~$"):
            _raw[os.path.splitext(_fn)[0]] = _read_csv(os.path.join(SOURCE, _fn))
    if not _raw:
        sys.exit(f"No CSV files found in folder: {SOURCE}")
    print(f"reading {len(_raw)} CSV files from folder: {SOURCE}")
elif os.path.isfile(SOURCE) and SOURCE.lower().endswith((".xlsx", ".xlsm", ".xls")):
    _raw = _read_book(SOURCE)
    print(f"reading combined workbook: {SOURCE}")
else:
    sys.exit(f"SOURCE not found: '{SOURCE}'. Pass a folder of CSVs or a combined .xlsx as the first argument.")

# Derive the adviser name from the file/sheet prefixes (e.g. "Simon Field_01. ..." -> "Simon Field")
# so the same script works for any adviser, naming the output and the pack title accordingly.
def _detect_adviser(names):
    for _nm in names:
        _m = re.match(r"^(.+?)_\d{2}[.\-]", str(_nm))
        if _m and _m.group(1).strip():
            return _m.group(1).strip()
    return os.path.splitext(os.path.basename(str(SOURCE).rstrip("/\\")))[0] or "Import"


ADVISER = _detect_adviser(list(_raw.keys()))
_safe = re.sub(r"[^A-Za-z0-9]+", "_", ADVISER).strip("_") or "Import"
OUT = os.path.join(OUTPUT_DIR, "higgs_Merge_Review.xlsx")
OUT_BT = os.path.join(OUTPUT_DIR, "higgs_Merge_Review_BT.xlsx")
print(f"adviser: {ADVISER}  ->  {OUT}")


def _strip_prefix(nm):
    # drop a leading "<Adviser Name>_" the combined workbook added, leaving "NN. Sheet"
    return re.sub(r"^.*?_(?=\d{2}[.\-])", "", str(nm)).strip()


# Excel truncates sheet names to 31 chars and de-dupes, so exact-name matching is unreliable.
# Match each source sheet to the internal name the builder expects by a keyword in the
# (possibly truncated) name, disambiguating the two Income & Expenditure tables by content:
# the ENTRIES table is the one carrying an 'income_or_expenditure' column.
def _classify_sheet(nm, df):
    n = _strip_prefix(nm).lower()
    cols = {str(c).strip().lower() for c in df.columns}
    if "income_or_expenditure" in cols:
        return "24 - Income and Expenditure Ent"
    pairs = [("person", "01 - Person Clients"), ("business", "02 - Business Clients"),
             ("trust", "03 - Trust Clients"), ("defined benefit", "05 - Defined Benefit Pensions"),
             ("db pension", "05 - Defined Benefit Pensions"),
             ("topup", "06 - Topups and Withdrawals"), ("top up", "06 - Topups and Withdrawals"),
             ("withdrawal", "06 - Topups and Withdrawals"), ("polic", "04 - Policies"),
             ("review", "07 - Reviews"), ("vulnerab", "11 - Vulnerabilities"),
             ("address", "12 - Address History"), ("segment", "13 - Segmentation History"),
             ("employ", "14 - Employment History"), ("will", "16 - Wills"),
             ("propert", "18 - Properties"), ("aml", "19 - AML History"),
             ("need", "22 - Needs and Objectives"), ("objective", "22 - Needs and Objectives"),
             ("contact", "25 - Client Contact Log"),
             ("circumstance", "26 - Circumstances and Goals"), ("goal", "26 - Circumstances and Goals"),
             ("attorney", "15 - Power of Attorney"), ("power of", "15 - Power of Attorney")]
    for kw, internal in pairs:
        if kw in n:
            return internal
    return None


eg = {}
for _name, _df in _raw.items():
    _internal = _classify_sheet(_name, _df)
    if _internal and _internal not in eg:
        eg[_internal] = _df
for _needed in ["01 - Person Clients", "02 - Business Clients", "03 - Trust Clients",
                "04 - Policies", "05 - Defined Benefit Pensions",
                "06 - Topups and Withdrawals", "07 - Reviews", "11 - Vulnerabilities",
                "12 - Address History", "13 - Segmentation History",
                "14 - Employment History", "15 - Power of Attorney", "16 - Wills",
                "18 - Properties", "19 - AML History", "22 - Needs and Objectives",
                "24 - Income and Expenditure Ent", "25 - Client Contact Log",
                "26 - Circumstances and Goals"]:
    if _needed not in eg:
        eg[_needed] = pd.DataFrame()
print("sheets matched:", ", ".join(sorted(k for k in eg if not eg[k].empty)))
# Robustness: ensure every source column the builders reference (via any row variable, not just
# the main loop one) exists on every sheet, so a structural difference between AR exports cannot
# raise KeyError (missing -> blank -> flagged). Container/lookup dicts are excluded.
try:
    _self = open(__file__, encoding="utf-8").read()
    _skip = {"eg", "tp", "COLS", "REQ", "FIRM", "DROPDOWNS", "FREQ", "PM", "PM_NORM",
             "UNDER", "REVIEW_CADENCE", "REVIEW_CLASS_TO_STATUS", "NUM_TO_INTERNAL",
             "MEET", "EMPMAP", "_pc", "_refbook", "SHEET_ALIASES"}
    _refcols = {col for var, col in re.findall(r'(\w+)\["([^"]+)"\]', _self) if var not in _skip}
    # also harvest snake_case names that appear as plain quoted strings (e.g. columns read via
    # r[x] inside comprehensions), since over-adding empty columns is harmless but a missing one throws
    _refcols |= set(re.findall(r'"([a-z][a-z0-9_]{2,})"', _self))
except Exception:
    _refcols = set()
for _k in list(eg):
    _df = eg[_k].copy()
    for _c in _refcols:
        if _c not in _df.columns:
            _df[_c] = ""
    eg[_k] = _df

# ---- Load v4 matches (pass --no-matches to skip for regression test) ---------
import uuid as _uuid_mod

_USE_MATCHES = _USE_MATCHES_EARLY
if _USE_MATCHES and os.path.isfile(MATCHES_FILE):
    _mdf = pd.read_excel(MATCHES_FILE, sheet_name="Matches", dtype=str)
    _mdf = _mdf.fillna("")
else:
    _mdf = pd.DataFrame()
    if _USE_MATCHES:
        print(f"WARNING: MATCHES_FILE not found: {MATCHES_FILE} — running without matches")
    else:
        print("Running without matches (--no-matches flag)")

# Index by vision_policy_uuid -> row dict
_matches_by_puuid = {}
_matches_by_pn = {}
_adv_person_uuids = set()   # Vision person UUIDs on adviser sheet (any status)
_new_client_rows = []       # row_status == 'New-client'
_adviser_only_rows = []     # row_status == 'Adviser-only'
_ambiguous_puuids = set()   # Vision plan UUIDs that are Ambiguous

for _, _mr in _mdf.iterrows():
    _rs = str(_mr.get("row_status", "")).strip()
    _vuuid = str(_mr.get("vision_policy_uuid", "")).strip()
    _pmid = str(_mr.get("person_match_uuid", "")).strip()
    if _rs == "Ambiguous":
        if _vuuid:
            _ambiguous_puuids.add(_vuuid)
    elif _rs == "Matched":
        if _vuuid:
            _matches_by_puuid[_vuuid] = _mr.to_dict()
        _pn_norm_key = str(_mr.get("adv_policy_number", "")).strip().upper()
        _pn_norm_key = re.sub(r"[\s.\-/]", "", _pn_norm_key)
        if _pn_norm_key.endswith(".0"):
            _pn_norm_key = _pn_norm_key[:-2]
        if _pn_norm_key:
            _matches_by_pn[_pn_norm_key] = _mr.to_dict()
        if _pmid:
            _adv_person_uuids.add(_pmid)
    elif _rs == "Adviser-only":
        _adviser_only_rows.append(_mr.to_dict())
        if _pmid:
            _adv_person_uuids.add(_pmid)
    elif _rs == "New-client":
        _new_client_rows.append(_mr.to_dict())

def _adv_override(puuid):
    """Return the adviser match row for this Vision plan, or None."""
    return _matches_by_puuid.get(str(puuid).strip())

print(f"Matches loaded: {len(_matches_by_puuid)} matched, {len(_adviser_only_rows)} adviser-only, {len(_new_client_rows)} new-client, {len(_ambiguous_puuids)} ambiguous excluded")

# Fix B: Build set of person UUIDs who have at least one fee-bearing plan on adviser sheet
_clients_with_fee = set()
for _, _mfr in _mdf.iterrows():
    _mfrs = str(_mfr.get("row_status", "")).strip()
    if _mfrs not in ("Matched", "Adviser-only"):
        continue
    _mfleg = str(_mfr.get("adv_legacy_label", "")).strip().lower()
    if "trail" in _mfleg or "commission" in _mfleg:
        continue
    try:
        _mffee = float(str(_mfr.get("adv_fee_pct", "")).strip())
    except (ValueError, TypeError):
        _mffee = 0.0
    if _mffee > 0:
        _mfpid = str(_mfr.get("person_match_uuid", "")).strip()
        if _mfpid:
            _clients_with_fee.add(_mfpid)
print(f"Clients with fee-bearing plan: {len(_clients_with_fee)}")

tp = _read_book(TEMPLATES)

# review-cadence decode for the export's 'classification' code (per AR)
REVIEW_CADENCE = {"1": "6 Monthly Suitability Reviews", "2": "12 Monthly",
                  "3": "24 Monthly (legacy)", "4": "No ongoing review agreed",
                  "1.0": "6 Monthly Suitability Reviews", "2.0": "12 Monthly",
                  "3.0": "24 Monthly (legacy)", "4.0": "No ongoing review agreed"}

# classification code -> Plannr Tier 2 client status. EDIT HERE if the level/cadence pairing
# differs. 1 = 6-monthly = Pioneer, 2 = 12-monthly = Navigator, 3 = 24-monthly legacy (no live
# tier -> left blank and flagged), 4 = no ongoing review = Transactional.
REVIEW_CLASS_TO_STATUS = {
    "1": "LEVEL 1 / PIONEER", "1.0": "LEVEL 1 / PIONEER",
    "2": "LEVEL 2 / NAVIGATOR", "2.0": "LEVEL 2 / NAVIGATOR",
    "3": "", "3.0": "",                       # 24-monthly legacy - confirm
    "4": "TRANSACTIONAL", "4.0": "TRANSACTIONAL",
}
# classification codes that imply an ongoing review (so a missing ongoing charge is a flag)
ONGOING_REVIEW_CLASSES = {"1", "1.0", "2", "2.0", "3", "3.0"}


# ---------------------------------------------------------------- helpers
def s(v):
    """Clean scalar -> string ('' for blank). Fix cp1252 mojibake, trim."""
    if v is None or (isinstance(v, float) and pd.isna(v)) or pd.isna(v):
        return ""
    if isinstance(v, pd.Timestamp):
        return v.strftime("%Y-%m-%d")
    v = str(v)
    for _bad, _good in (("â€“", "-"), ("â€”", "-"), ("â€™", "'"), ("â€˜", "'"),
                        ("â€œ", '"'), ("â€¦", "..."), ("Â£", "£"), ("Â", ""),
                        ("â€", '"')):
        v = v.replace(_bad, _good)
    v = v.strip()
    if re.fullmatch(r"-?\d+\.0", v):   # numeric read as float (e.g. 6684610684.0) -> integer text
        v = v[:-2]
    return "" if v.lower() in ("nan", "nat", "none") else v

def d8(v):
    """Date -> YYYY-MM-DD or ''."""
    v = s(v)
    if not v:
        return ""
    try:
        return pd.to_datetime(v).strftime("%Y-%m-%d")
    except Exception:
        return v  # leave as-is; will amber via validation if needed

def num(v):
    if v is None or pd.isna(v):
        return ""
    try:
        f = float(v)
        return int(f) if f == int(f) else round(f, 2)
    except Exception:
        return s(v)

def yn(v, yes=("yes", "1", "true", "y", "1.0"), no=("no", "0", "false", "n", "0.0")):
    t = s(v).lower()
    if t in yes:
        return "yes"
    if t in no:
        return "no"
    return ""

FREQ = {"monthly": "Monthly", "annually": "Annually", "quarterly": "Quarterly",
        "weekly": "Weekly", "single": "One Off", "one off": "One Off",
        "daily": "Daily", "fortnightly": "Fortnightly", "biannually": "Biannually"}

def freq(v):
    return FREQ.get(s(v).lower(), s(v))

# ---------------------------------------------------------------- provider directory
def _pnorm(x):
    x = str(x).lower().strip()
    x = x.replace("&", " and ")
    x = re.sub(r"\b(ltd|limited|plc|llp|uk|group)\b", "", x)
    x = re.sub(r"[^a-z0-9 ]", " ", x)
    return re.sub(r"\s+", " ", x).strip()

PROV_EXACT = {}     # normalised name -> (Name, UUID)
PROV_ROWS = []      # (normalised name, Name, UUID)
PROV_OK = False
_refbook = {}
try:
    _pf = glob.glob(REFDATA)
    if _pf:
        _refbook = _read_book(_pf[0])
        _pv = _refbook.get("Providers")
        if _pv is not None:
            for _, _r in _pv.iterrows():
                k = _pnorm(_r["Name"])
                if not k:
                    continue
                PROV_ROWS.append((k, str(_r["Name"]), str(_r["UUID"])))
                if k not in PROV_EXACT:
                    PROV_EXACT[k] = (str(_r["Name"]), str(_r["UUID"]))
            PROV_OK = bool(PROV_ROWS)
        print(f"reference data loaded from {_pf[0]} ({len(PROV_ROWS)} providers)")
    else:
        print(f"reference data not found ({REFDATA}) - providers and dropdowns fall back to built-in lists")
except Exception as _e:
    print("reference data load failed:", _e)


def ref_list(sheet, col, fallback):
    """Distinct values from a reference sheet column, preserving order; fallback if absent."""
    df = _refbook.get(sheet)
    if df is None or col not in df.columns:
        return fallback
    out, seen = [], set()
    for v in df[col].dropna().tolist():
        t = str(v).strip()
        if t and t.lower() != col.lower() and t not in seen:
            seen.add(t); out.append(t)
    return out or fallback


def match_provider(name):
    """Return (provider_uuid, suggestion_note, matched_exact). uuid only on an exact match."""
    nm = s(name)
    if not nm or not PROV_OK:
        return "", "", False
    n = _pnorm(nm)
    if n in PROV_EXACT:
        return PROV_EXACT[n][1], "", True
    cands = []
    if len(n) >= 4:
        for k, Name, _u in PROV_ROWS:
            if n in k or k in n:
                cands.append((difflib.SequenceMatcher(None, n, k).ratio(), Name))
    if not cands:
        for k in difflib.get_close_matches(n, [r[0] for r in PROV_ROWS], n=3, cutoff=0.86):
            cands.append((0.9, next(r[1] for r in PROV_ROWS if r[0] == k)))
    best = sorted({c[1]: c[0] for c in cands}.items(), key=lambda x: -x[1])[:3]
    if best:
        return "", "Provider not matched - closest in directory: " + "; ".join(b[0] for b in best) + " (correct the name or set provider_uuid)", False
    return "", "Provider '" + nm + "' not in directory - select manually", False

# ---------------------------------------------------------------- template contracts
def template_cols(name):
    return [c for c in tp[name].columns if not str(c).startswith("Unnamed")]

def mandatory(howto):
    ht = tp[howto]
    fc = [c for c in ht.columns if "field" in str(c).lower()][0]
    mc = [c for c in ht.columns if "mandatory" in str(c).lower()][0]
    return set(ht[ht[mc].astype(str).str.strip().str.lower() == "mandatory"][fc])

T = {  # tab -> (template sheet, howto sheet)
    "Clients": ("Clients Template", "Client How To Use"),
    "Plans": ("Plans Template", "Plans How To"),
    "Valuations": ("Valuations", "Valuations How To"),
    "Premiums": ("Premiums", "Premiums How To"),
    "Charges": ("Charges", "Charges How To"),
    "Income & Employment": ("Income & Employment", "Income & Employment How To Use"),
    "Expenses": ("Expenses", "Expenses How To"),
    "Notes": ("Notes", "Notes How To"),
}
COLS = {k: template_cols(v[0]) for k, v in T.items()}
MAND = {k: mandatory(v[1]) for k, v in T.items()}

VALID_TITLES = set(ref_list("Titles", "title", list(tp["Client Dropdowns"]["title"].dropna())))

# Valid plan sub_types per type slug, from the importer's Plans Dropdowns (authoritative for
# import; note the reference data file uses different display strings for some sub_types).
def _build_subtype_valid():
    try:
        pd2 = tp["Plans Dropdowns"]
    except Exception:
        return {}
    cols = list(pd2.columns)
    out = {}
    for i, c in enumerate(cols):
        if str(c).startswith("sub_type") and i > 0:
            parent = next((str(v).strip() for v in pd2[cols[i - 1]].dropna().tolist()), "")
            slug = parent.lower().replace(" ", "_")
            if not slug:
                continue
            vals = {str(v).strip().lower() for v in pd2[c].dropna().tolist() if str(v).strip()}
            out.setdefault(slug, set()).update(vals)
    return out

SUBTYPE_VALID = _build_subtype_valid()


def subtype_note(typ, sub):
    """Return a flag note if sub is not an accepted Plannr sub_type for this type, else ''."""
    if not sub or not SUBTYPE_VALID or typ not in SUBTYPE_VALID:
        return ""
    if str(sub).strip().lower() in SUBTYPE_VALID[typ]:
        return ""
    return f"sub_type '{sub}' is not an accepted Plannr sub-type for type '{typ}' - pick one from the Plans Dropdowns sheet"


# Beyond-side (grey) columns per tab — Ed ignores these
FIRM = {
    "Clients": {"uuid", "status_uuid", "status_name", "assigned_adviser_uuid",
                "assigned_administrator_uuid", "assigned_paraplanner_uuid",
                "introduced_by_uuid", "send_welcome_email", "is_prospect",
                "prospect_converted_at", "value",
                "anniversary_review_date", "lead_source", "write_to_whom",
                "risk_uuid", "risk_level", "risk_level_string", "risk_completed_at",
                "risk_completed_by", "risk_completed_by_uuid",
                "risk_assessment_provider", "risk_time_horizon", "groups",
                "do_not_contact_via_fax", "do_not_contact_via_mail",
                "do_not_contact_via_sms", "do_not_contact_via_email",
                "do_not_contact_via_telephone", "do_not_visit", "do_not_market"},
    "Plans": {"uuid", "client_uuids", "provider_uuid", "seller_uuid", "seller_name",
              "active_adviser_uuid", "contract_enquiry_account_uuid",
              "visible_to_circles_uuids", "wrap_uuid", "tags", "on_panel",
              "show_on_circles", "visible_on_client_portal",
              "include_in_client_valuation", "is_wrap_product", "is_group_scheme",
              "originally_sold_by_firm", "linked_policy_numbers",
              "proposal_reference"},
    "Valuations": {"plan_uuid", "sub_account_uuid"},
    "Premiums": {"uuid", "plan_uuid", "transferred_from_plan_uuid",
                 "transferred_from_plan_type", "transferred_from_plan_sub_type"},
    "Charges": {"uuid", "plan_uuid", "client_uuid", "adviser_uuid", "premium_uuid",
                "category", "recurring", "type", "service_type", "tax_exempt",
                "exclude_from_rmar", "rmar_type", "advice_type",
                "advice_area_deposit", "advice_area_non_regulated",
                "advice_area_non_rmar", "advice_area_non_investment_insurance",
                "advice_area_regulated_mortgage", "advice_area_retail_investment",
                "advice_area_other_fca_regulated", "advice_area_unknown"},
    "Income & Employment": {"client_uuid", "employment_uuid"},
    "Expenses": {"client_uuid", "life_stage_name", "is_custom_expense",
                 "external_reference_company", "external_reference_value"},
    "Notes": {"model_type", "account_uuid", "plan_uuid", "risk_uuid", "case_uuid",
              "task_uuid", "path", "is_pinned", "visible_to_clients",
              "external_reference_source", "external_reference_value"},
}

# ---------------------------------------------------------------- product map
# ceding product -> (type, sub_type, amber_note or None)
PM = {}
def pm(prods, typ, sub, note=None):
    for p in prods:
        PM[p] = (typ, sub, note)

pm(["Discretionary Fund Management - MPS", "Discretionary Fund Management - Bespoke",
    "Discretionary Managed", "Discretionary Managed Service",
    "Discretionary Portfolio Service", "Bespoke Discretionary Investment Portfolio",
    "Managed Portfolio Service"], "investment", "Discretionary Managed Service")
pm(["ISA - Stocks and Shares", "ISA Stocks & Shares", "Stocks and Shares ISA",
    "ISA - Stocks and Shares "], "investment", "ISA (Stocks And Shares)")
pm(["ISA"], "investment", "ISA (Stocks And Shares)", "Generic 'ISA' - confirm S&S vs Cash")
pm(["General Investment Account", "General Investment Account (GIA)",
    "General Investment Account - GIA"], "investment", "General Investment Account")
pm(["SIPP", "Self Invested Personal Pension", "Self Invested Personal Pension (SIPP)"],
   "pension", "SIPP")
pm(["SIPP - Trust Based"], "pension", "SIPP", "Trust-based - confirm sub-type (Family SIPP?)")
pm(["Personal Pension Plan (PPP)"], "pension", "Personal Pension Plan")
pm(["Pension"], "pension", "Unknown", "Generic 'Pension' - specify sub-type")
pm(["Executive Pension Plan"], "pension", "Executive Pension Plan")
pm(["Small Self Administered Scheme (SSAS)"], "pension", "SSAS")
pm(["Business Group Personal Pensions"], "pension", "Group Personal Pension")
pm(["Final Salary"], "pension", "Final Salary Scheme")
pm(["Investment Bond"], "investment", "Investment Bond")
pm(["Investment Bond - Offshore", "Offshore Investment Bond", "Offshore Bond",
    "Premiere Account Offshore Bond"], "investment", "Offshore Bond")
pm(["Wealth Preservation Account", "Wealth Preservation Trust"], "investment",
   "Offshore Bond", "CLI Wealth Preservation - bond in trust; confirm")
pm(["Evolution Bond - Loan Trusts"], "investment", "Offshore Bond",
   "Utmost Evolution in loan trust - confirm")
pm(["Discounted Gift Trust"], "investment", "Trust", "DGT - confirm underlying bond type")
pm(["Unit Trusts OEIC"], "investment", "OEIC / Unit Trust")
pm(["Enterprise Investment Scheme", "Enterprise Investment Scheme (EIS)"],
   "investment", "Enterprise Investment Scheme (Requlated)",
   "Confirm regulated vs non-regulated")
pm(["Venture Capital Trust (VCT)", "Titan VCT"], "investment",
   "Venture Capital Trust (Regulated)", "Confirm regulated vs non-regulated")
pm(["AIM - Investment"], "investment", "AIM Investment Tax Service",
   "Assumed AIM IHT portfolio - confirm")
pm(["BPR", "Business Property Relief", "ITS", "IHT Portfolio",
    "EPS - Estate Planning Service", "Estate Planning (IHT)"], "investment",
   "Investment Tax Service", "BPR/IHT service - confirm sub-type")
pm(["Wrap Account"], "investment", "Wrap")
pm(["Premier Account"], "investment", "Unknown", "Confirm product type")
pm(["Investment Plan"], "investment", "Unknown", "Generic - specify sub-type")
pm(["Level Term Assurance"], "protection_policy", "Term Protection (Level)")
pm(["Term Assurance", "Term Protection"], "protection_policy", "Term Protection")
pm(["Relevant Life Cover"], "protection_policy", "Relevant Life Policy")
pm(["Whole of Life"], "protection_policy", "Whole of Life")
pm(["Flexible Life Plan"], "protection_policy", "Flexible Plan")
pm(["Lifetime Annuity Plan"], "annuity", "Lifetime Annuity")
pm(["Pension Annuity"], "annuity", "Pension Annuity")
# ---- SDA (Steve Dunn) additional products ------------------------------------------
pm(["Discretionary Fund Management - MPS", "Discretionary Fund Management - Bespoke",
    "Bespoke Discretionary Investment Portfolio", "Discretionary Portfolio Service (STPS)",
    "DFM", "DFM - STPS - AIM/IHT", "DFM - STPS - EIS"], "investment",
   "Discretionary Managed Service")
pm(["Self Invested Personal Pension (SIPP)"], "pension", "SIPP")
pm(["Personal Pension Plan (PPP)", "Personal Pension", "Platform Pension"],
   "pension", "Personal Pension Plan")
pm(["STPS ISA"], "investment", "ISA (Stocks And Shares)", "STPS ISA - confirm S&S vs cash")
pm(["Venture Capital Trust (VCT)"], "investment", "Venture Capital Trust (Regulated)",
   "Confirm regulated vs non-regulated")
pm(["Onshore Investment Bond"], "investment", "Investment Bond",
   "Onshore bond - confirm sub-type")
pm(["Offshore Investment Bond"], "investment", "Offshore Bond")
pm(["Income Drawdown or Other flexible income product"], "pension", "Income Drawdown",
   "Confirm drawdown sub-type")
pm(["Retirement Account", "Retirement Investment Account", "Platform Pension"],
   "pension", "Personal Pension Plan", "Confirm pension wrapper")
pm(["Critical Illness: Decreasing Term with Life Cover"], "protection_policy",
   "Term Protection (Decreasing)", "CIC with life cover - confirm sub-type")
pm(["Life Cover Lump Sum: Decreasing Term Assurance"], "protection_policy",
   "Term Protection (Decreasing)")
pm(["Life Term Assurance"], "protection_policy", "Term Protection",
   "Confirm level vs decreasing")
pm(["Income Protection Plan"], "protection_policy", "Income Protection")
pm(["AIM - Investment"], "investment", "AIM Investment Tax Service",
   "Assumed AIM IHT portfolio - confirm")
pm(["Estate Planning Service", "Specialist Tax Portfolio"], "investment",
   "Investment Tax Service", "IHT/tax service - confirm sub-type")
pm(["Group Personal Pension Plan", "Group Personal Pension Plan (GPPP)"], "pension",
   "Group Personal Pension")
pm(["Enhanced Annuity"], "annuity", "Enhanced Annuity")
pm(["Pension Annuity"], "annuity", "Pension Annuity")
pm(["Lifetime Annuity"], "annuity", "Lifetime Annuity")
pm(["Transferred Policy"], "investment", "Unknown",
   "Generic 'Transferred Policy' - specify the actual product")

# ---- Higgs Financial Solutions Ltd PM block --------------------------------
pm(["Retirement Annuity Pensions", "Retirement Annuity Pension"], "pension", "s226 RAC")
pm(["Beneficiary Income Release"], "pension", "Income Drawdown (Beneficiary)")
pm(["SIPP CASH", "SIPP Bank Acc"], "pension", "SIPP", "Separate plan per D7")
pm(["SSAS Bank Acc"], "pension", "SSAS", "Separate plan per D7")
pm(["Free Standing AVC", "FSAVC"], "pension", "FSAVC")
pm(["Stakeholder Pension Plan (SHP)", "Stakeholder Pension"], "pension", "Stakeholder Individual")
pm(["Adaptable Pension Plan", "Flexible Pension Plan"], "pension", "Personal Pension Plan")
pm(["(Ex-Investec) Wealth SIPP"], "pension", "SIPP")
pm(["DFM (Within SIPP)", "SIPP DFM"], "pension", "SIPP")
pm(["SSAS DFM"], "pension", "SSAS")
pm(["Annuity"], "annuity", "Pension Annuity", "Note: confirm sub-type")
pm(["Annuity: Enhanced", "Open Market Enhanced Annuity"], "annuity", "Enhanced Pension Annuity")
pm(["Legacy", "Stocks And Shares", "Investment Products", "Investment Account"], "investment", "Unknown", "Note: confirm")
pm(["Calderwood", "Charterhouse", "Rathbones"], "investment", "Discretionary Managed Service")
pm(["flexible protection"], "protection_policy", "Flexible Plan", "Note: confirm")
pm(["Life and Critical Illness"], "protection_policy", "Term Protection (Level - Life & CI)", "Note: confirm")
pm(["Portfolio Bond", "Sterling Investment Bond", "Flexibond", "Estate Planning Bond", "Initial Charge Bond"], "investment", "Investment Bond", "Note: confirm onshore")
pm(["DFM (Within Bond)", "Bond DFM"], "investment", "Investment Bond", "Note: confirm onshore")
pm(["DGEPS"], "investment", "Investment Tax Service", "Note: confirm")
pm(["Family Income Benefit"], "protection_policy", "Term Protection (Family Income Protection)")
pm(["Decreasing Term Assurance"], "protection_policy", "Term Protection (Decreasing Term)")
pm(["Flexible Inheritor Plan"], "investment", "Offshore Bond", "CLI FIP - offshore bond in trust")
pm(["(Ex-Investec) AIM IHT ISA"], "investment", "ISA (AIM)")
pm(["Group Life Cover", "Business Group Life Cover"], "group_protection", "Group Life")
pm(["Income Drawdown", "Flexible Access Drawdown", "Flexi-Access Drawdown"], "pension", "Income Drawdown (Flexi-Access)")
pm(["(Ex-Investec) FTSE 100 3YR Deposit"], "investment", "Structured Plan (Deposit)")
pm(["Business Protection: Key Person Life Cover", "Business Protection: Key Person CI"], "protection_policy", "Term Protection (Key Person Assurance)")
pm(["Level Life Cover", "Level Term"], "protection_policy", "Term Protection (Level)")
pm(["Junior ISA", "JISA"], "investment", "Junior ISA (Stocks And Shares)")
pm(["Collective Investment"], "investment", "Collective Investment Account")
pm(["(Ex-Investec) AIM IHT GIA", "AIM IHT GIA"], "investment", "AIM Investment Tax Service")
pm(["Private Medical Insurance"], "general_insurance", "Private Medical Insurance")
pm(["Regular Premium Whole Life Assurance", "Whole of Life Policy", "Whole of Life Assurance", "Lump Sum Benefits: Whole of Life"], "protection_policy", "Whole of Life")
pm(["Gift Trust"], "investment", "Trust")
pm(["Multi Benefit Protection"], "protection_policy", "Multi-benefit Protection")
pm(["Open-ended investment company (OEIC)", "Unit Trusts", "OEICs"], "investment", "OEIC / Unit Trust")
pm(["Income Protection"], "protection_policy", "Income Protection")
pm(["Apollo VCT"], "investment", "Venture Capital Trust (Regulated)")
pm(["Shareholder Protection Level Life Cover"], "protection_policy", "Term Protection (Directors Share Protection)")
pm(["Liquidity Account"], "cash_account", "Cash Account")
pm(["Business Protection Term Assurance"], "protection_policy", "Term Protection (Level)")
pm(["GIA & ISA"], "investment", "Wrap", "D5: flag for adviser to split if unmatched")

# Fix invalid sub_types in existing PM entries
for _k, _v in list(PM.items()):
    _typ, _sub, _note = _v
    if _sub == "Term Protection (Decreasing)":
        PM[_k] = (_typ, "Term Protection (Decreasing Term)", _note)
    elif _sub == "Income Drawdown":
        PM[_k] = (_typ, "Income Drawdown (Flexi-Access)", _note)
    elif _sub == "Enhanced Annuity":
        PM[_k] = (_typ, "Enhanced Pension Annuity", _note)
    elif _sub == "Term Protection":
        PM[_k] = (_typ, "Term Protection (Level)", _note)

# Deliberately unmapped -> RED + flag:
NON_PLANS = {"Fee (no product)", "CAP ADJUSTMENT", "Fee"}

# Build a normalised lookup so messy variants (case, trailing space/tab, full stops,
# '(Ex-Investec)' prefixes, 'STPS' tags) resolve to the same mapping.
def _prodnorm(p):
    p = s(p).lower()
    p = re.sub(r"\(ex-investec\)", "", p)
    p = p.replace(".", " ").replace("&", " and ")
    p = re.sub(r"[^a-z0-9 ]", " ", p)
    return re.sub(r"\s+", " ", p).strip()


PM_NORM = {}
for _k, _v in PM.items():
    PM_NORM.setdefault(_prodnorm(_k), _v)
_NONPLAN_NORM = {_prodnorm(x) for x in NON_PLANS}


def pm_lookup(prod):
    """Return ('NONPLAN'|None) or (type, sub, note). Exact first, then normalised."""
    if prod in NON_PLANS:
        return "NONPLAN"
    if prod in PM:
        return PM[prod]
    n = _prodnorm(prod)
    if n in _NONPLAN_NORM:
        return "NONPLAN"
    return PM_NORM.get(n)

INCOME_TYPE = [
    (r"state pension", "Pension"), (r"pension", "Pension"),
    (r"salary|annual income|earnings", "Salary"), (r"dividend", "Dividend"),
    (r"bonus", "Bonus"), (r"rental|rent", "Rental income"),
    (r"investment|interest", "Investment income"), (r"benefit", "Benefit"),
]
def income_type(name):
    n = name.lower()
    for pat, t in INCOME_TYPE:
        if re.search(pat, n):
            return t, ""
    return "Other", name

# Map a source expense name to one of Plannr's three sections, using the category
# groupings configured in Plannr. Source names are kept as-is (custom expenses);
# this only decides the section. No match -> blank + red.
EXP_SECTION = [
    # explicit section words, if the source already labels them
    (r"\bessential\b", "Essential"),
    (r"\blifestyle\b", "Lifestyle"),
    (r"\bdiscretionary\b", "Discretionary"),
    # Discretionary: Regular Savings, Pension Contributions, Gifting, Gambling
    (r"saving|pension|\bgift|donation|charit|gambl|betting|\bbet\b|lottery|casino|wager", "Discretionary"),
    # Lifestyle: Hobbies/Club, Holidays, Socialising/Eating Out, Cigarettes/Tobacco/Vaping,
    #            Personal Care, Professional Fees, Subscriptions, Pets
    (r"hobb|\bclub|members|\bgym\b|leisure|\bsport", "Lifestyle"),
    (r"holiday|vacation", "Lifestyle"),
    (r"social|eating out|dining|restaurant|\bpub\b|takeaway|take-away|cinema|entertain", "Lifestyle"),
    (r"cigarett|tobacco|vap|smoking", "Lifestyle"),
    (r"personal care|hairdress|\bhair\b|beauty|grooming|cosmetic|toiletr|\bsalon", "Lifestyle"),
    (r"professional fee|accountant|solicitor|legal fee|advis(?:e|o)r fee", "Lifestyle"),
    (r"subscription|netflix|spotify|disney|streaming|magazine|newspaper", "Lifestyle"),
    (r"\bpet\b|\bpets\b|\bvet\b|veterin", "Lifestyle"),
    # Essential: Mortgage (main/other), Utilities, Food/Household, Travel/Commuting,
    #            Loans/Credit Cards/Other Debt, Insurance, Clothing, TV/Internet/Phone,
    #            Mobile Phones, Childcare/School Fees
    (r"mortgage|\brent\b", "Essential"),
    (r"utilit|\bgas\b|electric|\bwater\b|council tax|energy|heating", "Essential"),
    (r"food|grocer|household|supermarket", "Essential"),
    (r"travel|commut|\bfuel\b|petrol|diesel|transport|\btrain\b|\bbus\b|parking|\bcar\b", "Essential"),
    (r"\bloan|credit card|\bdebt\b|hire purchase|\bfinance\b|overdraft", "Essential"),
    (r"insurance|protection|life cover", "Essential"),
    (r"clothing|clothes|shoes|footwear", "Essential"),
    (r"\btv\b|internet|\bphone|broadband|landline|mobile", "Essential"),
    (r"childcare|childmind|\bnanny\b|au pair|\bschool|nursery|tuition|education", "Essential"),
]
def exp_section(name):
    n = name.lower()
    for pat, sec in EXP_SECTION:
        if re.search(pat, n):
            return sec
    return ""

# ---------------------------------------------------------------- gather clients
P = eg["01 - Person Clients"]
B = eg["02 - Business Clients"]
TR = eg["03 - Trust Clients"]

client_uuids = (set(P["client_uuid"].dropna()) | set(B["primary_contact_uuid"].dropna())
                | set(TR["client_uuid"].dropna()))

def latest(df, key, datecol):
    """dict key -> latest row by datecol"""
    out = {}
    for _, r in df.iterrows():
        k = s(r.get(key))
        if not k:
            continue
        dt = d8(r.get(datecol)) or "0000"
        if k not in out or dt >= out[k][0]:
            out[k] = (dt, r)
    return {k: v[1] for k, v in out.items()}

addr_cur = {}
for _, r in eg["12 - Address History"].iterrows():
    if s(r["status"]).lower() == "current":
        k = s(r["uuid"])
        prev = addr_cur.get(k)
        if prev is None or (d8(r["moved_in_date"]) or "0000") >= (d8(prev["moved_in_date"]) or "0000"):
            addr_cur[k] = r

wills_cur = {s(r["client_uuid"]): r for _, r in eg["16 - Wills"].iterrows()
             if s(r["status"]).lower() == "current"}
vuln = latest(eg["11 - Vulnerabilities"], "uuid", "date_of_assessment")
aml = latest(eg["19 - AML History"], "client_uuid", "date_performed")
seg_cur = {s(r["client_uuid"]): s(r["segment"])
           for _, r in eg["13 - Segmentation History"].iterrows()
           if s(r["segment_status"]).lower() == "current"}
rev = latest(eg["07 - Reviews"], "client_1_uuid", "date_of_reivew")
rev2 = latest(eg["07 - Reviews"], "client_2_uuid", "date_of_reivew")
for k, r in rev2.items():
    if k not in rev or (d8(r["date_of_reivew"]) or "0000") > (d8(rev[k]["date_of_reivew"]) or "0000"):
        rev[k] = r

# POA (sheet 15): client UUID in col A; merge col C (scope) + ", " + col E (attorneys)
poa_by_ref = {}
_poa = eg.get("15 - Power of Attorney")
if _poa is not None and not _poa.empty:
    # pick the column with the most UUID-like values (the real client-uuid column; columns
    # added by the safe-access reindex are empty and score zero)
    _best, _bestn = None, 0
    for _c in _poa.columns:
        _n = sum(1 for _v in _poa[_c].tolist()
                 if re.fullmatch(r"[0-9a-fA-F]{8}-[0-9a-fA-F-]{4,}", s(_v)))
        if _n > _bestn:
            _best, _bestn = _c, _n
    if _best is not None:
        for _, r in _poa.iterrows():
            cu = s(r[_best])
            if cu and cu not in poa_by_ref:
                poa_by_ref[cu] = "Power of attorney recorded in source - review details"

rows_clients, fl_clients = [], []   # row dicts, flag dicts {(col): 'RED'/'AMBER'}
flags_text = []
_title_email_changes = []  # Fix C: (ref_short, field, vision_val, adv_val) — Rule 9: ref only
# Stage 2 Edit H: person-level title/email from v4 matches file
# Index by person_match_uuid (Vision UUID); take first non-blank value per UUID
_adv_title_by_puuid = {}   # vision_client_uuid -> adv_title
_adv_email_by_puuid = {}   # vision_client_uuid -> adv_email
for _, _hr in _mdf.iterrows():
    _hp = str(_hr.get("person_match_uuid", "")).strip()
    if not _hp: continue
    _ht = str(_hr.get("adv_title", "")).strip()
    _he = str(_hr.get("adv_email", "")).strip()
    if _ht and _hp not in _adv_title_by_puuid:
        _adv_title_by_puuid[_hp] = _ht
    if _he and _hp not in _adv_email_by_puuid:
        _adv_email_by_puuid[_hp] = _he

# Pre-scan policies for ongoing charges so the client status (and its flag) can use it
# (clients are built before plans).
clients_with_ongoing = set()
for _, _pr in eg["04 - Policies"].iterrows():
    _f = s(_pr["ongoing_service_charge_frequency"]).lower()
    _p = num(_pr["ongoing_service_charge_percentage"])
    if (_f in ("monthly", "quarterly", "annually", "biannually", "fortnightly", "weekly")
            or (str(_p) not in ("", "0") and float(_p or 0) > 0)):
        for _o in (s(_pr["client_1_uuid"]), s(_pr["client_2_uuid"])):
            if _o:
                clients_with_ongoing.add(_o)

# mobiles cannot be imported via the importer -> shown for API load (first-wins per ref)
mobile_by_ref = {}
for _, r in P.iterrows():
    k = s(r["client_uuid"])
    if k and k not in mobile_by_ref:
        mobile_by_ref[k] = s(r["mobile_number"])

def base(cols):
    return {c: "" for c in cols}

corrupt_refs, dup_refs, seen = [], [], set()
_navigator_upgraded_count = 0
_navigator_stayed_below_count = 0
_health_dropped_count = 0        # Health flag suppressed (sheet 01 says good health)
_blanked_emails = []             # (externalref, original_email) for Fix-after-import
_nc_nameless_excluded = []       # new-client keys excluded for having no name at all

def _clean_email(email):
    """Return a single valid-format email or blank. Handles '-' sentinels and multi-email cells."""
    e = str(email).strip()
    if not e or e.lower() in ("-", "n/a", "na", "none", "nan"):
        return ""
    # Multi-email: space-separated or comma-separated with multiple @
    if (e.count("@") > 1) or (" " in e and "@" in e):
        candidates = re.split(r"[\s,;]+", e)
        for cand in candidates:
            cand = cand.strip()
            if "@" in cand and cand.count("@") == 1:
                return cand
        return ""
    if "@" not in e:
        return ""
    return e
for _, r in P.iterrows():
    c = base(COLS["Clients"])
    f, note = {}, []
    ref = s(r["client_uuid"])
    if ref in seen:
        dup_refs.append(ref)
        note.append("Duplicate row for this client ref in source")
        f["first_name"] = f.get("first_name", "AMBER")
    seen.add(ref)
    corrupted = not s(r["forename"]) and not s(r["surname"])
    if corrupted:
        corrupt_refs.append(ref)
        note.append("Source row corrupted on export (fields shifted) - please re-extract this client")
    c["type"] = "individual"
    _vis_title_orig = s(r["title"])
    title = _vis_title_orig
    if title and title not in VALID_TITLES:
        note.append(f"Title field contained stray text: '{title[:40]}'")
        title = ""
    # Fix H / Stage 2 Edit H: person-level title — adviser wins where both have a value
    _adv_title_h = _adv_title_by_puuid.get(ref, "")
    if _adv_title_h and _adv_title_h in VALID_TITLES:
        if not title:
            title = _adv_title_h
        elif title != _adv_title_h:
            # Fix C: track override for report
            _title_email_changes.append((ref[:8], "title", _vis_title_orig, _adv_title_h))
            title = _adv_title_h
    c["title"] = title
    fore = s(r["forename"]).split(" ", 1)
    c["first_name"], c["middle_name"] = fore[0], (fore[1] if len(fore) > 1 else "")
    c["last_name"] = s(r["surname"])
    c["maiden_name"] = s(r["maiden_name"])
    c["nickname"] = s(r["preferred_name"])
    c["gender"] = s(r["sex"])
    ni = s(r["national_insurance_number"]).replace(" ", "").upper()
    c["ni_number"] = ni
    if ni and not re.fullmatch(r"[A-Z]{2}\d{6}[A-D]?", ni):
        f["ni_number"] = "AMBER"; note.append("NI number not in standard format")
    c["marital_status"] = s(r["marital_status"])
    c["date_of_birth"] = d8(r["date_of_birth"])
    c["target_retirement_age"] = num(r["intended_retirement_age"])
    # Fix H / Stage 2 Edit H: person-level email — adviser wins where both have a value
    _vis_email = s(r["email"])
    _adv_email_h = _adv_email_by_puuid.get(ref, "")
    if _adv_email_h and _vis_email and _adv_email_h != _vis_email:
        _title_email_changes.append((ref[:8], "email", "vision-has-value", "adviser-override"))
    _raw_email = _adv_email_h if _adv_email_h else _vis_email
    _cleaned_email = _clean_email(_raw_email)
    # If adviser email was bad, try Vision's as fallback
    if not _cleaned_email and _adv_email_h and _vis_email and _adv_email_h != _vis_email:
        _cleaned_email = _clean_email(_vis_email)
    if _raw_email and not _cleaned_email:
        _blanked_emails.append((ref, _raw_email))
    c["email"] = _cleaned_email
    c["phone"] = s(r["home_number"]) or s(r["work_number"])
    nat = s(r["nationality"])
    nat2 = s(r["second_nationality"])
    c["nationalities"] = ", ".join(x for x in (nat, nat2) if x)
    c["country_of_residence"] = s(r["residency"])
    c["is_uk_resident"] = "yes" if s(r["residency"]) == "United Kingdom" else ""
    c["is_uk_domicile"] = "yes" if s(r["domocile"]) == "United Kingdom" else ""
    c["is_in_receipt_of_your_state_pension"] = yn(r["in_receipt_of_state_pension"])
    c["has_obtained_a_forecast"] = yn(r["state_pension_forecast_received"])
    yc = num(r["years_of_ni_contributions"])
    c["years_of_national_insurance_contributions"] = yc if yc != 0 else ""
    smk, vap = s(r["smokes_tobacco"]), s(r["vapes"])
    c["smoker_status"] = ("smoker" if smk in ("1", "1.0")
                          else "vaper" if vap in ("1", "1.0")
                          else "non-smoker" if smk in ("0", "0.0") else "")
    igh = s(r["in_good_health"])
    c["current_state_of_health"] = "good" if igh in ("1", "1.0") else ""
    if igh in ("0", "0.0"):
        f["current_state_of_health"] = "AMBER"; note.append("Source: not in good health - state fair/poor")
    hn = [s(r["health_notes"]), s(r["smoked_tobacco_details"])]
    c["health_notes"] = "; ".join(x for x in hn if x)
    c["client_regulatory_classification"] = "Retail Client"   # blanket: advised individuals
    f["client_regulatory_classification"] = "AMBER"
    c["is_prospect"] = "no"
    # Tier 2 status from the review-cadence classification; flag a 'level' client with no ongoing charge
    _cls = s(r["classification"])
    c["status_name"] = REVIEW_CLASS_TO_STATUS.get(_cls, "")
    if _cls in ("3", "3.0"):
        f["status_name"] = "AMBER"
        note.append("Classification 3 (24-monthly legacy) has no live service tier - set status manually")
    elif _cls in ONGOING_REVIEW_CLASSES and ref not in clients_with_ongoing:
        f["status_name"] = "AMBER"
        note.append(f"On an ongoing review cadence ({REVIEW_CADENCE.get(_cls, '?')}) but no ongoing charge found on any plan - confirm proposition or charge")
    # Fix B: minimum NAVIGATOR only for clients with a fee-paying plan (not trail commission)
    _pre_upgrade_status = c["status_name"]
    if _pre_upgrade_status not in ("LEVEL 1 / PIONEER", "LEVEL 2 / NAVIGATOR") and ref in _clients_with_fee:
        c["status_name"] = "LEVEL 2 / NAVIGATOR"
        f["status_name"] = "AMBER"
        note.append("Status upgraded to NAVIGATOR: client has fee-paying plan on adviser sheet")
        _navigator_upgraded_count = _navigator_upgraded_count + 1
    elif _pre_upgrade_status not in ("LEVEL 1 / PIONEER", "LEVEL 2 / NAVIGATOR"):
        _navigator_stayed_below_count = _navigator_stayed_below_count + 1
    # enrichment
    a = addr_cur.get(ref)
    if a is not None:
        c["address_1"], c["address_2"], c["address_3"] = s(a["address_line_1"]), s(a["address_line_2"]), s(a["address_line_3"])
        town, city = s(a["town"]), s(a["city"])
        c["address_4"] = town if (town and city) else ""
        c["town_city"] = city or town
        c["county"], c["postcode"], c["country"] = s(a["county"]), s(a["postcode"]), s(a["country"])
    w = wills_cur.get(ref)
    if w is not None:
        c["has_will"] = "yes"
        c["will_last_reviewed"] = d8(w["date_will_last_reviewed"])
        c["will_executors"] = s(w["executors"])
        wn = []
        if d8(w["date_will_set_up"]):
            wn.append(f"Set up {d8(w['date_will_set_up'])}")
        if s(w["beneficiaries"]):
            wn.append(f"Beneficiaries: {s(w['beneficiaries'])}")
        c["will_notes"] = "; ".join(wn)
    v = vuln.get(ref)
    if v is not None:
        c["vulnerability_last_assessment_date"] = d8(v["date_of_assessment"])
        inds = {"Life Events": s(v["major_life_events"]), "Resilience": s(v["low_resilience"]),
                "Capability": s(v["low_capacity"]),
                "Health": "Yes" if s(v["in_good_health"]).lower() == "no" else "No"}
        pos = [k for k, x in inds.items() if x.lower() == "yes"]
        vi = s(v["vulnerability_indicators"])
        # If Health is the only indicator but sheet 01 says good health, it's a Vision default — suppress
        _health_suppressed = False
        if "Health" in pos and igh in ("1", "1.0"):
            pos = [k for k in pos if k != "Health"]
            _health_suppressed = True
            note.append("Health vulnerability suppressed: in_good_health=1 on sheet 01 (likely Vision default on sheet 11) — verify with adviser")
            f["vulnerability_sub_type"] = "AMBER"
            _health_dropped_count += 1
        _VST_MAP = {"Health": "health", "Life Events": "life_events",
                    "Resilience": "resilience", "Capability": "capability"}
        if _health_suppressed and not pos and not vi:
            # Health was the only reason and it was spurious — blank both fields
            c["is_vulnerable"] = ""
            c["vulnerability_sub_type"] = ""
        else:
            c["is_vulnerable"] = "yes" if (pos or vi) else ("no" if any(x.lower() == "no" for x in inds.values()) else "")
            c["vulnerability_sub_type"] = _VST_MAP.get(pos[0], pos[0].lower().replace(" ", "_")) if pos else ""
        sup = s(v.get("GROUP_CONCAT(di.extra_supporting_actions SEPARATOR '|')", "")) \
            + ("; " + s(v.get("GROUP_CONCAT(di.detail_supporting_actions SEPARATOR '|')", "")) if s(v.get("GROUP_CONCAT(di.detail_supporting_actions SEPARATOR '|')", "")) else "")
        c["vulnerability_details"] = "; ".join(x for x in (vi, sup.strip("; ")) if x)
    am = aml.get(ref)
    if am is not None:
        c["checked_sanctions_at"] = d8(am["date_performed"])
        f["checked_sanctions_at"] = "AMBER"
    if ref in poa_by_ref:
        c["has_power_of_attorney"] = "yes"
        c["power_of_attorney_notes"] = poa_by_ref[ref]
    rv = rev.get(ref)
    if rv is not None:
        c["previous_review_date"] = d8(rv["date_of_reivew"])
        try:
            c["next_review_date"] = (pd.to_datetime(c["previous_review_date"]) + pd.DateOffset(years=1)).strftime("%Y-%m-%d")
            f["next_review_date"] = "AMBER"
        except Exception:
            pass
    rows_clients.append(c); fl_clients.append(f)
    flags_text.append("; ".join(note))

# Fix A (6d): generate stable business UUID per (pcu, business_name) — one person can own many businesses
_biz_uuid_map = {}  # (primary_contact_uuid, business_name) -> biz_entity_uuid
for _, _br in B.iterrows():
    _bpcu = s(_br["primary_contact_uuid"])
    _bname = s(_br["business_name"])
    _bkey = (_bpcu, _bname)
    if _bpcu and _bkey not in _biz_uuid_map:
        _biz_uuid_map[_bkey] = str(_uuid_mod.uuid5(
            _uuid_mod.NAMESPACE_DNS, "higgs_biz_" + _bname + "_" + _bpcu))
_biz_uuid_distinct = sum(1 for buuid in _biz_uuid_map.values() if buuid != list(_biz_uuid_map.keys())[0][0])
print(f"Business UUID map: {len(_biz_uuid_map)} business entries across {len({k[0] for k in _biz_uuid_map})} contacts")

biz_seen = set()
biz_primary_uuids = set()  # person UUIDs that are primary contacts for a business (person UUIDs)
for _, r in B.iterrows():
    pcu_raw = s(r["primary_contact_uuid"])
    _bname_r = s(r["business_name"])
    ref = _biz_uuid_map.get((pcu_raw, _bname_r), pcu_raw)   # new biz entity UUID per (pcu, name)
    biz_primary_uuids.add(pcu_raw)               # track person UUID for plan flagging
    if ref in biz_seen:
        dup_refs.append(ref); continue
    biz_seen.add(ref)
    c = base(COLS["Clients"]); f, note = {}, []
    c["type"] = "business"
    c["entity_name"] = s(r["business_name"])
    c["is_prospect"] = "no"
    f["client_regulatory_classification"] = "RED"
    pcu = pcu_raw   # person UUID stays as owner ref
    c["owners"] = pcu
    if pcu:
        f["owners"] = "AMBER"
        note.append(f"Owner = primary contact {s(r['primary_contact_name'])} (source UUID, remapped at import)")
    reg = [s(r.get(x, "")) for x in ("address_line_1", "address_line_2", "address_town",
                                     "address_city", "address_county", "address_postcode", "address_country")]
    cor = [s(r.get(x, "")) for x in ("correspondance_address_line_1", "correspondance_address_line_2",
                                     "correspondance_address_town", "correspondance_address_city",
                                     "correspondance_address_county", "correspondance_address_postcode",
                                     "correspondance_address_country")]
    use = reg if any(reg) else cor
    if use == cor and any(cor):
        note.append("Registered address blank - correspondence address used")
    c["address_1"], c["address_2"], c["address_4"] = use[0], use[1], use[2]
    c["town_city"], c["county"], c["postcode"], c["country"] = use[3] or use[2], use[4], use[5], use[6]
    if use[3]:
        c["address_4"] = use[2]
    c["phone"] = s(r["business_phone_number"])
    if s(r["vat_number"]):
        note.append(f"VAT number {s(r['vat_number'])} - no importer field, hold for record")
    rows_clients.append(c); fl_clients.append(f); flags_text.append("; ".join(note))

for _, r in TR.iterrows():
    c = base(COLS["Clients"]); f, note = {}, []
    c["type"] = "trust"
    c["entity_name"] = s(r["client_name"])
    c["is_prospect"] = "no"
    f["client_regulatory_classification"] = "RED"
    c["owners"] = ""
    f["owners"] = "RED"
    note.append("Trust requires an individual owner - supply the client UUID (none in source)")
    c["address_1"], c["address_2"] = s(r["address_line_1"]), s(r["address_line_2"])
    town, city = s(r["town"]), s(r["city"])
    c["address_4"] = town if (town and city) else ""
    c["town_city"] = city or town
    c["county"], c["postcode"], c["country"] = s(r["county"]), s(r["postcode"]), s(r["country"])
    rows_clients.append(c); fl_clients.append(f); flags_text.append("; ".join(note))

client_refs = ([s(r["client_uuid"]) for _, r in P.iterrows()]
               + [s(r["primary_contact_uuid"]) for _, r in B.iterrows() if s(r["primary_contact_uuid"]) in biz_seen]
               + [s(r["client_uuid"]) for _, r in TR.iterrows()])
# dedupe biz path may shorten; rebuild aligned refs
client_refs = []
seen2, biz_seen2 = set(), set()
for _, r in P.iterrows():
    client_refs.append(s(r["client_uuid"]))
for _, r in B.iterrows():
    _pcu_raw2 = s(r["primary_contact_uuid"])
    _bname_r2 = s(r["business_name"])
    ref = _biz_uuid_map.get((_pcu_raw2, _bname_r2), _pcu_raw2)   # Fix A: per (pcu, name)
    if ref not in biz_seen2:
        biz_seen2.add(ref); client_refs.append(ref)
for _, r in TR.iterrows():
    client_refs.append(s(r["client_uuid"]))

# ---------------------------------------------------------------- plans
rows_plans, fl_plans, plan_flags, plan_refs, plan_src = [], [], [], [], []
plan_owner_names, plan_product_src = [], []   # source values shown next to mapped fields
policy_val_rows = []                            # (ref, name, adv_value, adv_date) 4-tuple for wired plans
orphan_plan_owners = set()
not_on_his_list = []                            # (policy_uuid, name, provider, fee_pct) — off-sheet with fee
possible_dup_puuids = set()                     # Vision plans flagged POSSIBLE_DUP in v4
# Collect possible-dup UUIDs from matches file
for _, _pd_row in _mdf.iterrows():
    for _cand in str(_pd_row.get("candidates", "")).split("|"):
        _cand = _cand.strip()
        if _cand.startswith("POSSIBLE_DUP:"):
            for _pu in _cand[13:].split("|"):
                if _pu.strip():
                    possible_dup_puuids.add(_pu.strip())
POL = eg["04 - Policies"]
UNDER = {"yes": "yes", "servicing_agent": "yes", "information_authority": "yes", "none": "no"}


def own_split(owners):
    """client_uuids (comma-joined) + client_ownership_percentage (parallel, equal split)."""
    owners = [o for o in owners if o]
    if not owners:
        return "", ""
    if len(owners) == 1:
        return owners[0], "100"
    n = len(owners)
    each = round(100.0 / n, 2)
    pcts = [each] * n
    pcts[-1] = round(100.0 - each * (n - 1), 2)
    fmt = lambda x: str(int(x)) if float(x).is_integer() else str(x)
    return ",".join(owners), ",".join(fmt(p) for p in pcts)


def set_policy_number(c, f, note, raw, is_fixed_asset):
    pn = s(raw)
    if is_fixed_asset:
        c["policy_number"] = pn
        return
    if not pn:
        c["policy_number"] = "unknown"
        f["policy_number"] = "AMBER"
        note.append("No policy number in source - defaulted to 'unknown'; please supply the real number")
    else:
        c["policy_number"] = pn
        if "|" in pn:
            note.append("Multiple policy numbers - enter the single number the valuation feed must match")

_off_sheet_orphan_skipped = 0
for i, r in POL.iterrows():
    _puuid = s(r["policy_uuid"])
    _is_matched = _puuid in _matches_by_puuid
    _is_ambiguous = _puuid in _ambiguous_puuids
    _is_off_sheet = not _is_matched and not _is_ambiguous
    # For off-sheet plans: skip orphan-owned ones (no adviser row to supply owner)
    if _is_off_sheet:
        _c1_test = s(r["client_1_uuid"]); _c2_test = s(r["client_2_uuid"])
        _c1_ok = _c1_test and _c1_test in client_uuids
        _c2_ok = _c2_test and _c2_test in client_uuids
        if not _c1_ok and not _c2_ok:
            _off_sheet_orphan_skipped += 1
            continue
        # Track fee-paying off-sheet for "Not on his list" tab
        _vis_ocp = num(r["ongoing_service_charge_percentage"])
        if str(_vis_ocp) not in ("", "0") and float(_vis_ocp or 0) > 0:
            not_on_his_list.append((
                _puuid,
                s(r["product_specific_name"]) or s(r["product"]),
                s(r["primary_provider"]),
                _vis_ocp,
            ))
    c = base(COLS["Plans"]); f, note = {}, []
    _adv = _adv_override(_puuid) or {}
    # Stage 2: merge provider — use adviser's where it matches directory, else Vision's
    _adv_prov = str(_adv.get("adv_provider", "")).strip()
    _vis_prov = s(r["primary_provider"])
    _pu_match, _, _pe_match = match_provider(_adv_prov) if _adv_prov else ("", "", False)
    prod_raw = s(r["product"])
    # Stage 2: merge product — keep Vision's unless it maps to Unknown or different Plannr type
    _adv_type = str(_adv.get("adv_type", "")).strip()
    _vis_pm = pm_lookup(prod_raw)
    _adv_pm = pm_lookup(_adv_type) if _adv_type else None
    if _vis_pm and _vis_pm != "NONPLAN" and _vis_pm[1] != "Unknown":
        prod = prod_raw  # use Vision's product
        if _adv_pm and _adv_pm != "NONPLAN" and _adv_pm[0] != _vis_pm[0]:
            prod = _adv_type
            note.append(f"Product overridden: Vision '{prod_raw}' maps to {_vis_pm[0]}, adviser type '{_adv_type}' maps to {_adv_pm[0]} - used adviser")
    else:
        if _adv_pm and _adv_pm != "NONPLAN":
            prod = _adv_type
            note.append(f"Product overridden: Vision '{prod_raw}' maps to Unknown, using adviser type '{_adv_type}'")
        else:
            prod = prod_raw
    # Stage 2: owners — Vision's stand; replace orphan owners with adviser's matched person
    _c1 = s(r["client_1_uuid"]); _c2 = s(r["client_2_uuid"])
    _adv_person_uuid = str(_adv.get("person_match_uuid", "")).strip()
    if _c1 and _c1 not in client_uuids and _adv_person_uuid:
        _c1 = _adv_person_uuid
        note.append("Orphan owner replaced with adviser's matched person")
    if _c2 and _c2 not in client_uuids:
        _c2 = ""
        note.append("Orphan co-owner dropped (no Vision record)")
    owners = [_c1] + ([_c2] if _c2 else [])
    for o in owners:
        if o and o not in client_uuids:
            orphan_plan_owners.add(o)
            f["client_uuids"] = "AMBER"; note.append(f"Owner ref {o[:8]}... has no client record in this export")
    c["client_uuids"], c["client_ownership_percentage"] = own_split(owners)
    if len(owners) > 1:
        f["client_ownership_percentage"] = "AMBER"
        note.append("Joint plan - assumed equal ownership, confirm split")
    c["name"] = s(r["product_specific_name"]) or prod
    # Rule 5: his policy number replaces Vision's on every matched plan
    _adv_pn_r5 = str(_adv.get("adv_policy_number", "")).strip() if _is_matched else ""
    _pn_source = _adv_pn_r5 if _adv_pn_r5 else s(r["policy_numbers"])
    set_policy_number(c, f, note, _pn_source, is_fixed_asset=False)
    if _adv_pn_r5 and _adv_pn_r5 != s(r["policy_numbers"]):
        note.append(f"Policy number replaced: Vision '{s(r['policy_numbers'])[:30]}' → adviser '{_adv_pn_r5[:30]}'")

    # Stage 2: use adviser's provider if it matches directory, else Vision's
    if _pe_match and _adv_prov:
        c["provider"] = _adv_prov
        c["provider_uuid"] = _pu_match
        if _adv_prov.lower() != _vis_prov.lower():
            note.append(f"Provider overridden: adviser '{_adv_prov}' matched directory; Vision had '{_vis_prov}'")
    else:
        c["provider"] = _vis_prov if _vis_prov else _adv_prov
        _pu, _pn, _pe = match_provider(c["provider"])
        if _pe:
            c["provider_uuid"] = _pu
        elif c["provider"]:
            f["provider"] = "AMBER"
            if _pn:
                note.append(_pn)
    owner_nm = " & ".join(x for x in [s(r["client_1"]), s(r["client_2"])] if x)
    _pmr = pm_lookup(prod)
    if _pmr == "NONPLAN":
        f["type"] = f["sub_type"] = "RED"
        note.append(f"'{prod}' is not a plan - confirm whether this is a fee/adjustment record to drop or re-key")
    elif _pmr:
        typ, sub, n2 = _pmr
        c["type"], c["sub_type"] = typ, sub
        if n2:
            f["sub_type"] = "AMBER"; note.append(n2)
        _stn = subtype_note(c["type"], c["sub_type"])
        if _stn and "sub_type" not in f:
            f["sub_type"] = "AMBER"; note.append(_stn)
    else:
        f["type"] = f["sub_type"] = "RED"; note.append(f"No mapping for product '{prod}'")
    c["status"] = "Active"
    _sale, _xfer, _legacy = d8(r["date_of_sale"]), d8(r["date_of_transfer"]), d8(r["plan_start_date"])
    if _sale:
        c["opened_at"], _src = _sale, "date_of_sale"
    elif _xfer:
        c["opened_at"], _src = _xfer, "date_of_transfer"
    elif _legacy:
        c["opened_at"], _src = _legacy, "plan_start_date"
    else:
        c["opened_at"], _src = "2026-01-01", "defaulted 2026-01-01"
    f["opened_at"] = "AMBER"
    note.append(f"Start date from {_src} - AR to approve")
    agt = s(r["agency_transfer"]).lower()
    c["agency_in_your_name"] = "yes" if agt == "yes" else ("no" if agt == "no" else "")
    if _is_off_sheet or _is_ambiguous:
        # Off-sheet/ambiguous: no AR fee, force under_advice=no (Vision value may say yes for stale reasons)
        c["under_advice"] = "no"
        c["agency_in_your_name"] = "no"
        note.append("Off-sheet plan: not on adviser's sheet; under_advice set to no, no charge created")
    else:
        c["under_advice"] = UNDER.get(s(r["under_my_advice"]).lower(), "")
        # Ongoing charge present -> force under advice + agency in your name; agency-from = plan start.
        _ocf = s(r["ongoing_service_charge_frequency"]).lower()
        _ocp = num(r["ongoing_service_charge_percentage"])
        has_ongoing = (_ocf in ("monthly", "quarterly", "annually", "biannually", "fortnightly", "weekly")
                       or (str(_ocp) not in ("", "0") and float(_ocp or 0) > 0))
        if has_ongoing:
            c["under_advice"] = "yes"
            c["agency_in_your_name"] = "yes"
            for o in owners:
                if o:
                    clients_with_ongoing.add(o)
    if c["agency_in_your_name"] == "yes":
        c["agency_in_your_name_from"] = c["opened_at"]
        f["agency_in_your_name_from"] = "AMBER"
    c["show_on_circles"] = "yes"
    c["visible_on_client_portal"] = "yes"
    c["include_in_client_valuation"] = "yes"
    c["is_wrap_product"] = "yes" if prod == "Wrap Account" else "no"
    c["review_date"] = ""
    rows_plans.append(c); fl_plans.append(f); plan_flags.append("; ".join(note))
    plan_refs.append(s(r["policy_uuid"])); plan_src.append(f"04 - Policies r{i+2}")
    plan_owner_names.append(owner_nm); plan_product_src.append(prod)
    # Valuation: matched/ambiguous plans use adviser col K dated col W; off-sheet plans have no adviser row
    if _is_matched:
        _val_adv_str = str(_adv.get("adv_value", "")).strip()
        _val_date_str = str(_adv.get("last_updated", "")).strip()
        try:
            _val_adv = float(_val_adv_str) if _val_adv_str else None
        except ValueError:
            _val_adv = None
        if _val_adv and _val_adv >= 10:
            policy_val_rows.append((_puuid, c["name"], _val_adv, _val_date_str))

DB = eg["05 - Defined Benefit Pensions"]
for i, r in DB.iterrows():
    c = base(COLS["Plans"]); f, note = {}, []
    # This CRM's DB export can put an internal id in client_1_uuid rather than the client.
    # Prefer the owner column(s) that actually match the client master.
    _cands = [s(r["client_1_uuid"]), s(r["client_2_uuid"])]
    _matched = [o for o in _cands if o and o in client_uuids]
    owners = _matched if _matched else ([o for o in _cands if o] or [s(r["client_1_uuid"])])
    for o in owners:
        if o and o not in client_uuids:
            orphan_plan_owners.add(o); f["client_uuids"] = "AMBER"
            note.append(f"Owner ref {o[:8]}... has no client record in this export")
    c["client_uuids"], c["client_ownership_percentage"] = own_split(owners)
    if len(owners) > 1:
        f["client_ownership_percentage"] = "AMBER"
        note.append("Joint plan - assumed equal ownership, confirm split")
    c["name"] = s(r["scheme_name"])
    c["type"], c["sub_type"] = "pension", "Defined Benefit Pension"
    if subtype_note(c["type"], c["sub_type"]):
        f["sub_type"] = "AMBER"; note.append(subtype_note(c["type"], c["sub_type"]))
    c["status"] = "Active"
    set_policy_number(c, f, note, "", is_fixed_asset=False)
    c["provider"] = s(r["scheme_name"])
    _pu, _pn, _pe = match_provider(c["provider"])
    if _pe:
        c["provider_uuid"] = _pu
    f["provider"] = "AMBER"; note.append("Scheme name used as provider - confirm administrator")
    c["pension_plan_retirement_age"] = num(r["scheme_Retirement_age"])
    c["pension_projected_income"] = num(r["ROUND(db.full_pension,2)"])
    c["pension_tax_free_amount"] = num(r["ROUND(db.tax_free_cash,2)"])
    red = num(r["ROUND(db.reduced_pension,2)"])
    if red != "":
        note.append(f"Reduced pension {red} - no importer field, hold for fact-find")
    c["show_on_circles"] = c["visible_on_client_portal"] = c["include_in_client_valuation"] = "yes"
    c["is_wrap_product"] = "no"
    f["opened_at"] = "AMBER"
    rows_plans.append(c); fl_plans.append(f); plan_flags.append("; ".join(note))
    plan_refs.append(s(r["defined_benfit_uuid"])); plan_src.append(f"05 - DB Pensions r{i+2}")
    plan_owner_names.append(" & ".join(x for x in [s(r["client_1_name"]), s(r["client_2_name"])] if x))
    plan_product_src.append("DB Pension (sheet 05)")

# Pre-initialize charge lists so adviser-only block can append to them
rows_chg, fl_chg, chg_refs = [], [], []
PAID = {"client": "Client", "provider": "Provider", "dfm": "Provider", "platform": "Provider"}

# ---- Stage 2: Adviser-only plans (person in Vision, plan new to Vision) ----
_adv_only_added = 0
for _ao in _adviser_only_rows:
    _person_uuid = str(_ao.get("person_match_uuid", "")).strip()
    if not _person_uuid or _person_uuid not in client_uuids:
        continue
    _ao_type = str(_ao.get("adv_type", "")).strip()
    _ao_pm = pm_lookup(_ao_type)
    if not _ao_pm or _ao_pm == "NONPLAN":
        continue
    c = base(COLS["Plans"]); f, note = {}, []
    c["client_uuids"] = _person_uuid; c["client_ownership_percentage"] = "100"
    c["name"] = _ao_type
    _ao_typ, _ao_sub, _ao_note = _ao_pm
    c["type"] = _ao_typ; c["sub_type"] = _ao_sub
    if _ao_note:
        f["sub_type"] = "AMBER"; note.append(_ao_note)
    c["status"] = "Active"
    _ao_prov = str(_ao.get("adv_provider", "")).strip()
    c["provider"] = _ao_prov
    _ao_pu, _ao_pnote, _ao_pe = match_provider(_ao_prov)
    if _ao_pe:
        c["provider_uuid"] = _ao_pu
    elif _ao_prov:
        f["provider"] = "AMBER"
        if _ao_pnote:
            note.append(_ao_pnote)
    _ao_pn_raw = str(_ao.get("adv_policy_number", "")).strip()
    c["policy_number"] = _ao_pn_raw if _ao_pn_raw else "unknown"
    if not _ao_pn_raw:
        f["policy_number"] = "AMBER"; note.append("No policy number")
    c["opened_at"] = GO_LIVE; f["opened_at"] = "AMBER"
    c["show_on_circles"] = c["visible_on_client_portal"] = c["include_in_client_valuation"] = "yes"
    c["is_wrap_product"] = "no"
    c["under_advice"] = "yes"; c["agency_in_your_name"] = "yes"
    c["agency_in_your_name_from"] = GO_LIVE; f["agency_in_your_name_from"] = "AMBER"
    note.append("Adviser-only plan: added from adviser sheet, not in Vision")
    rows_plans.append(c); fl_plans.append(f); plan_flags.append("; ".join(note))
    _ao_ref = "adv_only_" + str(_uuid_mod.uuid5(_uuid_mod.NAMESPACE_DNS, f"higgs_{_ao.get('adv_row', '')}_ao"))
    plan_refs.append(_ao_ref); plan_src.append(f"Adviser sheet row {_ao.get('adv_row', '?')}")
    plan_owner_names.append(""); plan_product_src.append(_ao_type)
    # Valuation for adviser-only plan — 4-tuple (ref, name, value, date)
    _ao_val_str = str(_ao.get("adv_value", "")).strip()
    _ao_date_str = str(_ao.get("last_updated", "")).strip()
    try:
        _ao_val = float(_ao_val_str) if _ao_val_str else None
    except ValueError:
        _ao_val = None
    if _ao_val and _ao_val >= 10:
        policy_val_rows.append((_ao_ref, c["name"], _ao_val, _ao_date_str))
    # Charge for adviser-only plan
    _ao_fee_str = str(_ao.get("adv_fee_pct", "")).strip()
    _ao_legacy = str(_ao.get("adv_legacy_label", "")).strip().lower()
    try:
        _ao_fee = float(_ao_fee_str) if _ao_fee_str else None
    except ValueError:
        _ao_fee = None
    if _ao_fee and _ao_fee > 0 and "trail" not in _ao_legacy and "commission" not in _ao_legacy:
        cc = base(COLS["Charges"]); cf = {}
        cc["payment_origin"] = "Provider"
        _ao_fq = str(_ao.get("adv_frequency", "")).strip().lower()
        if "quarter" in _ao_fq:
            cc["frequency"] = "Quarterly"
        elif "annual" in _ao_fq:
            cc["frequency"] = "Annually"
        else:
            cc["frequency"] = "Monthly"
        if "invoice" in _ao_fq:
            cc["payment_origin"] = "Client"
        cc["type"] = "Fee"; cc["category"] = "ongoing_fee"; cc["recurring"] = "yes"
        cc["percentage"] = round(_ao_fee, 4); cc["percentage_frequency"] = "Annually"
        cc["tax_exempt"] = "yes"; cc["exclude_from_rmar"] = "no"
        cc["service_type"] = "Ongoing"; cc["start_date"] = GO_LIVE
        cf["start_date"] = "AMBER"
        rows_chg.append(cc); fl_chg.append(cf)
        chg_refs.append((_ao_ref, "Adviser-only plan charge"))
    _adv_only_added += 1

print(f"Adviser-only plans added: {_adv_only_added}")

# ---- Stage 2 Edit G: New-client wiring (Fix D: dedup by person identity) ----
# New-client rows: persons entirely new to Vision; generate stable UUID, add client + plan + val + chg
_new_client_uuids = {}   # adv_row_str -> generated uuid (for cross-referencing)
_new_client_added = 0
new_client_added_list = []  # for "Added" tab

# Fix D: dedup new-client rows by person identity (surname+forename+postcode)
_nc_dedup = {}    # identity key -> canonical row
_nc_by_key = {}   # identity key -> list of all rows (all plans for this person)
for _nc in _new_client_rows:
    _nk = (str(_nc.get("adv_surname", "")).strip().upper(),
           str(_nc.get("adv_forename", "")).strip().upper(),
           str(_nc.get("adv_postcode", "")).strip().upper())
    if _nk not in _nc_dedup:
        _nc_dedup[_nk] = _nc
    _nc_by_key.setdefault(_nk, []).append(_nc)
print(f"New-client rows: {len(_new_client_rows)} rows -> {len(_nc_dedup)} distinct persons")

for _nk, _canonical_nc in _nc_dedup.items():
    # Skip rows with no surname AND no forename — nothing to import
    if not _nk[0] and not _nk[1]:
        _adv_rows_skip = [str(_nr.get("adv_row", "")).strip() for _nr in _nc_by_key[_nk]]
        _nc_nameless_excluded.append({"adv_rows": _adv_rows_skip, "postcode": _nk[2]})
        continue
    # One stable ref per person
    _nc_ref = "nc_" + str(_uuid_mod.uuid5(_uuid_mod.NAMESPACE_DNS,
                                           "higgs_nc_" + _nk[0] + "_" + _nk[1] + "_" + _nk[2]))
    # Register all row numbers for this person
    for _nrow in _nc_by_key[_nk]:
        _nc_row_s = str(_nrow.get("adv_row", "")).strip()
        _new_client_uuids[_nc_row_s] = _nc_ref
    # Build ONE client record from canonical row
    c_nc = base(COLS["Clients"]); f_nc, note_nc = {}, []
    c_nc["type"] = "individual"
    # Title — from adv_title in v4 matches
    _nc_title = str(_canonical_nc.get("adv_title", "")).strip()
    if _nc_title and _nc_title in VALID_TITLES:
        c_nc["title"] = _nc_title
    elif _nc_title:
        c_nc["title"] = ""; note_nc.append(f"Title '{_nc_title[:20]}' not in valid list — confirm")
    # Name — from matches file (Rule 9: not printed)
    _nc_surname = str(_canonical_nc.get("adv_surname", "")).strip()
    _nc_forename = str(_canonical_nc.get("adv_forename", "")).strip()
    _nc_fore_parts = _nc_forename.split(" ", 1)
    c_nc["first_name"] = _nc_fore_parts[0]
    c_nc["middle_name"] = _nc_fore_parts[1] if len(_nc_fore_parts) > 1 else ""
    c_nc["last_name"] = _nc_surname
    if not _nc_surname:
        f_nc["first_name"] = "RED"; note_nc.append("Surname missing — supply from adviser sheet")
    if not c_nc["first_name"]:
        # Forename missing: use "Unknown" and list for adviser
        c_nc["first_name"] = "Unknown"
        f_nc["first_name"] = "AMBER"
        note_nc.append("first_name unknown — supply from adviser sheet")
    # Email
    _nc_email_raw = str(_canonical_nc.get("adv_email", "")).strip()
    _nc_email = _clean_email(_nc_email_raw)
    if _nc_email_raw and not _nc_email:
        _blanked_emails.append((_nc_ref, _nc_email_raw))
    c_nc["email"] = _nc_email
    # Address — parse from adv_address
    _nc_addr_raw = str(_canonical_nc.get("adv_address", "")).strip()
    _nc_postcode = str(_canonical_nc.get("adv_postcode", "")).strip()
    if _nc_addr_raw:
        _nc_parts = [x.strip() for x in _nc_addr_raw.split(",") if x.strip()]
        _nc_parts_no_pc = [p for p in _nc_parts if not re.search(r'[A-Z]{1,2}[0-9]', p.upper())]
        c_nc["address_1"] = _nc_parts_no_pc[0] if len(_nc_parts_no_pc) > 0 else ""
        c_nc["address_2"] = _nc_parts_no_pc[1] if len(_nc_parts_no_pc) > 1 else ""
        c_nc["town_city"] = _nc_parts_no_pc[-1] if len(_nc_parts_no_pc) > 2 else (_nc_parts_no_pc[1] if len(_nc_parts_no_pc) > 1 else "")
        c_nc["postcode"] = _nc_postcode
    c_nc["is_prospect"] = "no"
    c_nc["client_regulatory_classification"] = "Retail Client"; f_nc["client_regulatory_classification"] = "AMBER"
    # New client: Navigator only if any of their plans has a fee
    _nc_has_fee = any(
        float(str(_nr.get("adv_fee_pct", "0") or "0")) > 0 and
        "trail" not in str(_nr.get("adv_legacy_label", "")).lower() and
        "commission" not in str(_nr.get("adv_legacy_label", "")).lower()
        for _nr in _nc_by_key[_nk]
    )
    c_nc["status_name"] = "LEVEL 2 / NAVIGATOR" if _nc_has_fee else "TRANSACTIONAL"
    f_nc["status_name"] = "AMBER"
    note_nc.append("New client: added from adviser sheet, not in Vision source")
    rows_clients.append(c_nc); fl_clients.append(f_nc)
    flags_text.append("; ".join(note_nc))
    client_refs.append(_nc_ref)
    client_uuids.add(_nc_ref)
    _adv_person_uuids.add(_nc_ref)
    if _nc_has_fee:
        clients_with_ongoing.add(_nc_ref)
    new_client_added_list.append(_nc_ref)
    _new_client_added += 1
    # Plans — one per adviser row for this person
    for _nc_plan_row in _nc_by_key[_nk]:
        _nc_row = str(_nc_plan_row.get("adv_row", "")).strip()
        _nc_type = str(_nc_plan_row.get("adv_type", "")).strip()
        _nc_pm = pm_lookup(_nc_type)
        if _nc_pm == "NONPLAN":
            continue
        if not _nc_pm:
            _nc_pm = ("investment", "Unknown",
                      f"No PM mapping for '{_nc_type}' - confirm type and sub_type")
        cp = base(COLS["Plans"]); fp, notep = {}, []
        cp["client_uuids"] = _nc_ref; cp["client_ownership_percentage"] = "100"
        cp["name"] = _nc_type
        _nc_ptyp, _nc_psub, _nc_pnote = _nc_pm
        cp["type"] = _nc_ptyp; cp["sub_type"] = _nc_psub
        if _nc_pnote: fp["sub_type"] = "AMBER"; notep.append(_nc_pnote)
        cp["status"] = "Active"
        _nc_prov = str(_nc_plan_row.get("adv_provider", "")).strip()
        cp["provider"] = _nc_prov
        _nc_pu, _nc_pnote2, _nc_pe = match_provider(_nc_prov)
        if _nc_pe: cp["provider_uuid"] = _nc_pu
        elif _nc_prov: fp["provider"] = "AMBER"
        _nc_pn_raw = str(_nc_plan_row.get("adv_policy_number", "")).strip()
        cp["policy_number"] = _nc_pn_raw if _nc_pn_raw else "unknown"
        if not _nc_pn_raw: fp["policy_number"] = "AMBER"; notep.append("No policy number")
        cp["opened_at"] = GO_LIVE; fp["opened_at"] = "AMBER"
        cp["show_on_circles"] = cp["visible_on_client_portal"] = cp["include_in_client_valuation"] = "yes"
        cp["is_wrap_product"] = "no"
        cp["under_advice"] = "yes"; cp["agency_in_your_name"] = "yes"
        cp["agency_in_your_name_from"] = GO_LIVE; fp["agency_in_your_name_from"] = "AMBER"
        notep.append("New-client plan: added from adviser sheet")
        _nc_plan_ref = "nc_plan_" + str(_uuid_mod.uuid5(_uuid_mod.NAMESPACE_DNS, f"higgs_nc_plan_{_nc_row}"))
        rows_plans.append(cp); fl_plans.append(fp); plan_flags.append("; ".join(notep))
        plan_refs.append(_nc_plan_ref); plan_src.append(f"Adviser sheet row {_nc_row} (new client)")
        plan_owner_names.append(""); plan_product_src.append(_nc_type)
        # Valuation
        _nc_val_str = str(_nc_plan_row.get("adv_value", "")).strip()
        _nc_date_str = str(_nc_plan_row.get("last_updated", "")).strip()
        try: _nc_val = float(_nc_val_str) if _nc_val_str else None
        except ValueError: _nc_val = None
        if _nc_val and _nc_val >= 10:
            policy_val_rows.append((_nc_plan_ref, cp["name"], _nc_val, _nc_date_str))
        # Charge
        _nc_fee_str = str(_nc_plan_row.get("adv_fee_pct", "")).strip()
        _nc_legacy = str(_nc_plan_row.get("adv_legacy_label", "")).strip().lower()
        try: _nc_fee = float(_nc_fee_str) if _nc_fee_str else None
        except ValueError: _nc_fee = None
        if _nc_fee and _nc_fee > 0 and "trail" not in _nc_legacy and "commission" not in _nc_legacy:
            cc_nc = base(COLS["Charges"]); cf_nc = {}
            cc_nc["payment_origin"] = "Provider"
            _nc_fq = str(_nc_plan_row.get("adv_frequency", "")).strip().lower()
            if "quarter" in _nc_fq: cc_nc["frequency"] = "Quarterly"
            elif "annual" in _nc_fq: cc_nc["frequency"] = "Annually"
            else: cc_nc["frequency"] = "Monthly"
            if "invoice" in _nc_fq: cc_nc["payment_origin"] = "Client"
            cc_nc["type"] = "Fee"; cc_nc["category"] = "ongoing_fee"; cc_nc["recurring"] = "yes"
            cc_nc["percentage"] = round(_nc_fee, 4); cc_nc["percentage_frequency"] = "Annually"
            cc_nc["tax_exempt"] = "yes"; cc_nc["exclude_from_rmar"] = "no"
            cc_nc["service_type"] = "Ongoing"; cc_nc["start_date"] = GO_LIVE
            cf_nc["start_date"] = "AMBER"
            rows_chg.append(cc_nc); fl_chg.append(cf_nc)
            chg_refs.append((_nc_plan_ref, "New-client plan charge"))

print(f"New-client records added: {_new_client_added} persons")

PR = eg["18 - Properties"]
prop_val_rows = []
for i, r in PR.iterrows():
    c = base(COLS["Plans"]); f, note = {}, []
    owners = [s(x) for x in s(r["owner_uuids"]).split("|") if s(x)]
    for o in owners:
        if o not in client_uuids:
            orphan_plan_owners.add(o); f["client_uuids"] = "AMBER"
            note.append(f"Owner ref {o[:8]}... has no client record in this export")
    c["client_uuids"], c["client_ownership_percentage"] = own_split(owners)
    if len(owners) > 1:
        f["client_ownership_percentage"] = "AMBER"
        note.append("Joint asset - assumed equal ownership, confirm split")
    c["name"] = s(r["property_name"]) or "Property"
    c["type"], c["sub_type"] = "fixed_asset", "Property"
    if subtype_note(c["type"], c["sub_type"]):
        f["sub_type"] = "AMBER"; note.append(subtype_note(c["type"], c["sub_type"]))
    prop_owner_nm = s(r["owners"])
    st = s(r["property_status"])
    c["status"] = st if st in ("Active", "Inactive") else "Active"
    if st not in ("Active", "Inactive"):
        f["status"] = "AMBER"; note.append("Property status unknown - confirm")
    if c["status"] == "Inactive":
        c["inactive_at"] = d8(r["date_sold"])
        c["inactive_reason"] = "complete"
    c["fixed_asset_address_1"] = s(r["address"])
    c["fixed_asset_postcode"] = s(r["postcode"])
    c["show_on_circles"] = c["visible_on_client_portal"] = c["include_in_client_valuation"] = "yes"
    c["is_wrap_product"] = "no"
    rows_plans.append(c); fl_plans.append(f); plan_flags.append("; ".join(note))
    pref = s(r["property_uuid"])
    plan_refs.append(pref); plan_src.append(f"18 - Properties r{i+2}")
    plan_owner_names.append(prop_owner_nm); plan_product_src.append("Property (sheet 18)")
    cv = num(r["current_value"])
    if cv != "" and float(cv or 0) > 0 and c["status"] == "Active":
        prop_val_rows.append((pref, c["name"], cv))
    if s(r["has_current_mortgage"]).lower() == "yes":
        m = base(COLS["Plans"]); mf, mn = {}, []
        m["client_uuids"] = c["client_uuids"]
        m["client_ownership_percentage"] = c["client_ownership_percentage"]
        if len(owners) > 1:
            mf["client_ownership_percentage"] = "AMBER"
        m["name"] = f"Mortgage - {c['name']}"
        m["type"] = "mortgage"
        mt = s(r["mortgage_type"])
        m["sub_type"] = ("Mortgage Buy to Let" if mt == "Buy to Let" else "Mortgage")
        if mt == "Other":
            mf["sub_type"] = "AMBER"; mn.append("Mortgage type 'Other' - specify")
        if subtype_note(m["type"], m["sub_type"]) and "sub_type" not in mf:
            mf["sub_type"] = "AMBER"; mn.append(subtype_note(m["type"], m["sub_type"]))
        m["status"] = "Active"
        m["provider"] = s(r["mortgage_lender"])
        _mu, _mn, _me = match_provider(m["provider"])
        if _me:
            m["provider_uuid"] = _mu
        elif m["provider"]:
            mf["provider"] = "AMBER"
            if _mn:
                mn.append(_mn)
        set_policy_number(m, mf, mn, r["mortgage_reference"], is_fixed_asset=False)
        m["mortgage_outstanding_balance"] = num(r["outstanding_mortgage_amount"])
        m["mortgage_payment_amount"] = num(r["monthly_mortgage_payment_amount"])
        m["mortgage_payment_frequency"] = "Monthly" if m["mortgage_payment_amount"] != "" else ""
        rt = s(r["mortgage_rate_type"]).lower()
        m["mortgage_interest_type"] = rt if rt in ("fixed", "variable") else ""
        m["mortgage_interest_rate"] = num(r["mortgage_rate_percentage"])
        m["mortgage_initial_period_end"] = d8(r["mortgage_rate_end"])
        m["mortgage_repayment_type"] = ("repayment" if mt == "Repayment"
                                        else "interest_only" if mt == "Interest Only" else "")
        m["mortgage_repayment_term"] = num(r["remaining_term"])
        m["show_on_circles"] = m["visible_on_client_portal"] = m["include_in_client_valuation"] = "yes"
        m["is_wrap_product"] = "no"
        mf["opened_at"] = "AMBER"
        rows_plans.append(m); fl_plans.append(mf); plan_flags.append("; ".join(mn))
        plan_refs.append(pref + " (mortgage)"); plan_src.append(f"18 - Properties r{i+2}")
        plan_owner_names.append(prop_owner_nm); plan_product_src.append("Mortgage (sheet 18)")

# ---------------------------------------------------------------- valuations
rows_val, fl_val, val_refs, val_notes = [], [], [], []
for pref, nm, cv in prop_val_rows:
    c = base(COLS["Valuations"]); f = {}
    c["value"] = cv
    c["valuation_date"] = GO_LIVE
    f["valuation_date"] = "AMBER"
    rows_val.append(c); fl_val.append(f); val_refs.append((pref, nm))
    val_notes.append("Property current value; valuation date assumed = go-live")
# Stage 2: valuations from adviser col K (4-tuple: ref, name, value, date)
_val_under10 = 0
_val_used = 0
for entry in policy_val_rows:
    pref, nm, _adv_val, _adv_date_str = entry  # 4-tuples throughout
    if _adv_val < 10:
        _val_under10 += 1
        continue
    c = base(COLS["Valuations"]); f = {}
    c["value"] = round(_adv_val, 2)
    c["valuation_date"] = str(_adv_date_str)[:10] if _adv_date_str and len(str(_adv_date_str)) >= 10 else GO_LIVE
    if not _adv_date_str:
        f["valuation_date"] = "AMBER"
    rows_val.append(c); fl_val.append(f); val_refs.append((pref, nm))
    val_notes.append("Adviser col K value, dated col W")
    _val_used += 1
print(f"Valuations: policy_val_rows={len(policy_val_rows)}, used={_val_used}, under-£10={_val_under10}")

# ---------------------------------------------------------------- premiums
rows_prem, fl_prem, prem_refs, prem_notes = [], [], [], []
for i, r in POL.iterrows():
    if s(r["frequency"]).lower() in ("", "none"):
        continue
    amt = num(r["reoccurring_contribution"])
    if amt == "" or float(amt or 0) == 0:
        continue
    c = base(COLS["Premiums"]); f = {}
    c["amount"] = amt
    c["frequency"] = freq(r["frequency"])
    f["source"] = "AMBER"; f["payment_method"] = "AMBER"; f["start_date"] = "AMBER"
    rows_prem.append(c); fl_prem.append(f)
    prem_refs.append(s(r["policy_uuid"]))
    prem_notes.append("Recurring contribution on policy; source/method/start optional - add if known")

TU = eg["06 - Topups and Withdrawals"]
withdrawal_count = 0
for i, r in TU.iterrows():
    if s(r["topup_or_withdrawal"]).lower() == "withdrawal":
        withdrawal_count += 1
        continue
    c = base(COLS["Premiums"]); f, note = {}, []
    c["amount"] = num(r["topup_amount"])
    fq = s(r["reoccurring_contribution_frequency"]).lower()
    c["frequency"] = freq(fq) if fq not in ("", "none", "single") else "One Off"
    c["start_date"] = d8(r["date_of_topup_withdrawal"])
    if not c["start_date"]:
        f["start_date"] = "AMBER"
    c["date_advice_given"] = c["start_date"]
    f["source"] = "AMBER"; f["payment_method"] = "AMBER"
    rows_prem.append(c); fl_prem.append(f)
    prem_refs.append(s(r["policy_id"])); prem_notes.append("Top-up; source/method optional - add if known")

# ---------------------------------------------------------------- charges (Vision matched plans)
for i, r in POL.iterrows():
    _puuid_c = s(r["policy_uuid"])
    # Stage 2: skip off-sheet and ambiguous plans
    if _puuid_c not in _matches_by_puuid:
        continue
    _adv_c = _adv_override(_puuid_c) or {}
    # Stage 2: use adviser's col M fee; skip Trail Commission rows
    _adv_fee_str = str(_adv_c.get("adv_fee_pct", "")).strip()
    _legacy = str(_adv_c.get("adv_legacy_label", "")).strip().lower()
    if "trail" in _legacy or "commission" in _legacy:
        continue  # Trail commission row — no charge created
    try:
        _adv_fee = float(_adv_fee_str) if _adv_fee_str else None
    except ValueError:
        _adv_fee = None
    if _adv_fee is not None:
        pct = round(_adv_fee, 4)
    else:
        pct = num(r["ongoing_service_charge_percentage"])
    if pct == "" or float(pct or 0) == 0:
        continue
    c = base(COLS["Charges"]); f, note = {}, []
    # Stage 2: frequency from adviser's col P
    _adv_freq_raw = str(_adv_c.get("adv_frequency", "")).strip()
    _freq_lower = _adv_freq_raw.lower()
    if "quarter" in _freq_lower:
        c["frequency"] = "Quarterly"
    elif "annual" in _freq_lower:
        c["frequency"] = "Annually"
    else:
        c["frequency"] = "Monthly"
    # Payer: Invoice variants -> Client
    if "invoice" in _freq_lower:
        c["payment_origin"] = "Client"
    else:
        pb = s(r["ongoing_service_charge_paid_by"]).lower()
        c["payment_origin"] = PAID.get(pb, "Provider")
        if pb in ("dfm", "platform"):
            f["payment_origin"] = "AMBER"; note.append(f"Paid by {pb} - mapped to Provider, confirm")
        elif not PAID.get(pb):
            note.append(f"Payer '{pb or 'blank'}' not recognised - defaulted to Provider")
    c["type"] = "Fee"
    c["category"] = "ongoing_fee"
    c["recurring"] = "yes"
    c["percentage"] = pct
    c["percentage_frequency"] = "Annually"
    note.append("Charge taken " + c["frequency"].lower() + "; percentage is the annual figure")
    c["tax_exempt"] = "yes"   # Stage 2: all charges tax exempt (fee exc. VAT)
    c["exclude_from_rmar"] = "no"
    c["service_type"] = "Ongoing"
    c["start_date"] = GO_LIVE
    f["start_date"] = "AMBER"; note.append("Charge start date set to go-live - confirm")
    rows_chg.append(c); fl_chg.append(f)
    chg_refs.append((_puuid_c, "; ".join(note)))

# Any plan carrying an ongoing fee -> agency in your name + under advice = yes;
# agency_in_your_name_from defaults to the plan start date where blank.
fee_plan_refs = {ref for ref, _ in chg_refs if ref}
for i, c in enumerate(rows_plans):
    if plan_refs[i] in fee_plan_refs:
        c["agency_in_your_name"] = "yes"
        c["under_advice"] = "yes"
        fl_plans[i]["agency_in_your_name"] = "AMBER"
        fl_plans[i]["under_advice"] = "AMBER"
        if not c["agency_in_your_name_from"]:
            c["agency_in_your_name_from"] = c["opened_at"]
            fl_plans[i]["agency_in_your_name_from"] = "AMBER"
        plan_flags[i] = (plan_flags[i] + "; " if plan_flags[i] else "") + "Ongoing fee present: agency in your name and under advice set to yes"

rows_ie, fl_ie, ie_refs, ie_notes = [], [], [], []
IE = eg["24 - Income and Expenditure Ent"]
exp_rows = []
for i, r in IE.iterrows():
    kind = s(r["income_or_expenditure"]).lower()
    nm = s(r["entry_name"])
    amt = num(r["amount"])
    refs = [s(r["client_1_uuid"])] + ([s(r["client_2_uuid"])] if s(r["client_2_uuid"]) else [])
    if kind == "expenditure":
        exp_rows.append((refs, nm, amt, s(r["other_text"]), i))
        continue
    c = base(COLS["Income & Employment"]); f, note = {}, []
    c["name"] = nm
    t, other = income_type(nm)
    c["type"] = t
    c["type_other"] = other
    c["is_taxable"] = "no" if other.strip().lower() == "other income - non-taxable" else "yes"
    # Fix E (D6): divide annual source figure by 12, set frequency Monthly
    if amt != "" and str(amt) not in ("0",):
        try:
            c["amount"] = round(float(str(amt)) / 12, 2)
        except (ValueError, TypeError):
            c["amount"] = amt
    else:
        c["amount"] = amt
    if c["amount"] == "" or c["amount"] == 0:
        f["amount"] = "RED"
    c["amount_type"] = "Gross"
    f["amount_type"] = "AMBER"
    c["frequency"] = "Monthly"   # Fix E: D6 — source annual ÷12
    f["frequency"] = "AMBER"
    if len(refs) > 1:
        note.append("Joint income - confirm whose record")
    rows_ie.append(c); fl_ie.append(f)
    ie_refs.append(" | ".join(refs))
    ie_notes.append("Amount divided by 12 from annual source figure; frequency set Monthly per D6")

EMP = eg["14 - Employment History"]
emp_seen = set()
emp_history_notes = []   # (client_uuid, text) -> appended to Notes
EMPMAP = {"employed": "Employed", "self-employed": "Self Employed",
          "unemployed": "Unemployed", "retired": "Retired", "semi-retired": "Retired"}
for i, r in EMP.iterrows():
    eu = s(r["employment_entry_uuid"])
    key = (eu, s(r["client_uuid"]))
    if key in emp_seen:
        continue
    emp_seen.add(key)
    js = s(r["job_status"]).lower()
    biz = s(r["business_name"])
    role = s(r["job_role"])
    cur = s(r["employment_status"]).lower() == "current"
    if cur and js in ("employed", "self-employed"):
        # No salary amount in source — write as note (same as past employment)
        ds_emp = d8(r["date_started"])
        when_emp = f" (from {ds_emp})" if ds_emp else ""
        txt = f"Current employment: {role or 'role n/a'}" + (f" at {biz}" if biz else "") + when_emp
        txt += f"; status {EMPMAP.get(js, js)}"
        emp_history_notes.append((s(r["client_uuid"]), txt))
    else:
        # past employment / non-earning current status -> Notes (no income importer route)
        when = ""
        ds, df_ = d8(r["date_started"]), d8(r["date_finished"])
        if ds or df_:
            when = f" ({ds or '?'} to {df_ or 'present'})"
        status = EMPMAP.get(js, js or "")
        label = "Current" if cur else "Previous"
        txt = f"{label} employment: {role or 'role n/a'}" + (f" at {biz}" if biz else "") + when
        if status:
            txt += f"; status {status}"
        emp_history_notes.append((s(r["client_uuid"]), txt))

# ---------------------------------------------------------------- expenses
rows_exp, fl_exp, exp_refs, exp_notes = [], [], [], []
for refs, nm, amt, other, i in exp_rows:
    if amt == "" or float(amt or 0) == 0:
        continue   # drop zero / blank expenditure lines
    c = base(COLS["Expenses"]); f, note = {}, []
    c["expense_name"] = nm
    c["amount"] = amt
    note.append("Annual figure - make monthly per Guidance")
    if len(refs) > 1:
        f["amount"] = "AMBER"
        note.append("Joint expenditure - split equally across the joint clients unless told otherwise (Guidance)")
    c["life_stage_name"] = "Current Lifestyle"
    sec = exp_section(nm)
    c["life_stage_section_name"] = sec
    if sec:
        f["life_stage_section_name"] = "AMBER"
    else:
        f["life_stage_section_name"] = "RED"
        note.append("Section could not be inferred")
    c["is_custom_expense"] = "yes"   # source categories do not map to Plannr's
    if other:
        note.append(f"Detail: {other[:80]}")
    rows_exp.append(c); fl_exp.append(f)
    exp_refs.append(" | ".join(refs)); exp_notes.append("; ".join(note))

# ---------------------------------------------------------------- notes
rows_nt, fl_nt, nt_refs, nt_src = [], [], [], []
MEET = {"Face to Face Meeting": "meeting", "Video Call": "meeting",
        "Telephone Call": "call", "Email Correspondence": "email", "Desktop Review": "note"}
CL = eg["25 - Client Contact Log"]
for i, r in CL.iterrows():
    c = base(COLS["Notes"]); f = {}
    c["model_type"] = "account"
    c["type"] = MEET.get(s(r["meeting_type"]), "note")
    parts = [s(r["reason_for_contact"])]
    if s(r["time_spent_at_meeting"]):
        parts.append(f"Time spent: {s(r['time_spent_at_meeting'])}")
    if s(r["joint_with_spouse"]).lower() in ("yes", "1", "1.0"):
        parts.append("Joint with spouse")
    c["contents"] = " | ".join(x for x in parts if x) or f"{s(r['meeting_type'])} logged"
    c["is_pinned"] = "No"
    c["visible_to_clients"] = "No"
    c["created_at"] = d8(r["meeting_date_time"])
    rows_nt.append(c); fl_nt.append(f)
    nt_refs.append(s(r["contact_uuid"])); nt_src.append(f"25 - Contact Log r{i+2}")

CG = eg["26 - Circumstances and Goals"]
for i, r in CG.iterrows():
    c = base(COLS["Notes"]); f = {}
    c["model_type"] = "account"; c["type"] = "note"
    parts = []
    if s(r["current_circumstances"]):
        parts.append(f"Circumstances: {s(r['current_circumstances'])}")
    if s(r["objectives_moving_forward"]):
        parts.append(f"Objectives: {s(r['objectives_moving_forward'])}")
    if s(r["other_notes"]):
        parts.append(s(r["other_notes"]))
    c["contents"] = " | ".join(parts)
    if not c["contents"]:
        continue
    c["is_pinned"] = "No"; c["visible_to_clients"] = "No"
    c["created_at"] = d8(r["date_captured"])
    rows_nt.append(c); fl_nt.append(f)
    nt_refs.append(s(r["client_uuid"])); nt_src.append(f"26 - Circumstances r{i+2}")

NO = eg["22 - Needs and Objectives"]
for i, r in NO.iterrows():
    c = base(COLS["Notes"]); f = {}
    areas = []
    for col, label in [("preretirement_planning", "Pre-retirement"), ("protection", "Protection"),
                       ("estate_planning", "Estate planning"), ("savings_investments", "Savings/Investments"),
                       ("at_retirement", "At retirement"), ("other", "Other")]:
        if s(r[col]).lower() in ("yes", "1", "1.0", "true"):
            areas.append(label)
    parts = []
    if areas:
        parts.append("Needs identified: " + ", ".join(areas))
    for col in ("reason_notes", "addressed_later_notes", "declined_notes", "other_notes"):
        if s(r[col]):
            parts.append(s(r[col]))
    if not parts:
        continue
    c["model_type"] = "account"; c["type"] = "note"
    c["contents"] = "Needs & objectives - " + " | ".join(parts)
    c["is_pinned"] = "No"; c["visible_to_clients"] = "No"
    c["created_at"] = d8(r["date_captured"])
    rows_nt.append(c); fl_nt.append(f)
    nt_refs.append(s(r["client_1_uuid"])); nt_src.append(f"22 - Needs&Obj r{i+2}")

RV = eg["07 - Reviews"]
for i, r in RV.iterrows():
    c = base(COLS["Notes"]); f = {}
    c["model_type"] = "account"; c["type"] = "note"
    c["contents"] = (f"Review ({s(r['type_of_review'])}) - {s(r['outcome_of_review'])}"
                     + (f" | Policies reviewed: {s(r['reviewed_policy_uuids'])}" if s(r["reviewed_policy_uuids"]) else ""))
    c["is_pinned"] = "No"; c["visible_to_clients"] = "No"
    c["created_at"] = d8(r["date_of_reivew"])
    rows_nt.append(c); fl_nt.append(f)
    nt_refs.append(s(r["client_1_uuid"])); nt_src.append(f"07 - Reviews r{i+2}")

# employment history (records not carried as Salary income) -> notes
for cuid, txt in emp_history_notes:
    c = base(COLS["Notes"]); f = {}
    c["model_type"] = "account"; c["type"] = "note"
    c["contents"] = txt
    c["is_pinned"] = "No"; c["visible_to_clients"] = "No"
    rows_nt.append(c); fl_nt.append(f)
    nt_refs.append(cuid); nt_src.append("14 - Employment History")

# ---------------------------------------------------------------- workbook (dynamic)
RED = PatternFill("solid", start_color="F4CCCC", end_color="F4CCCC")
AMBER = PatternFill("solid", start_color="FCE5CD", end_color="FCE5CD")
GREY = PatternFill("solid", start_color="EFEFEF", end_color="EFEFEF")
GREYH = PatternFill("solid", start_color="D9D9D9", end_color="D9D9D9")
BLUE = PatternFill("solid", start_color="DEEBF7", end_color="DEEBF7")
BLUEH = PatternFill("solid", start_color="9DC3E6", end_color="9DC3E6")
GREEN = PatternFill("solid", start_color="D6E9CE", end_color="D6E9CE")
NAVY = PatternFill("solid", start_color="020849", end_color="020849")
HEADFONT = Font(bold=True, color="FFFFFF", name="Arial", size=10)
BODY = Font(name="Arial", size=10)
THIN = Border(bottom=Side(style="thin", color="CCCCCC"))
TITLEF = Font(bold=True, size=14, name="Arial", color="020849")
H2F = Font(bold=True, size=11, name="Arial", color="020849")
SMALL = Font(name="Arial", size=9, color="555555")

wb = Workbook()
wb.remove(wb.active)

# review cadence per client ref (person sheet only; blank for business/trust)
CADENCE_BY_REF = {}
for _, r in P.iterrows():
    CADENCE_BY_REF[s(r["client_uuid"])] = REVIEW_CADENCE.get(s(r["classification"]), "")

# ---- requirement rules per tab (drive live red highlight + counts) -------------
# kinds: plain | type_indiv | type_entity | if_value(other,vals) | marker(name)
REQ = {
    "Clients": [("client_regulatory_classification", "plain"),
                ("first_name", "type_indiv"), ("last_name", "type_indiv"),
                ("entity_name", "type_entity"), ("owners", "type_entity")],
    "Plans": [("type", "plain"), ("sub_type", "plain"), ("status", "plain"),
              ("policy_number", "contains_sep")],
    "Valuations": [("value", "plain")],
    "Premiums": [],
    "Charges": [],
    "Income & Employment": [("name", "plain"), ("type", "plain"), ("amount", "plain"),
                            ("amount_type", "plain"), ("frequency", "plain"),
                            ("type_other", "if_value", ("type", ("Other",))),
                            ("employment_job_title", "if_value", ("type", ("Salary", "Dividend", "Bonus", "Benefit"))),
                            ("employment_status", "if_value", ("type", ("Salary", "Dividend", "Bonus", "Benefit")))],
    "Expenses": [("expense_name", "plain"), ("amount", "plain"),
                 ("life_stage_section_name", "plain")],
    "Notes": [],
}

DROPDOWNS = {
    "Clients": {
        "client_regulatory_classification": ",".join(ref_list("RegulatoryClassification", "Regulatory Classification",
            ["Retail Client", "Professional Client", "Eligible Counterparty", "Not Applicable", "Unknown"])),
        "gender": ",".join(ref_list("Gender", "gender",
            ["Female", "Male", "Non-binary", "Transgender", "Intersex", "Other", "Unknown"])),
        "marital_status": ",".join(ref_list("MaritalStatus", "marital_status",
            ["Single", "Cohabiting", "Married", "Divorced", "Civil Partnership", "Dissolved Civil Partnership",
             "Widowed", "Separated", "Engaged", "In Relationship", "Unknown"])),
        "smoker_status": ",".join(ref_list("SmokerStatus", "smoker_status",
            ["non-smoker", "smoker", "vaper", "ex-smoker"])),
        "current_state_of_health": ",".join(ref_list("StateOfHealth", "current_state_of_health",
            ["excellent", "good", "fair", "poor"])),
    },
    "Plans": {
        "status": ",".join(ref_list("PlanStatus", "status", ["Active", "Proposed", "Inactive"])),
        "type": ",".join(ref_list("PlanTypeSubType", "ParentTypeValue",
            ["cash_account", "investment", "protection_policy", "pension", "annuity", "general_insurance",
             "group_protection", "mortgage", "equity_release", "loan_credit", "fixed_asset"])),
    },
    "Valuations": {},
    "Premiums": {
        "source": "Employee,Employer",
        "payment_method": "Salary Sacrifice,Salary Deduction,Direct Debit,Standing Order,Cheque,Bank Transfer",
        "frequency": ",".join(ref_list("PremiumFrequency", "Frequency",
            ["Daily", "Weekly", "Fortnightly", "Four Weekly", "Monthly", "Bimonthly", "Quarterly",
             "Biannually", "Annually", "Yearly", "Adhoc", "One Off"])),
    },
    "Charges": {
        "payment_origin": "Client,Provider",
        "frequency": ",".join(ref_list("ChargeFrequency", "Frequency",
            ["Daily", "Weekly", "Fortnightly", "Four Weekly", "Monthly", "Bimonthly", "Quarterly",
             "Biannually", "Annually", "Yearly"])),
    },
    "Income & Employment": {
        "amount_type": "Gross,Net",
        "type": "Salary,Dividend,Pension,Bonus,Benefit,Investment income,Rental income,Other",
        "frequency": "Daily,Weekly,Fortnightly,Four Weekly,Monthly,Bimonthly,Quarterly,Biannually,Annually,Adhoc,One off",
        "employment_status": "Employed,Self Employed,Unemployed,Student,Retired,House person",
    },
    "Expenses": {"life_stage_section_name": "Essential,Lifestyle,Discretionary"},
    "Notes": {},
}


def _blank(cl, r):
    return 'LEN(TRIM(' + cl + str(r) + '&""))=0'


def _pred(rule, colmap, r):
    """Excel boolean: TRUE when this cell is required-but-missing for row r."""
    col, kind = rule[0], rule[1]
    cl = colmap[col]
    b = _blank(cl, r)
    if kind == "plain":
        return b
    if kind == "type_indiv":
        tc = colmap["type"]
        return 'AND($' + tc + str(r) + '="individual",' + b + ')'
    if kind == "type_entity":
        tc = colmap["type"]
        return 'AND(OR($' + tc + str(r) + '="business",$' + tc + str(r) + '="trust"),' + b + ')'
    if kind == "if_value":
        other, vals = rule[2]
        oc = colmap[other]
        ors = "OR(" + ",".join('$' + oc + str(r) + '="' + v + '"' for v in vals) + ")"
        return 'AND(' + ors + ',' + b + ')'
    if kind == "marker":
        mc = colmap[rule[2][0]]
        return 'AND($' + mc + str(r) + '=1,' + b + ')'
    if kind == "contains_sep":
        return 'ISNUMBER(SEARCH("|",' + cl + str(r) + '))'
    return b


def _miss_formula(req, colmap, r):
    if not req:
        return "=0"
    return "=" + "+".join('(' + _pred(rule, colmap, r) + ')*1' for rule in req)


def write_tab(name, cols, rows, flags, helpers, helper_rows, firm,
              req=None, markers=None, dropdowns=None, src_after=None):
    req = req or []
    markers = markers or {}
    dropdowns = dropdowns or {}
    src_after = src_after or {}
    ws = wb.create_sheet(name[:31])
    helper_set = set(helpers)
    mk_names = list(markers.keys())
    # interleave source-reference columns after their mapped field
    layout = []   # (header, role, src_values_or_None)  role in {data, src}
    for col in cols:
        layout.append((col, "data", None))
        if col in src_after:
            spec = src_after[col]
            items = spec if isinstance(spec, list) else [spec]
            for sh, svals in items:
                layout.append((sh, "src", svals))
    headers = [h for h, _, _ in layout]
    allcols = headers + helpers + mk_names + ["_miss"]
    colmap = {c: get_column_letter(j) for j, c in enumerate(allcols, 1)}
    n = len(rows)
    last = n + 1 if n else 1
    SRCH = Font(bold=True, size=9, color="1F4E79")
    SRCB = Font(name="Arial", size=9, color="1F4E79", italic=True)
    for j, col in enumerate(allcols, 1):
        cell = ws.cell(1, j, col)
        cell.alignment = Alignment(vertical="center")
        role = layout[j - 1][1] if j - 1 < len(layout) else None
        if col == "_miss" or col in mk_names:
            cell.font = Font(bold=True, size=8, color="AAAAAA", name="Arial"); cell.fill = GREY
        elif role == "src":
            cell.font = SRCH; cell.fill = BLUE
        elif col.startswith("x_") or col in helper_set:
            cell.font = HEADFONT; cell.fill = BLUEH
        elif col in firm:
            cell.font = Font(bold=True, color="333333", name="Arial", size=10); cell.fill = GREYH
        else:
            cell.font = HEADFONT; cell.fill = NAVY
    amber_count = 0
    for i, (row, fl) in enumerate(zip(rows, flags), 2):
        for j, (header, role, svals) in enumerate(layout, 1):
            if role == "data":
                v = row.get(header, "")
                cell = ws.cell(i, j, v if v != "" else None)
                cell.font = BODY; cell.border = THIN
                if fl.get(header) == "AMBER":
                    cell.fill = AMBER; amber_count += 1
                elif header in firm:
                    cell.fill = GREY
            else:  # src reference
                sv = svals[i - 2] if i - 2 < len(svals) else ""
                cell = ws.cell(i, j, sv if sv else None)
                cell.font = SRCB; cell.fill = BLUE; cell.border = THIN
        base_j = len(layout)
        for k, hv in enumerate(helper_rows[i - 2], 1):
            cell = ws.cell(i, base_j + k, hv if hv else None)
            cell.font = Font(name="Arial", size=9, color="1F4E79"); cell.fill = BLUE; cell.border = THIN
        for k, mk in enumerate(mk_names, 1):
            ws.cell(i, base_j + len(helpers) + k, markers[mk][i - 2])
        ws.cell(i, len(allcols), value=_miss_formula(req, colmap, i))
    if n:
        for rule in req:
            if rule[0] not in colmap:
                continue
            cl = colmap[rule[0]]
            ws.conditional_formatting.add(cl + "2:" + cl + str(last),
                                          FormulaRule(formula=[_pred(rule, colmap, 2)], fill=RED))
        for col, opts in dropdowns.items():
            if col not in colmap:
                continue
            dv = DataValidation(type="list", formula1='"' + opts + '"', allow_blank=True)
            ws.add_data_validation(dv)
            dv.add(colmap[col] + "2:" + colmap[col] + str(last))
    ws.freeze_panes = "A2"
    vis = len(layout) + len(helpers)
    ws.auto_filter.ref = "A1:" + get_column_letter(vis) + str(max(last, 1))
    for j, col in enumerate(allcols, 1):
        ws.column_dimensions[get_column_letter(j)].width = max(12, min(40, len(str(col)) + 4))
    for k in range(vis + 1, len(allcols) + 1):
        ws.column_dimensions[get_column_letter(k)].hidden = True
    return {"name": name, "sheet": name[:31], "miss": colmap["_miss"],
            "last": last, "rows": n, "amber": amber_count}

# ---------------------------------------------------------------- linkage lookups (for blue helper columns)
def _client_display(c):
    return (s(c.get("entity_name")) or
            " ".join(x for x in [s(c.get("first_name")), s(c.get("middle_name")), s(c.get("last_name"))] if x)).strip()

CLIENT_NAME = {}
for _i in range(len(rows_clients)):
    _ref = client_refs[_i]
    if _ref and _ref not in CLIENT_NAME:
        CLIENT_NAME[_ref] = _client_display(rows_clients[_i])

CLIENT_LASTNAME = {}
for _i in range(len(rows_clients)):
    _ref = client_refs[_i]
    if _ref and rows_clients[_i].get("type") == "individual":
        _ln = s(rows_clients[_i].get("last_name", ""))
        if _ln and _ref not in CLIENT_LASTNAME:
            CLIENT_LASTNAME[_ref] = _ln

def names_for(refs_str, sep="|"):
    """Map a separated list of source client UUIDs to their display names."""
    out = []
    for r in str(refs_str).split(sep):
        r = r.strip()
        if r:
            out.append(CLIENT_NAME.get(r, r))
    return " & ".join(out)

PLAN_INFO = {}
for _i in range(len(rows_plans)):
    _owners = [o.strip() for o in str(rows_plans[_i].get("client_uuids", "")).split(",") if o.strip()]
    _onames = " & ".join(CLIENT_NAME.get(o, "") for o in _owners if CLIENT_NAME.get(o, ""))
    if not _onames:
        _onames = plan_owner_names[_i]
    PLAN_INFO[plan_refs[_i]] = {
        "name": rows_plans[_i].get("name", ""),
        "type": rows_plans[_i].get("type", ""),
        "sub_type": rows_plans[_i].get("sub_type", ""),
        "provider": rows_plans[_i].get("provider", ""),
        "policy_number": rows_plans[_i].get("policy_number", ""),
        "owner_names": _onames,
        "owner_refs": ", ".join(_owners),
    }

# Valuations link to a plan by policy number (column C). Pull it from the plan it belongs to.
for _i in range(len(rows_val)):
    _pi = PLAN_INFO.get(val_refs[_i][0], {})
    if _pi.get("type") == "fixed_asset":
        continue   # property valuations link by plan, not a policy number
    _pn = _pi.get("policy_number", "")
    rows_val[_i]["policy_number"] = _pn
    if (not _pn) or _pn == "unknown" or "|" in str(_pn):
        fl_val[_i]["policy_number"] = "AMBER"


def plan_idy(refs):
    """src_after spec: plan + owner identity columns for a list of source plan refs."""
    g = lambda k: [PLAN_INFO.get(r, {}).get(k, "") for r in refs]
    return [("x_plan_ref", list(refs)), ("x_plan_name", g("name")), ("x_plan_type", g("type")),
            ("x_plan_sub_type", g("sub_type")), ("x_provider", g("provider")),
            ("x_client_names", g("owner_names")), ("x_client_refs", g("owner_refs"))]


def client_idy(refs):
    """src_after spec: client name + source UUID for a list of (possibly pipe-joined) refs."""
    return [("x_client_ref", list(refs)), ("x_client_name", [names_for(r) for r in refs])]


def plan_cols(plan_ref):
    """Return [name, type, sub_type, provider, owner_names, owner_refs] for a source plan ref."""
    pi = PLAN_INFO.get(plan_ref, {})
    return [pi.get("name", ""), pi.get("type", ""), pi.get("sub_type", ""),
            pi.get("provider", ""), pi.get("owner_names", ""), pi.get("owner_refs", "")]

# ---------------------------------------------------------------- static text blocks
MAPPING = [
    ("Clients", "type", "01/02/03 sheet", "individual / business / trust by source sheet"),
    ("Clients", "first_name / middle_name", "01.forename", "split on first space"),
    ("Clients", "last_name / entity_name", "01.surname / 02.business_name / 03.client_name", ""),
    ("Clients", "nickname", "01.preferred_name", "known-as column"),
    ("Clients", "gender", "01.sex", "direct (confirm)"),
    ("Clients", "ni_number", "01.national_insurance_number", "spaces stripped, uppercased; non-standard flagged amber"),
    ("Clients", "marital_status", "01.marital_status", "carried as-is (all values valid in Plannr)"),
    ("Clients", "smoker_status", "01.smokes_tobacco + vapes", "1=smoker; vapes=vaper; 0=non-smoker"),
    ("Clients", "current_state_of_health", "01.in_good_health", "1=good; 0 flagged amber (set fair/poor)"),
    ("Clients", "phone", "01 home>work landline", "mobile is NOT importable -> x_mobile_via_api (API load)"),
    ("Clients", "owners", "02.primary_contact_uuid (business) / (none for trust)", "individual owner; business pre-filled (source UUID), trust red - supply"),
    ("Clients", "has_power_of_attorney / power_of_attorney_notes", "15 POA", "col C (scope) + ', ' + col E (attorneys); has_power_of_attorney set yes"),
    ("Clients", "client_regulatory_classification", "NOT in export", "AR to supply (red). NB 01.classification is review cadence, see x_review_cadence"),
    ("Clients", "x_review_cadence", "01.classification", "1=6mo, 2=12mo, 3=24mo(legacy), 4=none (reference only)"),
    ("Clients", "externalreference", "01.client_uuid", "source UUID -> Plannr custom field 'externalreference' (model_type client)"),
    ("Clients", "address_*", "12 Address History (current)", "latest current address"),
    ("Clients", "has_will / will_*", "16 Wills (Current)", "executors, last reviewed; set-up date & beneficiaries to will_notes"),
    ("Clients", "is_vulnerable / vulnerability_*", "11 Vulnerability (latest)", "inferred from indicators (confirm)"),
    ("Clients", "checked_sanctions_at", "19 AML (latest date_performed)", "AML date used (confirm)"),
    ("Clients", "previous_review_date / next_review_date", "07 Reviews (latest)", "next = previous + 1 year (amber). Tags not used."),
    ("Plans", "client_uuids", "04 client_1/2_uuid; 18 owner_uuids", "SOURCE refs (remapped). 'owner name(s) [source]' shows who owns it"),
    ("Plans", "client_ownership_percentage", "derived", "one owner = 100; two or more = equal split (e.g. 50,50) - confirm"),
    ("Plans", "type / sub_type", "04.product", "mapped per product table below; 'product [source]' shows original; sub_type checked against importer list"),
    ("Plans", "policy_number", "04.policy_numbers / 18.mortgage_reference", "missing -> 'unknown' (amber), except fixed assets; multiple numbers flagged red"),
    ("Plans", "name", "04.product_specific_name else product", ""),
    ("Plans", "provider / provider_uuid", "04.primary_provider", "matched to Plannr provider directory; exact match sets provider_uuid, else amber + suggestions"),
    ("Plans", "opened_at", "04.date_of_sale > date_of_transfer > 01-01-2026", "all amber - AR to approve (source in x_flags)"),
    ("Plans", "agency_in_your_name / under_advice", "04 + ongoing fee", "set yes where an ongoing fee exists; agency_from = plan start date"),
    ("Plans", "DB pensions", "05", "pension / Defined Benefit Pension; full pension -> projected income; tax_free_cash -> tax-free amount"),
    ("Plans", "fixed assets / mortgages", "18 Property", "one fixed_asset per property; mortgage row where has_current_mortgage=Yes"),
    ("Plans", "externalreference", "04.policy_uuid / 05 / 18 uuid", "source UUID -> Plannr custom field 'externalreference' (model_type plan)"),
    ("Valuations", "value", "18.current_value + 04.contribution", "active property values + policy contribution (confirm contribution is current value)"),
    ("Valuations", "valuation_date", "go-live 26-06-2026", "defaulted (amber) - confirm"),
    ("Charges", "rows", "04 ongoing_service_charge_*", "Fee / ongoing_fee / Ongoing; monthly unless quarterly; percentage is the annual figure; start_date = go-live"),
    ("Income & Employment", "income rows", "21 Income & Expenditure (income)", "type inferred; Monthly; amount_type Gross/Net red; is_taxable yes unless 'Other Income - Non-Taxable'"),
    ("Income & Employment", "salary rows", "14 employment (current employed/self-employed)", "type Salary; amount + gross/net red; job title from col F, 'unknown' (amber) where blank"),
    ("Income & Employment", "(other employment)", "14 employment (past / non-earning)", "no income route -> written to Notes"),
    ("Expenses", "rows", "21 Income & Expenditure (expenditure)", "is_custom_expense yes; zero lines dropped; amounts over 1500 flagged (could be annual); monthly"),
    ("Notes", "rows", "23 contact log, 24 circumstances, 21 needs, 07 reviews, 14 employment history", "model_type account; created_at from source dates"),
]

GUIDE = [
    ("Clients",
     ["client_regulatory_classification - required for every client. Values: Retail Client, Professional Client, Eligible Counterparty, Not Applicable, Unknown.",
      "  (The export 'classification' is a review-cadence code, shown for context in x_review_cadence - it is NOT the regulatory classification.)",
      "first_name / last_name - required for individuals (already filled from the export; red only if missing).",
      "entity_name - required for Business and Trust clients (already filled).",
      "owners - Business and Trust clients need an individual client as owner. Business: pre-filled from the source primary contact (a source UUID, remapped at import). Trust: blank and red - supply the individual client UUID."],
     ["gender mapped from the export 'sex' field - confirm.",
      "marital status carried across as-is (all values are valid in Plannr) - confirm.",
      "smoker status and state of health derived from the source yes/no flags - confirm; set state of health to fair/poor where 'not in good health'.",
      "vulnerability flag and sub-type inferred from indicators - confirm.",
      "power of attorney pulled from sheet 15 (scope + attorneys merged into power_of_attorney_notes) - confirm.",
      "next_review_date set to one year after the previous review date - confirm. Tags are not carried (not used).",
      "AML date used as the sanctions-check date - confirm.",
      "NI numbers not in the standard format are highlighted - check.",
      "phone holds the landline only. Mobile numbers cannot be loaded by the importer - they are in x_mobile_via_api and load by API.",
      "externalreference holds the source client UUID and imports into the Plannr 'externalreference' custom field. Leave it as-is."]),
    ("Plans",
     ["type / sub_type - blank where the source product had no Plannr mapping (e.g. 'Fee (no product)', 'CAP ADJUSTMENT'). Map to a type/sub_type or remove the row.",
      "  Plan types: cash_account, investment, protection_policy, pension, annuity, general_insurance, group_protection, mortgage, equity_release, loan_credit, fixed_asset. Sub-types: see the Plans Dropdowns sheet in the Plannr template.",
      "  Mapped sub-types are checked against the Plannr importer list; any not recognised for their type are flagged amber.",
      "policy_number - where the source held more than one number (e.g. 'A | B | C') the cell is red. Valuation feeds need ONE exact number - delete the others, leaving the matching policy number.",
      "  Missing numbers default to 'unknown' (amber) for everything except fixed assets - replace with the real number. Fixed assets (property) need no policy number."],
     ["status defaulted to Active - confirm.",
      "client_uuids / ownership: owners are the SOURCE CRM UUIDs (remapped at import). 'owner name(s) [source]' shows who the plan belongs to. One owner = 100%; two or more = an equal split (e.g. 50,50) - confirm.",
      "opened_at = date_of_sale, else date_of_transfer, else 01-01-2026 - approve the start date (source shown in x_flags).",
      "agency_in_your_name and under_advice are set to yes for any plan carrying an ongoing fee; agency_in_your_name_from defaults to the plan start date - confirm.",
      "provider matched to the Plannr directory: an exact match fills provider_uuid; otherwise the name is amber with the closest directory entries suggested in x_flags - pick the right one.",
      "DB pensions: scheme name used as provider - confirm administrator; reduced pension is noted in x_flags only (no Plannr field).",
      "externalreference holds the source plan UUID and imports straight into the Plannr 'externalreference' custom field. Leave it as-is."]),
    ("Valuations",
     ["value - required. Property current value and (where present) the policy 'contribution' figure are carried across; red only if missing."],
     ["valuation_date defaulted to the go-live date (26-06-2026) - confirm.",
      "Rows built from policy 'contribution' are flagged: confirm this is a current value, not total contributions paid. See x_note.",
      "The blue columns identify the plan each valuation belongs to (source plan UUID, name, type, sub_type, provider) and its owner client(s) by name and source UUID."]),
    ("Charges",
     ["Nothing is required. fee, percentage and frequency are filled."],
     ["payer mapped: client -> Client; provider / platform / DFM -> Provider - confirm.",
      "frequency assumed Monthly unless the source said quarterly; the percentage is the annual figure.",
      "VAT mapped to tax_exempt - confirm. start_date set to the go-live date (26-06-2026) - confirm.",
      "The blue columns identify the plan each charge belongs to (source plan UUID, name, type, sub_type, provider) and its owner client(s) by name and source UUID."]),
    ("Income & Employment",
     ["IMPORTANT - every income amount is an ANNUAL figure from the source. The script does NOT alter amounts; convert each to a MONTHLY figure before import and set frequency to Monthly.",
      "amount_type - defaulted to Gross (amber). Change any that are Net.",
      "amount - enter where blank (e.g. current-employment salary).",
      "name / type - already filled. frequency is set to Annually to match the source; change to Monthly once the amounts are monthly.",
      "employment_status - required when income type is Salary, Dividend, Bonus or Benefit. If a 'Dividend' line is investment income, change its type to 'Investment income' to clear it.",
      "type_other - required only when type = Other."],
     ["is_taxable assumed yes for all income, except where type_other is 'Other Income - Non-Taxable' (set no) - confirm.",
      "employment_job_title taken from sheet 14 column F for current employment; set to 'unknown' (amber) where the source had none.",
      "Past employment and non-earning current roles are not income lines, so they are written to Notes instead of here.",
      "The blue columns show which client each line belongs to, by name and source UUID."]),
    ("Expenses",
     ["IMPORTANT - every expenditure amount is an ANNUAL figure from the source. The script does NOT alter amounts; convert each to MONTHLY before import.",
      "IMPORTANT - any JOINT expenditure (a line with two owners, flagged amber) will be split EQUALLY across the joint clients unless you tell us otherwise.",
      "life_stage_section_name - blank where the section could not be inferred from the name. Set Essential, Lifestyle or Discretionary.",
      "amount - enter where blank."],
     ["is_custom_expense is set to yes (the source keeps its own expense names rather than Plannr's standard categories).",
      "zero-value lines are dropped.",
      "each expense keeps its original name and is mapped to one of your three Plannr sections (Essential / Lifestyle / Discretionary) where possible; where it genuinely cannot be placed the section is blank and red. Life stage set to 'Current Lifestyle'.",
      "The blue columns show which client each expense belongs to, by name and source UUID."]),
    ("Notes",
     ["Nothing to complete."],
     ["Built from contact log, circumstances & goals, needs & objectives, reviews, and employment history.",
      "The blue columns show which client each note belongs to, by name and source UUID."]),
]

CONFIRMS = [
    "Plan start dates approved (date_of_sale / date_of_transfer / 01-01-2026 fallback)",
    "Plan ownership splits confirmed (one owner = 100%, joint = equal split)",
    "Multiple-policy-number plans reduced to the single feed-matching number",
    "Missing policy numbers ('unknown') replaced with the real number",
    "Provider matches confirmed; unmatched names resolved",
    "Ongoing-fee plans: agency in your name / under advice = yes is correct",
    "Valuations confirmed (contribution figures are current values, date = go-live)",
    "Income Gross/Net set; is_taxable assumptions confirmed",
    "Expenses: section set; figures over 1500 confirmed as monthly",
    "Business/Trust owners set to an individual client",
    "Power of attorney and next review date (previous + 1 year) confirmed",
    "Charge / valuation go-live date (26-06-2026) confirmed; mobiles to load via API",
]

QUESTIONS = [
    "DB Pension primary owner: 05.client_1_uuid does not match the client master (client_2_uuid does). Confirm the correct primary owner.",
    "Orphan owner references: some plan owner UUIDs have no client record in this export (see each plan's x_flags). Confirm the missing clients.",
    "Multiple policy numbers: many plans carry several numbers in one field (A | B | C). Supply the single number each valuation feed must match.",
    "Policy 'contribution': used as a current valuation here. Confirm it is a current value and not total contributions paid.",
    "Plan start dates: not supplied; derived from date_of_sale / date_of_transfer with 01-01-2026 fallback. Confirm or supply true start dates.",
    "Property addresses & mortgage references: sheet 18 address, postcode and mortgage_reference are empty for every row. Supply if held elsewhere.",
    "Trust owners: trusts carry no individual owner in the source. Supply the individual client for each trust.",
    "Provider matching: exact names matched to the Plannr directory; unmatched names are amber with the closest entries suggested - pick the right one.",
    "Premiums are not in this pack (no source/method/start dates). The AR keys recurring premiums manually after import.",
]


def _pick(lst, idxs):
    return [lst[i] for i in idxs]


def build_book(out_path, title, keep):
    global wb, metas, META
    wb = Workbook()
    wb.remove(wb.active)
    metas = []
    ci = keep["Clients"]
    metas.append(write_tab("Clients", COLS["Clients"], _pick(rows_clients, ci), _pick(fl_clients, ci),
                           ["externalreference", "x_review_cadence", "x_mobile_via_api", "x_flags"],
                           [[client_refs[i], CADENCE_BY_REF.get(client_refs[i], ""),
                             mobile_by_ref.get(client_refs[i], ""), flags_text[i]] for i in ci],
                           FIRM["Clients"], REQ["Clients"], {}, DROPDOWNS["Clients"]))
    pi = keep["Plans"]
    metas.append(write_tab("Plans", COLS["Plans"], _pick(rows_plans, pi), _pick(fl_plans, pi),
                           ["externalreference", "x_source", "x_move_to_business", "x_flags"],
                           [[plan_refs[i], plan_src[i], plan_move_biz[i], plan_flags[i]] for i in pi],
                           FIRM["Plans"], REQ["Plans"], {}, DROPDOWNS["Plans"],
                           src_after={"client_uuids": ("owner name(s) [source]", _pick(plan_owner_names, pi)),
                                      "sub_type": ("product [source]", _pick(plan_product_src, pi))}))
    vi = keep["Valuations"]
    metas.append(write_tab("Valuations", COLS["Valuations"], _pick(rows_val, vi), _pick(fl_val, vi),
                           ["x_plan_ref", "x_plan_name", "x_plan_type", "x_plan_sub_type", "x_provider",
                            "x_client_names", "x_client_refs", "x_note"],
                           [[val_refs[i][0]] + plan_cols(val_refs[i][0]) + [val_notes[i]] for i in vi],
                           FIRM["Valuations"], REQ["Valuations"], {}, DROPDOWNS["Valuations"]))
    gi = keep["Charges"]
    metas.append(write_tab("Charges", COLS["Charges"], _pick(rows_chg, gi), _pick(fl_chg, gi),
                           ["x_flags"],
                           [[chg_refs[i][1]] for i in gi],
                           FIRM["Charges"], REQ["Charges"], {}, DROPDOWNS["Charges"],
                           src_after={"adviser_uuid": plan_idy([chg_refs[i][0] for i in gi])}))
    ii = keep["Income & Employment"]
    metas.append(write_tab("Income & Employment", COLS["Income & Employment"], _pick(rows_ie, ii), _pick(fl_ie, ii),
                           ["x_note"],
                           [[ie_notes[i]] for i in ii],
                           FIRM["Income & Employment"], REQ["Income & Employment"], {}, DROPDOWNS["Income & Employment"],
                           src_after={"client_uuid": client_idy([ie_refs[i] for i in ii])}))
    xi = keep["Expenses"]
    metas.append(write_tab("Expenses", COLS["Expenses"], _pick(rows_exp, xi), _pick(fl_exp, xi),
                           ["x_note"],
                           [[exp_notes[i]] for i in xi],
                           FIRM["Expenses"], REQ["Expenses"], {}, DROPDOWNS["Expenses"],
                           src_after={"client_uuid": client_idy([exp_refs[i] for i in xi])}))
    ni = keep["Notes"]
    metas.append(write_tab("Notes", COLS["Notes"], _pick(rows_nt, ni), _pick(fl_nt, ni),
                           ["x_client_ref", "x_client_name", "x_source"],
                           [[nt_refs[i], names_for(nt_refs[i]), nt_src[i]] for i in ni],
                           FIRM["Notes"], REQ["Notes"], {}, DROPDOWNS["Notes"]))
    META = {m["name"]: m for m in metas}

    # field mapping
    mp = wb.create_sheet("Field Mapping")
    mp.sheet_view.showGridLines = False
    mp.append(["Tab", "Plannr column", "Source", "Transform / default"])
    for c in mp[1]:
        c.font = HEADFONT
        c.fill = NAVY
    for row in MAPPING:
        mp.append(row)
    mp.append([])
    mp.append(["Product mapping (04.product -> Plannr)", "", "", ""])
    mp.cell(row=mp.max_row, column=1).font = Font(bold=True, name="Arial")
    mp.append(["Ceding product", "Plannr type", "Plannr sub_type", "Note"])
    for c in mp[mp.max_row]:
        c.font = Font(bold=True, name="Arial", size=10)
    for p in sorted(PM):
        t, sub, n2 = PM[p]
        mp.append([p, t, sub, n2 or ""])
    for p in sorted(NON_PLANS):
        mp.append([p, "(blank)", "(blank)", "Not a plan - map or remove the row (red in Plans)"])
    for col, w in zip("ABCD", (46, 26, 42, 70)):
        mp.column_dimensions[col].width = w
    for rowcells in mp.iter_rows(min_row=2):
        for c in rowcells:
            if c.font is None or not c.font.bold:
                c.font = Font(name="Arial", size=10)

    # guide
    gd = wb.create_sheet("Guide", 0)
    gd.sheet_view.showGridLines = False
    gr = 1
    gd.cell(gr, 1, "HOW TO COMPLETE EACH SHEET").font = TITLEF
    gr += 2
    for tab, reds, ambers in GUIDE:
        if tab not in META:
            continue
        m = META.get(tab, {})
        gd.cell(gr, 1, tab + "  (" + str(m.get("rows", 0)) + " rows)").font = H2F
        gr += 1
        gd.cell(gr, 1, "RED - must complete:").font = Font(bold=True, name="Arial", size=10, color="9C0006")
        gr += 1
        for ln in reds:
            gd.cell(gr, 1, ln).font = BODY
            gr += 1
        gd.cell(gr, 1, "AMBER - confirm or correct:").font = Font(bold=True, name="Arial", size=10, color="9C6500")
        gr += 1
        for ln in ambers:
            gd.cell(gr, 1, ln).font = BODY
            gr += 1
        gr += 1
    gd.column_dimensions["A"].width = 140

    # dashboard
    dash = wb.create_sheet("START HERE", 0)
    dash.sheet_view.showGridLines = False
    R = 1
    dash.cell(R, 1, title).font = TITLEF
    R += 1
    dash.cell(R, 1, "Work through the coloured cells on each tab. The counts below update live as you fill them in.").font = SMALL
    R += 2
    dash.cell(R, 1, "KEY").font = H2F
    R += 1
    for label, fill in [("RED - required for import, missing. Please complete.", RED),
                        ("AMBER - we assumed or derived this, or it is optional. Confirm or correct.", AMBER),
                        ("GREY - Beyond populates (system, adviser, ownership links). Please ignore.", GREY),
                        ("BLUE columns - read-only context for you: source client/plan UUIDs, client and plan names, provider and plan type, and the externalreference migration key. Do not edit or delete.", BLUE)]:
        sw = dash.cell(R, 1, " ")
        sw.fill = fill
        dash.cell(R, 2, label).font = BODY
        R += 1
    R += 1
    dash.cell(R, 1, "PROGRESS (live)").font = H2F
    R += 1
    for j, h in enumerate(["Section", "Fields to complete", "Rows needing attention", "Assumptions to confirm"], 1):
        c = dash.cell(R, j, h)
        c.font = HEADFONT
        c.fill = NAVY
    R += 1
    first_data = R
    for m in metas:
        dash.cell(R, 1, m["name"]).font = BODY
        if m["rows"]:
            rng = "'" + m["sheet"] + "'!" + m["miss"] + "2:" + m["miss"] + str(m["last"])
            dash.cell(R, 2, "=SUM(" + rng + ")").font = BODY
            dash.cell(R, 3, '=COUNTIF(' + rng + ',">0")').font = BODY
        else:
            dash.cell(R, 2, 0).font = BODY
            dash.cell(R, 3, 0).font = BODY
        dash.cell(R, 4, m["amber"]).font = BODY
        R += 1
    last_data = R - 1
    dash.cell(R, 1, "TOTAL").font = Font(bold=True, name="Arial", size=10)
    dash.cell(R, 2, "=SUM(B" + str(first_data) + ":B" + str(last_data) + ")").font = Font(bold=True, name="Arial", size=11)
    dash.cell(R, 3, "=SUM(C" + str(first_data) + ":C" + str(last_data) + ")").font = Font(bold=True, name="Arial", size=10)
    dash.cell(R, 4, "=SUM(D" + str(first_data) + ":D" + str(last_data) + ")").font = Font(bold=True, name="Arial", size=10)
    tfc = "B" + str(R)
    dash.conditional_formatting.add(tfc, FormulaRule(formula=[tfc + "=0"], fill=GREEN))
    dash.conditional_formatting.add(tfc, FormulaRule(formula=[tfc + ">0"], fill=AMBER))
    R += 2
    dash.cell(R, 1, "CONFIRMATIONS - tick when reviewed").font = H2F
    R += 1
    tally_row = R
    dash.cell(R, 1, "Confirmed:").font = Font(bold=True, name="Arial", size=10)
    R += 1
    conf_first = R
    dv_yn = DataValidation(type="list", formula1='"Yes,No"', allow_blank=True)
    dash.add_data_validation(dv_yn)
    for item in CONFIRMS:
        dash.cell(R, 1, item).font = BODY
        a = dash.cell(R, 2, "")
        a.font = BODY
        a.fill = AMBER
        dv_yn.add(a)
        R += 1
    conf_last = R - 1
    dash.cell(tally_row, 2,
              '=COUNTIF(B' + str(conf_first) + ':B' + str(conf_last) + ',"Yes")&" of "&' + str(len(CONFIRMS))).font = Font(bold=True, name="Arial", size=11)
    R += 2
    dash.cell(R, 1, "QUESTIONS FOR THE DATA SOURCE (not AR actions)").font = H2F
    R += 1
    for q in QUESTIONS:
        dash.cell(R, 1, "- " + q).font = BODY
        R += 1
    dash.column_dimensions["A"].width = 96
    for col in ("B", "C", "D"):
        dash.column_dimensions[col].width = 22

    wb.calculation.fullCalcOnLoad = True
    wb.save(out_path)
    return {m["name"]: m["rows"] for m in metas}


def _write_extra_sheets(out_path):
    """Append Ambiguous, Not on his list, Possible duplicates, and Added sheets to the workbook."""
    from openpyxl import load_workbook as _lw
    _wb = _lw(out_path)
    # Fix F: Ambiguous rows — awaiting AR decision, excluded from files
    _amb_rows_list = [_mr for _, _mr in _mdf.iterrows()
                      if str(_mr.get("row_status", "")).strip() == "Ambiguous"]
    if _amb_rows_list:
        _ws_amb = _wb.create_sheet("Ambiguous")
        _amb_col_names = ["adv_row", "adv_provider", "adv_policy_number", "adv_type",
                          "adv_value", "adv_fee_pct", "adv_frequency", "adv_legacy_label",
                          "person_match_status", "person_match_method",
                          "plan_match_status", "candidates"]
        _ws_amb.append(_amb_col_names)
        for _amr in _amb_rows_list:
            _ws_amb.append([str(_amr.get(_c, "")) for _c in _amb_col_names])
        print(f"Ambiguous tab: {len(_amb_rows_list)} rows")
    # Not on his list
    if not_on_his_list:
        _ws = _wb.create_sheet("Not on his list")
        _ws.append(["vision_policy_uuid", "plan_name", "provider", "vis_fee_pct"])
        for row in not_on_his_list:
            _ws.append(list(row))
    # Possible duplicates (Vision plans that rule 4 couldn't pair)
    if possible_dup_puuids:
        _ws2 = _wb.create_sheet("Possible duplicates")
        _ws2.append(["vision_policy_uuid", "product", "provider", "vis_fee_pct"])
        for _pd_pu in sorted(possible_dup_puuids):
            _pd_pol = {}
            for _, _pr in POL.iterrows():
                if s(_pr["policy_uuid"]) == _pd_pu:
                    _pd_pol = _pr.to_dict(); break
            _ws2.append([_pd_pu, _pd_pol.get("product", ""), _pd_pol.get("primary_provider", ""),
                         _pd_pol.get("ongoing_service_charge_percentage", "")])
    # Added (new clients)
    if new_client_added_list:
        _ws3 = _wb.create_sheet("Added")
        _ws3.append(["generated_ref", "adv_row", "adv_type", "adv_provider", "adv_policy_number"])
        for _nc in _new_client_rows:
            _nc_row = str(_nc.get("adv_row", "")).strip()
            _ws3.append([
                "nc_" + str(_uuid_mod.uuid5(_uuid_mod.NAMESPACE_DNS, f"higgs_nc_{_nc_row}")),
                _nc_row,
                str(_nc.get("adv_type", "")),
                str(_nc.get("adv_provider", "")),
                str(_nc.get("adv_policy_number", "")),
            ])
    _wb.save(out_path)


# ---------------------------------------------------------------- business plan flagging (Edit I)
# 45 plans owned by business primary contacts: flag x_move_to_business for group_protection
# and any plan whose product contains "key person", "business protection", or "relevant life"
_biz_flag_keywords = ("key person", "business protection", "relevant life")
plan_move_biz = []
for _pi, _pc in enumerate(rows_plans):
    _pc_owners = [o.strip() for o in str(_pc.get("client_uuids", "")).split(",") if o.strip()]
    _is_biz_owned = any(o in biz_primary_uuids for o in _pc_owners)
    if not _is_biz_owned:
        plan_move_biz.append("")
        continue
    _pc_type = _pc.get("type", "")
    _pc_prod = plan_product_src[_pi].lower()
    _move = (_pc_type == "group_protection" or
             any(kw in _pc_prod for kw in _biz_flag_keywords))
    plan_move_biz.append("YES" if _move else "")

_biz_pu_sample = list(biz_primary_uuids)[:3]
_biz_plan_total = sum(1 for _bpi, _bpc in enumerate(rows_plans)
                     if any(o.strip() in biz_primary_uuids for o in str(_bpc.get("client_uuids","")).split(",")))
_biz_plan_move = sum(1 for v in plan_move_biz if v == "YES")
print(f"Business plans: {_biz_plan_total} owned by biz primary contacts; {_biz_plan_move} flagged x_move_to_business=YES")

# ---------------------------------------------------------------- partition: individuals vs business/trust
biztrust_refs = {client_refs[i] for i, c in enumerate(rows_clients) if c["type"] in ("business", "trust")}

plan_bt = []
for i in range(len(rows_plans)):
    owners = [o.strip() for o in str(rows_plans[i].get("client_uuids", "")).split(",") if o.strip()]
    plan_bt.append(any(o in biztrust_refs for o in owners))
plan_ref_bt = {plan_refs[i]: plan_bt[i] for i in range(len(rows_plans))}


def _split(n, predicate):
    ind = [i for i in range(n) if not predicate(i)]
    bt = [i for i in range(n) if predicate(i)]
    return ind, bt


c_ind, c_bt = _split(len(rows_clients), lambda i: rows_clients[i]["type"] in ("business", "trust"))
p_ind, p_bt = _split(len(rows_plans), lambda i: plan_bt[i])
v_ind, v_bt = _split(len(rows_val), lambda i: plan_ref_bt.get(val_refs[i][0], False))
g_ind, g_bt = _split(len(rows_chg), lambda i: plan_ref_bt.get(chg_refs[i][0], False))
i_ind, i_bt = _split(len(rows_ie), lambda i: any(x.strip() in biztrust_refs for x in str(ie_refs[i]).split("|")))
x_ind, x_bt = _split(len(rows_exp), lambda i: any(x.strip() in biztrust_refs for x in str(exp_refs[i]).split("|")))
n_ind, n_bt = _split(len(rows_nt), lambda i: nt_refs[i] in biztrust_refs)

keep_ind = {"Clients": c_ind, "Plans": p_ind, "Valuations": v_ind, "Charges": g_ind,
            "Income & Employment": i_ind, "Expenses": x_ind, "Notes": n_ind}
keep_bt = {"Clients": c_bt, "Plans": p_bt, "Valuations": v_bt, "Charges": g_bt,
           "Income & Employment": i_bt, "Expenses": x_bt, "Notes": n_bt}

r_ind = build_book(OUT, f"{ADVISER.upper()} - PLANNR IMPORT PACK (INDIVIDUALS)", keep_ind)
_write_extra_sheets(OUT)
print("individuals ->", OUT, r_ind)

if any(keep_bt[t] for t in keep_bt):
    r_bt = build_book(OUT_BT, f"{ADVISER.upper()} - PLANNR IMPORT PACK (BUSINESS & TRUST)", keep_bt)
    _write_extra_sheets(OUT_BT)
    print("business/trust ->", OUT_BT, r_bt)
else:
    print("no business/trust clients - second pack not produced")

print("corrupt persons:", len(corrupt_refs), "| dup refs:", len(set(dup_refs)),
      "| orphan plan owners:", len(orphan_plan_owners),
      "| withdrawals skipped:", withdrawal_count,
      "| employment notes:", len(emp_history_notes))
print(f"Off-sheet orphan plans skipped: {_off_sheet_orphan_skipped}")
print(f"Not on his list (fee-paying off-sheet): {len(not_on_his_list)}")
print(f"Possible duplicates: {len(possible_dup_puuids)}")
print(f"New clients added: {_new_client_added}")
print(f"Title/email overrides: {len(_title_email_changes)} changes "
      f"(title:{sum(1 for x in _title_email_changes if x[1]=='title')}, "
      f"email:{sum(1 for x in _title_email_changes if x[1]=='email')})")
print(f"Navigator reconciliation: {len(_clients_with_fee)} clients have fee-bearing plan; "
      f"{_navigator_upgraded_count} actually upgraded to NAVIGATOR; "
      f"{_navigator_stayed_below_count} below NAVIGATOR with no fee (stayed Transactional/blank)")
print(f"Emails blanked (invalid/multi-email, Vision fallback tried): {len(_blanked_emails)}")
print(f"Health vulnerability suppressed (sheet 01 says good health): {_health_dropped_count}")
print(f"New-client rows excluded (no surname AND no forename): {len(_nc_nameless_excluded)}")
if _nc_nameless_excluded:
    for _ne in _nc_nameless_excluded:
        print(f"  Excluded adv_rows={_ne['adv_rows']} postcode={_ne['postcode']}")

# ---- Write clean importer files (01-09) ----------------------------------------
import csv as _csv_mod

IMPORTER_DIR = os.path.join(OUTPUT_DIR, "importer")
os.makedirs(IMPORTER_DIR, exist_ok=True)


def _write_importer(num, name, cols, rows_data, extra_cols=None):
    """Write one importer file as xlsx + csv. extra_cols: list of (col_name, values_list)."""
    extra_cols = extra_cols or []
    all_cols = list(cols) + [_ec for _ec, _ in extra_cols]
    _prefix = f"{num:02d}" if isinstance(num, int) else str(num)
    csv_path = os.path.join(IMPORTER_DIR, f"{_prefix} {name}.csv")
    def _sv(v): return "" if v is None else v
    with open(csv_path, "w", newline="", encoding="utf-8") as _cf:
        _wr = _csv_mod.writer(_cf)
        _wr.writerow(all_cols)
        for _ri, _row in enumerate(rows_data):
            _data = [_sv(_row.get(_c)) if isinstance(_row, dict) else "" for _c in cols]
            for _, _vals in extra_cols:
                _data.append(_vals[_ri] if _ri < len(_vals) else "")
            _wr.writerow(_data)
    from openpyxl import Workbook as _WB2
    _xwb = _WB2()
    _xws = _xwb.active
    _xws.title = name[:31]
    _xws.append(all_cols)
    for _ri, _row in enumerate(rows_data):
        _data = [_sv(_row.get(_c)) if isinstance(_row, dict) else "" for _c in cols]
        for _, _vals in extra_cols:
            _data.append(_vals[_ri] if _ri < len(_vals) else "")
        _xws.append(_data)
    _xwb.save(os.path.join(IMPORTER_DIR, f"{_prefix} {name}.xlsx"))
    return len(rows_data)


# 01 Clients — send_welcome_email=No
_imp_clients = [dict(rows_clients[i]) for i in keep_ind["Clients"]]
for _rc in _imp_clients:
    _rc["send_welcome_email"] = "No"
_imp_client_refs = [client_refs[i] for i in keep_ind["Clients"]]
n01 = _write_importer(1, "Clients", COLS["Clients"], _imp_clients,
                      extra_cols=[("externalreference", _imp_client_refs)])

# Add "Fix after import" sheet to 01 Clients.xlsx
_fix_xlsx = os.path.join(IMPORTER_DIR, "01 Clients.xlsx")
if os.path.isfile(_fix_xlsx):
    from openpyxl import load_workbook as _lw_fix
    _fix_wb = _lw_fix(_fix_xlsx)
    _fix_ws = _fix_wb.create_sheet("Fix after import")
    _fix_ws.append(["Issue", "externalreference", "original_value", "action"])
    # Blanked emails
    for _fe_ref, _fe_orig in _blanked_emails:
        _fix_ws.append(["blanked_email", _fe_ref, _fe_orig, "Supply correct email address"])
    # Health vulnerability suppressed — build list from client rows
    for _i2, _cr2 in enumerate(rows_clients):
        if str(_cr2.get("is_vulnerable", "")) == "" and _cr2.get("type") == "individual":
            _cr2_ref = client_refs[_i2] if _i2 < len(client_refs) else ""
            _cr2_note = flags_text[_i2] if _i2 < len(flags_text) else ""
            if "Health vulnerability suppressed" in _cr2_note and _cr2_ref:
                _fix_ws.append(["health_suppressed", _cr2_ref, "Health (sheet 11 default)",
                                 "Confirm vulnerability status with client"])
    # Unknown first names
    for _i3, _cr3 in enumerate(rows_clients):
        if str(_cr3.get("first_name", "")) == "Unknown":
            _cr3_ref = client_refs[_i3] if _i3 < len(client_refs) else ""
            _fix_ws.append(["unknown_first_name", _cr3_ref, "", "Supply correct first name"])
    # Nameless excluded rows
    for _ne2 in _nc_nameless_excluded:
        _fix_ws.append(["nameless_excluded", "", str(_ne2["adv_rows"]),
                         "No name found — add client manually in Plannr"])
    _fix_wb.save(_fix_xlsx)
    print(f"  Fix after import tab: {_fix_ws.max_row - 1} items")

# 02 Plans
_imp_plans = [dict(rows_plans[i]) for i in keep_ind["Plans"]]
_imp_plan_refs_2 = [plan_refs[i] for i in keep_ind["Plans"]]
n02 = _write_importer(2, "Plans", COLS["Plans"], _imp_plans,
                      extra_cols=[("externalreference", _imp_plan_refs_2)])

# 03 Valuations
n03 = _write_importer(3, "Valuations", COLS["Valuations"],
                      [rows_val[i] for i in keep_ind["Valuations"]])

# 04 Premiums — filter to individuals only (by plan ownership)
_prem_ind_idx = [i for i in range(len(rows_prem))
                 if not plan_ref_bt.get(prem_refs[i], False)]
n04 = _write_importer(4, "Premiums", COLS["Premiums"],
                      [rows_prem[i] for i in _prem_ind_idx])

# 05 Charges
n05 = _write_importer(5, "Charges", COLS["Charges"],
                      [rows_chg[i] for i in keep_ind["Charges"]])

# 06 Income & Employment — inject client_uuid from ie_refs
_imp_ie = []
for i in keep_ind["Income & Employment"]:
    _ie = dict(rows_ie[i])
    _ie["client_uuid"] = ie_refs[i].split("|")[0].strip()
    _imp_ie.append(_ie)
n06 = _write_importer(6, "Income & Employment", COLS["Income & Employment"], _imp_ie)

# 07 Expenses — joint rows split equally; inject client_uuid
_raw_exp_ind = [(rows_exp[i], exp_refs[i]) for i in keep_ind["Expenses"]]
_split_exp_rows = []
for _er, _eref in _raw_exp_ind:
    _refs_list = [_r.strip() for _r in str(_eref).split("|") if _r.strip()]
    if len(_refs_list) <= 1:
        _out = dict(_er)
        _out["client_uuid"] = _refs_list[0] if _refs_list else ""
        _split_exp_rows.append(_out)
    else:
        try:
            _total_amt = float(str(_er.get("amount", "") or 0))
        except (ValueError, TypeError):
            _total_amt = 0.0
        _half1 = round(_total_amt / 2, 2)
        _half2 = round(_total_amt - _half1, 2)
        for _ru, _amt_s in zip(_refs_list[:2], [_half1, _half2]):
            _out = dict(_er)
            _out["client_uuid"] = _ru
            _out["amount"] = _amt_s
            _split_exp_rows.append(_out)
_exp_chunk_size = 5000
_exp_n_files = max(1, -(-len(_split_exp_rows) // _exp_chunk_size))
n07 = 0
for _chunk_i in range(_exp_n_files):
    _chunk_rows = _split_exp_rows[_chunk_i * _exp_chunk_size: (_chunk_i + 1) * _exp_chunk_size]
    _write_importer(7, f"Expenses ({_chunk_i + 1} of {_exp_n_files})", COLS["Expenses"], _chunk_rows)
    n07 += len(_chunk_rows)

# 08 Notes — inject account_uuid from nt_refs
_imp_notes = []
for i in keep_ind["Notes"]:
    _nr = dict(rows_nt[i])
    _nr["account_uuid"] = nt_refs[i]
    _imp_notes.append(_nr)
n08 = _write_importer(8, "Notes", COLS["Notes"], _imp_notes)

# 09 Circles — build household groups from joint plans, I&E, and expense rows
_circ_parent = {}

def _circ_find(x):
    _circ_parent.setdefault(x, x)
    if _circ_parent[x] != x:
        _circ_parent[x] = _circ_find(_circ_parent[x])
    return _circ_parent[x]

def _circ_union(a, b):
    _circ_parent[_circ_find(a)] = _circ_find(b)

for _rp in rows_plans:
    _pu = [_u.strip() for _u in str(_rp.get("client_uuids", "")).split(",") if _u.strip()]
    for _j in range(1, len(_pu)):
        _circ_union(_pu[0], _pu[_j])

for _ref_ie in ie_refs:
    _pu = [_u.strip() for _u in str(_ref_ie).split("|") if _u.strip()]
    for _j in range(1, len(_pu)):
        _circ_union(_pu[0], _pu[_j])

for _ref_ex in exp_refs:
    _pu = [_u.strip() for _u in str(_ref_ex).split("|") if _u.strip()]
    for _j in range(1, len(_pu)):
        _circ_union(_pu[0], _pu[_j])

from collections import defaultdict as _defdict
_circ_groups = _defdict(set)
for _cu in _circ_parent:
    _circ_groups[_circ_find(_cu)].add(_cu)

_ind_ref_set = set(client_refs[i] for i in keep_ind["Clients"])
rows_circles = []
for _croot, _cmembers in sorted(_circ_groups.items()):
    _known = sorted(m for m in _cmembers if m in _ind_ref_set)
    if len(_known) < 2:
        continue
    _surnames = [CLIENT_LASTNAME.get(m, "") for m in _known if CLIENT_LASTNAME.get(m)]
    _surname = max(set(_surnames), key=_surnames.count) if _surnames else "Household"
    rows_circles.append({
        "uuid": "",
        "name": f"{_surname} Household",
        "account_uuids": ",".join(_known),
        "group_uuids": "",
    })

_circles_cols = ["uuid", "name", "account_uuids", "group_uuids"]
n09 = _write_importer(9, "Circles", _circles_cols, rows_circles)

# BT importer CSVs (for r3 plans + r4 valuations/charges)
_bt_plans = [dict(rows_plans[i]) for i in keep_bt["Plans"]]
_bt_plan_refs = [plan_refs[i] for i in keep_bt["Plans"]]
n_bt_plans = _write_importer("02_BT", "Plans BT", COLS["Plans"], _bt_plans,
                              extra_cols=[("externalreference", _bt_plan_refs)])

n_bt_val = _write_importer("03_BT", "Valuations BT", COLS["Valuations"],
                            [rows_val[i] for i in keep_bt["Valuations"]])

n_bt_chg = _write_importer("05_BT", "Charges BT", COLS["Charges"],
                            [rows_chg[i] for i in keep_bt["Charges"]])

print(f"\nImporter files written to {IMPORTER_DIR}")
print(f"  01 Clients:              {n01}")
print(f"  02 Plans:                {n02}")
print(f"  03 Valuations:           {n03}")
print(f"  04 Premiums:             {n04}")
print(f"  05 Charges:              {n05}")
print(f"  06 Income & Employment:  {n06}")
print(f"  07 Expenses (post-split):{n07}  (from {len(_raw_exp_ind)} pre-split, {_exp_n_files} files)")
print(f"  08 Notes:                {n08}")
print(f"  09 Circles:              {n09}")
print(f"  02 Plans BT:             {n_bt_plans}")
print(f"  03 Valuations BT:        {n_bt_val}")
print(f"  05 Charges BT:           {n_bt_chg}")
