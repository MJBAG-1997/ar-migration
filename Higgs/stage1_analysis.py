"""
Stage 1 Steps 3+4: Match adviser sheet to Vision export, propose mapping.
Rule 9: NO client names in stdout output.
"""
import os, re, sys
import pandas as pd
import openpyxl
from collections import defaultdict, Counter

DATA_ROOT = r"C:\Users\markj\OneDrive - Beyond Advice\Technology - Documents\workspace\AR Migration\Higgs"
VISION_SOURCE = os.path.join(DATA_ROOT, "AA Data")
ADVISER_XLSX = os.path.join(DATA_ROOT, "Higgs Master.xlsx")
STAGE1_DIR = os.path.join(DATA_ROOT, "stage1")
os.makedirs(STAGE1_DIR, exist_ok=True)
OUT_XLSX = os.path.join(STAGE1_DIR, "higgs_matches.xlsx")

# ── normalisation helpers ────────────────────────────────────────────────────

BLANK_TOKENS = {"tbc", "n/a", "unknown", "-", "", "none", "nan"}

def _blank(v):
    return str(v).strip().lower() in BLANK_TOKENS

def _nname(v):
    """Normalise a name: lower, strip titles, remove punctuation."""
    TITLES = {"mr","mrs","ms","miss","dr","prof","rev","sir","lord","lady","master"}
    v = str(v or "").lower()
    v = re.sub(r"[^a-z0-9 ]", " ", v)
    parts = [p for p in v.split() if p not in TITLES]
    return " ".join(parts).strip()

def _npol(v):
    """Normalise a policy number: upper, remove spaces/hyphens/slashes/dots, keep leading zeros."""
    if _blank(v):
        return ""
    v = str(v).upper()
    # If it was read as a float (e.g. 123456.0), strip .0
    if re.fullmatch(r"-?\d+\.0", v):
        v = v[:-2]
    v = re.sub(r"[\s\-/\.]", "", v)
    return v.strip()

def _strip_leading_zeros(v):
    return v.lstrip("0") if v else v

def _pnorm(x):
    x = str(x or "").lower().strip()
    x = x.replace("&", " and ")
    x = re.sub(r"\b(ltd|limited|plc|llp|uk|group)\b", "", x)
    x = re.sub(r"[^a-z0-9 ]", " ", x)
    return re.sub(r"\s+", " ", x).strip()

# ── product → Plannr type map for col-J types ────────────────────────────────

ADV_TYPE_MAP = {
    "sipp": "pension", "sipp dfm": "pension", "sipp bank acc": "pension",
    "sipp wrapper": "pension", "ssas dfm": "pension", "ssas bank acc": "pension",
    "pension": "pension", "retirement account": "pension",
    "dfm": "investment", "mps": "investment", "bond dfm": "investment",
    "dgeps": "investment",
    "isa": "investment", "jisa": "investment",
    "gia": "investment", "gia & isa": "investment",
    "bond": "investment", "offshore bond": "investment",
    "unit trust": "investment", "oeic": "investment",
    "vct": "investment", "eis": "investment",
    "iht": "investment", "investment plan": "investment",
    "dfm (within bond)": "investment", "dfm (within sipp)": "pension",
    "cash account": "cash_account",
    "liquidity": "investment",
}

def adv_plannr_type(t):
    return ADV_TYPE_MAP.get(str(t or "").lower().strip().rstrip(), "")

# ── REVIEW_CLASS_TO_STATUS (from builder) ────────────────────────────────────
REVIEW_CLASS_TO_STATUS = {
    "1": "LEVEL 1 / PIONEER", "1.0": "LEVEL 1 / PIONEER",
    "2": "LEVEL 2 / NAVIGATOR", "2.0": "LEVEL 2 / NAVIGATOR",
    "3": "", "3.0": "",
    "4": "TRANSACTIONAL", "4.0": "TRANSACTIONAL",
}

# ── Read adviser xlsx ────────────────────────────────────────────────────────

print("Reading adviser xlsx...")
wb = openpyxl.load_workbook(ADVISER_XLSX, data_only=True)
ws = wb["Sheet1"]

# detect pct-format columns
pct_cols = set()
for rn in range(2, min(ws.max_row+1, 20)):
    for ci in range(1, ws.max_column+1):
        cell = ws.cell(rn, ci)
        if cell.number_format and "%" in (cell.number_format or "") and cell.value is not None:
            pct_cols.add(ci)

adv_rows = []
int_pn_count = 0
for rn in range(2, ws.max_row+1):
    def cv(ci):
        return ws.cell(rn, ci).value

    surname = str(cv(3) or "").strip()
    forename = str(cv(4) or "").strip()
    if not surname and not forename:
        continue  # skip blank rows

    pn_raw = cv(9)
    pn_is_int = isinstance(pn_raw, (int, float)) and not isinstance(pn_raw, bool)
    if pn_is_int:
        int_pn_count += 1
    pn_str = "" if pn_raw is None else str(int(pn_raw)) if isinstance(pn_raw, float) and pn_raw == int(pn_raw) else str(pn_raw)

    # fee pct: col M (13), already decimal (0.0055 = 0.55%)
    fee_raw = cv(13)
    fee_pct = None
    if isinstance(fee_raw, (int, float)) and fee_raw is not None:
        fee_pct = round(float(fee_raw) * 100, 4)  # 0.0055 → 0.55
    elif isinstance(fee_raw, str) and fee_raw.strip() not in ("", "Flat Fee"):
        try:
            fee_pct = round(float(fee_raw.replace("%",""))*100, 4)
        except:
            fee_pct = None

    adv_rows.append({
        "adv_row": rn,
        "adv_surname": surname,
        "adv_forename": forename,
        "adv_address": str(cv(1) or "").strip(),
        "adv_ownership": str(cv(5) or "").strip(),
        "adv_provider": str(cv(8) or "").strip(),
        "adv_policy_number_raw": pn_str,
        "adv_policy_number_is_int": pn_is_int,
        "adv_type": str(cv(10) or "").strip(),
        "adv_value": cv(11),
        "adv_fee_pct": fee_pct,
        "adv_frequency": str(cv(16) or "").strip(),
        "adv_x_label": str(cv(24) or "").strip(),
    })

print(f"Adviser rows loaded: {len(adv_rows)}")
print(f"Policy numbers stored as int (leading-zero risk): {int_pn_count}")

# ── Read Vision sheets ────────────────────────────────────────────────────────

def read_csv(path):
    try:
        return pd.read_csv(path, dtype=str, keep_default_na=False, encoding="utf-8-sig")
    except UnicodeDecodeError:
        return pd.read_csv(path, dtype=str, keep_default_na=False, encoding="latin-1")

print("Reading Vision CSVs...")
vis_files = {os.path.splitext(f)[0]: os.path.join(VISION_SOURCE, f)
             for f in sorted(os.listdir(VISION_SOURCE)) if f.lower().endswith(".csv")}

def find_sheet(keyword):
    for k, p in vis_files.items():
        if keyword.lower() in k.lower():
            return read_csv(p)
    return pd.DataFrame()

persons_df = find_sheet("01. Person")
policies_df = find_sheet("04. Polic")
addr_df = find_sheet("12. Address")
seg_df = find_sheet("13. Segment")

def s(v):
    if v is None: return ""
    v = str(v).strip()
    return "" if v.lower() in ("nan","nat","none") else v

# ── Build Vision person index ─────────────────────────────────────────────────

# Latest current address per client
addr_postcode = {}  # client_uuid → postcode
for _, r in addr_df.iterrows():
    if s(r.get("status","")).lower() == "current":
        cu = s(r.get("uuid",""))
        pc = s(r.get("postcode","")).replace(" ","").upper()
        if cu and pc:
            # keep latest by moved_in_date
            prev = addr_postcode.get(cu)
            if prev is None:
                addr_postcode[cu] = pc

# Person lookup
vis_persons = {}  # client_uuid → dict
for _, r in persons_df.iterrows():
    cu = s(r.get("client_uuid",""))
    if not cu: continue
    dob = s(r.get("date_of_birth",""))
    ni = s(r.get("national_insurance_number","")).replace(" ","").upper()[:8]
    surn = _nname(s(r.get("surname","")))
    fore = _nname(s(r.get("forename","")))
    pc = addr_postcode.get(cu, "")
    vis_persons[cu] = {
        "surname": surn, "forename": fore, "dob": dob, "ni": ni,
        "postcode": pc, "client_uuid": cu,
    }

# ── Build Vision policy index ─────────────────────────────────────────────────

vis_policies = {}  # policy_uuid → dict
# Map normalised policy numbers to policy_uuid(s)
pn_to_puuid = defaultdict(list)  # norm_pn → [policy_uuid, ...]

for _, r in policies_df.iterrows():
    pu = s(r.get("policy_uuid",""))
    if not pu: continue
    raw_pns = s(r.get("policy_numbers",""))
    norm_pns = [_npol(p) for p in raw_pns.split("|") if not _blank(_npol(p))]
    vis_policies[pu] = {
        "policy_uuid": pu,
        "policy_numbers_raw": raw_pns,
        "norm_policy_numbers": norm_pns,
        "provider": s(r.get("primary_provider","")),
        "provider_norm": _pnorm(s(r.get("primary_provider",""))),
        "product": s(r.get("product","")),
        "product_specific_name": s(r.get("product_specific_name","")),
        "client_1_uuid": s(r.get("client_1_uuid","")),
        "client_2_uuid": s(r.get("client_2_uuid","")),
        "client_1_name": s(r.get("client_1","")),
        "client_2_name": s(r.get("client_2","")),
        "under_advice": s(r.get("under_my_advice","")),
        "fee_pct": s(r.get("ongoing_service_charge_percentage","")),
        "fee_freq": s(r.get("ongoing_service_charge_frequency","")),
    }
    for npn in norm_pns:
        if npn:
            pn_to_puuid[npn].append(pu)

# Also map stripped-leading-zeros variants
pn_nozero_to_puuid = defaultdict(list)
for npn, puuids in pn_to_puuid.items():
    nz = _strip_leading_zeros(npn)
    if nz != npn:
        for pu in puuids:
            pn_nozero_to_puuid[nz].append(pu)

# ── Build person name→uuid lookups ────────────────────────────────────────────

# surname+fore → [uuids]
name_to_uuid = defaultdict(list)
# surname+fore+postcode → [uuids]
name_pc_to_uuid = defaultdict(list)

for cu, p in vis_persons.items():
    key = (p["surname"], p["forename"])
    name_to_uuid[key].append(cu)
    if p["postcode"]:
        key_pc = (p["surname"], p["forename"], p["postcode"])
        name_pc_to_uuid[key_pc].append(cu)

# ── Baseline pack Plannr types (from baseline output) ────────────────────────
# We'll classify Vision policies using pm_lookup (simplified)

PM_SIMPLE = {
    # vision product keyword → plannr type
    "sipp": "pension", "ssas": "pension", "personal pension": "pension",
    "executive pension": "pension", "group personal pension": "pension",
    "defined benefit": "pension", "income drawdown": "pension",
    "annuity": "annuity",
    "isa": "investment", "unit trust": "investment", "oeic": "investment",
    "investment bond": "investment", "offshore bond": "investment",
    "general investment": "investment", "discretionary": "investment",
    "aim": "investment", "vct": "investment", "eis": "investment",
    "bpr": "investment", "iht": "investment", "wrap": "investment",
    "bond": "investment",
    "term assurance": "protection_policy", "term protection": "protection_policy",
    "whole of life": "protection_policy", "income protection": "protection_policy",
    "relevant life": "protection_policy", "life cover": "protection_policy",
    "flexible life": "protection_policy",
    "property": "fixed_asset",
    "mortgage": "mortgage",
}

def vis_plannr_type(product):
    p = product.lower()
    for kw, t in PM_SIMPLE.items():
        if kw in p:
            return t
    return "unknown"

# ── Matching ──────────────────────────────────────────────────────────────────

# Track which Vision policy UUIDs have been matched
matched_vision_puuids = set()

results = []

for row in adv_rows:
    adv_pn = _npol(row["adv_policy_number_raw"])
    adv_pn_nz = _strip_leading_zeros(adv_pn)
    adv_surn = _nname(row["adv_surname"])
    adv_fore = _nname(row["adv_forename"])
    # extract postcode from address col A (last token that looks like postcode)
    pc_match = re.search(r"\b([A-Z]{1,2}\d{1,2}[A-Z]?\s*\d[A-Z]{2})\b", row["adv_address"].upper())
    adv_pc = pc_match.group(1).replace(" ","") if pc_match else ""

    match_status = "Adviser-only"
    match_method = ""
    vision_puuid = ""
    vision_pns = ""
    vision_product = ""
    vision_provider = ""
    vision_under_advice = ""
    vision_fee_pct = ""
    candidates = ""

    # ── 1. Policy number exact match
    if adv_pn and adv_pn in pn_to_puuid:
        puuids = pn_to_puuid[adv_pn]
        if len(puuids) == 1:
            vision_puuid = puuids[0]
            match_status = "Matched"
            match_method = "policy_number"
        else:
            match_status = "Ambiguous"
            match_method = "policy_number"
            candidates = "; ".join(puuids)

    # ── 2. Policy number with leading zeros stripped
    elif adv_pn_nz and adv_pn_nz in pn_nozero_to_puuid:
        puuids = pn_nozero_to_puuid[adv_pn_nz]
        if len(puuids) == 1:
            vision_puuid = puuids[0]
            match_status = "Matched"
            match_method = "policy_number_nozero"
        else:
            match_status = "Ambiguous"
            match_method = "policy_number_nozero"
            candidates = "; ".join(puuids)

    # ── 3. No policy number → owner+provider+type
    elif not adv_pn:
        adv_prov_n = _pnorm(row["adv_provider"])
        adv_ptype = adv_plannr_type(row["adv_type"])
        # get Vision owner UUIDs for this person
        owner_cands = set()
        key = (adv_surn, adv_fore)
        for cu in name_to_uuid.get(key, []):
            owner_cands.add(cu)
        if adv_pc and not owner_cands:
            key_pc = (adv_surn, adv_fore, adv_pc)
            for cu in name_pc_to_uuid.get(key_pc, []):
                owner_cands.add(cu)
        if owner_cands:
            type_prov_matches = []
            for pu, pol in vis_policies.items():
                if pol["client_1_uuid"] not in owner_cands and pol["client_2_uuid"] not in owner_cands:
                    continue
                if _pnorm(pol["provider"]) != adv_prov_n:
                    continue
                vt = vis_plannr_type(pol["product"])
                if adv_ptype and vt != adv_ptype and adv_ptype != "investment" and vt != "investment":
                    continue
                type_prov_matches.append(pu)
            if len(type_prov_matches) == 1:
                vision_puuid = type_prov_matches[0]
                match_status = "Matched"
                match_method = "owner_provider_type"
            elif len(type_prov_matches) > 1:
                match_status = "Ambiguous"
                match_method = "owner_provider_type"
                candidates = "; ".join(type_prov_matches)

    # ── Person matching for unmatched rows ─────────────────────────────────
    if match_status == "Adviser-only":
        # Try name+postcode first
        if adv_pc:
            key_pc = (adv_surn, adv_fore, adv_pc)
            uuids = name_pc_to_uuid.get(key_pc, [])
            if len(uuids) == 1:
                # found a person — no policy yet, mark as New-client (adviser-only plan)
                match_status = "Adviser-only"
                match_method = "name_postcode"
            elif len(uuids) > 1:
                match_status = "Ambiguous"
                match_method = "name_postcode"
                candidates = "; ".join(uuids)

        if match_status == "Adviser-only" and not match_method:
            # Try name only
            key = (adv_surn, adv_fore)
            uuids = name_to_uuid.get(key, [])
            if len(uuids) == 1:
                # check for near-miss Vision plan with same owner+provider
                cu = uuids[0]
                adv_prov_n = _pnorm(row["adv_provider"])
                near_miss = [pu for pu, pol in vis_policies.items()
                             if (pol["client_1_uuid"] == cu or pol["client_2_uuid"] == cu)
                             and _pnorm(pol["provider"]) == adv_prov_n]
                if near_miss:
                    match_status = "Ambiguous"
                    match_method = "name_only"
                    candidates = "; ".join(near_miss)
                else:
                    match_method = "name_only"  # still adviser-only
            elif len(uuids) > 1:
                match_status = "Ambiguous"
                match_method = "name_only"
                candidates = "; ".join(uuids)
            else:
                # No match at all → New-client
                match_status = "New-client"
                match_method = ""

    # ── Near-miss check before final Adviser-only ─────────────────────────
    if match_status == "Adviser-only" and adv_pn:
        # Check if any Vision person shares surname + postcode (near-miss)
        adv_prov_n = _pnorm(row["adv_provider"])
        if adv_pc:
            near_persons = [cu for cu, p in vis_persons.items()
                            if p["surname"] == adv_surn and p["postcode"] == adv_pc]
            if near_persons:
                near_plans = [pu for pu, pol in vis_policies.items()
                              if pol["client_1_uuid"] in near_persons
                              and _pnorm(pol["provider"]) == adv_prov_n]
                if near_plans:
                    match_status = "Ambiguous"
                    match_method = "near_miss_prov"
                    candidates = "; ".join(near_plans)

    # Populate Vision details
    if vision_puuid and vision_puuid in vis_policies:
        pol = vis_policies[vision_puuid]
        vision_pns = pol["policy_numbers_raw"]
        vision_product = pol["product_specific_name"] or pol["product"]
        vision_provider = pol["provider"]
        vision_under_advice = pol["under_advice"]
        vision_fee_pct = pol["fee_pct"]
        matched_vision_puuids.add(vision_puuid)

    results.append({
        "adv_row": row["adv_row"],
        "adv_surname": row["adv_surname"],
        "adv_forename": row["adv_forename"],
        "adv_provider": row["adv_provider"],
        "adv_policy_number": row["adv_policy_number_raw"],
        "adv_type": row["adv_type"],
        "adv_value": row["adv_value"],
        "adv_fee_pct": row["adv_fee_pct"],
        "adv_frequency": row["adv_frequency"],
        "match_status": match_status,
        "match_method": match_method,
        "vision_policy_uuid": vision_puuid,
        "vision_policy_numbers": vision_pns,
        "vision_product": vision_product,
        "vision_provider": vision_provider,
        "vision_under_advice": vision_under_advice,
        "vision_fee_pct": vision_fee_pct,
        "candidates": candidates,
        "Decision": "",
    })

# ── Output xlsx ───────────────────────────────────────────────────────────────

df_out = pd.DataFrame(results)
df_out.to_excel(OUT_XLSX, index=False)
print(f"\nMatches saved to: {OUT_XLSX}")

# ── Summary stats (Rule 9: no names) ─────────────────────────────────────────

print(f"\n{'='*60}")
print(f"STAGE 1 REPORT — HIGGS FINANCIAL SOLUTIONS LTD")
print(f"{'='*60}")

print(f"\nTotal adviser rows processed: {len(results)}")
print(f"Policy numbers stored as int (leading-zero risk): {int_pn_count}")
no_pn = sum(1 for r in adv_rows if not _npol(r["adv_policy_number_raw"]))
print(f"Rows with no policy number: {no_pn}")

print(f"\n--- Match status ---")
status_counts = Counter(r["match_status"] for r in results)
for k, v in sorted(status_counts.items()):
    print(f"  {k}: {v}")

print(f"\n--- Match method ---")
method_counts = Counter(r["match_method"] for r in results)
for k, v in sorted(method_counts.items()):
    print(f"  {k!r}: {v}")

print(f"\n--- Adviser Types (col J) ---")
type_counts = Counter(r["adv_type"].strip() for r in adv_rows)
for k, v in sorted(type_counts.items(), key=lambda x: -x[1]):
    planned = adv_plannr_type(k)
    print(f"  {k!r:30s} {v:4d}  → {planned or 'UNMAPPED'}")

print(f"\n--- Frequency distribution (col P) ---")
freq_counts = Counter(r["adv_frequency"] for r in adv_rows)
for k, v in sorted(freq_counts.items(), key=lambda x: -x[1]):
    print(f"  {k!r:20s} {v}")

print(f"\n--- Fee % range (col M × 100) ---")
fees = [r["adv_fee_pct"] for r in adv_rows if r["adv_fee_pct"] is not None and r["adv_fee_pct"] > 0]
if fees:
    fees_s = pd.Series(fees)
    print(f"  Min: {fees_s.min():.4f}%  Max: {fees_s.max():.4f}%  Median: {fees_s.median():.4f}%")
    print(f"  Flat Fee / non-numeric: {sum(1 for r in adv_rows if r['adv_fee_pct'] is None)}")

print(f"\n--- Value range (col K) ---")
vals = [r["adv_value"] for r in adv_rows if r["adv_value"] is not None and isinstance(r["adv_value"], (int,float))]
if vals:
    vs = pd.Series(vals)
    print(f"  Min: £{vs.min():,.0f}  Max: £{vs.max():,.0f}  Median: £{vs.median():,.0f}")

print(f"\n--- Col X legacy labels ---")
x_counts = Counter(r["adv_x_label"] for r in adv_rows if r["adv_x_label"])
for k, v in sorted(x_counts.items(), key=lambda x: -x[1]):
    print(f"  {k!r}: {v}")

print(f"\n--- Adviser types with no PM mapping ---")
unmapped = [(k, v) for k, v in type_counts.items() if not adv_plannr_type(k) and k]
for k, v in sorted(unmapped, key=lambda x: -x[1]):
    print(f"  {k!r}: {v}")

# ── Vision plans NOT on adviser sheet ────────────────────────────────────────

unmatched_vis = {pu: pol for pu, pol in vis_policies.items() if pu not in matched_vision_puuids}
print(f"\n--- Vision plans NOT on adviser sheet: {len(unmatched_vis)} ---")

by_type = defaultdict(lambda: {"total":0, "under_advice":0, "has_fee":0, "has_agency":0})
for pu, pol in unmatched_vis.items():
    vt = vis_plannr_type(pol["product"])
    by_type[vt]["total"] += 1
    if pol["under_advice"].lower() in ("yes","servicing_agent","information_authority"):
        by_type[vt]["under_advice"] += 1
    fee = pol["fee_pct"]
    try:
        if fee and float(fee) > 0:
            by_type[vt]["has_fee"] += 1
    except:
        pass

for vt, ct in sorted(by_type.items()):
    print(f"  {vt:30s} total={ct['total']:4d}  under_advice={ct['under_advice']:3d}  has_fee={ct['has_fee']:3d}")

was_advised_total = sum(1 for pol in unmatched_vis.values()
                        if pol["under_advice"].lower() in ("yes","servicing_agent","information_authority"))
was_fee_total = 0
for pol in unmatched_vis.values():
    try:
        if pol["fee_pct"] and float(pol["fee_pct"]) > 0:
            was_fee_total += 1
    except:
        pass

print(f"  TOTAL with under_advice: {was_advised_total}")
print(f"  TOTAL with fee: {was_fee_total}")

# ── Plans on adviser sheet with no fee from either source ────────────────────

print(f"\n--- Plans on adviser sheet with no fee from either source ---")
no_fee_count = 0
no_fee_matched = 0
for r in results:
    adv_fee = r["adv_fee_pct"]
    vis_fee = r["vision_fee_pct"]
    adv_has_fee = adv_fee is not None and adv_fee > 0
    try:
        vis_has_fee = vis_fee and float(vis_fee) > 0
    except:
        vis_has_fee = False
    if not adv_has_fee and not vis_has_fee:
        no_fee_count += 1
        if r["match_status"] == "Matched":
            no_fee_matched += 1
print(f"  All rows: {no_fee_count}  (of which matched: {no_fee_matched})")

# ── Service levels in Vision segmentation ────────────────────────────────────

print(f"\n--- Service levels (Vision sheet 13) ---")
if not seg_df.empty:
    seg_cur = seg_df[seg_df.get("segment_status","").str.lower() == "current"] if "segment_status" in seg_df.columns else seg_df
    if "segment" in seg_df.columns:
        seg_counts = Counter(seg_df["segment"].str.strip().tolist())
        for k, v in sorted(seg_counts.items(), key=lambda x: -x[1]):
            if k:
                print(f"  {k!r}: {v}")
    else:
        print("  No 'segment' column in segmentation sheet")
else:
    print("  Segmentation sheet empty or not found")

# ── Vision client service classification ─────────────────────────────────────
print(f"\n--- Vision client classification codes → status ---")
if "classification" in persons_df.columns:
    cls_counts = Counter(persons_df["classification"].str.strip().tolist())
    for k, v in sorted(cls_counts.items(), key=lambda x: -x[1]):
        if k:
            status = REVIEW_CLASS_TO_STATUS.get(k, "NOT MAPPED")
            print(f"  class={k!r}: {v} → {status!r}")

# ── Income and expenditure amounts ───────────────────────────────────────────

ie_csv = None
for fn in sorted(os.listdir(VISION_SOURCE)):
    if "23." in fn or "income and expend" in fn.lower():
        ie_csv = os.path.join(VISION_SOURCE, fn)
        break
if not ie_csv:
    for fn in sorted(os.listdir(VISION_SOURCE)):
        if "income" in fn.lower() and "expenditure" in fn.lower() and "entries" not in fn.lower():
            ie_csv = os.path.join(VISION_SOURCE, fn)
            break

ie_ent_csv = None
for fn in sorted(os.listdir(VISION_SOURCE)):
    if "income and expenditure" in fn.lower() and "entries" not in fn.lower():
        ie_ent_csv = os.path.join(VISION_SOURCE, fn)
        break
if not ie_ent_csv:
    for fn in sorted(os.listdir(VISION_SOURCE)):
        path = os.path.join(VISION_SOURCE, fn)
        df_tmp = pd.read_csv(path, dtype=str, keep_default_na=False, nrows=2, encoding="utf-8-sig")
        if "income_or_expenditure" in [c.lower() for c in df_tmp.columns]:
            ie_ent_csv = path
            break

if ie_ent_csv:
    ie_df = pd.read_csv(ie_ent_csv, dtype=str, keep_default_na=False, encoding="utf-8-sig")
    print(f"\n--- Income & Expenditure amounts (Vision sheet 24) ---")
    if "amount" in ie_df.columns:
        try:
            amounts = pd.to_numeric(ie_df["amount"], errors="coerce").dropna()
            income_amounts = amounts[ie_df.reindex(amounts.index)["income_or_expenditure"].str.lower() == "income"] if "income_or_expenditure" in ie_df.columns else amounts
            exp_amounts = amounts[ie_df.reindex(amounts.index)["income_or_expenditure"].str.lower() == "expenditure"] if "income_or_expenditure" in ie_df.columns else pd.Series()
            print(f"  Income rows: {len(income_amounts)}  Median: £{income_amounts.median():,.0f}  >£1500: {(income_amounts>1500).sum()}")
            print(f"  Expenditure rows: {len(exp_amounts)}  Median: £{exp_amounts.median():,.0f}  >£1500: {(exp_amounts>1500).sum()}")
        except Exception as e:
            print(f"  Error: {e}")

# ── Proposed mapping table ────────────────────────────────────────────────────

print(f"""
{'='*60}
STEP 4 — PROPOSED COLUMN MAPPING
{'='*60}

Adviser sheet is plan-level (one row per plan).
Each row maps to Vision sheet 04 (Policies) via policy number.

Columns map as follows (adviser col → Vision sheet.column or x_override):

  Adviser → Vision/Builder
  ─────────────────────────────────────────────────────────────
  A  Address            → x_override_address_1 / town_city / postcode (client address update on sheet 12)
  B  Title              → x_override_title (client field)
  C  Surname            → (matching key only; updates client if Vision blank)
  D  Forename           → (matching key only)
  E  Couple/Single      → not mapped (ownership indicator; used as context)
  F  Age                → not mapped (derived from DOB; no importer field)
  G  Email              → x_override_email (client field)
  H  Provider           → sheet 04: primary_provider (x_override_primary_provider)
  I  Policy Number      → sheet 04: policy_numbers (matching key)
  J  Type               → sheet 04: product (x_override_product / type mapping)
  K  Current Value      → x_value (builder step 6b: valuation from adviser sheet)
  L  Household AUM      → not mapped (aggregated total; no importer field)
  M  AR Fee % excl VAT  → sheet 04: ongoing_service_charge_percentage (×100 to get %)
  N  AR Fee % incl VAT  → not mapped (use excl. VAT version for Plannr)
  O  AR Fee £ incl VAT  → DECISION NEEDED: no £ fee route in builder (see below)
  P  Frequency          → sheet 04: ongoing_service_charge_frequency
  Q  VIP % excl VAT     → not mapped (platform/provider rebate; not a builder field)
  R  VIP % incl VAT     → not mapped
  S  VIP £ incl VAT     → not mapped
  T  Provider % Fee     → not mapped (provider's own fee, not AR fee)
  U  Provider % incl V  → not mapped
  V  Provider £ incl V  → not mapped
  W  Last Updated       → x_valuation_date (paired with col K value)
  X  Legacy label       → not mapped (Inherited/Trail Commission = info only)

  DECISIONS NEEDED:
  D1. Col O (AR Fee £ incl. VAT): no route for a £ fee amount in the builder's charge
      logic, which only handles %. List the plans where only a £ figure is available
      (col M is blank/zero) for adviser to confirm equivalent % or enter manually.
  D2. Col J types with no PM mapping (see above) → propose builder PM additions below.
  D3. Col L (Household AUM) and col E (Couple/Single) are informational only; not imported.
  D4. Adviser fee % is excl. VAT; Plannr tax_exempt=Yes assumed unless col O has VAT marker.
""")

# ── Products needing PM additions ────────────────────────────────────────────

print("--- Products/types needing builder PM additions ---")
PROPOSED_PM = {
    "DGEPS":            ("investment", "Discretionary Managed Service", "DGEPS = Discretionary Growth & Estate Planning Service; confirm sub-type"),
    "Bond DFM":         ("investment", "Discretionary Managed Service", "Bond wrapper with DFM mandate; confirm sub-type"),
    "GIA & ISA":        ("investment", "General Investment Account", "Combined GIA+ISA account; two plans in one row - needs splitting"),
    "SIPP DFM":         ("pension", "SIPP", "SIPP with DFM mandate"),
    "SSAS DFM":         ("pension", "SSAS", "SSAS with DFM mandate"),
    "SSAS Bank Acc":    ("pension", "SSAS", "SSAS bank account component"),
    "SIPP Bank Acc":    ("pension", "SIPP", "SIPP bank account component"),
    "SIPP Wrapper":     ("pension", "SIPP", "SIPP wrapper; confirm sub-type"),
    "DFM (Within Bond)":("investment", "Discretionary Managed Service", "DFM within bond wrapper; type should follow the bond"),
    "DFM (Within SIPP)":("pension", "SIPP", "DFM within SIPP; type should follow the pension wrapper"),
    "IHT":              ("investment", "Investment Tax Service", "IHT/BPR portfolio; confirm sub-type"),
    "Liquidity":        ("investment", "General Investment Account", "Liquidity account; confirm type"),
    "Investment Plan":  ("investment", "Unknown", "Generic; specify sub-type"),
    "Cash Account":     ("cash_account", "Cash", "Cash/bank account; confirm type in Plannr"),
    "MPS":              ("investment", "Discretionary Managed Service", "Model Portfolio Service"),
    "JISA":             ("investment", "ISA (Stocks And Shares)", "Junior ISA"),
    "VCT":              ("investment", "Venture Capital Trust (Regulated)", "Confirm regulated vs non-regulated"),
    "EIS":              ("investment", "Enterprise Investment Scheme (Requlated)", "Confirm regulated vs non-regulated"),
    "Retirement Account": ("pension", "Personal Pension Plan", "Retirement account - confirm wrapper type"),
}

for k, (typ, sub, note) in PROPOSED_PM.items():
    cnt = type_counts.get(k, type_counts.get(k.strip(), 0))
    if cnt:
        print(f"  {k!r:30s} n={cnt}  → type={typ!r}, sub_type={sub!r}")
        print(f"    Note: {note}")

print(f"\nDone.")
