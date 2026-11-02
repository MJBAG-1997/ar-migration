"""
Higgs Stage 2 — resolve ambiguous rows by rule, produce higgs_matches_v4.xlsx.
Rule 9: no client names in output.
"""
import re, os, collections, datetime
import pandas as pd
import openpyxl
from openpyxl import Workbook

DATA_ROOT = r"C:\Users\markj\OneDrive - Beyond Advice\Technology - Documents\workspace\AR Migration\Higgs"
VISION_SRC = os.path.join(DATA_ROOT, "AA Data")
MATCHES_V3 = os.path.join(DATA_ROOT, "stage1", "higgs_matches_v3.xlsx")
MATCHES_V4 = os.path.join(DATA_ROOT, "stage1", "higgs_matches_v4.xlsx")
ADV_XLSX   = os.path.join(DATA_ROOT, "Higgs Master.xlsx")

def _pnorm(x):
    x = str(x).lower().strip()
    x = x.replace('&', ' and ')
    x = re.sub(r'\b(ltd|limited|plc|llp|uk|group)\b', '', x)
    x = re.sub(r'[^a-z0-9 ]', ' ', x)
    return re.sub(r'\s+', ' ', x).strip()

def extract_postcode(addr):
    m = re.search(r'\b([A-Z]{1,2}[0-9][0-9A-Z]?\s*[0-9][A-Z]{2})\b', str(addr).upper())
    return m.group(1).replace(' ', '').upper() if m else ''

# ── load matches v3 ────────────────────────────────────────────────────────────
mdf = pd.read_excel(MATCHES_V3, sheet_name="Matches", dtype=str).fillna("")

# ── load adviser xlsx for supplemental columns ────────────────────────────────
wb_adv = openpyxl.load_workbook(ADV_XLSX, data_only=True)
ws_adv = wb_adv['Sheet1']
# Index: excel_row_number → {last_updated, adv_email, adv_title, adv_address, adv_postcode}
adv_row_data = {}
for rn in range(2, ws_adv.max_row + 1):
    vals = [ws_adv.cell(rn, c).value for c in range(1, 30)]
    if all(v is None or (isinstance(v, str) and not v.strip()) for v in vals):
        continue
    addr_raw = str(vals[0] or '').strip()
    title_raw = str(vals[1] or '').strip()
    email_raw = str(vals[6] or '').strip()
    w_raw = vals[22]  # col W = last updated
    if hasattr(w_raw, 'date'):
        last_upd = w_raw.date().isoformat()
    elif isinstance(w_raw, datetime.datetime):
        last_upd = w_raw.date().isoformat()
    elif isinstance(w_raw, datetime.date):
        last_upd = w_raw.isoformat()
    else:
        last_upd = ''
    adv_row_data[rn] = {
        'last_updated': last_upd,
        'adv_email': email_raw,
        'adv_title': title_raw,
        'adv_address': addr_raw,
        'adv_postcode': extract_postcode(addr_raw),
    }

# Add supplemental columns to mdf
for col in ('last_updated', 'adv_email', 'adv_title', 'adv_address', 'adv_postcode'):
    mdf[col] = ''
for idx, row in mdf.iterrows():
    try:
        rn = int(float(str(row.get('adv_row', ''))))
    except (ValueError, TypeError):
        continue
    supp = adv_row_data.get(rn, {})
    for col in ('last_updated', 'adv_email', 'adv_title', 'adv_address', 'adv_postcode'):
        mdf.at[idx, col] = supp.get(col, '')

print(f"v3 loaded: {len(mdf)} rows, supplemental columns added from adviser xlsx")

# ── load Vision policies ──────────────────────────────────────────────────────
p04 = pd.read_csv(os.path.join(VISION_SRC, "Phil Thompson_04. Policies.csv"),
                  dtype=str, keep_default_na=False, encoding='utf-8-sig')
p02 = pd.read_csv(os.path.join(VISION_SRC, "Phil Thompson_02. Business-Clients.csv"),
                  dtype=str, keep_default_na=False, encoding='utf-8-sig')
p03 = pd.read_csv(os.path.join(VISION_SRC, "Phil Thompson_03. Trust-Clients.csv"),
                  dtype=str, keep_default_na=False, encoding='utf-8-sig')
p01 = pd.read_csv(os.path.join(VISION_SRC, "Phil Thompson_01. Person-Clients.csv"),
                  dtype=str, keep_default_na=False, encoding='utf-8-sig')

vis_policies = {r['policy_uuid'].strip(): r.to_dict() for _, r in p04.iterrows()}
all_client_uuids = (set(p01['client_uuid'].str.strip()) |
                    set(p02['primary_contact_uuid'].str.strip()) |
                    set(p03['client_uuid'].str.strip()))
biz_person_uuids = set(p02['primary_contact_uuid'].str.strip())

print(f"Vision loaded: {len(vis_policies)} policies, {len(all_client_uuids)} client UUIDs, "
      f"{len(biz_person_uuids)} business primary contacts")

# ── PM sub_type lookup ────────────────────────────────────────────────────────
PM_SUB = {
    'sipp': 'SIPP', 'sipp dfm': 'SIPP', 'sipp wrapper': 'SIPP Wrapper',
    'sipp bank acc': 'SIPP', 'sipp cash': 'SIPP', 'dfm (within sipp)': 'SIPP',
    'ssas': 'SSAS', 'ssas dfm': 'SSAS', 'ssas bank acc': 'SSAS',
    'pension': 'Personal Pension Plan', 'retirement account': 'Personal Pension Plan',
    'annuity': 'Pension Annuity', 'annuity: enhanced': 'Enhanced Pension Annuity',
    'isa': 'ISA (Stocks And Shares)', 'jisa': 'Junior ISA (Stocks And Shares)',
    'gia': 'General Investment Account', 'dfm': 'Discretionary Managed Service',
    'mps': 'Discretionary Managed Service', 'rathbones': 'Discretionary Managed Service',
    'offshore bond': 'Offshore Bond', 'bond': 'Investment Bond', 'investment bond': 'Investment Bond',
    'flexible access drawdown': 'Income Drawdown (Flexi-Access)',
    'flexi-access drawdown': 'Income Drawdown (Flexi-Access)',
    'income drawdown': 'Income Drawdown (Flexi-Access)',
}

def sub_type_from_adv(adv_type):
    return PM_SUB.get(str(adv_type).lower().strip(), '')

# ── rebuild provider families ─────────────────────────────────────────────────
alias_counts = collections.defaultdict(lambda: collections.defaultdict(int))
for _, row in mdf.iterrows():
    if row.get('plan_match_method') == 'policy_number' and row.get('vision_policy_uuid'):
        adv_p = _pnorm(row.get('adv_provider', ''))
        vis_p = _pnorm(vis_policies.get(row['vision_policy_uuid'], {}).get('primary_provider', ''))
        if adv_p and vis_p:
            alias_counts[adv_p][vis_p] += 1

provider_families = {}
for adv_p, vis_counts in alias_counts.items():
    aliases = {vp for vp, cnt in vis_counts.items() if cnt >= 2}
    provider_families[adv_p] = aliases if aliases else {adv_p}

def get_prov_family(adv_prov):
    adv_pn = _pnorm(adv_prov)
    return provider_families.get(adv_pn, {adv_pn})

# ── Rule 6 — dedup: multiple adviser rows matching same Vision plan ────────────
puuid_to_matched_rows = collections.defaultdict(list)
for idx, row in mdf.iterrows():
    if row['row_status'] == 'Matched' and row['vision_policy_uuid']:
        puuid_to_matched_rows[row['vision_policy_uuid']].append(idx)

dedup_splits = 0
for puuid, idxs in puuid_to_matched_rows.items():
    if len(idxs) > 1:
        for extra_idx in idxs[1:]:
            mdf.at[extra_idx, 'row_status'] = 'Adviser-only'
            mdf.at[extra_idx, 'plan_match_status'] = ''
            mdf.at[extra_idx, 'plan_match_method'] = 'dedup_split'
            mdf.at[extra_idx, 'vision_policy_uuid'] = ''
            mdf.at[extra_idx, 'vision_policy_numbers'] = ''
            mdf.at[extra_idx, 'vision_product'] = ''
            mdf.at[extra_idx, 'vision_provider'] = ''
            mdf.at[extra_idx, 'vision_under_advice'] = ''
            mdf.at[extra_idx, 'vision_fee_pct'] = ''
            dedup_splits += 1
print(f"Rule 6 - Dedup splits: {dedup_splits}")

# ── Rule 4 — resolve owner+family ambiguous rows ──────────────────────────────
already_matched = set(mdf[mdf['row_status'] == 'Matched']['vision_policy_uuid'].tolist())
already_matched.discard('')

ambig_family = mdf[(mdf['row_status'] == 'Ambiguous') &
                   (mdf['plan_match_method'].str.contains('owner_family', na=False))].copy()

rule4_matched_subtype = 0
rule4_matched_value = 0
rule4_new_plans = 0
rule4_possible_dups = 0

# Group by (person_match_uuid, provider_family_frozenset, broad_type)
BROAD_PM = {
    'sipp': 'pension', 'ssas': 'pension', 'pension': 'pension', 'drawdown': 'pension',
    'rac': 'pension', 'fsavc': 'pension', 'stakeholder': 'pension', 'avc': 'pension',
    'retirement': 'pension', 'annuity': 'annuity',
    'isa': 'investment', 'bond': 'investment', 'investment': 'investment',
    'trust': 'investment', 'oeic': 'investment', 'discretionary': 'investment',
    'vct': 'investment', 'eis': 'investment', 'iht': 'investment',
    'cash': 'investment', 'liquidity': 'investment', 'wrap': 'investment',
    'offshore': 'investment', 'portfolio': 'investment', 'rathbone': 'investment',
    'collective': 'investment', 'aim': 'investment', 'general investment': 'investment',
    'protection': 'protection', 'life': 'protection', 'critical': 'protection',
    'income protect': 'protection', 'whole of life': 'protection', 'group': 'protection',
    'medical insurance': 'general_insurance',
}

def broad_type(s):
    sl = str(s).lower()
    for k, v in BROAD_PM.items():
        if k in sl:
            return v
    return 'other'

groups = collections.defaultdict(list)
for idx, row in ambig_family.iterrows():
    person_uuid = row.get('person_match_uuid', '').strip()
    adv_prov = row.get('adv_provider', '').strip()
    adv_broad = broad_type(row.get('adv_type', ''))
    fam = frozenset(get_prov_family(adv_prov))
    groups[(person_uuid, fam, adv_broad)].append(idx)

def get_vision_candidates(group_idxs):
    cands = set()
    for idx in group_idxs:
        for c in str(mdf.at[idx, 'candidates']).split('|'):
            c = c.strip()
            if c and c != 'nan' and not c.startswith('POSSIBLE_DUP:'):
                cands.add(c)
    return [c for c in cands if c not in already_matched and c in vis_policies]

for (person_uuid, fam, adv_broad), idxs in groups.items():
    vision_cands = get_vision_candidates(idxs)
    if not vision_cands:
        for idx in idxs:
            mdf.at[idx, 'row_status'] = 'Adviser-only'
            mdf.at[idx, 'plan_match_status'] = ''
            mdf.at[idx, 'plan_match_method'] = 'owner_family_no_cand'
            mdf.at[idx, 'vision_policy_uuid'] = ''
            rule4_new_plans += 1
        continue

    adv_rows_list = [(idx, mdf.loc[idx]) for idx in idxs]
    paired = {}  # adv_idx → vis_uuid (sub_type matched)
    vis_paired = set()

    for adv_idx, adv_row in adv_rows_list:
        adv_sub = sub_type_from_adv(adv_row.get('adv_type', ''))
        adv_sub_lower = adv_sub.lower()
        for vc in vision_cands:
            if vc in vis_paired or adv_idx in paired:
                continue
            vis_prod = vis_policies.get(vc, {}).get('product', '').lower()
            if adv_sub_lower and (adv_sub_lower in vis_prod or vis_prod in adv_sub_lower):
                paired[adv_idx] = vc
                vis_paired.add(vc)
                break

    unpaired_adv = [idx for idx in idxs if idx not in paired]
    unpaired_vis = [vc for vc in vision_cands if vc not in vis_paired]

    def get_adv_val(idx):
        try: return float(mdf.at[idx, 'adv_value'])
        except: return 0.0

    def get_vis_val(puuid):
        try: return float(vis_policies.get(puuid, {}).get('contribution', 0))
        except: return 0.0

    value_pairs = list(zip(
        sorted(unpaired_adv, key=get_adv_val, reverse=True),
        sorted(unpaired_vis, key=get_vis_val, reverse=True)
    ))

    for adv_idx, vis_uuid in value_pairs:
        paired[adv_idx] = vis_uuid
        vis_paired.add(vis_uuid)

    for adv_idx, vis_uuid in paired.items():
        method = 'owner_family_subtype' if adv_idx not in [i for i, _ in value_pairs] else 'owner_family_value_rank'
        mdf.at[adv_idx, 'row_status'] = 'Matched'
        mdf.at[adv_idx, 'plan_match_status'] = 'Matched'
        mdf.at[adv_idx, 'plan_match_method'] = method
        mdf.at[adv_idx, 'vision_policy_uuid'] = vis_uuid
        pol = vis_policies.get(vis_uuid, {})
        mdf.at[adv_idx, 'vision_policy_numbers'] = pol.get('policy_numbers', '')
        mdf.at[adv_idx, 'vision_product'] = pol.get('product', '')
        mdf.at[adv_idx, 'vision_provider'] = pol.get('primary_provider', '')
        mdf.at[adv_idx, 'vision_under_advice'] = pol.get('under_my_advice', '')
        mdf.at[adv_idx, 'vision_fee_pct'] = pol.get('ongoing_service_charge_percentage', '')
        already_matched.add(vis_uuid)
        if method == 'owner_family_subtype':
            rule4_matched_subtype += 1
        else:
            rule4_matched_value += 1

    remaining_adv = [idx for idx in idxs if idx not in paired]
    for idx in remaining_adv:
        mdf.at[idx, 'row_status'] = 'Adviser-only'
        mdf.at[idx, 'plan_match_status'] = ''
        mdf.at[idx, 'plan_match_method'] = 'owner_family_unmatched'
        mdf.at[idx, 'vision_policy_uuid'] = ''
        rule4_new_plans += 1

    remaining_vis = [vc for vc in vision_cands if vc not in vis_paired]
    rule4_possible_dups += len(remaining_vis)
    if remaining_vis and idxs:
        first_idx = idxs[0]
        existing = mdf.at[first_idx, 'candidates']
        mdf.at[first_idx, 'candidates'] = (existing + '|' if existing else '') + 'POSSIBLE_DUP:' + '|'.join(remaining_vis)

print(f"Rule 4 - owner+family: {rule4_matched_subtype} matched by sub_type, "
      f"{rule4_matched_value} matched by value rank, {rule4_new_plans} new plans, "
      f"{rule4_possible_dups} possible-dup Vision plans")

# ── Rule 7 — new-client rows (already labelled, just count) ───────────────────
new_client_count = len(mdf[mdf['row_status'] == 'New-client'])
print(f"Rule 7 - New-client rows: {new_client_count}")

# ── Rule 8 — unwired adviser-only rows ────────────────────────────────────────
rule8_biz = 0
rule8_still_ambig = 0
for idx, row in mdf.iterrows():
    if row['row_status'] != 'Adviser-only':
        continue
    puuid = str(row.get('person_match_uuid', '')).strip()
    if not puuid or puuid in all_client_uuids:
        continue
    # UUID in file but not in Vision clients — check if it's a business primary contact
    if puuid in biz_person_uuids:
        mdf.at[idx, 'plan_match_method'] = str(mdf.at[idx, 'plan_match_method']) + '_biz_keyed'
        rule8_biz += 1
    else:
        mdf.at[idx, 'row_status'] = 'Ambiguous'
        mdf.at[idx, 'plan_match_status'] = 'Ambiguous'
        rule8_still_ambig += 1

print(f"Rule 8 - unwired adviser-only: {rule8_biz} business-keyed, {rule8_still_ambig} still ambiguous")

# ── Final status ──────────────────────────────────────────────────────────────
final_counts = mdf['row_status'].value_counts()
print(f"\nFinal v4 row status:")
for k, v in final_counts.items():
    print(f"  {k:20s} {v:4d}")
print(f"  {'TOTAL':20s} {len(mdf):4d}")

still_ambig = mdf[mdf['row_status'] == 'Ambiguous']
print(f"\nRemaining ambiguous: {len(still_ambig)}")
print("  Breakdown by plan_match_method:")
print(still_ambig['plan_match_method'].value_counts().to_string())

# ── write v4 ─────────────────────────────────────────────────────────────────
os.makedirs(os.path.dirname(MATCHES_V4), exist_ok=True)
out_wb = Workbook()
out_ws = out_wb.active
out_ws.title = "Matches"
headers = list(mdf.columns)
out_ws.append(headers)
for _, row in mdf.iterrows():
    out_ws.append([str(v) if v is not None else '' for v in row.tolist()])
out_wb.save(MATCHES_V4)
print(f"\nv4 written: {MATCHES_V4} ({len(mdf)} rows)")
