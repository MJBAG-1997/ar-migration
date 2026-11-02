"""
Stage 1 matching v2 — Higgs Financial Solutions Ltd
Rule 9: no client names in stdout.
"""
import re, os, sys, datetime, collections
import pandas as pd
import openpyxl
from openpyxl import Workbook

# ── paths ────────────────────────────────────────────────────────────────────
DATA_ROOT    = r"C:\Users\markj\OneDrive - Beyond Advice\Technology - Documents\workspace\AR Migration\Higgs"
VISION_SRC   = os.path.join(DATA_ROOT, "AA Data")
ADV_XLSX     = os.path.join(DATA_ROOT, "Higgs Master.xlsx")
MATCHES_OUT  = os.path.join(DATA_ROOT, "stage1", "higgs_matches_v2.xlsx")
TODAY        = datetime.date(2026, 11, 2)

os.makedirs(os.path.dirname(MATCHES_OUT), exist_ok=True)

# ── normalisation ─────────────────────────────────────────────────────────────
def norm_name(s):
    s = str(s).lower().strip()
    for t in ('mr','mrs','ms','miss','dr','prof','rev','sir','lady','lord'):
        s = re.sub(r'\b' + t + r'\b\.?', '', s)
    s = re.sub(r'[^a-z0-9 ]', ' ', s)
    return re.sub(r'\s+', ' ', s).strip()

def norm_entity(s):
    s = str(s).lower().strip()
    for w in ('limited','ltd','llp','plc','trust','the','llc','inc'):
        s = re.sub(r'\b' + w + r'\b\.?', '', s)
    s = re.sub(r'[^a-z0-9 ]', ' ', s)
    return re.sub(r'\s+', ' ', s).strip()

def norm_pn(s):
    s = str(s).strip().upper()
    if s.endswith('.0'):
        s = s[:-2]
    return re.sub(r'[\s\.\-/]', '', s)

def extract_postcode(addr):
    m = re.search(r'\b([A-Z]{1,2}[0-9][0-9A-Z]?\s*[0-9][A-Z]{2})\b', str(addr).upper())
    return m.group(1).replace(' ', '').upper() if m else ''

def pn_shape(pn):
    return re.sub(r'[A-Z]', 'A', re.sub(r'[0-9]', '9', pn))

def _pnorm(x):
    x = str(x).lower().strip()
    x = x.replace('&', ' and ')
    x = re.sub(r'\b(ltd|limited|plc|llp|uk|group)\b', '', x)
    x = re.sub(r'[^a-z0-9 ]', ' ', x)
    return re.sub(r'\s+', ' ', x).strip()

def calc_age(dob_dt, ref_date):
    if dob_dt is None or pd.isna(dob_dt):
        return None
    try:
        d = dob_dt.date() if hasattr(dob_dt, 'date') else dob_dt
        return (ref_date - d).days // 365
    except Exception:
        return None

# ── read Vision CSVs ──────────────────────────────────────────────────────────
def read_csv(fname):
    path = os.path.join(VISION_SRC, fname)
    try:
        return pd.read_csv(path, dtype=str, keep_default_na=False, encoding='utf-8-sig')
    except UnicodeDecodeError:
        return pd.read_csv(path, dtype=str, keep_default_na=False, encoding='latin-1')

p01 = read_csv("Phil Thompson_01. Person-Clients.csv")
p02 = read_csv("Phil Thompson_02. Business-Clients.csv")
p03 = read_csv("Phil Thompson_03. Trust-Clients.csv")
p04 = read_csv("Phil Thompson_04. Policies.csv")
p12 = read_csv("Phil Thompson_12. Address History.csv")
print(f"Vision loaded: {len(p01)} persons, {len(p02)} businesses, {len(p03)} trusts, {len(p04)} policies, {len(p12)} addresses")

# ── Phase 2a: Vision indexes ──────────────────────────────────────────────────
all_client_uuids = set(p01['client_uuid'].tolist()) | set(p02['primary_contact_uuid'].tolist()) | set(p03['client_uuid'].tolist())

vision_persons = {}
for _, r in p01.iterrows():
    uid = r['client_uuid'].strip()
    fore = r['forename'].strip()
    sur  = r['surname'].strip()
    ni   = re.sub(r'\s','', r['national_insurance_number'].strip()).upper()
    dob_raw = r['date_of_birth'].strip()
    dob_dt = pd.to_datetime(dob_raw, errors='coerce')
    vision_persons[uid] = {
        'forename': fore, 'surname': sur, 'ni': ni,
        'dob_dt': dob_dt,
        'age_today': calc_age(dob_dt, TODAY),
    }

# surname+forename → [uuid]
vpSF = collections.defaultdict(list)
vpSI = collections.defaultdict(list)   # surname+initial
vpSA = collections.defaultdict(list)   # surname+age_today

for uid, v in vision_persons.items():
    sf_key = (norm_name(v['surname']), norm_name(v['forename']))
    vpSF[sf_key].append(uid)
    ini = norm_name(v['forename'])[:1]
    if ini:
        vpSI[(norm_name(v['surname']), ini)].append(uid)
    if v['age_today'] is not None:
        vpSA[(norm_name(v['surname']), v['age_today'])].append(uid)

# current postcode per client (latest current address)
vision_current_postcode = {}
addr_rows = []
for _, r in p12.iterrows():
    uid = r['uuid'].strip()
    st  = r['status'].strip().lower()
    pc  = r['postcode'].strip().upper().replace(' ','')
    md  = r['moved_in_date'].strip()
    addr_rows.append((uid, st, pc, md))

addr_rows.sort(key=lambda x: x[3])
for uid, st, pc, md in addr_rows:
    if st == 'current' and pc:
        vision_current_postcode[uid] = pc

# surname+postcode → [uuid]
vpSPC = collections.defaultdict(list)
for uid, pc in vision_current_postcode.items():
    if uid in vision_persons and pc:
        vpSPC[(norm_name(vision_persons[uid]['surname']), pc)].append(uid)

# postcode → set of uuids
clients_at_postcode = collections.defaultdict(set)
for uid, pc in vision_current_postcode.items():
    if pc:
        clients_at_postcode[pc].add(uid)

# businesses and trusts
vision_businesses = {}
for _, r in p02.iterrows():
    bn = norm_entity(r['business_name'].strip())
    if bn:
        vision_businesses[bn] = r['primary_contact_uuid'].strip()

vision_trusts = {}
for _, r in p03.iterrows():
    tn = norm_entity(r['client_name'].strip())
    if tn:
        vision_trusts[tn] = r['client_uuid'].strip()

# plan indexes
vision_policies = {}
vision_pn_index = collections.defaultdict(list)   # norm_pn → [policy_uuid]
vision_pn_nozero = collections.defaultdict(list)  # nozero → [policy_uuid]

for _, r in p04.iterrows():
    puid = r['policy_uuid'].strip()
    vision_policies[puid] = r.to_dict()
    pns_raw = r['policy_numbers'].strip()
    if pns_raw:
        for part in pns_raw.split('|'):
            npn = norm_pn(part.strip())
            if npn:
                vision_pn_index[npn].append(puid)
                nz = npn.lstrip('0')
                if len(nz) >= 3 and nz != npn:
                    vision_pn_nozero[nz].append(puid)

print(f"Vision indexes: {len(vision_persons)} persons, {len(vision_pn_index)} policy-number keys")

# ── Phase 1: Load adviser xlsx ────────────────────────────────────────────────
wb_adv = openpyxl.load_workbook(ADV_XLSX, data_only=True)
ws_adv = wb_adv['Sheet1']

def cell_val(ws, row, col):
    v = ws.cell(row, col).value
    if v is None:
        return None
    if isinstance(v, float) and v == int(v):
        # check if it's an integer stored as float
        pass
    return v

# read raw rows
raw_rows = []
for rn in range(2, ws_adv.max_row + 1):
    vals = [ws_adv.cell(rn, c).value for c in range(1, 30)]
    if all(v is None or (isinstance(v, str) and not v.strip()) for v in vals):
        continue
    raw_rows.append({'_row': rn, '_vals': vals})

print(f"Adviser xlsx: {len(raw_rows)} non-blank rows")

# col indices (0-based in _vals list): A=0,B=1,...,L=11,M=12,N=13,P=15,W=22,X=23
CI = {
    'A':0,'B':1,'C':2,'D':3,'E':4,'F':5,'G':6,
    'H':7,'I':8,'J':9,'K':10,'L':11,'M':12,'N':13,
    'O':14,'P':15,'Q':16,'R':17,'S':18,'T':19,'U':20,'V':21,
    'W':22,'X':23,
}

def gv(row, col_letter):
    return row['_vals'][CI[col_letter]]

# ── detect groups ─────────────────────────────────────────────────────────────
# group starts when col L is non-null
groups = []
cur_group = []
for row in raw_rows:
    l_val = gv(row, 'L')
    if l_val is not None and l_val != '':
        if cur_group:
            groups.append(cur_group)
        cur_group = [row]
    else:
        cur_group.append(row)
if cur_group:
    groups.append(cur_group)

print(f"Adviser groups (households): {len(groups)}")

# ── fill down A,B,E,G within groups, extract fields ──────────────────────────
processed = []
for gid, grp in enumerate(groups, 1):
    # fill-down values: first non-null within group
    fill_A = next((gv(r,'A') for r in grp if gv(r,'A') is not None), None)
    fill_B = next((gv(r,'B') for r in grp if gv(r,'B') is not None), None)
    fill_E = next((gv(r,'E') for r in grp if gv(r,'E') is not None), None)
    fill_G = next((gv(r,'G') for r in grp if gv(r,'G') is not None), None)

    for row in grp:
        # extract fields
        sur = str(gv(row,'C') or '').strip()
        fore = str(gv(row,'D') or '').strip()
        addr = str(fill_A or '').strip()
        pc = extract_postcode(addr)

        age_raw = str(gv(row,'F') or '').strip()
        try:
            adv_age = int(float(age_raw)) if age_raw else None
        except (ValueError, TypeError):
            adv_age = None

        pn_raw = gv(row,'I')
        if isinstance(pn_raw, float):
            pn_raw = str(int(pn_raw))
        elif pn_raw is not None:
            pn_raw = str(pn_raw).strip()
        else:
            pn_raw = ''
        pn_norm = norm_pn(pn_raw) if pn_raw else ''

        m_val = gv(row,'M')
        ar_fee = float(m_val)*100 if isinstance(m_val,(int,float)) else None

        n_val = gv(row,'N')
        ar_fee_incl = float(n_val)*100 if isinstance(n_val,(int,float)) else None

        q_val = gv(row,'Q')
        ar_vip = float(q_val)*100 if isinstance(q_val,(int,float)) else None

        t_val = gv(row,'T')
        ar_prov_fee = float(t_val)*100 if isinstance(t_val,(int,float)) else None

        k_val = gv(row,'K')
        cur_val = float(k_val) if isinstance(k_val,(int,float)) else None

        w_val = gv(row,'W')
        if hasattr(w_val,'date'):
            last_updated = w_val.date()
        elif isinstance(w_val, datetime.datetime):
            last_updated = w_val.date()
        elif isinstance(w_val, datetime.date):
            last_updated = w_val
        else:
            last_updated = None

        processed.append({
            '_row': row['_row'],
            'group_id': gid,
            'couple_single': str(fill_E or '').strip(),
            'surname': sur,
            'forename': fore,
            'address': addr,
            'adv_postcode': pc,
            'adv_age': adv_age,
            'email': str(fill_G or '').strip(),
            'provider': str(gv(row,'H') or '').strip(),
            'pn_raw': pn_raw,
            'pn_norm': pn_norm,
            'adv_type': str(gv(row,'J') or '').strip(),
            'cur_val': cur_val,
            'ar_fee': ar_fee,
            'ar_fee_incl': ar_fee_incl,
            'ar_vip': ar_vip,
            'ar_prov_fee': ar_prov_fee,
            'frequency': str(gv(row,'P') or '').strip(),
            'last_updated': last_updated,
            'legacy_label': str(gv(row,'X') or '').strip(),
            # match results (filled later)
            'person_match_status': '', 'person_match_method': '', 'person_match_uuid': '',
            'plan_match_status': '', 'plan_match_method': '',
            'vision_policy_uuid': '', 'vision_policy_numbers': '', 'vision_product': '',
            'vision_provider': '', 'vision_under_advice': '', 'vision_fee_pct': '',
            'row_status': '', 'candidates': '',
        })

print(f"Processed rows: {len(processed)}")

# ── Phase 2b index: entity names ─────────────────────────────────────────────
# (used later per row)

# ── Phase 3: policy number matching ──────────────────────────────────────────
BLANK_PNS = {'', 'UNKNOWN', 'TBC', 'N/A', '-', '?', 'NONE'}

for row in processed:
    pn = row['pn_norm']
    if not pn or pn in BLANK_PNS:
        continue
    # exact
    matches = vision_pn_index.get(pn, [])
    if len(matches) == 1:
        row['plan_match_status'] = 'Matched'
        row['plan_match_method'] = 'policy_number'
        row['vision_policy_uuid'] = matches[0]
        pol = vision_policies[matches[0]]
        row['vision_policy_numbers'] = pol.get('policy_numbers','')
        row['vision_product'] = pol.get('product','')
        row['vision_provider'] = pol.get('primary_provider','')
        row['vision_under_advice'] = pol.get('under_my_advice','')
        row['vision_fee_pct'] = pol.get('ongoing_service_charge_percentage','')
        continue
    if len(matches) > 1:
        row['plan_match_status'] = 'Ambiguous'
        row['plan_match_method'] = 'policy_number'
        row['candidates'] = '|'.join(matches)
        continue
    # no-leading-zeros
    stripped = pn.lstrip('0')
    if len(stripped) >= 3 and stripped != pn:
        nz_matches = vision_pn_nozero.get(stripped, [])
        if len(nz_matches) == 1:
            row['plan_match_status'] = 'Matched'
            row['plan_match_method'] = 'policy_number_nozero'
            row['vision_policy_uuid'] = nz_matches[0]
            pol = vision_policies[nz_matches[0]]
            row['vision_policy_numbers'] = pol.get('policy_numbers','')
            row['vision_product'] = pol.get('product','')
            row['vision_provider'] = pol.get('primary_provider','')
            row['vision_under_advice'] = pol.get('under_my_advice','')
            row['vision_fee_pct'] = pol.get('ongoing_service_charge_percentage','')
            continue
        if len(nz_matches) > 1:
            row['plan_match_status'] = 'Ambiguous'
            row['plan_match_method'] = 'policy_number_nozero'
            row['candidates'] = '|'.join(nz_matches)
            continue
    # containment
    if len(pn) >= 6:
        cont_matches = []
        for vpn_str, puids in vision_pn_index.items():
            if len(vpn_str) < 6:
                continue
            shorter = pn if len(pn) <= len(vpn_str) else vpn_str
            longer  = vpn_str if len(pn) <= len(vpn_str) else pn
            if len(shorter) >= 6 and shorter in longer:
                cont_matches.extend(puids)
        cont_matches = list(dict.fromkeys(cont_matches))  # unique
        if len(cont_matches) == 1:
            row['plan_match_status'] = 'Matched'
            row['plan_match_method'] = 'policy_number_containment'
            row['vision_policy_uuid'] = cont_matches[0]
            pol = vision_policies[cont_matches[0]]
            row['vision_policy_numbers'] = pol.get('policy_numbers','')
            row['vision_product'] = pol.get('product','')
            row['vision_provider'] = pol.get('primary_provider','')
            row['vision_under_advice'] = pol.get('under_my_advice','')
            row['vision_fee_pct'] = pol.get('ongoing_service_charge_percentage','')
        elif len(cont_matches) > 1:
            row['plan_match_status'] = 'Ambiguous'
            row['plan_match_method'] = 'policy_number_containment'
            row['candidates'] = '|'.join(cont_matches[:5])

# ── Phase 4: Person matching per group ───────────────────────────────────────

# pre-compute: for "unique on both sides" check, count distinct person keys across all adviser rows
all_adv_person_keys = collections.Counter()
for row in processed:
    if row['forename']:
        all_adv_person_keys[(norm_name(row['surname']), norm_name(row['forename']))] += 1

# group rows by group_id
from itertools import groupby
group_rows = {}
for row in processed:
    group_rows.setdefault(row['group_id'], []).append(row)

person_match_results = {}   # (group_id, norm_surname, norm_forename) → {status, method, uuid, candidates}

for gid, grp in group_rows.items():
    # household anchor: any policy-number match → anchor uuids
    anchor_uuids = set()
    for row in grp:
        if row['plan_match_status'] == 'Matched' and row['vision_policy_uuid']:
            pol = vision_policies.get(row['vision_policy_uuid'], {})
            for fld in ('client_1_uuid','client_2_uuid'):
                uid = pol.get(fld,'').strip()
                if uid and uid in all_client_uuids:
                    anchor_uuids.add(uid)

    # extend anchor by postcode
    anchor_postcodes = {vision_current_postcode.get(uid,'') for uid in anchor_uuids if vision_current_postcode.get(uid,'')}
    household_candidates = set(anchor_uuids)
    for apc in anchor_postcodes:
        if apc:
            household_candidates |= clients_at_postcode.get(apc, set())

    # distinct persons in this group
    distinct_persons = {}   # (norm_sur, norm_fore) → {'sur_raw', 'fore_raw', 'age', 'postcode', 'ref_date'}
    for row in grp:
        sur_raw = row['surname']
        fore_raw = row['forename']
        if not sur_raw:
            continue
        ns = norm_name(sur_raw)
        nf = norm_name(fore_raw) if fore_raw else ''
        key = (ns, nf)
        if key not in distinct_persons:
            ref_date = row['last_updated'] or TODAY
            distinct_persons[key] = {
                'sur_raw': sur_raw, 'fore_raw': fore_raw,
                'ns': ns, 'nf': nf,
                'age': row['adv_age'],
                'postcode': row['adv_postcode'],
                'ref_date': ref_date,
            }

    # match each distinct person
    for key, p in distinct_persons.items():
        ns, nf = key
        result_key = (gid, ns, nf)
        if result_key in person_match_results:
            continue

        # entity match (no forename)
        if not nf:
            ent_key = norm_entity(p['sur_raw'])
            biz_match = vision_businesses.get(ent_key)
            trust_match = vision_trusts.get(ent_key)
            if biz_match or trust_match:
                uid = biz_match or trust_match
                person_match_results[result_key] = {'status':'Entity','method':'entity_name','uuid':uid,'candidates':''}
            else:
                person_match_results[result_key] = {'status':'New','method':'','uuid':'','candidates':''}
            continue

        # rung 2: surname+forename+postcode
        pc = p['postcode']
        matched = None
        method = ''
        candidates_list = []

        if pc:
            cands = vpSPC.get((ns, pc), [])
            # filter by forename
            cands_fore = [u for u in cands if norm_name(vision_persons[u]['forename']) == nf]
            if not cands_fore:
                # try initial match
                ini = nf[:1]
                cands_fore = [u for u in cands if norm_name(vision_persons[u]['forename'])[:1] == ini]
            if household_candidates:
                hc_filtered = [u for u in cands_fore if u in household_candidates]
                if hc_filtered:
                    cands_fore = hc_filtered
            if len(cands_fore) == 1:
                matched = cands_fore[0]; method = 'name_postcode'
            elif len(cands_fore) > 1:
                candidates_list = cands_fore

        # rung 3: surname+forename+age
        if matched is None and not candidates_list:
            ref_date = p['ref_date']
            adv_age = p['age']
            if adv_age is not None:
                cands_sf = vpSF.get((ns, nf), [])
                age_cands = []
                for u in cands_sf:
                    va = calc_age(vision_persons[u]['dob_dt'], ref_date)
                    if va is not None and abs(va - adv_age) <= 1:
                        age_cands.append(u)
                if household_candidates:
                    hc_f = [u for u in age_cands if u in household_candidates]
                    if hc_f:
                        age_cands = hc_f
                if len(age_cands) == 1:
                    matched = age_cands[0]; method = 'name_age'
                elif len(age_cands) > 1:
                    candidates_list = age_cands

        # rung 4: surname+initial+age
        if matched is None and not candidates_list:
            ini = nf[:1]
            adv_age = p['age']
            if ini and adv_age is not None:
                cands_si = vpSI.get((ns, ini), [])
                age_cands2 = []
                for u in cands_si:
                    va = calc_age(vision_persons[u]['dob_dt'], p['ref_date'])
                    if va is not None and abs(va - adv_age) <= 1:
                        age_cands2.append(u)
                if household_candidates:
                    hc_f = [u for u in age_cands2 if u in household_candidates]
                    if hc_f:
                        age_cands2 = hc_f
                if len(age_cands2) == 1:
                    matched = age_cands2[0]; method = 'initial_age'
                elif len(age_cands2) > 1:
                    candidates_list = age_cands2

        # rung 5: surname+forename alone if unique both sides
        if matched is None and not candidates_list:
            vis_cands = vpSF.get((ns, nf), [])
            adv_count = all_adv_person_keys.get((ns, nf), 0)
            if len(vis_cands) == 1 and adv_count == 1:
                matched = vis_cands[0]; method = 'name_only'
            elif len(vis_cands) > 1:
                candidates_list = vis_cands

        # rung 6: near-miss (→ Ambiguous, not New)
        near_miss_cands = []
        if matched is None:
            ini = nf[:1]
            # same surname + same initial
            if ini:
                nm_cands = vpSI.get((ns, ini), [])
                if nm_cands:
                    near_miss_cands.extend(nm_cands)
            # same surname + same postcode
            if pc:
                nm_pc = vpSPC.get((ns, pc), [])
                near_miss_cands.extend(nm_pc)
            near_miss_cands = list(dict.fromkeys(near_miss_cands))

        # rung 7: household anchor fallback
        if matched is None and not candidates_list and not near_miss_cands and household_candidates:
            ini = nf[:1]
            hc_fore = [u for u in household_candidates
                       if u in vision_persons and
                       (norm_name(vision_persons[u]['forename']) == nf or
                        (ini and norm_name(vision_persons[u]['forename'])[:1] == ini)) and
                       norm_name(vision_persons[u]['surname']) == ns]
            if len(hc_fore) == 1:
                matched = hc_fore[0]; method = 'household_anchor'
            elif len(hc_fore) > 1:
                candidates_list = hc_fore

        if matched:
            person_match_results[result_key] = {'status':'Matched','method':method,'uuid':matched,'candidates':''}
        elif candidates_list:
            person_match_results[result_key] = {'status':'Ambiguous','method':method or 'multi_candidate','uuid':'','candidates':'|'.join(candidates_list[:5])}
        elif near_miss_cands:
            person_match_results[result_key] = {'status':'Ambiguous','method':'near_miss','uuid':'','candidates':'|'.join(near_miss_cands[:5])}
        else:
            person_match_results[result_key] = {'status':'New','method':'','uuid':'','candidates':''}

# ── assign person results to rows ─────────────────────────────────────────────
for row in processed:
    ns = norm_name(row['surname'])
    nf = norm_name(row['forename']) if row['forename'] else ''
    res = person_match_results.get((row['group_id'], ns, nf))
    if res:
        row['person_match_status'] = res['status']
        row['person_match_method'] = res['method']
        row['person_match_uuid'] = res['uuid']
        if res['candidates'] and not row['candidates']:
            row['candidates'] = res['candidates']
    else:
        # entity or no surname
        if not row['surname']:
            row['person_match_status'] = 'New'

# ── Phase 5: row_status ───────────────────────────────────────────────────────
for row in processed:
    ps = row['person_match_status']
    pls = row['plan_match_status']
    if pls == 'Matched':
        row['row_status'] = 'Matched'
    elif pls == 'Ambiguous':
        row['row_status'] = 'Ambiguous'
    elif ps == 'Ambiguous':
        row['row_status'] = 'Ambiguous'
    elif ps == 'Matched' or ps == 'Entity':
        row['row_status'] = 'Adviser-only'
    elif ps == 'New':
        row['row_status'] = 'New-client'
    else:
        row['row_status'] = 'New-client'

# ── Phase 6: statistics ───────────────────────────────────────────────────────
print()
print("=" * 70)
print("MATCHING RESULTS")
print("=" * 70)

total_rows = len(processed)
total_groups = len(groups)
# distinct persons (across all groups)
all_distinct = set()
for row in processed:
    if row['surname']:
        all_distinct.add((norm_name(row['surname']), norm_name(row['forename']) if row['forename'] else ''))
print(f"Total adviser rows:       {total_rows}")
print(f"Total household groups:   {total_groups}")
print(f"Distinct person keys:     {len(all_distinct)}")

# row status
status_ct = collections.Counter(r['row_status'] for r in processed)
print()
print("Row status:")
for k in ('Matched','Adviser-only','Ambiguous','New-client'):
    print(f"  {k:15s} {status_ct[k]:4d}")
print(f"  {'TOTAL':15s} {total_rows:4d}")

# plan match method
plan_method_ct = collections.Counter(r['plan_match_method'] for r in processed if r['plan_match_status']=='Matched')
print()
print("Plan match methods (Matched rows):")
for k,v in plan_method_ct.most_common():
    print(f"  {k:35s} {v:4d}")
plan_amb = sum(1 for r in processed if r['plan_match_status']=='Ambiguous')
plan_none = sum(1 for r in processed if not r['plan_match_status'] or r['plan_match_status']=='')
print(f"  {'Ambiguous plan':35s} {plan_amb:4d}")
print(f"  {'No plan match':35s} {plan_none:4d}")

# person match method
print()
print("Person match methods:")
pm_ct = collections.Counter(
    (r['person_match_status'], r['person_match_method'])
    for r in processed
    # deduplicate by group+person key
)
# better: per distinct person
distinct_person_results = {}
for row in processed:
    ns = norm_name(row['surname'])
    nf = norm_name(row['forename']) if row['forename'] else ''
    key = (row['group_id'], ns, nf)
    if key not in distinct_person_results:
        distinct_person_results[key] = (row['person_match_status'], row['person_match_method'])

pm_distinct = collections.Counter(distinct_person_results.values())
for (st, meth), cnt in sorted(pm_distinct.items(), key=lambda x: -x[1]):
    print(f"  {st:10s} {meth or '(none)':30s} {cnt:4d}")

# Vision plans NOT on adviser sheet
print()
print("Vision plans NOT on adviser sheet:")
# Plans on his sheet = those with a Matched plan
matched_vision_puuids = set(r['vision_policy_uuid'] for r in processed if r['vision_policy_uuid'])
all_vision_puuids = set(p04['policy_uuid'].tolist())
off_sheet = all_vision_puuids - matched_vision_puuids
print(f"  Total Vision policies: {len(all_vision_puuids)}")
print(f"  Matched (on sheet):    {len(matched_vision_puuids)}")
print(f"  Off sheet:             {len(off_sheet)}")

off_under = sum(1 for pu in off_sheet
                if vision_policies[pu].get('under_my_advice','').lower() in ('yes','servicing_agent','information_authority'))
off_fee = sum(1 for pu in off_sheet
              if vision_policies[pu].get('ongoing_service_charge_percentage','').strip() not in ('','0','0.0','0.00'))
print(f"  Off sheet & under_advice (yes equiv): {off_under}")
print(f"  Off sheet & has_fee:   {off_fee}")

# Plans on his list with no fee
no_fee_rows = [r for r in processed
               if (r['ar_fee'] is None or r['ar_fee'] == 0)
               and r['vision_fee_pct'] in ('','0','0.0','0.00','None')]
print()
print(f"Plans on his list with no fee (either source): {len(no_fee_rows)}")

# ── Provider top-15 ───────────────────────────────────────────────────────────
print()
print("Provider top-15 (by adviser rows, _pnorm key):")

# group adviser rows by _pnorm(provider)
adv_by_prov = collections.defaultdict(list)
for row in processed:
    adv_by_prov[_pnorm(row['provider'])].append(row)

# group Vision rows by _pnorm(primary_provider)
vis_by_prov = collections.defaultdict(list)
for _, r in p04.iterrows():
    vis_by_prov[_pnorm(r['primary_provider'].strip())].append(r.to_dict())

# pick top 15 by adviser count
top_provs = sorted(adv_by_prov.items(), key=lambda x: -len(x[1]))[:15]

print(f"  {'Provider (_pnorm)':30s} {'Adv':>5} {'Vis':>5} {'MatchedByPN':>11} {'AdvShapes':40s} {'VisNoPN':>7}")
print(f"  {'-'*30} {'-'*5} {'-'*5} {'-'*11} {'-'*40} {'-'*7}")
for pnorm, adv_rows in top_provs:
    n_adv = len(adv_rows)
    vis_rows = vis_by_prov.get(pnorm, [])
    n_vis = len(vis_rows)
    n_matched = sum(1 for r in adv_rows if r['plan_match_method'].startswith('policy_number'))
    # policy number shapes
    adv_shapes = collections.Counter(pn_shape(r['pn_norm']) for r in adv_rows if r['pn_norm'])
    top_adv_shapes = ','.join(f"{s}×{c}" for s,c in adv_shapes.most_common(3))
    # Vision: blank policy number
    vis_no_pn = sum(1 for r in vis_rows if not r.get('policy_numbers','').strip())
    print(f"  {pnorm[:30]:30s} {n_adv:5d} {n_vis:5d} {n_matched:11d} {top_adv_shapes[:40]:40s} {vis_no_pn:7d}")

# ── Fee check ─────────────────────────────────────────────────────────────────
print()
print("Fee check (matched plans where Vision has non-zero fee and col M non-null):")
matched_rows = [r for r in processed if r['row_status'] == 'Matched']
fee_check_rows = []
for r in matched_rows:
    vfee_str = r['vision_fee_pct']
    try:
        vfee = float(vfee_str)
    except (ValueError, TypeError):
        continue
    if vfee == 0 or r['ar_fee'] is None:
        continue
    fee_check_rows.append({
        'vision_fee': vfee,
        'M': r['ar_fee'],
        'N': r['ar_fee_incl'],
        'T': r['ar_prov_fee'],
        'Q': r['ar_vip'],
    })

if fee_check_rows:
    m_match = sum(1 for fc in fee_check_rows if fc['M'] is not None and abs(fc['M'] - fc['vision_fee']) < 0.01)
    n_match = sum(1 for fc in fee_check_rows if fc['N'] is not None and abs(fc['N'] - fc['vision_fee']) < 0.01)
    t_match = sum(1 for fc in fee_check_rows if fc['T'] is not None and abs(fc['T'] - fc['vision_fee']) < 0.01)
    mq_match = sum(1 for fc in fee_check_rows
                   if fc['M'] is not None and fc['Q'] is not None
                   and abs(fc['M'] + fc['Q'] - fc['vision_fee']) < 0.01)
    no_match = sum(1 for fc in fee_check_rows
                   if (fc['M'] is None or abs(fc['M'] - fc['vision_fee']) >= 0.01)
                   and (fc['N'] is None or abs(fc['N'] - fc['vision_fee']) >= 0.01)
                   and (fc['T'] is None or abs(fc['T'] - fc['vision_fee']) >= 0.01)
                   and (fc['M'] is None or fc['Q'] is None or abs(fc['M'] + fc['Q'] - fc['vision_fee']) >= 0.01))
    total_fc = len(fee_check_rows)
    print(f"  Rows checked:           {total_fc}")
    print(f"  M == Vision fee:        {m_match} ({m_match*100//total_fc if total_fc else 0}%)")
    print(f"  N == Vision fee:        {n_match} ({n_match*100//total_fc if total_fc else 0}%)")
    print(f"  T == Vision fee:        {t_match} ({t_match*100//total_fc if total_fc else 0}%)")
    print(f"  M+Q == Vision fee:      {mq_match} ({mq_match*100//total_fc if total_fc else 0}%)")
    print(f"  No match at all:        {no_match}")
else:
    print("  No rows qualify for fee check")

# ── VAT analysis ──────────────────────────────────────────────────────────────
print()
print("VAT analysis (ratio N/M for rows where both M>0 and N>0):")
vat_rows = [r for r in processed if r['ar_fee'] and r['ar_fee'] > 0
            and r['ar_fee_incl'] and r['ar_fee_incl'] > 0]
no_vat = sum(1 for r in vat_rows if abs(r['ar_fee_incl']/r['ar_fee'] - 1.0) < 0.01)
vat_20 = sum(1 for r in vat_rows if abs(r['ar_fee_incl']/r['ar_fee'] - 1.2) < 0.01)
other_vat = [(r['ar_fee_incl']/r['ar_fee']) for r in vat_rows
             if abs(r['ar_fee_incl']/r['ar_fee'] - 1.0) >= 0.01
             and abs(r['ar_fee_incl']/r['ar_fee'] - 1.2) >= 0.01]
print(f"  Total rows:    {len(vat_rows)}")
print(f"  No VAT (~1.0): {no_vat}")
print(f"  20% VAT (~1.2):{vat_20}")
if other_vat:
    other_unique = sorted(set(round(x,4) for x in other_vat))
    print(f"  Other ratios:  {other_unique}")
else:
    print(f"  Other ratios:  none")

# ── Valuation comparison ──────────────────────────────────────────────────────
print()
print("Valuation comparison (matched plans: adviser K vs Vision contribution):")
val_comp = []
for r in matched_rows:
    if not r['vision_policy_uuid']:
        continue
    pol = vision_policies.get(r['vision_policy_uuid'], {})
    try:
        vc = float(pol.get('contribution', 0) or 0)
    except (ValueError, TypeError):
        vc = 0
    ak = r['cur_val']
    if vc and vc > 0 and ak and ak > 0:
        val_comp.append(ak / vc)

if val_comp:
    val_comp_sorted = sorted(val_comp)
    n = len(val_comp_sorted)
    median_r = val_comp_sorted[n//2]
    mean_r = sum(val_comp_sorted)/n
    within_10 = sum(1 for x in val_comp_sorted if 0.8 <= x <= 1.25)
    print(f"  Pairs compared:      {n}")
    print(f"  Median ratio:        {median_r:.3f}")
    print(f"  Mean ratio:          {mean_r:.3f}")
    print(f"  Within 0.8–1.25:     {within_10} ({within_10*100//n if n else 0}%)")
    if 0.8 <= median_r <= 1.25:
        print(f"  Verdict: Vision contribution likely current value — keep")
    else:
        print(f"  Verdict: Vision contribution not aligned — drop Vision contribution valuation rows")
else:
    print("  No comparable pairs found")

# ── Orphan owner analysis ─────────────────────────────────────────────────────
print()
print("Orphan owner analysis:")
orphan_plan_uuids = []
for _, r in p04.iterrows():
    c1 = r['client_1_uuid'].strip()
    if c1 and c1 not in all_client_uuids:
        orphan_plan_uuids.append(r['policy_uuid'].strip())

print(f"  Plans with orphan client_1_uuid:  {len(orphan_plan_uuids)}")
# How many of those are on his list (matched by number)
matched_by_pn = set(r['vision_policy_uuid'] for r in processed
                    if r['plan_match_method'].startswith('policy_number'))
orphan_on_sheet = sum(1 for pu in orphan_plan_uuids if pu in matched_by_pn)
print(f"  Of those, on adviser sheet:       {orphan_on_sheet}")

# ── Phase 7: write matches xlsx ───────────────────────────────────────────────
out_wb = Workbook()
out_ws = out_wb.active
out_ws.title = "Matches"

HEADERS = [
    'adv_row','group_id','couple_single',
    'adv_surname','adv_forename','adv_provider','adv_policy_number',
    'adv_type','adv_value','adv_fee_pct','adv_frequency','adv_legacy_label',
    'person_match_status','person_match_method','person_match_uuid',
    'plan_match_status','plan_match_method',
    'vision_policy_uuid','vision_policy_numbers','vision_product',
    'vision_provider','vision_under_advice','vision_fee_pct',
    'row_status','candidates','Decision'
]
out_ws.append(HEADERS)

for row in processed:
    def ts(v):
        if v is None:
            return ''
        s = str(v)
        if s.startswith('='):
            s = ' ' + s
        return s
    out_ws.append([
        ts(row['_row']), ts(row['group_id']), ts(row['couple_single']),
        ts(row['surname']), ts(row['forename']), ts(row['provider']), ts(row['pn_raw']),
        ts(row['adv_type']),
        ts(round(row['cur_val'],2) if row['cur_val'] else ''),
        ts(round(row['ar_fee'],4) if row['ar_fee'] is not None else ''),
        ts(row['frequency']), ts(row['legacy_label']),
        ts(row['person_match_status']), ts(row['person_match_method']), ts(row['person_match_uuid']),
        ts(row['plan_match_status']), ts(row['plan_match_method']),
        ts(row['vision_policy_uuid']), ts(row['vision_policy_numbers']), ts(row['vision_product']),
        ts(row['vision_provider']), ts(row['vision_under_advice']), ts(row['vision_fee_pct']),
        ts(row['row_status']), ts(row['candidates']), ''
    ])

out_wb.save(MATCHES_OUT)
print()
print(f"Matches written to: {MATCHES_OUT}")
print()
print("Done.")
