# HFP AR Migration — Handover Note
**Date:** 2026-10-13  
**Firm slug:** `hfp` (Hoban Financial Planning Ltd)  
**Session:** signoff pack build (no API calls; snapshot-only)

---

## Current files

| File | Description |
|---|---|
| `snapshot_20261012T090906Z/` | Deduplicated Plannr extract (plans, charges, accounts, account_details, reference). Do not re-run. |
| `hfp_signoff_pack_v6.xlsx` | Internal import pack (current version). Not for distribution. |
| `hfp_review_pack_v2.xlsx` | AR review pack sent to Jim Hoban. |
| `mappings_hfp.xlsx` | CWM→BA mappings: users, client_statuses, groups, tags, service_levels, providers, expenditure (185 rows). |
| `Hoban_Financial_Planning_Ltd_260922_Payment_Rates_113727875.xlsx` | Transact/Nucleus provider workbook. |
| `Jim list  - Rathbones.xlsx` | Rathbones provider statement (52 plans). |

---

## Decisions applied this session

| Decision | Detail |
|---|---|
| **Pence fix** | All Plannr `amount` fields are stored in pence. Divided by 100 everywhere (valuations, IAI, income, expenses). |
| **Placeholder policy numbers** | Values matching TBC, Unknown, N/A, blank-ish strings blanked in Plans tab; original text kept in note column. Plans of relevant types added to "No policy number" exceptions (active only). |
| **Dedup by UUID** | plans_dedup.json and charges_dedup.json already deduplicated; used as authoritative source. |
| **Groups** | All client accounts migrated to group `default`. |
| **D18 workbook charges** | Every plan matched in the Transact/Nucleus workbook with a positive rate gets one ongoing charge: start_date 2026-11-02, payment_origin=Provider, D15 RMAR/VAT rules. Replaces any Plannr charge on the same plan. |
| **69 zero-rate workbook plans** | Not given a charge. Section C only with AR_comment "Workbook rate zero or missing: confirm the rate, or no charge". |
| **Rathbones charges** | 17 matched plans with an existing Plannr charge: charge built at Plannr rate, percentage_frequency=Yearly, frequency=Quarterly, start_date=2027-01-31, payment_origin=Provider, notes "Rathbones: annual %, paid quarterly, first payment 31/01/2027". 16 matched plans with no Plannr charge: Section C for AR to provide rate. |
| **Expenditure mapping** | Source section+line mapped to BA destination via `mappings_hfp.xlsx` Expenditure tab (185 rows). Unmatched lines marked CUSTOM. |
| **Statuses** | Plans with status=active included; inactive/transferred/cancelled excluded from review pack Client book and active-only exception filters. |
| **Circles** | 92 of 98 snapshot circles included (6 excluded: 5 empty, 1 with members but no plans). Group = "default" for all. |
| **Tags** | Not processed this session. |
| **Vision IFP** | Not processed this session — open question (see below). |

---

## Key counts

| Metric | Count |
|---|---|
| Clients | 375 |
| Plans (total) | 1,864 |
| Active plans | 1,563 (in review pack Client book) |
| Charges (all) — workbook | 441 |
| Charges (all) — Rathbones | 17 |
| Charges (all) — Plannr carried over | 11 |
| **Charges (all) total** | **469** |
| **Total indicative annual income** | **£409,693** |
| Valuations rows | 1,528 |
| Income & Employment rows | 316 |
| Expenses rows | 2,381 |
| Circles | 92 |
| Dependants (unique persons) | 250 |
| Relationships (unique pairs) | 190 |
| Premiums | 173 |

**Dependants note:** 249 unique dependants in the 2026-10-12 snapshot. 1 additional found only via the live `/api/v2/dependants` download (UUID `89a694c4-4ad7-4ff3-98de-eaffdaef9688`; parent accounts `34aed1c8`, `ff64f4f4`, `0e3a8af6`). Total importer rows: 250. The figure 367 = total account-dependant links (each dependant counted once per linked parent account); it is not a unique count. 250 is not a per-page cap — the API returns up to 250 per page; use `filter[uuid]` with `include=accounts` and follow `next` pagination links to retrieve beyond 250.

### Indicative annual income by charge group

| Group | Count | Total valuation | VW avg rate | Total IAI |
|---|---|---|---|---|
| Workbook | 441 | £62,670,578 | 0.595% | £372,664 |
| Rathbones | 17 | £4,250,259 | 0.542% | £23,040 |
| Plannr carried over | 11 | £3,612,819 | 0.387% | £13,989 |
| **Total** | **469** | **£70,533,656** | | **£409,693** |

### Exceptions by type (v6 internal pack, 368 total)

| Category | Count |
|---|---|
| No policy number (active plans only) | 239 |
| No status, no plans | 55 |
| Rathbones plan not on Rathbones list | 53 |
| Active agency plan not in workbook | 11 |
| Account in multiple circles | 4 |
| Invalid postcode | 3 |
| Duplicate policy number | 2 |
| Fund number collision (anomaly) | 1 |

### Charges sign-off (v6, 671 rows + 3 section headers)

| Section | Count | Description |
|---|---|---|
| A | 458 | Workbook-confirmed (441 Transact/Nucleus + 17 Rathbones) |
| B | 11 | Plannr ongoing — active plans, no workbook confirmation |
| C | 199 | No confirmed charge: AR to provide rate or confirm no fee |

### Plannr 330-charge reconciliation

| Disposition | Count |
|---|---|
| Replaced by workbook charge (Section A) | 294 |
| Replaced by Rathbones charge | 17 |
| Kept in Section B | 11 |
| Dropped — plan inactive or transferred | 6 |
| Dropped — zero-rate workbook plan (Section C) | 4 |
| **Total** | **332** |

Note: snapshot contains 332 active ongoing charges vs 330 counted in v2 build. The 2-charge gap is under investigation.

---

## Rathbones match summary

| | Count |
|---|---|
| Rathbones plans in Plannr | 86 |
| Matched (file + Plannr) | 33 |
| → With existing Plannr charge → Section A | 17 |
| → No Plannr charge → Section C | 16 |
| Rathbones-only (file, not in Plannr) → Provider-only plans tab | 19 |
| Plannr-only Rathbones (not in file) → Exceptions | 53 |
| Fund number collision (anomaly) | 1 |

---

## Not yet done

- **Review pack not sent.** `hfp_review_pack_v2.xlsx` is built but not yet emailed to Jim Hoban. Return date not filled in on the Read me tab.
- **AR importer files (v2) built 2026-10-14.** 10 separate .xlsx files plus `hfp_Data_Analysis_v2.docx` in DATA_ROOT. Awaiting AR sign-off (charges Section C rates, provider-not-in-Plannr decisions, notes choice).
- **Block 4 (validate) not run.** No validation pass against the BA destination environment.
- **ZZ test not run.** No test import to BA staging.
- **No destination writes made.** All sessions snapshot/export-only; nothing written to the BA system.
- **Tags not processed.** Tag mapping not applied.

---

## Open questions for the AR (review pack)

1. **Fees to confirm (210 plans):** Charges sign-off Sections B and C require AR confirmation or a rate to be provided.
2. **Plans to check (325 plans):** No policy number (239), duplicates (2), Rathbones not on list (53), agency not in workbook (11), no adviser assigned (20).
3. **20 plans with no adviser:** Appear as "To confirm" in Summary. All show £0 value and £0 IAI — may be prospect or test plans. AR to assign adviser or confirm for removal.
4. **Notes — all or nothing:** Plannr holds 4,774 task notes and 681 case notes. BA importer is all-or-nothing per note type. Confirm whether to import all, none, or a filtered set.
5. **Vision IFP:** Not processed. Confirm scope — is Vision IFP data in scope for this migration?
6. **332 vs 330 charge gap:** 2 active Plannr charges unaccounted for in the v2 build. Investigate before final import.
