# HFP AR Migration — Handover Note
**Date:** 2026-10-14  
**Firm slug:** `hfp` (Hoban Financial Planning Ltd)  
**Session:** exports-based AR review pack, INFO columns, two post-review fixes

---

## Status

| | |
|---|---|
| **Pack sent** | 14 Oct 2026 |
| **AR returns pack** | 26 Oct 2026 (received) |
| **Call with the AR** | 15 Oct 2026 |
| **Target go-live** | 2 Nov 2026 |

**Next steps (in order):** ZZ test (needs Mark's approval) → waves W1–W8.

Build import files complete (27 Oct 2026). W1 validates clean against destination. See build session below.

---

## What was sent

The AR guide is **Hoban_Data_Extract_Analysis.pdf** (written outside this session).  
`hfp_Data_Analysis_v2.docx` was **not** sent.

The ten workbooks sent (all in `DATA_ROOT = C:\dev\AR Migration`):

| File | Sheet / purpose | Rows |
|---|---|---|
| `Plannr_Clients.xlsx` | Clients | 375 |
| `Plannr_Plans.xlsx` | Plans | 1,871 |
| `Plannr_Plans.xlsx` | Provider_plans_not_in_Plannr | 51 |
| `Plannr_Charges.xlsx` | Charges | 675 |
| `Plannr_Circles.xlsx` | Circles | 92 |
| `Plannr_Relationships.xlsx` | Relationships | 190 |
| `Plannr_Dependants.xlsx` | Dependants | 250 |
| `Plannr_Valuations.xlsx` | Valuations | 1,528 |
| `Plannr_Income_Employment.xlsx` | Income & Employment | 350 |
| `Plannr_Expenditure.xlsx` | Expenditure | 2,381 |
| `Plannr_Premiums.xlsx` | Premiums | 173 |

---

## Rules applied since 13 Oct

| Rule | Detail |
|---|---|
| **Data sources** | Exports (`exports_20261014/`) as the base; the 2026-10-12 snapshot (`snapshot_20261012T090906Z/`) for relationships, dependants, valuations and income. |
| **Review and INFO: columns** | Every sheet has a `Review` column (AR-editable flags) and `INFO:` reference columns (blue header; ignored at import). UUIDs moved to far right. |
| **Placeholder policy numbers** | Placeholder values (TBC, N/A, Unknown, etc.) restored to `policy_number` and flagged "Placeholder policy number" in Review. |
| **No policy number exclusions** | Active plans with no policy number flagged only if type is NOT in: `fixed_asset`, `cash_account`, `mortgage`, `loan_credit`, `general_insurance`, `group_protection`. |
| **Adviser fallback** | Plan adviser = plan's own `active_adviser_uuid`; if blank, falls back to the first owner's `assigned_adviser_uuid`. |
| **Blank status** | Clients with no Plannr status set to `Level 2 / Navigator`. |
| **under_advice / agency** | Plans matched on the provider statement or Rathbones list: `under_advice=True`, `agency_in_your_name=True`; originals kept in `ORIGINAL_` columns. |
| **Charges** | Three sections: A = provider-statement confirmed (458 plans); B = Plannr ongoing (11 plans); C = rate needed (206 plans). `basis` column records the source. |
| **Dependants — shared** | 73 dependants appear under two clients; `account_uuids` holds all parent accounts pipe-separated; `INFO: Client 2` added. |
| **Provider tab matching** | Re-matched on surname + first name/initial, ignoring titles/case/punctuation; 22 of 51 rows now matched (was 2). |

---

## What the guide promises the AR (the import must honour these)

| Promise | Implementation note |
|---|---|
| A deleted row means leave it out, including anything linked to a deleted client. | Filter rows where the AR has deleted or left blank the key identifier before building import files. |
| `INFO:` columns are ignored. | Drop all columns whose header starts `INFO:` before import. |
| A client or adviser name typed in the `Review` column is a correction for us to apply. | Parse Review for free-text names before dropping the column. |
| `Rate needed` rows left blank get no charge. | Omit Charges rows where `basis=Rate needed` and amount is blank. |
| Protection rows may come back with an amount and frequency instead of a percentage. | Accept both formats in the Charges importer. |
| `Add to Plannr` rows use the `Client in your Plannr` column. | Use the named client UUID when creating the plan from the Provider tab. |
| Notes show under our system user with a header line. | Import notes attributed to the migration system user; prefix each with a header line identifying the source. |
| No welcome emails. | Set `send_welcome_email=false` on all client creates. |

---

## Import notes (for whoever builds the import scripts)

- Drop `Review`, `INFO:`, `ORIGINAL_` and `basis` columns before building import files.
- Multi-value UUID cells use `|` (pipe) as separator in these workbooks; convert to `,` (comma) for the BA importer.
- `client_ownership_percentage` must be filled on joint plans (two owners).
- External references: `external_reference_company = data_migration`; `external_reference_value` = source Plannr UUID.

---

## AR decisions (confirmed 26 Oct 2026)

| Decision | Detail |
|---|---|
| **Notes** | Bring all task and case notes across onto the client record. |
| **Tags** | Not needed — do not migrate tags. |
| **Plans `on_panel`** | Remove every answer; leave `on_panel` blank on all plans at import. 497 plans currently have a value (427 "No", 70 "Yes"). |
| **Source Plannr frozen** | The AR's old Plannr is now read-only. The 26 Oct pull is the final extract. |

---

## 26 Oct 2026 session — AR review results

Fresh exports pulled to `exports_20261026/`. All 7 reports succeeded (client, plan, charges, circle, premium, expenditure, note) plus dependants (249 rows with `include=accounts`). Detail in `hfp_changes_20261026.xlsx`.

**Step 1b — Valuations (plan_export current_value / current_value_at)**

| | |
|---|---|
| Plans with changed value or date | 522 |
| Value + date changed | 491 |
| Date only changed | 31 |
| Multi-owner plans (not split) | 0 |
| Total value before (changed plans, £) | 92,941,916 |
| Total value after (changed plans, £) | 83,151,806 |

Unit verified: snapshot pence ÷ 100 = export pounds (5/5 spot checks).

**Step 2a — AR edits (RETURNED vs SENT)**

| Sheet | Added | Deleted | Changed cells | Format-only |
|---|---|---|---|---|
| Clients | 0 | 0 | 0 | 0 |
| Plans | 0 | 0 | 0 | 0 |
| Provider_plans_not_in_Plannr | 0 | 0 | 80 | 0 |
| Charges | 0 | 0 | 0 | 0 |

Notes: the AR returned only Clients, Plans (×3), and Charges sheets. The three returned Plans files merged cleanly (0 intra-conflicts). "Trail Fees (Excluding Transact & Rathbones).xlsx" was returned as supplementary data — not one of the ten sent workbooks.

**Step 2b — What the AR wrote**

| Item | Count |
|---|---|
| Review entries on Clients | 16 |
| Review entries on Plans | 67 |
| Provider "Add to Plannr" rows | 32 (all have Client in your Plannr filled) |
| Provider "Already in Plannr" rows | 19 |
| Charges — Rate needed rows filled in | 149 |
| Charges — Rate needed rows still blank | 57 |
| Charges — Protection rows with new amount | 60 |
| Charges rows deleted | 0 |

**Step 2c — CRM changes (26 Oct export vs 14 Oct export)**

| Entity | Old | New | Added | Deleted | Changed records |
|---|---|---|---|---|---|
| clients | 375 | 358 | 8 | 25 | 217 (incl. 178 `aum` recalc) |
| plans | 1,871 | 1,869 | 11 | 13 | 572 (incl. 522 valuation) |
| charges | 372 | 373 | 2 | 1 | 18 |
| circles | 92 | 87 | 1 | 6 | 6 membership changes |
| premiums | 173 | 171 | 0 | 2 | 3 |
| notes | 8,922 | 8,874 | 43 | 91 | 170 |
| expenditure | 2,381 | 2,381 | 0 | 0 | 0 |
| dependants | 250 | 249 | 2 | 3 | 0 |

**Step 2d — Conflicts:** 0 (no field changed by both the AR and the CRM).

---

## Open items

| Item | Note |
|---|---|
| **Vision IFP** | The AR's old network. Every reference to Vision IFP must be removed. Not yet swept from any sheet. |
| **Board-task notes** | 2,104 notes are workflow templates with no client link. They stay behind in the old Plannr. |
| **Charges — 57 Rate needed rows still blank** | AR did not fill these in; they will be omitted at import. |
| **25 clients deleted from CRM** | Review before import — confirm these are genuine deletions, not errors. |

---

## API lessons learned

| Lesson | Detail |
|---|---|
| `circle_export` structure | One row per member, not one row per circle. Build a circle-to-members map before using. |
| `note_export` includes board-task notes | Workflow template notes with no client link are in the export; filter on `account_uuid` being present. |
| `valuation_report` returned 422 | Used snapshot data instead. Do not retry without investigating the Plannr API error. |
| `income_and_employment_export` server failure | Used snapshot data instead. |
| CSV exports omit times from dates | Use `GET /api/v1/plans/{uuid}` if timestamp precision is needed. |
| `/api/v2/dependants` pagination | Returns at most 250 per page; follow the `next` link. Use `include=accounts` to get parent account UUIDs in one call. |
| Importer map call | `POST /api/v1/import/map/{uuid}` requires `{"headers": {col: col, …}}` (identity mapping). Calling with `{}` returns 422. Must wait ~8s after upload before calling map — headers aren't parsed immediately. |
| `archived_reason` required | When status maps to ARCHIVED at dest, `archived_reason` is required. Valid value: `"Left firm - Firm decision"`. Source had no values set. |
| Dest has no groups | `ccd7f98b` (HFP sub-account) has no client groups configured. Leave `groups` blank — any UUID causes "selected groups is invalid". |
| Non-UK phones rejected | Plannr rejects non-+44 phone numbers with "invalid number" error. Eight non-+44 numbers blanked; listed in "Manual phones" tab in hfp_build_lists_20261027.xlsx for post-import manual entry. |
| Rate-needed charge matching | Returned Charges workbook is in different row order from sent. Match by `external_reference_value` dict, not row index. |
| Rate-needed charge count | AR filled 47 non-zero rates (not 149 — the earlier figure was from a row-index comparison that picked up noise from row reordering). |
| Valuations zero+no-date | 338 plans had zero value and no valuation_date in snapshot — skip these; they never had a valuation. Plans that legitimately went to zero (have a date) are kept. |

---

## 27 Oct 2026 session — Build import files

Script: `scripts/build_import_20261027.py`  
Output: `import_20261027/` (12 CSV files) + `hfp_build_lists_20261027.xlsx`

**Pre-build check 1 — missing clients / plans**

All 25 clients and 13 plans missing from 26 Oct export confirmed deleted (0 alive on source). 18 of the 25 deleted clients were flagged "No status, no plans" in the v6 signoff pack.

**Pre-build check 2 — AR edits recount**

| Sheet | Changed cells |
|---|---|
| Clients | 0 |
| Plans | 0 |
| Provider_plans_not_in_Plannr | 80 |
| Charges | 6376 |

**Wave file row counts**

| File | Rows |
|---|---|
| W1_clients.csv | 272 |
| W2_clients.csv | 86 |
| W2_circles.csv | 87 |
| W3_relationships.csv | 247 |
| W3_dependants.csv | 249 |
| W3_income.csv | 316 |
| W4_plans.csv | 1 |
| W5_plans.csv | 1893 (inc 25 Add-to-Plannr) |
| W6_premiums.csv | 171 |
| W6_valuations.csv | 1,526 (338 zero+no-date skipped, 156 zero-or-no-date loaded) |
| W7_charges.csv | 420 (373 sec A/B, 47 sec C filled; 159 sec C omitted) |
| W8_notes.csv | 6,796 (601 account, 4,784 task re-homed, 676 case re-homed, 663 plan, 26 risk, 44 circle, 2 charge) |

**W1 validate result (post fix-pass):** `complete`, 0 errors (import UUID `bd1ba206-a335-4600-8173-3d6e8b7834c8`).

Fix pass before load (fix items 1–8):
1. Notes: use `account_uuids` fallback; task/case re-homed to client account with header line.
2. Charges reconciled: 675 sent → 420 loaded (373 sec A/B, 47 sec C). AR filled 47 non-zero rates (earlier "149" figure was from flawed row-index comparison).
3. pre_check_2 charges: 128 real changed cells after UUID-based matching (was 6,376 noise).
4. Valuations: 338 zero+no-date rows skipped; 1,526 loaded.
5. Relationships: 247 rows, 190 unique undirected pairs.
6. Archived_reason: DECEASED → "Deceased" (3); ARCHIVED/INACTIVE → "Left firm - Firm decision" (47).
7. Phones: 8 non-+44 phones blanked; listed in "Manual phones" tab in hfp_build_lists_20261027.xlsx for post-import entry.
8. Trail Fees: 3 sheets (177/25/153 rows), no policy_number column — not loaded.

Pre-fix validation errors:
1. Groups UUID blanked (dest account `ccd7f98b` has no groups)
2. `archived_reason = "Left firm - Firm decision"` / "Deceased"
3. Non-UK (+33) phones blanked
