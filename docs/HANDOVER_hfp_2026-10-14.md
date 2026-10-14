# HFP AR Migration — Handover Note
**Date:** 2026-10-14  
**Firm slug:** `hfp` (Hoban Financial Planning Ltd)  
**Session:** exports-based AR review pack, INFO columns, two post-review fixes

---

## Status

| | |
|---|---|
| **Pack sent** | 14 Oct 2026 |
| **AR returns pack** | by Friday 23 Oct 2026 |
| **Call with the AR** | 15 Oct 2026 |
| **Target go-live** | 2 Nov 2026 |

**Next steps (in order):** AR's answers → Block 4 clients validate run → ZZ test (needs Mark's approval) → waves W1–W8.

No API calls and no destination writes were made in this session.

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

## Open items

| Item | Note |
|---|---|
| **Vision IFP** | The AR's old network. Every reference to Vision IFP must be removed. Not yet swept from any sheet. |
| **Tags** | Not in the sheets. Decision pending. |
| **Board-task notes** | 2,104 notes are workflow templates with no client link. They stay behind in the old Plannr. |

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
