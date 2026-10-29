# HFP Migration Handover — 2026-10-28

Branch: `feature/ar-migration-hfp`
Destination account: `ccd7f98b-d6a4-46ad-bab3-242ab42826eb`

---

## Current state — W1–W8 stored (2026-10-29)

| Wave | Expected | Created | Notes |
|---|---|---|---|
| W1_clients | 272 | 272 | |
| W2_clients | 84 | 84 | |
| W2_circles | 87 | 87 | |
| W3_relationships | 182 | 147 | 35 dropped — see read-only checks below |
| W3_dependants | 249 | 249 | stored twice — site total 498, all counts 2× |
| W3_income | 316 | 316 | |
| expenditure | — | 118 | |
| W4_plans | 1 | 1 | |
| W5_plans | 1900 | 1900 | |
| W6_valuations | 1526 | 1526 | |
| W6_premiums | 171 | 0 | **NEVER IMPORTED** — not in import history |
| W7_charges | 542 | 542 | |
| W8_notes | 6757 | 6757 | submitted from v2; spot-checks confirm notes present |

Total: 356 clients, 1901 plans, 6757 notes.

**Premiums gap:** W6_premiums.csv (171 rows, source plan UUIDs) was never submitted. It does not appear in GET /api/v1/import history. Premiums must be imported before sign-off.

**Dependants doubled:** W3_dependants was stored 5 times across two migration days. Site now has 498 dependants (2× expected 249). All spot-check clients are 2×: C1=4, C4=12, C6=4. Duplicates must be removed before sign-off.

---

## W8 run results (2026-10-29)

**Script:** `scripts/run_w8_notes.py`
**File submitted:** `import_20261027/W8_notes_final_v2_rewritten.csv` (source UUIDs rewritten to dest)

### Pre-check

| Check | Result |
|---|---|
| Blank created_at | 0 |
| Earliest created_at | 01/02/2023 |
| Latest created_at | 31/10/2025 |

### Submit

- Status: `complete`, errors: 0, failed: 0
- Importer returns 0 for all row counts (Plannr API quirk — consistent with every other wave)
- Spot-checks confirm notes are on the destination: C5=174 ✓, C6=148 (147 expected), C1=15 (9 expected)
- Extras on C1 (+6) and C6 (+1) are likely from 3 earlier "notes store complete" runs on 2026-10-27 already in the import history

### Notes in source fetch but NOT in the file (8)

The all-notes fetch found 6,765 unique note UUIDs; the v2 file has 6,757. The 8 extras were in the source system but absent from the original W8_notes.csv export:

| Note UUID | Created |
|---|---|
| `e023e1f4-30b6-464f-9cf7-0331dc17ffa3` | 2024-06-25T15:46:27+01:00 |
| `3542f738-1973-49f5-8ab6-53089d0ae2e6` | 2026-10-28T11:40:07+00:00 |
| `2b1a9cc8-1b3f-45ac-9c2f-68f5befefca7` | 2026-10-28T11:40:41+00:00 |
| `7d41cba6-5af2-4422-bb8b-578fc79d39b8` | 2026-10-28T11:42:53+00:00 |
| `a13f20bd-3bbb-4263-91a3-3743d3d0c902` | 2026-10-28T11:46:37+00:00 |
| `06726324-9446-4d03-aa8e-1cf46cca9f85` | 2026-10-28T12:28:45+00:00 |
| `cdfe3196-4c81-4da8-8055-8cea0bf55d4b` | 2026-10-28T12:29:11+00:00 |
| `d4992c2a-0c8c-4d04-97d6-f1521ba114fa` | 2026-10-29T10:40:37+00:00 |

Seven were created on migration day (2026-10-28) or today; one dates from June 2024 and was missed in the source export.

---

## Reconciliation 2026-10-29 (`scripts/reconcile_20261029.py`)

### Check 1 — Notes: total count and source refs

`GET /api/v1/client/{uuid}/all-notes` does **not** return `external_reference_company`/`external_reference_value` in the note object. The note keys are: `account, author_name, contents, created_at, edited, is_pinned, notable, notable_type, path, type, updated_at, uuid, visible_to_clients`. Source ref matching is impossible via this endpoint.

| Metric | Result |
|---|---|
| Dest notes total | 6,861 |
| W8 file submitted | 6,757 |
| Extra notes | **104** — from 3 notes store jobs on 2026-10-27 |
| Source refs found in dest | 0 (ext-ref not in API) |
| Notes with no data_migration ref | 6,861 (all — API limitation) |
| Duplicate source refs | 0 (cannot assess) |

C1: 15 dest vs 9 in v2 = **+6 extra** from Oct-27 imports.
C6: 148 dest vs 147 in v2 = **+1 extra** from Oct-27 imports.

### Check 2 — Note dates

Cannot do per-note date comparison: external references not exposed by API, so dest notes cannot be matched to v2 source rows.

Source (v2 CSV) date range: **2022-04-12 to 2026-10-26**.

Dest note date range not computed (no matching basis).

### Check 3 — Duplicates by type

`/api/v1/client/{uuid}/relationships` → HTTP 404 for every client (correct endpoint unknown).
`/api/v1/client/{uuid}/income-and-employment` → HTTP 404 for every client (correct endpoint unknown).
External references not returned by circles or dependants list endpoints either.

| Type | Dest | Expected | Delta | Twins |
|---|---|---|---|---|
| circles | 89 | 87 | **+2** | cannot identify (no ext-ref in API) |
| relationships | unknown | 147 | unknown | endpoint 404 |
| dependants | 498 | 249 | **+249** (2×) | cannot pair (no ext-ref) |
| income | unknown | 316 | unknown | endpoint 404 |
| valuations | 1,525 | 1,526 | **−1** | plan `d8cc1c33` returned HTTP 502 during fetch |
| charges | 542 | 542 | 0 | ✓ |

### Check 4 — Dependant duplicate pairs

All 249 duplicates confirmed by count (498 = 2×249), but individual pairs cannot be listed: the `/api/v2/dependants` endpoint does not return external_reference fields. Correct Plannr endpoint or UI-based approach needed to identify which of each pair to delete.

### Check 5 — Expenditure: 2,381 source rows vs 114 dest fields

Dest expenditure form has **23 fields** (from GET /api/v1/expenditure-structure/section). Source export has **182 distinct (section, type) combinations**.

| Category | Count |
|---|---|
| Custom — (section, type) not in dest 23-field form | **2,229** |
| Amount ≤ 0 (skipped) | **34** |
| No dest client UUID | **0** |
| Overwritten (later row same client+field) | **4** |
| Written (last write for client+field, dest non-zero confirmed by check 9) | **114** |
| Total | **2,381** |

114 writes × 1 field each = 114 non-zero dest fields on 69 clients (matching check 9 result). 4 overwrites = 4 additional POST calls whose values were replaced = 118 total POST calls (matching the "118 records" in the state table). 19 source clients had only custom or zero rows → 0 fields written for them.

The 2,229 custom rows (93.6%) are granular line items (e.g. "Essential - Home / Council Tax", "Essential - Vehicles / Car Insurance") that do not match any of the 23 fixed fields in the destination expenditure form. They were silently skipped at import. Client data exists in the source export but is not held in the destination system.

---

## Post-W8 read-only checks (2026-10-29)

### Check 6 — Import history (GET /api/v1/import)

101 jobs total. All jobs show 0 for rows/created/failed — Plannr API does not return row counts. Full type+mode+status table is in the script output log. Summary by type × mode stored:

| Type | Stores |
|---|---|
| clients | multiple |
| circles | multiple |
| relationships | multiple |
| dependants | 5 |
| income_and_employment | 1 |
| plans | multiple |
| valuations | 1 |
| charges | 1 |
| notes | 4 (3 on 2026-10-27, 1 on 2026-10-29) |
| **premiums** | **0 — never stored** |

### Check 7 — Dependants (GET /api/v2/dependants)

| Client | Expected | Actual |
|---|---|---|
| Site total | 249 | 498 (2×) |
| C1 (`587bbc91`) | 2 | 4 |
| C4 (`c80da0e1`) | 6 | 12 |
| C6 (`26697c8a`) | 2 | 4 |

Endpoint `GET /api/v2/dependants?filter[account_uuids]={uuid}` (v2, singular path was wrong in v1 spot-check).

### Check 8 — Premiums (GET /api/v1/plans/{uuid}/premiums)

All 171 plans in W6_premiums.csv checked. Total premiums on destination: **0**. Premiums were never imported (confirmed by import history). W6_premiums.csv uses source plan UUIDs — needs rewrite to dest before submission.

### Check 9 — Expenditure (all clients)

| | Lines | Clients |
|---|---|---|
| Source (non-zero) | 1817 | 88 |
| Destination (non-zero) | 114 | 69 |

All 88 source clients differ. The destination API returns expenditure as a fixed-category form (sections/fields); the source export has one row per granular line item. The counts are structurally incomparable — the 114 dest lines are the non-zero fields in the form, not a count of source categories. 19 clients have 0 non-zero fields in the destination despite having source expenditure. Full per-client list is in the script output.

### Check 10 — Unknown Provider plans

Total plans with "Unknown Provider": **522**

| Plan type | Count | Provider required? |
|---|---|---|
| fixed_asset | 205 | no |
| cash_account | 149 | no |
| investment | 50 | **yes** |
| mortgage | 47 | **yes** |
| pension | 32 | **yes** |
| protection_policy | 23 | **yes** |
| group_protection | 8 | no |
| annuity | 1 | **yes** |
| general_insurance | 2 | **yes** |
| loan_credit | 5 | no |

**155 plans with Unknown Provider are types where provider is required.** Thursday clean-up task 1 (previously estimated 429 total — actual is 522).

### Check 11 — under_advice raw API value

602 plans in W5 have `under_advice=no`. The destination API stores these as `False` (Python boolean). The spot-check comparison was showing `actual=` (empty string) because of `str(False or "")` evaluation — that was a script artefact, not a data error.

---

## W8 notes file — step 1c check results

**File to use: `import_20261027/W8_notes_final_v2.csv`** (W8_notes_final.csv superseded — see resolution below)

| Check | Result | Status |
|---|---|---|
| Total rows | 6757 | ✓ |
| Unresolved UUIDs | 0 | ✓ |
| Rows on a joint account UUID | 0 | ✓ |
| Rows with visible_to_clients=Yes | 0 | ✓ |

---

## W8 notes — staff-account resolution (2026-10-29)

The 4,160 rows in W8_notes_final.csv with staff-account UUIDs were resolved by fetching
`GET /api/v1/client/{uuid}/all-notes` on the source for all 356 migrated clients (W1+W2),
building a map of source note UUID → source client UUID(s), then applying the rule:
under one migrated client → use that client's source UUID; under more than one →
first individual client; under none → drop.

| Metric | Count |
|---|---|
| Re-homed (staff UUID → migrated client) | 4122 |
| Multi-client notes (used first individual) | 2259 |
| Dropped (not under any migrated client) | 38 |
| Cross-check disagreements | 0 |

Dropped by staff account UUID:

| Source UUID | Dropped |
|---|---|
| `2153bb3c-7b76-4643-86a3-a9c53acab0b0` | 17 |
| `db0ab7f5-8e52-4c89-a915-29fefeeca0d4` | 9 |
| `5a8bc24f-8ff9-4671-ba49-4a5c8bfc5b2b` | 7 |
| `6dbac8ff-90d2-4dc8-b60d-26567773a68e` | 3 |
| `bc00b121-9dcc-4283-bcb2-a75697fa249d` | 1 |
| `2ce06c3f-6929-49f9-bfcf-946debc81209` | 1 |

All 38 dropped are model_type=account. They remain in W8_notes.csv (source of truth).

193 joint account notes (in W8_notes_final.csv) were redirected to the first individual
owner's source UUID. Trust and business account notes remain on their own source UUIDs.
circle_uuid column removed. All source UUIDs in account_uuid — W8 runner must rewrite
to dest UUIDs before submission.

---

## W8 run rules

- Wait for explicit "go ahead W8" — do not start without it
- Submit `W8_notes_final_v2.csv` (not W8_notes_final.csv)
- W8 runner must rewrite source UUIDs → dest UUIDs (client_map + plan_map) before upload
- Stop at the first failure or unmatched record
- Never re-submit a wave that may have stored (notes cannot be deleted)
- Pre-flight: employee list must show exactly 272 W1 clients resolved

---

## Read-only checks still owed

### W3 relationships — 35 dropped rows

The state file records `grace=25, self_links=10`. Split not yet confirmed:
- 10 self-links: verified against 12-Oct snapshot (source noise, not real relationships)
- 25 grace-dropped: need to confirm these split into:
  - Already linked (relationship existed from a prior import)
  - Links to excluded records (accounts not in scope)
  - Any other reason

**Check:** read `import_20261027/W3_relationships.csv` (182 rows). For each dropped row, look up `account_uuid` and `related_uuid` in the dest to confirm the relationship already exists or the linked record is not in `client_map`.

### Expenditure — lines and clients

Source export (`exports_20261026/Client Expenditure Export HFP Migration 2026-10-27.csv`): 2381 rows.
Created: 118 records (via API, not importer — API posts one record per client, not per line item).

**Check:** confirm 118 matches the number of distinct clients in the source export who have expenditure; confirm total line items per client match the source.

---

## Five spot-check client UUIDs (destination)

| Label | Destination UUID |
|---|---|
| joint client | `28cfd18c-89ca-421e-bfff-47e4db249198` |
| trust/business | `06550252-391b-40c0-834f-f6f9bca8f65a` |
| Rathbones plan | `26697c8a-e94b-4eeb-b021-d338793e63f7` |
| task notes | `a598e1a7-151a-4cb6-a774-caaf918ef7b5` |
| Add-to-Plannr plan | `a20b5b1c-dfb9-4cf3-bd06-eb93c45e1341` |

---

## Setup on a new PC

**Python:** 3.14 (py launcher). Install packages:

```
py -3 -m pip install azure-identity azure-keyvault-secrets requests
```

**Azure sign-in** (required for Key Vault — run once):

```
az login
```

Or if using a managed identity / device code flow, `DefaultAzureCredential` picks it up automatically once `az login` is complete.

**DATA_ROOT** is hard-coded in the script:

```
IMPORT_DIR = r"C:\dev\AR Migration\import_20261027"
```

If the path differs on the home PC, edit lines 28 and the `_client_export` path at the W7 fix call (search for `Client Export HFP Migration`).

---

## Run W8

```
cd "C:\dev\AR Migration"
py -3 scripts/run_import_20261027.py --wave W8_notes --mode store
```

*(No `--mode validate` step is needed — the notes importer validates on store.)*

**Before running:** resolve the 4160 unresolved notes (see above). The runner will STOP on the first unmatched UUID unless a pre-filter drops them.

---

## Rebuilding UUID maps if state files are missing (D11)

If `state/client_map.json` or `state/plan_map.json` are missing, the runner rebuilds them automatically on startup by calling:

- `GET /api/v1/clients?include=external_references` → all dest clients with source UUIDs
- `GET /api/v1/plans?include=external_references` → all dest plans with source ext_refs

This is `resolve_client_uuids(dest)` and `resolve_plan_uuids(dest)` in the script. No manual action needed — just run the script normally.

---

## Thursday plans clean-up list

Location: generated at runtime by `_plans_cleanup_report` in the final report output.

The last run's output is in `import_20261027/import_run_20261027.log` (search for "PLANS CLEAN-UP LIST").

Four items:
1. **429 plans set to Unknown Provider** — identify correct providers from the 26 Oct plan export
2. **46 plans with blank contract_enquiry_account_uuid** — 39 → Jim Hoban dest UUID, 7 → Katie Hoban dest UUID (full list in the final report)
3. **1 plan: type=investment → type=pension, sub_type=sipp** — src `prov-042-648-472-IH01853021`, dest `fff4250d-a0c6-4d97-8fec-b81b3483904d`
4. **21 joint-entity-owned plans** — ownership → individual owners (full dest UUID list in the final report)

---

## State files needed to run W8 (DATA_ROOT = `C:\dev\AR Migration\import_20261027`)

| File | Size | Purpose |
|---|---|---|
| `state/wave_state.json` | 702 B | Wave progress — W8 not started |
| `state/client_map.json` | 29 KB | src→dest client UUIDs (356 entries) |
| `state/plan_map.json` | 157 KB | src→dest plan UUIDs (1901 entries) |
| `state/w1_client_uuid_map.json` | 22 KB | W1 individual client map (legacy) |
| `W8_notes_final_v2.csv` | ~4,250 KB | **W8 submission file** (staff notes resolved) |
| `W8_notes_final.csv` | 4,262 KB | Previous version — superseded, keep for reference |
| `W8_notes.csv` | 4,267 KB | Original notes (reference only) |
| All W1–W7 source CSVs | — | Not needed for W8 but keep for re-checks |

Thursday's plans clean-up files stay on the office PC.
