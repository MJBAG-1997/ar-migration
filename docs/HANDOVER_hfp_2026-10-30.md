# HFP Migration Handover — 2026-10-30

Branch: `feature/ar-migration-hfp`
Destination account: `ccd7f98b-d6a4-46ad-bab3-242ab42826eb`
Source account: `db0ab7f5-8e52-4c89-a915-29fefeeca0d4`

---

## What is done

### All-waves run table

| Wave | Expected | Stored | Status |
|---|---|---|---|
| W1_clients | 272 | 272 | ✓ |
| W2_clients | 84 | 84 | ✓ |
| W2_circles | 87 | 87 | ✓ (+2 empty, see prior handover) |
| W3_relationships | 182 | 177 undirected pairs | 5 short; all within 356 clients |
| W3_dependants | 249 | 249 | ✓ |
| W3_income | 316 | 316 | ✓ |
| expenditure_standard | 114 | 114 | ✓ |
| expenditure_custom | 1703 | 1703 | ✓ |
| W4_plans | 1 | 1 | ✓ |
| W5_plans | 1900 | 1900 | ✓ |
| W6_valuations | 1526 | 1525 | −1 (502 during count) |
| W6_premiums | 171 rows | 173 | minor discrepancy |
| W7_charges | 542 | 542 | ✓ |
| W8_notes | 6757 | 6861 | +104 from Oct-27 partial runs |
| B1 provider clear | — | 115 + 210 = 325 cleared | 197 remain in AR list |
| E3 task notes | 8 | 8 confirmed on dest | ✓ |
| E5 flexi plan | — | old plan deleted | ✓ |
| E6 joints | 73 | 73 deleted | ✓ |
| C documents | 1,110 | 1,093 copied | 17 failed (see below) |

### Morning checks — 2026-10-30

| Check | Result |
|---|---|
| E3 notes on dest | 8 / 8 found |
| Duplicate files (dest > src) | 0 |
| Date mismatches (original_created_at) | 0 — nothing to fix |
| Unknown Provider plans cleared | 407 → 210 cleared; 50 refused 422; 197 remain |

---

## What is left

### 1. 17 failed document files

Full list in `import_20261027/state/failed_files.json`.

| Failure type | Count | Cause | Action |
|---|---|---|---|
| upload 403 | 13 | Plannr/WAF rejects filename — `.aspx` extension, spaces, parentheses | Sanitize filename before upload (strip `.aspx`, truncate long names); re-run `taskC_20261029.py` (state file skips the 1,093 already done) |
| download 403 | 2 | Both on source client `4a05c772` — files inaccessible on HFP Plannr | Investigate source account permissions; if unresolvable, note as lost |
| download 404 | 1 | File deleted from source — dest client `d2ef7c74` | Cannot recover |
| folder create | 1 | Folder could not be created on dest client `9eab42b4` | Re-run `taskC_20261029.py`; folder creation will be retried automatically |

**How to retry:**
1. `git pull` on the work PC to get the latest `taskC_20261029.py`.
2. Add filename sanitization to `upload_one()` before the `upload_data["name"]` assignment — strip or replace `.aspx`, parentheses, and leading/trailing spaces; truncate to 200 chars.
3. Run `python -I scripts/taskC_20261029.py` — it loads `file_state.json` and skips the 1,093 already-copied files, retrying only the missing ones.
4. Count after: state file should grow from 1,093 to at most 1,110 (minus the 3 unrecoverable: 1×404 and 2×403 from source).

### 2. 197 Unknown Provider plans — AR workbook

Tab: **Unknown Provider** in `hfp_AR_actions_20261029.xlsx` (available via `git pull`).

| Category | Count | Action |
|---|---|---|
| Financial products (pensions, ISAs, investments, protection) | 147 | Advisers assign provider via UI |
| Mortgage/property_holdings/trust — Plannr refused null (422) | 50 | Advisers assign correct lender/provider via UI |

Plannr enforces `provider_uuid` as required on specific mortgage and trust plans that have associated lender data. These 50 cannot be null-cleared; the adviser must pick the correct provider in the UI.

### 3. Old Plannr notes — not moved

52 notes remain on the source Plannr (`import_20261027/state/e2_remainder_notes.json`).
All 52 are `author=IT SYSTEM` — system-generated (type: plan/client/charge).
Listed in the **Old Plannr Notes** tab of the workbook.
Review to confirm none contain adviser-authored content before the source account is decommissioned.

### 4. Other open items (from prior handover)

| # | Item |
|---|---|
| 5 relationships short | 177 undirected pairs on dest vs 182 in source CSV — 5 pairs missing; all within 356 clients |
| 2 empty circles | `5ec4636c`, `fbadcb0c` — delete via UI if confirmed unwanted |
| Premiums 173 vs 167 unique plans | Minor; investigate if sign-off requires exact count |
| Workbook tabs | No Charge Rate, Income Defaults, Placeholder Policy Nos, Phone Numbers, Provider List Plans — review and act via UI |

---

## Standing rules (CLAUDE.md)

### Imports
- Every import goes through the runner's import function: create, map headers, start, poll, then count records on the destination. Never write new import code.
- If an import's progress hasn't moved for 5 minutes, stop waiting. GET the import and report `status`, `mapping`, `total_jobs`, `progress` and `failed_jobs`, then act on what the fields show.
- Never create a second import for the same data while one exists that isn't complete or deleted.

### Waiting and reporting
- Never wait on anything for more than 5 minutes without reporting.
- If a step is blocked, carry on with the steps that don't depend on it, then come back to it.
- No diagnosis without evidence: quote the response field or error that shows the cause.

### Writes
- Count before, write once, count after. If after minus before isn't the expected number, STOP.
- The API takes pennies; the importer takes pounds. Before any bulk write, write one record, read it back and compare with the source.
- Only one script talks to Plannr at a time.

---

## State files — manual transfer required (gitignored)

All under `import_20261027/state/`. Transfer the whole directory.

| File | Size | Purpose |
|---|---|---|
| `file_state.json` | 89 KB | Copy manifest: src_uuid → dest_uuid (1,093 entries) |
| `folder_state.json` | 123 KB | Folder UUID cache keyed `{acct_uuid}\|path` |
| `b1_ar_list.json` | 26 KB | 197 remaining Unknown Provider plans |
| `plan_map.json` | 154 KB | src ext_ref → dest plan UUID (1,901 entries) |
| `client_map.json` | 29 KB | src client UUID → dest client UUID (356 entries) |
| `e2_remainder_notes.json` | 11 KB | 52 old Plannr notes (IT SYSTEM) not migrated |
| `exp_custom_state.json` | 256 KB | Expenditure custom lines state (1,703 entries) |
| `failed_files.json` | — | 17 failed document files with failure type and dest client |
| `e1_premium_state.json` | <1 KB | Premium import state |
| `b3_flexi_details.json` | 11 KB | Flexi plan details |
| `b6_clear_list.json` | 21 KB | Joint accounts deleted list |
| `wave_state.json` | 1 KB | Wave progress flags |

Also transfer: `import_20261027/W8_notes_final_v2.csv` (4.1 MB) — notes import source file, not in git.

Everything else is in the repo and available via `git pull origin feature/ar-migration-hfp`.
