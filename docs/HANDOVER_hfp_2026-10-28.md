# HFP Migration Handover — 2026-10-28

Branch: `feature/ar-migration-hfp`
Destination account: `ccd7f98b-d6a4-46ad-bab3-242ab42826eb`

---

## Current state — W1–W7 stored, W8 not started

| Wave | Expected | Created | Notes |
|---|---|---|---|
| W1_clients | 272 | 272 | |
| W2_clients | 84 | 84 | |
| W2_circles | 87 | 87 | |
| W3_relationships | 182 | 147 | 35 dropped — see read-only checks below |
| W3_dependants | 249 | 249 | |
| W3_income | 316 | 316 | |
| expenditure | — | 118 | |
| W4_plans | 1 | 1 | |
| W5_plans | 1900 | 1900 | |
| W6_valuations | 1526 | 1526 | |
| W7_charges | 542 | 542 | |

Total: 356 clients, 1901 plans. W8 (6795 notes) awaits explicit go-ahead.

---

## W8 notes file — step 1c check results

File to use: `import_20261027/W8_notes_final.csv` (4,262 KB)

| Check | Result | Status |
|---|---|---|
| Total rows | 6795 | ✓ |
| model_type=account | 6132 | |
| model_type=plan | 663 | |
| Rows on a joint account UUID | 0 | ✓ |
| Rows with visible_to_clients=Yes | 0 (9 fixed) | ✓ |
| Blank / unresolved UUIDs | **4160** | ✗ |

**Unresolved UUID detail** — 4160 notes on 6 source UUIDs that are not in client_map (adviser/admin accounts, not imported as clients):

| Source UUID | Notes | Identity |
|---|---|---|
| `db0ab7f5-8e52-4c89-a915-29fefeeca0d4` | 2659 | Jim Hoban (adviser) |
| `6dbac8ff-90d2-4dc8-b60d-26567773a68e` | 899 | Valerie Hoban (admin) |
| `2153bb3c-7b76-4643-86a3-a9c53acab0b0` | 307 | Katie Hoban (adviser) |
| `5a8bc24f-8ff9-4671-ba49-4a5c8bfc5b2b` | 293 | Tommie Hoban (adviser) |
| `bc00b121-9dcc-4283-bcb2-a75697fa249d` | 1 | Unknown (likely deleted account) |
| `2ce06c3f-6929-49f9-bfcf-946debc81209` | 1 | Unknown (likely deleted account) |

**W8 cannot run until these 4160 rows are resolved.** Options: redirect to a firm-level account UUID, or drop them with a pre-filter. Decision needed before go-ahead.

193 joint account notes were redirected to the first individual owner's source UUID (resolved to dest UUID at W8 run time). Trust and business account notes remain on their own account UUIDs. circle_uuid column removed.

---

## W8 run rules

- Wait for explicit "go ahead W8" — do not start without it
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
| `W8_notes_final.csv` | 4,262 KB | W8 submission file |
| `W8_notes.csv` | 4,267 KB | Original notes (reference only) |
| All W1–W7 source CSVs | — | Not needed for W8 but keep for re-checks |

Thursday's plans clean-up files stay on the office PC.
