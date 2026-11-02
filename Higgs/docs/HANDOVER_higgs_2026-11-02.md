# Higgs AR migration — handover 2 Nov 2026

## Status

R1–R4 are in. R5 and the plans retry are written but not uploaded.

| Round | Contents | Status |
|---|---|---|
| R1 | 01 Clients (individuals) | ✅ uploaded |
| R2 | 02 Clients BT (businesses/trusts) | ✅ 27 uploaded; 46 held back |
| R3 | 03 Plans | ✅ uploaded (1,597 in Plannr; 36 failed — see below) |
| R4 | 04 Valuations + 05 Charges | ✅ uploaded |
| R5 | 06 I&E + 07 Expenses + 08 Notes + 09 Circles | ⏳ files ready, not uploaded |
| Retry | upload/R5/03 Plans_RETRY.csv (18 rows) | ⏳ ready, not uploaded |

## In Plannr now

| Record type | Count | Notes |
|---|---|---|
| Main clients | 501 | 502 sent; 1 duplicate refused, remapped via supplement |
| BT clients | 27 uploaded | 46 held back (39 trusts no owner; 7 businesses owner outside scope) |
| Plans | 1,597 | 36 failed; 18 in retry file; 13 already loaded as property (prior run); 5 dropped (non-plan rows) |
| Plans by type | inv 695 / prop 465 / pen 280 / prot\_policy 84 / mort 49 / ann 21 / cash 2 / GI 1 | 465 property plans have no provider |
| Valuations | 888 sent | 12 dropped (no link field); 17 BT held-back dropped; 27 policy-number-only (may need manual link) |
| Charges | 363 | 369 source − 6 dropped (linked to failed plans) |
| Income & Employment | not yet | 2,924 rows in upload/R3 |
| Expenses | not yet | 19,841 rows (5 files) in upload/R3 |
| Notes | not yet | 4,506 rows in upload/R3 |
| Circles | not yet | 93 rows in upload/R3 |

## Check first tomorrow — before any upload

### a) BT client count
The upload file has 27 rows; the BT export found 20 matching externalreferences.
Confirm: are all 27 of our uploaded BT clients in Plannr? Are exactly 46 still held back (39 trusts, 7 businesses)?
Export BT clients from Plannr, filter by externalreference to our 27 source refs, count.

### b) Valuation count
888 rows were sent in R4. The `plannr_exports/Migration Export Plan _ Value.csv` (1,601 rows) is a **plan list** with current values attached (`abstract_type`: asset 1,467 / protection 85 / liability 49) — it is not a valuation record list.
To count actual valuation records: export valuations from Plannr (separate export type), count rows, and check for duplicates (same plan UUID + date + value appearing more than once). If any validation run stored valuation records before the real upload, they will show as duplicates.

### c) 13 property plans from a prior validation run
The failed-plans report listed 13 plans as failures (rows 23–35). Investigation showed these were actually uploaded in an earlier run as `type=fixed_asset / sub_type=Property`, stored with clean externalreferences (no `(mortgage)` suffix), and with no provider set. They are in the Migration Export Plans as fixed_asset rows.
Confirm: search the Plannr plan export for these 13 externalreferences and verify none appear twice.
Externalrefs: `f21181cf, fa323bdf, 0b8cf71d, 348e730e, 3329f2bc, b4c6bfa4, 4240fe5a, 344e628b, 00d10022, a1f666a9, cf2f51c7, 2584bfea, 6ac6e8a4`.

### d) R5 source vs upload differences — resolve before uploading notes
Notes **cannot be changed after import**. Resolve all gaps before uploading.

| File | Source (importer) | Upload (R3) | Gap | Likely cause |
|---|---|---|---|---|
| 06 Income & Employment | 3,038 | 2,924 | −114 | Records linked to clients outside migration scope or held-back BT |
| 07 Expenses | 21,381 | 19,841 | −1,540 | Same: some client refs unresolvable at build time |
| 08 Notes | 4,660 | 4,506 | −154 | Same |

Verify by re-running `python higgs_remap.py r3` with verbose output and comparing unresolvable-ref counts to these gaps. If the explanation is held-back clients, no action needed. If any records are missing for non-held-back clients, investigate before uploading notes.

### e) Unknown Provider sheet correction
Property / fixed\_asset plans in Plannr do not require a provider — Plannr accepted all 465 with no provider set. Remove the 13 property rows from the "Unknown Provider Plans" sheet in `higgs_AR_actions_20261102.xlsx`; only the 2 Shrop Council DB pensions should remain there.

## Next, in order

1. **Plans retry** — upload `upload/R5/03 Plans_RETRY.csv` (18 rows)
2. **Plan export** — drop fresh Plannr plan export into `plannr_exports/`
3. **Valuations + charges for retry plans** — run the valuation/charge writer against the new plan export (27 valuations linked to retry externalrefs + any charges)
4. **R5** — upload in this order:
   - `upload/R3/06 Income and Employment.csv` (2,924)
   - `upload/R3/07 Expenses (1 of 5).csv` through `(5 of 5).csv` (19,841)
   - `upload/R3/08 Notes.csv` (4,506) — **after verifying gap in check d above**
   - `upload/R3/09 Circles.csv` (93)
5. **Full reconciliation** — export every record type from Plannr and compare counts
6. **Mobiles by API** — phone number cleanup via Plannr API

## For the adviser

Deliver `higgs_AR_actions_20261102.xlsx`:
- **Failed Plans** (36) — adviser to fix type/provider in Plannr or re-upload
- **BT Held Back** (46) — supply Plannr owner UUID for each, then upload via R2 importer
- **Val Manual Link** (27) — link each valuation to its plan in Plannr manually (plans failed upload)
- **Dropped (non-plan)** (5) — fee/adjustment placeholders; no action needed
- **Unknown Provider Plans** (2 after e above) — DB pensions uploaded as "Unknown Provider"; adviser to set correct pension scheme

Also deliver:
- **24 Ambiguous rows** (review pack "Ambiguous" sheet) — adviser decision needed
- **183 fee-paying plans not on his list** — from the review pack "Not on his list" sheet
- **119 possible duplicates** — from the review pack "Possible duplicates" sheet
- **Fix after import** (Failed Plans + BT Held Back + Unknown Provider as above)

**Not for adviser yet:**
- `Val Manual Link` is superseded by the plans retry — once retry uploads, most of these 27 should auto-link via plan_uuid. Send only what remains after retry.
- **Premiums (142)** — held back pending a decision on manual vs automated linking.

## Rules agreed in this session

| Rule | Detail |
|---|---|
| Merge | Joint Vision records → Plannr circle; don't merge into one client |
| Under management | "Under management" flag taken from adviser's own list, not Vision flag |
| No VAT | All charges net; VAT not stored in Plannr |
| Charges from | 2 Nov 2026 (charges start date for new Plannr records) |
| Vision contribution | Not used as a plan value; valuation only from valuation records |
| Retail Client | Applied to all clients including businesses and trusts |
| Navigator | Set for all fee-paying clients |
| Health vulnerability | Suppressed — not imported (no equivalent Plannr field) |
| Email handling | If cell has multiple addresses, first address used; remainder go to "Fix after import" |
| Unknown first names | "Unknown" used as first\_name placeholder; flagged in handover |
| Business key | `entity_name` is the key field for BT clients (not first/last name) |
| Round order | R1 → R2 → R3 → R4 → retry → R5 → reconciliation → mobiles |
| Remap | `higgs_remap.py` swaps source externalrefs → Plannr UUIDs; re-run each round after fresh export |

## Files and scripts

```
DATA_ROOT = Higgs/   (this folder)

Scripts
  higgs_remap.py                      Reference-swap and round assembly. Run: python higgs_remap.py r<n>
  build_plannr_import_pack_v3.py      Builder (stage 2). Reads Vision exports, writes stage2_build/importer/
  higgs_resolve_v4.py                 Resolver (stage 1). Reads raw Vision export, writes stage2_build/

Source data
  stage2_build/importer/              Clean importer CSVs (output of builder)
  stage2_build/higgs_Merge_Review.xlsx      Helper columns for charges + valuations
  stage2_build/higgs_Merge_Review_BT.xlsx   BT client data

Upload files
  upload/R1/  01 Clients.csv
  upload/R2/  02 Clients BT.csv  |  02 Clients_HELD_BACK.csv
  upload/R3/  03 Plans.csv  |  06 I&E  |  07 Expenses (1-5)  |  08 Notes  |  09 Circles
  upload/R4/  04 Valuations.csv  |  05 Charges.csv
  upload/R5/  03 Plans_RETRY.csv   (and xlsx)

Exports from Plannr (put fresh exports here, remap picks latest by mtime)
  plannr_exports/Migration Clients.csv
  plannr_exports/BT Client Export - 2nd Nov 2026.csv
  plannr_exports/Migration Export Plans.csv
  plannr_exports/providers_live.csv

Adviser output
  higgs_AR_actions_20261102.xlsx      (DATA_ROOT, not the parent AR Migration folder)

To run the next round after a fresh Plannr export:
  python higgs_remap.py r5     (not yet in the script — R5 = re-run r3 against latest client map)
  Note: r3 in the script rebuilds Plans + I&E + Expenses + Notes + Circles.
        R5 is conceptually a second pass of r3 for the remaining files; the files in upload/R3 are already built.
        For the valuation/charge pass after retry: run python higgs_remap.py r4 with the new plan export in plannr_exports/.
```
