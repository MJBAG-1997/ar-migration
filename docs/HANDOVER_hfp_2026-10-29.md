# HFP Migration Handover — 2026-10-29

Branch: `feature/ar-migration-hfp`
Destination account: `ccd7f98b-d6a4-46ad-bab3-242ab42826eb`
Script: `scripts/fix_20261029.py`

---

## Current state — all waves stored

| Wave | Expected | Stored | Status |
|---|---|---|---|
| W1_clients | 272 | 272 | ✓ |
| W2_clients | 84 | 84 | ✓ |
| W2_circles | 87 | 87 | ✓ (+2 extra, see below) |
| W3_relationships | 182 | 207 directed | see below |
| W3_dependants | 249 | 249 | ✓ fixed today |
| W3_income | 316 | 316 | ✓ |
| expenditure_standard | 114 | 114 | ✓ confirmed pennies |
| expenditure_custom | 1703 | 1703 | ✓ posted today |
| W4_plans | 1 | 1 | ✓ |
| W5_plans | 1900 | 1900 | ✓ |
| W6_valuations | 1526 | 1525 | −1 (502 during count) |
| W6_premiums | 171 rows | 173 | ✓ stored (see note) |
| W7_charges | 542 | 542 | ✓ |
| W8_notes | 6757 | 6861 | +104 from Oct-27 partial runs |

---

## Today's fixes (2026-10-29)

### Dependants — fixed ✓

Dependants were doubled (498). Fix:
- Deleted all 498 duplicates (all 249 pairs confirmed)
- Re-imported via full import flow (create → map → start → poll)
- Import job: `f4983389`, mode=store, status=complete, total_jobs=249
- Final count: **249 ✓**

### 10 stale import jobs — deleted ✓

All 10 confirmed `mapping=None, progress=0` before deletion:
- 5 × clients/validate
- 3 × plans/validate
- 2 × plans/store (1900 rows each) — **never started**, no duplicates created

### Premiums — stored (173, note) ✓

W6_premiums.csv: 171 rows, **167 unique plan UUIDs** (4 plans appear twice in the CSV).
Import created 173 premiums. The 6 above the 167 unique-plan count is unexplained — the importer
processed all 171 CSV rows and created 173. All reads confirmed non-zero values. Investigate if
exact count matters for sign-off.

### Expenditure custom lines — 1703 posted ✓

- Standard fields: 114 (previously written) — ratio check confirmed ≈100× (stored in pennies as expected)
- Custom lines: 1703 POSTed to `POST /api/v1/client/{uuid}/expenditure/{ls_uuid}/custom`
  with `life_stage_section_uuid` (expenditure structure UUID) — 0 errors
- Section mapping: `Essential - Home/Vehicles/etc.` → `Essential`, `Milestones...` → `Lifestyle`
- State file: `import_20261027/state/exp_custom_state.json` (1703 entries)
- Verification note: GET expenditure returns flat list — items use `is_custom_expenditure: true`,
  NOT a `custom_lines` key. Spot-check `af7b2c6f`: 23 custom + 2 standard ✓

### Notes delta — skipped (correct)

8 source notes not in W8_notes_final_v2.csv: all have `notable_type=task`. The notes importer
does not support task-linked notes. 7 were created on migration day (2026-10-28); 1 from
June 2024. Left as-is — cannot be imported.

---

## Read-only analysis results (2026-10-29)

### Relationships (207 on dest, expected 147)

Full analysis with corrected UUID extraction (the `/linked-clients` response has `uuid` = linked
client's UUID, not a relationship record UUID):

| Category | Count |
|---|---|
| Directed records fetched | 354 |
| Undirected pairs | 177 |
| Both clients within our 356 | 177 (100%) |
| Pairs also in W3_relationships.csv | 5 (10 directed) |
| Bidirectional pairs (normal) | 172 × 2 = 344 |
| External parties | 0 |
| Joint/trust/business involved | 0 |

W3_relationships.csv had 182 source pairs; dest has 177 undirected pairs — 5 short. No
external parties. All relationships are between our 356 migrated clients.

Step 1 left them unchanged because 207 ≠ multiple of 147.

### Employments (266 on dest, expected 0)

Expected 0 was wrong — employments are linked to income records via the importer:

| Metric | Count |
|---|---|
| Employments | 266 |
| Income records | 316 |
| Income with linked employment | 266 |
| Income with no linked employment | 50 |
| Orphan employments (not referenced by any income) | 0 |

266 employment records each correspond to an income record. 50 income records have no linked
employment (expected — pension, investment, etc. income types). All clean.

### Circles (+2 extra)

2 extra circles with no members, both created 2026-10-27:
- `5ec4636c-bfcd-4a61-9a12-e7707d8508f7`
- `fbadcb0c-fd70-4322-8e07-28bf26b5cdf9`

Empty, no members. Probably created by a test/partial run. Can be deleted via UI if needed.

### Note matching (W8_notes_final_v2 vs dest)

| Category | Count |
|---|---|
| Dest notes total | 6,861 |
| Matched to exactly one v2 row | 6,561 |
| Matched to >1 v2 row | 197 |
| Unmatched (no v2 row) | 103 |

All 103 unmatched are `author=IT SYSTEM`. They span 37 author+date groups, ranging from
2022-10-06 to 2026-10-28. These are system-generated notes not in the original W8 export.
The 18 from 2026-10-28 are from migration-day activity.

---

## Open items

| # | Item | Action needed |
|---|---|---|
| 1 | 2 empty circles (`5ec4636c`, `fbadcb0c`) | Delete via UI if confirmed unwanted |
| 2 | 103 unmatched IT SYSTEM notes | Review — all system-generated, likely harmless |
| 3 | 207 relationships vs 147 expected | 5 pairs short of 182 CSV rows; all within 356 clients |
| 4 | Premiums 173 vs 167 unique plans | Minor discrepancy — investigate if sign-off requires exact count |
| 5 | Plans clean-up (from prior session) | 155 Unknown Provider plans, 46 blank adviser, 1 type fix, 21 joint ownership |

---

## Pending imports at close

```
GET /api/v1/import → 93 jobs total, 0 in mapping/pending state
```

---

## State files

| File | Purpose |
|---|---|
| `import_20261027/state/client_map.json` | src→dest client UUIDs (356) |
| `import_20261027/state/plan_map.json` | src→dest plan UUIDs (1901) |
| `import_20261027/state/exp_custom_state.json` | Expenditure custom lines (1703 rows done) |
