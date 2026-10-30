# HFP Migration Completion — 2026-10-29

## Run summary

| Section | Status | Notes |
|---------|--------|-------|
| B1 | Complete | 115 plans SET provider=null; 407 on AR list |
| E1 | Complete | 4 missing premiums posted; total 171 |
| E3 | Complete | 8 task notes imported; all 8 confirmed on dest clients (2026-10-30) |
| E5 | Complete | Flexi drawdown → personal pension; old plan deleted |
| E6 | Complete | 73 joint records deleted; 0 skipped (expenditure unverified) |
| E7 | Complete | 0 stale imports deleted |
| E8 | Complete | AR workbook saved to hfp_AR_actions_20261029.xlsx |
| C  | Complete | 1,093 / 1,110 files copied; 17 failed (see below) |

## Morning checks — 2026-10-30

| Check | Result |
|-------|--------|
| E3 verification | 8 / 8 notes found on dest clients |
| Duplicates | 0 clients with more files in dest than source |
| Date mismatches | 0 (original_created_at already matches created_at on all 1,093 copies) |
| Unknown Provider clearing | 407 → 210 cleared (provider_uuid set null); 50 refused (422 — Plannr enforces provider on mortgages/property_holdings/trust); 197 remain in workbook |

## C — 17 failed files

| failure | count | cause |
|---------|-------|-------|
| upload 403 | 13 | Plannr dest rejected (filenames with .aspx, spaces, parentheses) |
| download 403 | 2 | Source files inaccessible — client 4a05c772 (both) |
| download 404 | 1 | File deleted on source — client d2ef7c74 |
| folder create | 1 | Could not create/find folder — client 9eab42b4 |

These 17 files were not retried; retry deferred to post-go-live.

## Outstanding AR actions
See `hfp_AR_actions_20261029.xlsx`:
- **197 plans** need providers assigned or investigation (Tab: Unknown Provider)
  - 147 are financial products (pensions, ISAs, investments, protection) needing a real provider
  - 50 are mortgages/property_holdings/trust that Plannr refused to null (422 — provider required)
- Review income defaults, charge rates, policy numbers, phone numbers

## Skipped joints (expenditure sync unverified)
None
