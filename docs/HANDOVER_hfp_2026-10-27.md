# HFP Import — Handover 2026-10-27

Run date: 2026-10-27. Branch: `feature/ar-migration-hfp`.
Destination account: `ccd7f98b-d6a4-46ad-bab3-242ab42826eb` (HFP sub-account).

---

## Current state

| Wave | Status | Rows | Notes |
|------|--------|------|-------|
| W1 — individual clients | **Stored** | 272 created | All 272 source→dest UUIDs resolved |
| W2 — joint/company clients | **Validate failed** | 2 failures | See decision needed below |
| W2 — circles | Not started | — | |
| W3 — relationships | Not started | — | |
| W3 — dependants | Not started | — | |
| W3 — income & employment | Not started | — | |
| Expenditure (API) | Not started | — | |
| W4 — ATP plan | Not started | — | Awaiting W3 + expenditure clean |
| W5 — plans | Not started | — | 1,901 total (W4+W5); Vision sweep due first |
| W6 — valuations | Not started | — | 1,526 rows; AUM £323,134,346 |
| W7 — charges | Not started | — | Awaiting separate go-ahead |
| W8 — notes | Not started | — | Awaiting separate go-ahead |

Nothing beyond W1 has been written to the destination.

---

## Decision needed before W2 can proceed

Two business-type accounts in W2_clients.csv have an empty `owners` field in the source data. Plannr's API rejects them with `"The owners must be an array"` — business and joint accounts require at least one individual client as owner.

Both are internal adviser accounts (email-style entity names). Source ext-ref UUIDs:

- `5a38c4d6-02cd-4b00-ad38-2a9d35971c45` — adviser who has **no individual client record** in W1
- `9450d9c3-3b7e-4815-80e5-79c185be85df` — adviser who **does** have an individual client record in W1

**Options:**

A. Remove both rows from `W2_clients.csv` and continue. (Safest if they're internal-only accounts with no client data.)

B. For `9450d9c3-…` only: add the adviser's W1 individual client UUID as owner before re-running. For `5a38c4d6-…`: remove (no W1 record exists). The W1 UUID to use for option B is at key `d1f4f77b-d582-4fc9-92ce-4cf8719d77b1` in `state/w1_client_uuid_map.json`.

C. Keep both and import without owners — not possible; API will always reject.

Instruct the next session which option to take. The runner will stop again at W2 validation until this is resolved.

---

## Remaining waves and approval rules

Approved (carry through if each wave is 0 failed rows and 0 unmatched records):

1. **W2** — joint/company clients (86 rows, 2 currently failing); circles (87 rows)
2. **W3** — relationships (182), dependants (249), income & employment (316)
3. **Expenditure** — via direct API POST, not import endpoint
4. **Vision sweep** — before W4: count Vision cells by column in W4–W7 files (excluding notes columns); blank any whole-cell Vision firm name. Report counts.
5. **W4** (1 row — Add-to-Plannr plan) and **W5** (1,900 rows); UUID rewrite on `client_uuids` (hard) and `visible_to_circles_uuids` (soft — blank unmatched, don't stop)
6. **W6** valuations (1,526 rows); UUID rewrite on `plan_uuid` (hard)
7. **STOP** after W6 with full report.

**W7 (542 charges) and W8 (6,796 notes): await separate explicit go-ahead.** Do not start without a new "go ahead W7 and W8" instruction.

**Stop rules:** stop at the first wave with any failed rows or any unmatched UUID. Any unmatched record stops the run entirely.

**Transact ATP plans (26 of 28 Transact plans):** go in without valuations — approved 2026-10-27. A valuation feed will populate values once the plans are in.

---

## Input file locations (DATA_ROOT = `C:\dev\AR Migration`)

| File | Path | Rows | Size |
|------|------|------|------|
| W1 individuals | `import_20261027\W1_clients.csv` | 272 | 161 KB |
| W2 joint/company | `import_20261027\W2_clients.csv` | 86 | 36 KB |
| W2 circles | `import_20261027\W2_circles.csv` | 87 | 14 KB |
| W3 relationships | `import_20261027\W3_relationships.csv` | 182 | 16 KB |
| W3 dependants | `import_20261027\W3_dependants.csv` | 249 | 42 KB |
| W3 income | `import_20261027\W3_income.csv` | 316 | 104 KB |
| W4 ATP plan | `import_20261027\W4_plans.csv` | 1 | 5 KB |
| W5 plans | `import_20261027\W5_plans.csv` | 1,900 | 886 KB |
| W6 valuations | `import_20261027\W6_valuations.csv` | 1,526 | 100 KB |
| W6 premiums | `import_20261027\W6_premiums.csv` | 171 | 18 KB |
| W7 charges | `import_20261027\W7_charges.csv` | 542 | 121 KB |
| W8 notes | `import_20261027\W8_notes.csv` | 6,796 | 4.3 MB |
| Expenditure export | `exports_20261026\Client Expenditure Export HFP Migration 2026-10-27.csv` | — | — |
| W1 UUID map | `import_20261027\state\w1_client_uuid_map.json` | 272 entries | 22 KB |
| Run state | `import_20261027\state\run_state.json` | — | 1.3 KB |

Source exports (read-only reference): `exports_20261026\` (26 Oct) and `exports_20261014\` (14 Oct).

Build script (rebuilds import CSVs from exports): `scripts\build_import_20261027.py`

---

## Resume command

From a terminal in `C:\dev\AR Migration`:

```
python scripts\run_import_20261027.py
```

The runner auto-detects that W1 is already stored (checks whether all 272 W1 source UUIDs resolve against destination external references). It will skip W1 and go straight to W2. Pre-flight employee guard runs every time.

**Before resuming:** resolve the W2_clients decision above (edit `W2_clients.csv` to remove or fix the two failing rows), then run.

---

## Rebuilding the UUID map without the state file

If `state/w1_client_uuid_map.json` is missing, the runner rebuilds it live at startup by calling
`GET /api/v1/client?include=external_references` (paginated, `per_page=200`) and matching
`external_references[*].third_party == "data_migration"`. The `resolve_client_uuids()` function in
`scripts/run_import_20261027.py` does exactly this; it also runs automatically at every pre-flight.

To rebuild manually:

```python
import json, requests
from azure.identity import DefaultAzureCredential
from azure.keyvault.secrets import SecretClient

sc = SecretClient(vault_url="https://kv-plannr-ingestion.vault.azure.net/",
                  credential=DefaultAzureCredential())
pat = sc.get_secret("plannr-prod-pat").value

s = requests.Session()
s.headers.update({
    "Authorization": f"Bearer {pat}",
    "Accept": "application/json",
    "X-PLANNR-ACCOUNT-UUID": "ccd7f98b-d6a4-46ad-bab3-242ab42826eb",
})

uuid_map, page = {}, 1
while True:
    r = s.get("https://api.plannrcrm.com/api/v1/client",
              params={"page": page, "per_page": 200, "include": "external_references"})
    body = r.json()
    for c in body.get("data", []):
        for ref in (c.get("external_references") or []):
            if ref.get("third_party") == "data_migration":
                uuid_map[ref["reference"]] = c["uuid"]
    if page >= body.get("meta", {}).get("last_page", 1):
        break
    page += 1

with open(r"C:\dev\AR Migration\import_20261027\state\w1_client_uuid_map.json", "w") as f:
    json.dump(uuid_map, f, indent=2)
```

---

## Technical notes for next session

- **Import type strings** (Plannr API): `clients`, `circles`, `relationships`, `dependants`, `income_and_employment`, `plans`, `valuations`, `charges`, `notes` — all plural. `client` (singular) returns HTTP 422 "The selected type is invalid."
- **External references** are not included in GET `/api/v1/client` by default — must pass `?include=external_references`.
- **W2_clients rewrite:** `owners` field contains source client UUIDs; rewritten to dest UUIDs before upload. Rewritten file: `W2_clients_rewritten.csv` (regenerated each run).
- **W4/W5 rewrite fields:** `client_uuids` (hard — stop on unmatch), `visible_to_circles_uuids` (soft — blank on unmatch, do not stop).
- **W6 BOM:** read with `encoding="utf-8-sig"` to strip BOM from `plan_uuid` column.
- **Expenditure API:** `POST /api/v1/client/{dest_uuid}/expenditure/{life_stage_uuid}/{field_uuid}`, amount in pence. Structure from `GET /api/v1/expenditure-structure/section`.
- **Manual phones:** 8 non-+44 numbers blanked from W1; listed in `hfp_build_lists_20261027.xlsx` "Manual phones" tab for post-import entry.

---

## Previous handovers

- `docs/HANDOVER_hfp_2026-10-14.md` — build pack and AR review
- `docs/HANDOVER_hfp_2026-10-13.md` — initial analysis
