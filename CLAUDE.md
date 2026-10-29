# HFP Migration — Standing Rules

## Imports

- Every import goes through the runner's import function: create, map headers, start, poll, then count records on the destination. Never write new import code.
- If an import's progress hasn't moved for 5 minutes, stop waiting. GET the import and report `status`, `mapping`, `total_jobs`, `progress` and `failed_jobs`, then act on what the fields show (usually not mapped or not started).
- Never create a second import for the same data while one exists that isn't complete or deleted.

## Waiting and reporting

- Never wait on anything for more than 5 minutes without reporting.
- If a step is blocked, carry on with the steps that don't depend on it, then come back to it.
- No diagnosis without evidence: quote the response field or error that shows the cause. Never blame Plannr without one.

## Writes

- Count before, write once, count after. If after minus before isn't the expected number, STOP.
- The API takes pennies; the importer takes pounds. Before any bulk write, write one record, read it back and compare with the source.
- Only one script talks to Plannr at a time.
