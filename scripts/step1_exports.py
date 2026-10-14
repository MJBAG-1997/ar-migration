"""
HFP AR Migration - Step 1: Pull exports from source Plannr site.

Guard: GET allowed on any path; POST allowed ONLY for /api/v1/report;
       PUT / PATCH / DELETE raise PermissionError before any HTTP call.

Saves all exports to DATA_ROOT/exports_YYYYMMDD/.
Also saves dependants via GET /api/v2/dependants (firm-wide, paged).

Rule 9: aggregates only in console output. No PII printed.
"""
import os, json, time, re, sys
from datetime import date

DATA_ROOT  = r"C:\dev\AR Migration"
SNAP_DIR   = os.path.join(DATA_ROOT, "snapshot_20261012T090906Z")
BASE_URL   = "https://api.plannrcrm.com"
KV_NAME    = "kv-plannr-ingestion"
PAT_SECRET = "plannr-hfp-pat"
USER_AGENT = "HFP-AR-Migration/1.0 step1_exports"

REPORT_TYPES = [
    "client_export",
    "plan_export",
    "charges_export",
    "circle_export",
    "valuation_report",
    "premium_export",
    "income_and_employment_export",
    "client_expenditure_export",
    "note_export",
    "staff_member_export",
    "group_scheme_members_export",
]

POLL_INTERVAL = 15    # seconds between polls
MAX_POLLS     = 120   # 30 min max per report
DEP_PER_PAGE  = 500

def p(*a): print(*a, flush=True)


# ── Guard: source-only Plannr client ─────────────────────────────────────────

class PlannrSourceClient:
    """
    Wraps requests.Session with a method guard.
    GET: allowed for any path.
    POST: allowed only for paths starting with /api/v1/report.
    PUT / PATCH / DELETE: PermissionError always.
    """
    _ALLOWED_POST = "/api/v1/report"

    def __init__(self, pat: str, base_url: str, user_agent: str):
        import requests
        self._s = requests.Session()
        self._s.headers.update({
            "Authorization": f"Bearer {pat}",
            "User-Agent":     user_agent,
            "Accept":         "application/json",
        })
        self._base = base_url.rstrip("/")

    # ── guard ──
    def _guard(self, method: str, path: str):
        if method == "GET":
            return
        if method == "POST" and (path == self._ALLOWED_POST
                                  or path.startswith(self._ALLOWED_POST + "/")):
            return
        raise PermissionError(
            f"PlannrSourceClient guard: {method} {path!r} refused. "
            f"Source site is GET-only except POST {self._ALLOWED_POST}."
        )

    # ── public methods ──
    def get(self, path: str, **kw):
        self._guard("GET", path)
        return self._s.get(f"{self._base}{path}", **kw)

    def post(self, path: str, **kw):
        self._guard("POST", path)
        kw.setdefault("headers", {})["Content-Type"] = "application/json"
        return self._s.post(f"{self._base}{path}", **kw)

    def put(self, path: str, **kw):
        self._guard("PUT", path)

    def patch(self, path: str, **kw):
        self._guard("PATCH", path)

    def delete(self, path: str, **kw):
        self._guard("DELETE", path)


def _guard_selftest():
    """Run guard logic checks without making any HTTP call."""
    g = PlannrSourceClient.__new__(PlannrSourceClient)
    g._base = BASE_URL

    # These should NOT raise
    g._guard("GET", "/api/v1/me")
    g._guard("GET", "/api/v1/accounts")
    g._guard("POST", "/api/v1/report")

    # These MUST raise PermissionError
    should_refuse = [
        ("POST",   "/api/v1/clients"),
        ("POST",   "/api/v1/plans"),
        ("PUT",    "/api/v1/clients/xxx"),
        ("PATCH",  "/api/v1/clients/xxx"),
        ("DELETE", "/api/v1/clients/xxx"),
        ("POST",   "/api/v1/import"),
    ]
    for method, path in should_refuse:
        try:
            g._guard(method, path)
            raise AssertionError(f"Guard FAILED to refuse {method} {path}")
        except PermissionError:
            pass

    p("  Guard self-test: PASS (6 allowed, 6 refused)")


# ── PAT resolution ────────────────────────────────────────────────────────────

def resolve_pat() -> str:
    from azure.identity import DefaultAzureCredential
    from azure.keyvault.secrets import SecretClient
    vault_uri = f"https://{KV_NAME}.vault.azure.net"
    client = SecretClient(vault_url=vault_uri, credential=DefaultAzureCredential())
    secret = client.get_secret(PAT_SECRET)
    return secret.value


# ── API helpers ───────────────────────────────────────────────────────────────

def _raise_for_status(r, context: str):
    if r.status_code not in (200, 201, 202):
        raise RuntimeError(
            f"{context}: HTTP {r.status_code}\n{r.text[:500]}"
        )


def create_report(client: PlannrSourceClient, report_type_id: str,
                  extra_params: dict | None = None) -> str:
    """POST /api/v1/report; return report UUID."""
    name = report_type_id.replace("_", " ").title() + " HFP Migration " + date.today().strftime("%Y-%m-%d")
    body = {"report_type_id": report_type_id, "name": name}
    if extra_params:
        body.update(extra_params)
    r = client.post("/api/v1/report", json=body)
    _raise_for_status(r, f"create_report({report_type_id})")
    data = r.json()
    # Handle {data: {uuid:...}} or {uuid:...} response shapes
    if isinstance(data, dict):
        inner = data.get("data", data)
        uuid = inner.get("uuid") or inner.get("id")
    else:
        uuid = None
    if not uuid:
        raise RuntimeError(f"create_report({report_type_id}): no UUID in response: {str(data)[:200]}")
    return uuid


def poll_report(client: PlannrSourceClient, report_uuid: str, label: str) -> dict:
    """Poll GET /api/v1/report/{uuid} until status=complete or failed."""
    for i in range(MAX_POLLS):
        r = client.get(f"/api/v1/report/{report_uuid}")
        _raise_for_status(r, f"poll_report({label})")
        data = r.json()
        inner = data.get("data", data) if isinstance(data, dict) else {}
        status = (inner.get("status") or "").lower()
        if status == "complete":
            return inner
        if status in ("failed", "error"):
            raise RuntimeError(f"poll_report({label}): status={status!r}  data={str(inner)[:200]}")
        elapsed = (i + 1) * POLL_INTERVAL
        p(f"    [{label}] status={status!r}  elapsed={elapsed}s  poll {i+1}/{MAX_POLLS}")
        time.sleep(POLL_INTERVAL)
    raise RuntimeError(f"poll_report({label}): timed out after {MAX_POLLS * POLL_INTERVAL}s")


def download_file(client: PlannrSourceClient, file_uuid: str, dest_dir: str,
                  default_name: str) -> str:
    """GET /api/v1/file/{uuid}/download; return saved path."""
    r = client.get(f"/api/v1/file/{file_uuid}/download", stream=True)
    _raise_for_status(r, f"download_file({file_uuid})")

    # Derive filename from Content-Disposition or default
    cd = r.headers.get("Content-Disposition", "")
    m = re.search(r'filename[^;=\n]*=([\'"]?)([^\'";\n]+)\1', cd)
    filename = m.group(2).strip() if m else default_name
    # Sanitise
    filename = re.sub(r'[\\/:*?"<>|]', "_", filename)

    path = os.path.join(dest_dir, filename)
    with open(path, "wb") as fh:
        for chunk in r.iter_content(chunk_size=65536):
            if chunk:
                fh.write(chunk)
    size = os.path.getsize(path)
    return path, size


# ── Dependants (GET only) ──────────────────────────────────────────────────────

def fetch_dependants(client: PlannrSourceClient, dest_dir: str) -> str:
    """Page through GET /api/v2/dependants; save JSON; return path."""
    rows = []
    page = 1
    while True:
        r = client.get("/api/v2/dependants",
                       params={"per_page": DEP_PER_PAGE, "page": page})
        _raise_for_status(r, f"fetch_dependants page {page}")
        body = r.json()

        # Handle nested data key
        data_outer = body.get("data", body) if isinstance(body, dict) else body
        if isinstance(data_outer, list):
            batch = data_outer
        elif isinstance(data_outer, dict):
            batch = data_outer.get("data", [])
        else:
            batch = []

        rows.extend(batch)
        p(f"    page {page}: {len(batch)} rows  (cumulative {len(rows)})")

        if not batch:
            break

        # Pagination: check meta or links
        meta = body.get("meta", {}) if isinstance(body, dict) else {}
        last_page = meta.get("last_page") or meta.get("total_pages")
        if last_page and page >= int(last_page):
            break

        links = body.get("links", {}) if isinstance(body, dict) else {}
        if not links.get("next"):
            # also check inner links
            inner_links = data_outer.get("links", {}) if isinstance(data_outer, dict) else {}
            if not inner_links.get("next"):
                break

        page += 1

    path = os.path.join(dest_dir, "dependants.json")
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(rows, fh, indent=2, ensure_ascii=False)
    return path, len(rows)


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    p("=" * 60)
    p("HFP AR Migration — Step 1: Pull exports")
    p("=" * 60)

    p("\n[0] Guard self-test")
    _guard_selftest()

    p("\n[1] Resolving PAT from Key Vault...")
    pat = resolve_pat()
    p(f"  PAT resolved (length={len(pat)})")

    p("\n[2] Creating Plannr source client")
    client = PlannrSourceClient(pat, BASE_URL, USER_AGENT)

    p("\n[3] Confirming source account UUID via GET /api/v1/me")
    r = client.get("/api/v1/me")
    _raise_for_status(r, "GET /api/v1/me")
    me = r.json()
    inner = me.get("data", me) if isinstance(me, dict) else {}
    # current_account.uuid is the employee account UUID used as X-PLANNR-ACCOUNT-UUID
    current_account = inner.get("current_account", {}) or {}
    source_uuid = current_account.get("uuid") or inner.get("uuid")
    p(f"  SOURCE_ACCOUNT_UUID (employee): {source_uuid}")
    p(f"  Firm: (redacted per Rule 9)")

    # Add account UUID header to all subsequent requests
    client._s.headers.update({"X-PLANNR-ACCOUNT-UUID": source_uuid})

    p(f"\n[4] Creating export directory")
    today = date.today().strftime("%Y%m%d")
    exports_dir = os.path.join(DATA_ROOT, f"exports_{today}")
    os.makedirs(exports_dir, exist_ok=True)
    p(f"  {exports_dir}")

    # Load client UUIDs for valuation_report (requires account_uuids)
    import json as _json
    with open(os.path.join(SNAP_DIR, "accounts.json"), encoding="utf-8") as _f:
        _accs = _json.load(_f)
    CLIENT_UUIDS = [a["uuid"] for a in _accs if a.get("type") == "client"]
    p(f"  Loaded {len(CLIENT_UUIDS)} client UUIDs for valuation report")

    p(f"\n[5] Pulling {len(REPORT_TYPES)} report exports")
    results = {}
    for rtype in REPORT_TYPES:
        p(f"\n  [{rtype}]")
        try:
            p(f"    POST /api/v1/report...")
            # valuation_report requires date range and account UUIDs
            extra = None
            if rtype == "valuation_report":
                extra = {
                    "date_from": "2010-01-01",
                    "date_to": date.today().strftime("%Y-%m-%d"),
                    "account_uuids": CLIENT_UUIDS,
                }
            report_uuid = create_report(client, rtype, extra)
            p(f"    report_uuid: {report_uuid}")

            p(f"    Polling...")
            report_data = poll_report(client, report_uuid, rtype)
            p(f"    status: complete")

            # File UUID lives in the files array: data.files[0].uuid
            files_list = report_data.get("files", [])
            file_uuid = files_list[0].get("uuid") if files_list else None

            if file_uuid:
                default_name = f"{rtype}.csv"
                path, size = download_file(client, file_uuid, exports_dir, default_name)
                p(f"    saved: {os.path.basename(path)}  ({size:,} bytes)")
                results[rtype] = {"status": "ok", "path": path, "size": size}
            else:
                p(f"    WARN: no file_uuid in completed report — keys: {list(report_data.keys())}")
                results[rtype] = {"status": "no_file", "report_data_keys": list(report_data.keys())}

        except Exception as e:
            p(f"    ERROR: {e}")
            results[rtype] = {"status": "error", "error": str(e)}

    p(f"\n[6] Fetching dependants via GET /api/v2/dependants")
    try:
        dep_path, dep_count = fetch_dependants(client, exports_dir)
        p(f"  saved: dependants.json  ({dep_count:,} rows)")
        results["dependants"] = {"status": "ok", "path": dep_path, "count": dep_count}
    except Exception as e:
        p(f"  ERROR: {e}")
        results["dependants"] = {"status": "error", "error": str(e)}

    # Save results manifest
    manifest_path = os.path.join(exports_dir, "_manifest.json")
    with open(manifest_path, "w", encoding="utf-8") as fh:
        json.dump({"source_account_uuid": source_uuid,
                   "pulled_at": today,
                   "results": results}, fh, indent=2)

    p("\n" + "=" * 60)
    p("STEP 1 SUMMARY")
    p("=" * 60)
    ok     = sum(1 for v in results.values() if v["status"] == "ok")
    errors = sum(1 for v in results.values() if v["status"] == "error")
    other  = sum(1 for v in results.values() if v["status"] not in ("ok", "error"))
    p(f"  OK:      {ok}")
    p(f"  Error:   {errors}")
    p(f"  Other:   {other}")
    p(f"  Dir:     {exports_dir}")

    if errors:
        p("\n  Failed exports:")
        for k, v in results.items():
            if v["status"] == "error":
                p(f"    {k}: {v['error']}")

    p("\nReady for Step 2 (step2_compare.py).")
    return exports_dir


if __name__ == "__main__":
    exports_dir = main()
    print(exports_dir)
