"""
HFP AR Migration — Step 1 (26 Oct 2026): Fresh pull from source Plannr site.

Pulls only the exports that succeeded on 14 Oct:
  client_export, plan_export, charges_export, circle_export,
  premium_export, client_expenditure_export, note_export

Also fetches GET /api/v2/dependants?include=accounts (paged, following next link).

Skips: valuation_report (422 on 14 Oct), income_and_employment_export (server fail),
       relationships (per session instructions).

Guard: POST /api/v1/report only; everything else GET.
Rule 9: aggregates only in console output.

Saves to DATA_ROOT/exports_20261026/.
"""
import os, json, time, re, sys
from datetime import date

DATA_ROOT   = r"C:\dev\AR Migration"
OUT_DIR     = os.path.join(DATA_ROOT, "exports_20261026")
BASE_URL    = "https://api.plannrcrm.com"
KV_NAME     = "kv-plannr-ingestion"
PAT_SECRET  = "plannr-hfp-pat"
USER_AGENT  = "HFP-AR-Migration/1.0 step1_exports_20261026"

REPORT_TYPES = [
    "client_export",
    "plan_export",
    "charges_export",
    "circle_export",
    "premium_export",
    "client_expenditure_export",
    "note_export",
]

POLL_INTERVAL = 15
MAX_POLLS     = 120

def p(*a): print(*a, flush=True)


class PlannrSourceClient:
    _ALLOWED_POST = "/api/v1/report"

    def __init__(self, pat, base_url, user_agent):
        import requests
        self._s = requests.Session()
        self._s.headers.update({
            "Authorization": f"Bearer {pat}",
            "User-Agent":     user_agent,
            "Accept":         "application/json",
        })
        self._base = base_url.rstrip("/")

    def _guard(self, method, path):
        if method == "GET":
            return
        if method == "POST" and (path == self._ALLOWED_POST
                                  or path.startswith(self._ALLOWED_POST + "/")):
            return
        raise PermissionError(
            f"PlannrSourceClient guard: {method} {path!r} refused."
        )

    def get(self, path, **kw):
        self._guard("GET", path)
        return self._s.get(f"{self._base}{path}", **kw)

    def post(self, path, **kw):
        self._guard("POST", path)
        kw.setdefault("headers", {})["Content-Type"] = "application/json"
        return self._s.post(f"{self._base}{path}", **kw)

    def put(self, path, **kw):   self._guard("PUT", path)
    def patch(self, path, **kw): self._guard("PATCH", path)
    def delete(self, path, **kw): self._guard("DELETE", path)


def _raise_for_status(r, context):
    if r.status_code not in (200, 201, 202):
        raise RuntimeError(f"{context}: HTTP {r.status_code}\n{r.text[:500]}")


def create_report(client, report_type_id):
    name = report_type_id.replace("_", " ").title() + " HFP Migration " + date.today().strftime("%Y-%m-%d")
    r = client.post("/api/v1/report", json={"report_type_id": report_type_id, "name": name})
    _raise_for_status(r, f"create_report({report_type_id})")
    data = r.json()
    inner = data.get("data", data) if isinstance(data, dict) else {}
    uuid = inner.get("uuid") or inner.get("id")
    if not uuid:
        raise RuntimeError(f"create_report({report_type_id}): no UUID in response: {str(data)[:200]}")
    return uuid


def poll_report(client, report_uuid, label):
    for i in range(MAX_POLLS):
        r = client.get(f"/api/v1/report/{report_uuid}")
        _raise_for_status(r, f"poll_report({label})")
        data = r.json()
        inner = data.get("data", data) if isinstance(data, dict) else {}
        status = (inner.get("status") or "").lower()
        if status == "complete":
            return inner
        if status in ("failed", "error"):
            raise RuntimeError(f"poll_report({label}): status={status!r}")
        p(f"    [{label}] status={status!r}  elapsed={(i+1)*POLL_INTERVAL}s")
        time.sleep(POLL_INTERVAL)
    raise RuntimeError(f"poll_report({label}): timed out")


def download_file(client, file_uuid, dest_dir, default_name):
    r = client.get(f"/api/v1/file/{file_uuid}/download", stream=True)
    _raise_for_status(r, f"download_file({file_uuid})")
    cd = r.headers.get("Content-Disposition", "")
    m = re.search(r'filename[^;=\n]*=([\'"]?)([^\'";\n]+)\1', cd)
    filename = m.group(2).strip() if m else default_name
    filename = re.sub(r'[\\/:*?"<>|]', "_", filename)
    path = os.path.join(dest_dir, filename)
    with open(path, "wb") as fh:
        for chunk in r.iter_content(chunk_size=65536):
            if chunk:
                fh.write(chunk)
    return path, os.path.getsize(path)


def fetch_dependants(client, dest_dir):
    """Page through GET /api/v2/dependants?include=accounts following next link."""
    rows = []
    url = "/api/v2/dependants"
    params = {"per_page": 250, "include": "accounts"}
    page = 1
    while True:
        r = client.get(url, params=params)
        _raise_for_status(r, f"fetch_dependants page {page}")
        body = r.json()

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

        # Follow the next link
        links = body.get("links", {}) if isinstance(body, dict) else {}
        next_url = links.get("next")
        if not next_url:
            inner_links = data_outer.get("links", {}) if isinstance(data_outer, dict) else {}
            next_url = inner_links.get("next")
        if not next_url:
            meta = body.get("meta", {}) if isinstance(body, dict) else {}
            last_page = meta.get("last_page") or meta.get("total_pages")
            if last_page and page >= int(last_page):
                break

        if next_url:
            # next_url is an absolute URL; strip base and params for our client
            from urllib.parse import urlparse, parse_qs
            parsed = urlparse(next_url)
            url = parsed.path
            params = {k: v[0] for k, v in parse_qs(parsed.query).items()}
            page += 1
        else:
            break

    path = os.path.join(dest_dir, "dependants.json")
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(rows, fh, indent=2, ensure_ascii=False)
    return path, len(rows)


def resolve_pat():
    from azure.identity import DefaultAzureCredential
    from azure.keyvault.secrets import SecretClient
    vault_uri = f"https://{KV_NAME}.vault.azure.net"
    sc = SecretClient(vault_url=vault_uri, credential=DefaultAzureCredential())
    return sc.get_secret(PAT_SECRET).value


def main():
    p("=" * 60)
    p("HFP AR Migration — Step 1 (26 Oct 2026): Pull exports")
    p("=" * 60)

    p("\n[1] Resolving PAT...")
    pat = resolve_pat()
    p(f"  PAT resolved (length={len(pat)})")

    p("\n[2] Creating source client")
    client = PlannrSourceClient(pat, BASE_URL, USER_AGENT)

    p("\n[3] Confirming source account UUID")
    r = client.get("/api/v1/me")
    _raise_for_status(r, "GET /api/v1/me")
    me = r.json()
    inner = me.get("data", me) if isinstance(me, dict) else {}
    current_account = inner.get("current_account", {}) or {}
    source_uuid = current_account.get("uuid") or inner.get("uuid")
    p(f"  SOURCE_ACCOUNT_UUID: {source_uuid}")
    client._s.headers.update({"X-PLANNR-ACCOUNT-UUID": source_uuid})

    p(f"\n[4] Output directory: {OUT_DIR}")
    os.makedirs(OUT_DIR, exist_ok=True)

    p(f"\n[5] Pulling {len(REPORT_TYPES)} report exports")
    results = {}
    for rtype in REPORT_TYPES:
        p(f"\n  [{rtype}]")
        try:
            report_uuid = create_report(client, rtype)
            p(f"    report_uuid: {report_uuid}")
            report_data = poll_report(client, report_uuid, rtype)
            p(f"    status: complete")
            files_list = report_data.get("files", [])
            file_uuid = files_list[0].get("uuid") if files_list else None
            if file_uuid:
                path, size = download_file(client, file_uuid, OUT_DIR, f"{rtype}.csv")
                p(f"    saved: {os.path.basename(path)}  ({size:,} bytes)")
                results[rtype] = {"status": "ok", "path": path, "size": size}
            else:
                p(f"    WARN: no file_uuid — keys: {list(report_data.keys())}")
                results[rtype] = {"status": "no_file"}
        except Exception as e:
            p(f"    ERROR: {e}")
            results[rtype] = {"status": "error", "error": str(e)}

    p(f"\n[6] Fetching dependants (include=accounts)")
    try:
        dep_path, dep_count = fetch_dependants(client, OUT_DIR)
        p(f"  saved: dependants.json  ({dep_count:,} rows)")
        results["dependants"] = {"status": "ok", "path": dep_path, "count": dep_count}
    except Exception as e:
        p(f"  ERROR: {e}")
        results["dependants"] = {"status": "error", "error": str(e)}

    manifest_path = os.path.join(OUT_DIR, "_manifest.json")
    with open(manifest_path, "w", encoding="utf-8") as fh:
        json.dump({
            "source_account_uuid": source_uuid,
            "pulled_at": date.today().strftime("%Y%m%d"),
            "results": results,
        }, fh, indent=2)

    p("\n" + "=" * 60)
    ok = sum(1 for v in results.values() if v["status"] == "ok")
    err = sum(1 for v in results.values() if v["status"] == "error")
    p(f"  OK: {ok}  Error: {err}  Dir: {OUT_DIR}")
    if err:
        for k, v in results.items():
            if v["status"] == "error":
                p(f"    FAILED: {k}: {v['error']}")
    p("Done.")


if __name__ == "__main__":
    main()
