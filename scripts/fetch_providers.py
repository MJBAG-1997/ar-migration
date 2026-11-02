"""Fetch current Plannr providers list and save to Higgs/ as providers_live.csv."""
import csv, json, sys, os, datetime
from azure.identity import DefaultAzureCredential
from azure.keyvault.secrets import SecretClient
import requests

BASE_URL          = "https://api.plannrcrm.com"
DEST_ACCOUNT_UUID = "ccd7f98b-d6a4-46ad-bab3-242ab42826eb"
OUT_PATH          = r"C:\Users\markj\OneDrive - Beyond Advice\Technology - Documents\workspace\AR Migration\Higgs\providers_live.csv"

def p(*a): print(*a, flush=True)

sc  = SecretClient(vault_url="https://kv-plannr-ingestion.vault.azure.net/",
                   credential=DefaultAzureCredential())
pat = sc.get_secret("plannr-prod-pat").value
s   = requests.Session()
s.headers.update({"Authorization": f"Bearer {pat}", "Accept": "application/json",
                  "X-PLANNR-ACCOUNT-UUID": DEST_ACCOUNT_UUID})

def get_paged(path, max_pages=100, params=None):
    items, page = [], 1
    while page <= max_pages:
        pr = dict(params or {}); pr["page"] = page; pr["per_page"] = 200
        r = s.get(f"{BASE_URL}{path}", params=pr)
        if r.status_code != 200:
            p(f"  [WARN] GET {path} -> {r.status_code}: {r.text[:200]}")
            break
        body = r.json()
        if isinstance(body, list):
            items.extend(body); break
        data = body.get("data", [])
        if isinstance(data, list): items.extend(data)
        elif isinstance(data, dict): items.append(data)
        meta = body.get("meta", {}) if isinstance(body, dict) else {}
        last = meta.get("last_page") or 1
        p(f"  Page {page}/{last} — {len(items)} providers so far")
        if page >= last: break
        page += 1
    return items

p("Fetching providers from Plannr...")
providers = get_paged("/api/v1/providers")
p(f"Total providers: {len(providers)}")

if providers:
    # Show keys from first record
    p(f"Fields: {list(providers[0].keys())}")

    with open(OUT_PATH, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=["uuid", "name", "created_at", "updated_at"])
        writer.writeheader()
        for prov in providers:
            writer.writerow({
                "uuid": prov.get("uuid", ""),
                "name": prov.get("name", ""),
                "created_at": prov.get("created_at", ""),
                "updated_at": prov.get("updated_at", ""),
            })
    p(f"Saved to {OUT_PATH}")
else:
    p("No providers returned — check auth/account UUID")
