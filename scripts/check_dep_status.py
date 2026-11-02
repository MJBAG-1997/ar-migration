"""One-shot: print import status + current dependant count. No looping."""
import json, requests, sys
from azure.identity import DefaultAzureCredential
from azure.keyvault.secrets import SecretClient

sc = SecretClient(vault_url="https://kv-plannr-ingestion.vault.azure.net/",
                  credential=DefaultAzureCredential())
pat = sc.get_secret("plannr-prod-pat").value
s = requests.Session()
s.headers.update({"Authorization": f"Bearer {pat}", "Accept": "application/json",
                  "X-PLANNR-ACCOUNT-UUID": "ccd7f98b-d6a4-46ad-bab3-242ab42826eb"})

r = s.get("https://api.plannrcrm.com/api/v1/import/f4983389-5e72-40fb-b1a3-165fefffba92")
d = r.json().get("data", r.json())
print("Import status:", d.get("status"))
print("Import keys:", sorted(d.keys()))

client_map = json.load(open(r"C:\dev\AR Migration\import_20261027\state\client_map.json"))
seen = set()
for du in client_map.values():
    pg = 1
    while True:
        r2 = s.get("https://api.plannrcrm.com/api/v2/dependants",
                   params={"filter[account_uuids]": du, "page": pg, "per_page": 200})
        if r2.status_code != 200:
            break
        b2 = r2.json()
        for rec in b2.get("data", []):
            seen.add(rec.get("uuid", ""))
        if pg >= (b2.get("meta", {}).get("last_page") or 1):
            break
        pg += 1
print("Current dependant count:", len(seen))
