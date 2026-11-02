import sys, os, json, requests, time
from azure.identity import DefaultAzureCredential
from azure.keyvault.secrets import SecretClient

sc = SecretClient(vault_url="https://kv-plannr-ingestion.vault.azure.net/",
                  credential=DefaultAzureCredential())
pat = sc.get_secret("plannr-prod-pat").value
s = requests.Session()
s.headers.update({"Authorization": f"Bearer {pat}", "Accept": "application/json",
                  "X-PLANNR-ACCOUNT-UUID": "ccd7f98b-d6a4-46ad-bab3-242ab42826eb"})

imp_uuid = "f4983389-5e72-40fb-b1a3-165fefffba92"
deadline = time.time() + 1800
while time.time() < deadline:
    r = s.get(f"https://api.plannrcrm.com/api/v1/import/{imp_uuid}")
    d = r.json().get("data", r.json())
    st = str(d.get("status") or "").lower()
    print(time.strftime("%H:%M:%S"), " status=" + st, flush=True)
    if st in ("complete", "failed", "error"):
        break
    time.sleep(30)
else:
    print("STILL NOT COMPLETE after 30 min")
    sys.exit(2)

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
print("Dependants after import:", len(seen), " (expected 249)")
