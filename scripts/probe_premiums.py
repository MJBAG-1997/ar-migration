"""Probe premium counts — find extras vs file."""
import csv, json, sys, time, requests
from collections import defaultdict
from azure.identity import DefaultAzureCredential
from azure.keyvault.secrets import SecretClient

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

BASE = "https://api.plannrcrm.com"
ACCT = "ccd7f98b-d6a4-46ad-bab3-242ab42826eb"
IMP  = r"C:\dev\AR Migration\import_20261027"

sc  = SecretClient(vault_url="https://kv-plannr-ingestion.vault.azure.net/",
                   credential=DefaultAzureCredential())
pat = sc.get_secret("plannr-prod-pat").value
s   = requests.Session()
s.headers.update({"Authorization": f"Bearer {pat}", "Accept": "application/json",
                  "X-PLANNR-ACCOUNT-UUID": ACCT})

def lcsv(p):
    with open(p, newline="", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))

prem_rw = lcsv(IMP + r"\W6_premiums_rewritten.csv")
prem_exp = defaultdict(int)
for row in prem_rw:
    dp = (row.get("plan_uuid") or "").strip()
    if dp:
        prem_exp[dp] += 1

total = 0; extras = 0; missing_count = 0
for dp, fc in sorted(prem_exp.items()):
    r = s.get(f"{BASE}/api/v1/plans/{dp}/premiums")
    if r.status_code != 200:
        print(f"WARN {dp[:8]} HTTP {r.status_code}"); continue
    body  = r.json()
    prems = body.get("data", body) if isinstance(body, dict) else body
    if not isinstance(prems, list): prems = []
    dc = len(prems)
    total += dc
    if dc > fc:
        extras += dc - fc
        print(f"EXTRA: plan={dp[:8]} file={fc} dest={dc}")
        for pr in prems:
            amt = pr.get("amount") or {}
            amt_v = amt.get("amount","?") if isinstance(amt,dict) else str(amt)
            puid  = pr.get("uuid","?")[:8]
            freq  = pr.get("frequency","?")
            start = pr.get("start_date","?")
            cr    = str(pr.get("created_at",""))[:19]
            print(f"  prem={puid} amt={amt_v} freq={freq} start={start} created={cr}")
    elif dc < fc:
        missing_count += fc - dc

print(f"\nTotal dest premiums: {total}")
print(f"Extras (dest > file): {extras}")
print(f"Missing (file > dest): {missing_count}")
print(f"Plans in CSV: {len(prem_exp)}")
print("=== DONE ===")
