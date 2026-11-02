import csv
missing = {'1a3bcc01-9a16-44c5-9009-c3f506d41c65', '33ad5cba-60f8-4a71-96f2-384856892bc3', '6b2a019e-2f41-476e-9c0e-74bce1065d6d'}
with open('import_20261027/W6_premiums_rewritten.csv', newline='', encoding='utf-8') as f:
    for row in csv.DictReader(f):
        dp = (row.get('plan_uuid') or '').strip()
        if dp in missing:
            amt = row.get('amount', '')
            freq = row.get('frequency', '')
            start = row.get('start_date', '')
            print(f"plan={dp[:8]}  amt={amt}  freq={freq}  start={start}")
