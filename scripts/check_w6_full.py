import csv
missing = {'1a3bcc01-9a16-44c5-9009-c3f506d41c65'}
with open('import_20261027/W6_premiums_rewritten.csv', newline='', encoding='utf-8') as f:
    for row in csv.DictReader(f):
        dp = (row.get('plan_uuid') or '').strip()
        if dp in missing:
            print({k: v for k, v in row.items() if v})
