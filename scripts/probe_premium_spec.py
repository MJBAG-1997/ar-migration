import json
with open(r'C:/Users/markj/Downloads/plannrcrm-openapi-source (1).json') as f:
    spec = json.load(f)
paths = spec.get('paths', {})
for p, methods in paths.items():
    if 'premium' in p.lower():
        for m, details in methods.items():
            if not isinstance(details, dict):
                continue
            print(m.upper(), p)
            rb = details.get('requestBody', {})
            if rb:
                content = rb.get('content', {})
                for ct, cv in content.items():
                    schema = cv.get('schema', {})
                    props = schema.get('properties', {})
                    required = schema.get('required', [])
                    print('  required:', required)
                    for k, v in props.items():
                        if not isinstance(v, dict):
                            continue
                        desc = v.get('description', '')
                        typ = v.get('type', '?')
                        print('  ' + k + ': ' + typ + ' - ' + desc)
