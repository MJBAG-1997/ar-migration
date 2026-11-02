import json
with open(r'C:/Users/markj/Downloads/plannrcrm-openapi-source (1).json') as f:
    spec = json.load(f)
paths = spec.get('paths', {})
p = '/api/v1/plans/{plan_uuid}/premiums'
methods = paths.get(p, {})
details = methods.get('post', {})
rb = details.get('requestBody', {})
content = rb.get('content', {})
for ct, cv in content.items():
    schema = cv.get('schema', {})
    props = schema.get('properties', {})
    required = schema.get('required', [])
    print('required:', required)
    for k, v in props.items():
        if not isinstance(v, dict):
            continue
        desc = v.get('description', '')
        typ = v.get('type', '?')
        enum = v.get('enum', [])
        print(k + ': ' + typ + (' enum=' + str(enum) if enum else '') + (' - ' + desc if desc else ''))
