import json
with open(r'C:/Users/markj/Downloads/plannrcrm-openapi-source (1).json') as f:
    spec = json.load(f)

def resolve_ref(ref, spec):
    parts = ref.lstrip('#/').split('/')
    node = spec
    for p in parts:
        node = node[p]
    return node

paths = spec.get('paths', {})
p = '/api/v1/plans/{plan_uuid}/premiums'
details = paths[p]['post']
rb = details.get('requestBody', {})
content = rb.get('content', {})
for ct, cv in content.items():
    schema = cv.get('schema', {})
    if '$ref' in schema:
        schema = resolve_ref(schema['$ref'], spec)
    props = schema.get('properties', {})
    required = schema.get('required', [])
    print('required:', required)
    for k, v in props.items():
        if not isinstance(v, dict):
            continue
        if '$ref' in v:
            v = resolve_ref(v['$ref'], spec)
        desc = v.get('description', '')
        typ = v.get('type', '?')
        enum = v.get('enum', [])
        print(k + ': ' + typ + (' enum=' + str(enum) if enum else '') + (' - ' + desc if desc else ''))
