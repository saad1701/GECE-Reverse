
# Stdlib-only Phase 7 builder: loads inputs, normalizes entries, integrity checks, cross-maps, writes outputs.
import os, json, re
from collections import defaultdict

REQUIRED = ['id','name','type','owner','tags','path','registry_ref']
EXP_NAMES = set(['xmlMod','modGeceExport','modSched','modRTFPrintFromExcel'])

def load_json_or_ndjson(path):
    with open(path,'r',encoding='utf-8') as f:
        raw = f.read()
    try:
        return json.loads(raw)
    except Exception:
        lines = [l for l in raw.splitlines() if l.strip()]
        objs = []
        for l in lines:
            try:
                objs.append(json.loads(l))
            except Exception:
                pass
        return objs

def normalize_entry(e):
    out = dict(e)
    if not out.get('registry_ref'):
        cid = out.get('canonical_id') or out.get('id')
        out['registry_ref'] = str(cid) if cid is not None else None
    if not isinstance(out.get('tags'), list):
        if out.get('tags') is None:
            out['tags'] = []
        else:
            out['tags'] = [str(out.get('tags'))]
    for k in ['id','name','type','owner','path','bucket','registry_ref']:
        if k in out and isinstance(out.get(k), str):
            out[k] = out.get(k).strip()
    return out

def owner_tag_ok(e):
    b = e.get('bucket')
    t = e.get('type')
    owner = e.get('owner')
    tags = set(e.get('tags') or [])
    if b == 'sheets':
        return (t == 'sheet' and owner == 'data-team' and 'sheet' in tags)
    if b == 'vba':
        exp = e.get('name') in EXP_NAMES
        return (t == 'vba_module' and ((exp and owner == 'gece-exports') or ((not exp) and owner == 'gece-core')) and 'vba' in tags)
    if b == 'forms':
        return (t == 'form' and owner == 'ui-team' and 'form' in tags and 'vba-ui' in tags)
    return True

def name_tokens(s):
    if not s:
        return []
    import re
    s2 = re.sub(r'[^A-Za-z0-9_]+', ' ', s)
    return [t for t in s2.split() if t]

def main():
    # inputs
    reg_paths = ['registry/registry_links_merged.json','registry_links_merged.json']
    reg_raw = None
    for p in reg_paths:
        if os.path.exists(p):
            reg_raw = load_json_or_ndjson(p)
            break
    if isinstance(reg_raw, dict):
        reg_links = reg_raw.get('links', [])
    else:
        reg_links = reg_raw or []

    # normalize
    normalized = [normalize_entry(e) for e in reg_links]

    # integrity
    violations = {'missing_fields': [], 'duplicate_ids': [], 'duplicate_registry_refs': [], 'duplicate_bucket_path': [], 'owner_tag_rule_violations': []}
    seen_id = {}
    seen_ref = {}
    seen_bp = {}
    for e in normalized:
        miss = [k for k in REQUIRED if e.get(k) in [None, '', []]]
        if miss:
            violations['missing_fields'].append({'id': e.get('id'), 'registry_ref': e.get('registry_ref'), 'missing': miss})
        i, r = e.get('id'), e.get('registry_ref')
        bp = (e.get('bucket'), e.get('path'))
        if i is not None:
            if i in seen_id:
                violations['duplicate_ids'].append({'id': i})
            else:
                seen_id[i] = True
        if r is not None:
            if r in seen_ref:
                violations['duplicate_registry_refs'].append({'registry_ref': r})
            else:
                seen_ref[r] = True
        if all(bp):
            if bp in seen_bp:
                violations['duplicate_bucket_path'].append({'bucket': bp[0], 'path': bp[1]})
            else:
                seen_bp[bp] = True
        if not owner_tag_ok(e):
            violations['owner_tag_rule_violations'].append({'id': e.get('id'), 'registry_ref': e.get('registry_ref')})

    # dedupe (keep first)
    kept_ref = set()
    kept_bp = set()
    final_index = []
    for e in normalized:
        r = e.get('registry_ref')
        bp = (e.get('bucket'), e.get('path'))
        if r in kept_ref:
            continue
        if all(bp) and bp in kept_bp:
            continue
        kept_ref.add(r)
        if all(bp):
            kept_bp.add(bp)
        final_index.append(e)

    # cross-map
    sheets = [e for e in final_index if e.get('type') == 'sheet']
    modules = [e for e in final_index if e.get('type') in ['vba_module','module','vba']]
    forms = [e for e in final_index if e.get('type') == 'form']

    sheet_to_modules = defaultdict(list)
    module_to_sheets = defaultdict(list)
    form_to_modules = defaultdict(list)

    for s in sheets:
        stoks = name_tokens(s.get('name'))
        for m in modules:
            mblob = ((m.get('name') or '') + ' ' + (m.get('path') or '')).lower()
            hit = any([(t or '').lower() in mblob for t in stoks])
            if (not hit) and s.get('path') and s.get('path') in (m.get('path') or ''):
                hit = True
            if hit:
                sid, mid = s.get('id'), m.get('id')
                if mid not in sheet_to_modules[sid]:
                    sheet_to_modules[sid].append(mid)
                if sid not in module_to_sheets[mid]:
                    module_to_sheets[mid].append(sid)

    for f in forms:
        ftoks = name_tokens(f.get('name'))
        for m in modules:
            mblob = ((m.get('name') or '') + ' ' + (m.get('path') or '')).lower()
            hit = any([(t or '').lower() in mblob for t in ftoks])
            if (not hit) and f.get('path') and f.get('path') in (m.get('path') or ''):
                hit = True
            if hit:
                fid, mid = f.get('id'), m.get('id')
                if mid not in form_to_modules[fid]:
                    form_to_modules[fid].append(mid)

    crossmap = {
        'sheet_to_modules': dict(sheet_to_modules),
        'module_to_sheets': dict(module_to_sheets),
        'form_to_modules': dict(form_to_modules)
    }

    with open('registry_crossmap.json','w',encoding='utf-8') as f:
        json.dump(crossmap, f, ensure_ascii=False, indent=2)

    with open('registry_index.json','w',encoding='utf-8') as f:
        json.dump(final_index, f, ensure_ascii=False, indent=2)

if __name__ == '__main__':
    main()
