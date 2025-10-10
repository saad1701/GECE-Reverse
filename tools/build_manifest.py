#!/usr/bin/env python3
import json, pathlib

def load_json(path):
    for enc in ("utf-8","utf-8-sig","latin-1"):
        try: return json.loads(path.read_text(encoding=enc, errors="ignore"))
        except Exception: pass
    return None

def load_recursive(root):
    arr=[]
    for f in pathlib.Path(root).rglob("*.json"):
        if f.name in ("forms_index.json","vba_index.json"):  # handled separately
            continue
        obj = load_json(f)
        if obj is None:   # fallback: make a minimal record by filename
            arr.append({"id": f"{pathlib.Path(root).name}:{f.stem}", "path": str(f.as_posix())})
            continue
        if isinstance(obj, list): arr += obj
        elif isinstance(obj, dict): arr.append(obj)
    return arr

sheets = load_recursive("data/sheets")

# Prefer compiled index for VBA; else try any *.json; else empty list
vba_idx = pathlib.Path("data/vba/vba_index.json")
if vba_idx.exists():
    v = load_json(vba_idx) or {}
    vba = v.get("vba", [])
else:
    vba = load_recursive("data/vba")

forms_idx = pathlib.Path("data/forms/forms_index.json")
forms = (load_json(forms_idx) or {}).get("forms", []) if forms_idx.exists() else []

manifest = {"sheets": sheets, "vba": vba, "forms": forms, "ui": []}
out = pathlib.Path("manifest/manifest_curr.json")
out.write_text(json.dumps(manifest, indent=2, ensure_ascii=False), encoding="utf-8")
print(f"✓ manifest_curr.json created with {len(sheets)} sheets, {len(vba)} vba modules, {len(forms)} forms")
