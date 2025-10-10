#!/usr/bin/env python3
import json, sys, re
def load(p): 
    with open(p,'r',encoding='utf-8') as f: return json.load(f)
def norm(s): 
    return re.sub(r'\s+',' ', s.strip().lower()) if isinstance(s,str) else s
def index(reg):
    ix={}
    for r in reg:
        keys=set(filter(None,[r.get("id"),r.get("path"),r.get("name"),r.get("sheet"),r.get("module")]))
        for k in list(keys): ix[norm(k)] = r
        for a in r.get("aliases",[]): ix[norm(a)] = r
    return ix
def flatten(manifest):
    out=[]
    for bucket in ("sheets","vba","forms","ui"):
        for it in manifest.get(bucket,[]): out.append({"bucket":bucket, **it})
    return out
def main(manifest_path, *registry_paths):
    manifest = load(manifest_path)
    regs=[]
    for p in registry_paths:
        try: regs += load(p)
        except FileNotFoundError: pass
    R = index(regs)
    links=[]; unresolved=[]
    for it in flatten(manifest):
        candidates=[it.get("id"),it.get("path"),it.get("name"),it.get("sheet"),it.get("module")]
        found=None
        for c in candidates:
            if c and norm(c) in R: found = R[norm(c)]; break
        rec={"bucket":it["bucket"],"id":it.get("id") or it.get("path") or it.get("name"),
             "name":it.get("name"),"path":it.get("path"),
             "registry_ref": found.get("id") if found else None,
             "registry_meta": {k:v for k,v in (found or {}).items() if k!="aliases"}}
        (links if found else unresolved).append(rec)
    print(json.dumps({"links":links,"unresolved":unresolved,
                      "stats":{"linked":len(links),"unresolved":len(unresolved)}},
                     indent=2, ensure_ascii=False))
if __name__=="__main__":
    if len(sys.argv)<2:
        print("Usage: registry_linker.py manifest_curr.json [registry/*.json ...]"); sys.exit(2)
    main(*sys.argv[1:])
