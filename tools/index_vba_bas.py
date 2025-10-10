#!/usr/bin/env python3
import pathlib, re, json
root = pathlib.Path("data/vba")
mods=[]
for f in root.glob("*.bas"):
    txt = f.read_text(encoding="latin-1", errors="ignore")
    mname = None
    m = re.search(r'Attribute\s+VB_Name\s*=\s*"([^"]+)"', txt)
    mname = m.group(1) if m else f.stem
    procs=[]
    for kind,name in re.findall(r'^\s*(Sub|Function)\s+([A-Za-z0-9_]+)', txt, re.M):
        procs.append({"name":name,"kind":kind})
    mods.append({
        "id": f"vba:{mname}",
        "type":"vba_module",
        "module": mname,
        "path": str(f.as_posix()),
        "procedures": procs
    })
out={"vba":mods}
pathlib.Path("data/vba/vba_index.json").write_text(json.dumps(out,indent=2,ensure_ascii=False),encoding="utf-8")
print("✓ Indexed", len(mods), "VBA modules to data/vba/vba_index.json")
