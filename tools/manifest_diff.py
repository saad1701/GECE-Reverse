#!/usr/bin/env python3
import json, hashlib, sys, pathlib, datetime
def load(p): 
    with open(p,'r',encoding='utf-8') as f: return json.load(f)
def index_by_id(items):
    idx={}
    for it in items:
        k = it.get("id") or it.get("path")
        if not k:
            h = hashlib.sha256(json.dumps(it,sort_keys=True).encode()).hexdigest()[:16]
            k = f"auto:{h}"
        idx[k]=it
    return idx
def cmp(a,b):
    ca=json.dumps(a,sort_keys=True); cb=json.dumps(b,sort_keys=True)
    return hashlib.sha256(ca.encode()).hexdigest()==hashlib.sha256(cb.encode()).hexdigest()
def flatten(m):
    out=[]
    for k in ("sheets","vba","forms","ui"):
        for it in m.get(k,[]): 
            it2=dict(it); it2["_bucket"]=k; out.append(it2)
    return out
def main(prev_path, curr_path, out_json, out_html):
    prev, curr = load(prev_path), load(curr_path)
    A, B = index_by_id(flatten(prev)), index_by_id(flatten(curr))
    added=[B[k] for k in B.keys()-A.keys()]
    removed=[A[k] for k in A.keys()-B.keys()]
    changed=[]; unchanged=[]
    for k in A.keys() & B.keys():
        if cmp(A[k],B[k]): unchanged.append(B[k])
        else: changed.append({"before":A[k],"after":B[k]})
    diff={"generated_at":datetime.datetime.utcnow().isoformat()+"Z",
          "counts":{"added":len(added),"removed":len(removed),
                    "changed":len(changed),"unchanged":len(unchanged)},
          "added":added,"removed":removed,"changed":changed}
    with open(out_json,"w",encoding="utf-8") as f: json.dump(diff,f,indent=2,ensure_ascii=False)
    def row(cells): return "<tr>"+"".join(f"<td><pre>{c}</pre></td>" for c in cells)+"</tr>"
    html=["<!doctype html><meta charset='utf-8'><title>Manifest Diff Report</title>",
          "<style>body{font:14px system-ui}table{border-collapse:collapse;width:100%}td,th{border:1px solid #ddd;padding:6px;vertical-align:top}pre{margin:0;white-space:pre-wrap;word-break:break-word}</style>",
          f"<h1>Manifest Diff Report</h1><p>Generated: {diff['generated_at']}</p>",
          "<h2>Summary</h2>",
          f"<ul><li>Added: {len(added)}</li><li>Removed: {len(removed)}</li><li>Changed: {len(changed)}</li><li>Unchanged: {len(unchanged)}</li></ul>"]
    if added:
        html+=["<h2>Added</h2><table><tr><th>Bucket</th><th>ID/Path</th><th>Object</th></tr>"]
        for it in added: html.append(row([it.get('_bucket',''), it.get('id') or it.get('path',''), json.dumps(it,indent=2,ensure_ascii=False)]))
        html.append("</table>")
    if removed:
        html+=["<h2>Removed</h2><table><tr><th>Bucket</th><th>ID/Path</th><th>Object</th></tr>"]
        for it in removed: html.append(row([it.get('_bucket',''), it.get('id') or it.get('path',''), json.dumps(it,indent=2,ensure_ascii=False)]))
        html.append("</table>")
    if changed:
        html+=["<h2>Changed</h2><table><tr><th>Bucket</th><th>ID/Path</th><th>Before</th><th>After</th></tr>"]
        for ch in changed:
            before, after = ch["before"], ch["after"]
            idp = after.get('id') or after.get('path') or before.get('id') or before.get('path','')
            html.append(row([after.get('_bucket', before.get('_bucket','')), idp,
                             json.dumps(before,indent=2,ensure_ascii=False),
                             json.dumps(after,indent=2,ensure_ascii=False)]))
        html.append("</table>")
    with open(out_html,"w",encoding="utf-8") as f: f.write("\n".join(html))
if __name__=="__main__":
    if len(sys.argv)<5:
        print("Usage: manifest_diff.py manifest_prev.json manifest_curr.json manifest_diff.json manifest_report.html"); sys.exit(2)
    main(*sys.argv[1:5])
