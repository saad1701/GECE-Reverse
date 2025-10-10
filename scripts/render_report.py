#!/usr/bin/env python3
# Render a self-contained HTML report from manifest_diff.json
import argparse, json, html

def escape(s): return html.escape(str(s), quote=True)
def section(title, inner): return f"<section><h2>{escape(title)}</h2>{inner}</section>"

def render_file_diff(file_path, data):
    parts = [f"<h3>{escape(file_path)}</h3>"]
    for key in ["added","removed","changed"]:
        if data.get(key) not in (None, {}, []):
            parts.append(f"<h4>{escape(key.title())}</h4><pre>{escape(json.dumps(data[key], indent=2))}</pre>")
    return "<div class='file-diff'>" + "".join(parts) + "</div>"

def render_breaking(breaking):
    if not breaking: return "<p>No breaking changes detected.</p>"
    items = [f"<li><strong>{escape(b.get('file',''))}</strong> path {escape(b.get('path',''))} type {escape(b.get('rule_type',''))} match {escape(b.get('match_prefix',''))}</li>" for b in breaking]
    return "<ul>" + "".join(items) + "</ul>"

def main():
    ap = argparse.ArgumentParser(); ap.add_argument("--in", dest="inp", default="manifest_diff.json"); ap.add_argument("--out", dest="outp", default="manifest_report.html"); args = ap.parse_args()
    data = json.load(open(args.inp, "r", encoding="utf-8"))
    files = data.get("files", {}); breaking = data.get("breaking", [])
    head = "<!doctype html><html><head><meta charset='utf-8'><title>Manifest Diff Report</title><style>body{font-family:Arial,Helvetica,sans-serif;margin:24px;} h1{margin-top:0;} section{margin-bottom:24px;} pre{background:#f7f7f9;border:1px solid #e1e1e8;padding:10px;white-space:pre-wrap;word-wrap:break-word;} .file-diff{margin-bottom:20px;} .muted{color:#666;}</style></head><body>"
    title = f"<h1>Manifest Diff Report</h1><p class='muted'>Base: {escape(data.get('base',''))} | Head: {escape(data.get('head',''))}</p>"
    files_html = "<div>" + "".join(render_file_diff(fp, fd) for fp, fd in files.items()) + "</div>"
    html_out = head + title + section("Files", files_html) + section("Breaking Changes", render_breaking(breaking)) + "</body></html>"
    open(args.outp, "w", encoding="utf-8").write(html_out); print("Report written to", args.outp)

if __name__ == "__main__":
    main()
