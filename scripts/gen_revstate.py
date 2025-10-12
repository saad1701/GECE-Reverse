#!/usr/bin/env python3
import os, sys, hashlib, json, pathlib, datetime
BASE_DIR = pathlib.Path(__file__).resolve().parents[1]
TARGETS = [BASE_DIR / "manifest", BASE_DIR / "registry"]
OUT_JSON = BASE_DIR / "phase13-registry_hashes.json"
OUT_MD = BASE_DIR / "REVSTATE.md"
def sha256_file(p: pathlib.Path) -> str:
    h = hashlib.sha256()
    with open(p, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()
def walk_files(root: pathlib.Path):
    for d, _, files in os.walk(root):
        for fn in files:
            p = pathlib.Path(d) / fn
            if ".git" in p.parts:
                continue
            yield p
def main():
    entries = []
    for t in TARGETS:
        if not t.exists():
            continue
        for p in walk_files(t):
            rel = p.relative_to(BASE_DIR).as_posix()
            entries.append({"path": rel, "size": p.stat().st_size, "sha256": sha256_file(p)})
    entries.sort(key=lambda x: x["path"])
    stamp = datetime.datetime.utcnow().isoformat() + "Z"
    payload = {"generated_utc": stamp, "files": entries}
    OUT_JSON.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    lines = ["# REVSTATE — Frozen baselines (Phase 13)", f"Generated (UTC): {stamp}", "", "| Path | Size (bytes) | SHA-256 |", "|------|---------------|---------|"]
    for e in entries:
        lines.append(f"| {e['path']} | {e['size']} | {e['sha256']} |")
    OUT_MD.write_text("\n".join(lines) + "\n", encoding="utf-8")
if __name__ == "__main__":
    sys.exit(main())
