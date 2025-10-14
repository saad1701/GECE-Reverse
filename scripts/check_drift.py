#!/usr/bin/env python3
import hashlib, json, os, sys
from pathlib import Path

IGNORE_EXT = {".zip", ".7z", ".exe", ".pdf", ".png", ".jpg", ".jpeg"}
REGISTRY_DIR = Path("registry")
BASELINE_JSON = Path("phase13-registry_hashes.json")

def hash_file(p):
    h = hashlib.sha256()
    with open(p, "rb") as f:
        for chunk in iter(lambda: f.read(8192), b""):
            h.update(chunk)
    return h.hexdigest()

def collect_hashes(root):
    files = []
    if not root.exists():
        return {}
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames.sort(); filenames.sort()
        for name in filenames:
            ext = os.path.splitext(name)[1].lower()
            if ext in IGNORE_EXT: continue
            path = Path(dirpath) / name
            rel = str(path.relative_to(root)).replace(os.sep, "/")
            files.append((rel, hash_file(path)))
    return {k: v for k, v in files}

def main():
    fail = "--fail-on-drift" in sys.argv
    if not BASELINE_JSON.exists():
        print("Baseline missing")
        sys.exit(1 if fail else 0)
    baseline = json.loads(BASELINE_JSON.read_text(encoding="utf-8"))
    current = collect_hashes(REGISTRY_DIR)

    base_keys = set(baseline.keys()); cur_keys = set(current.keys())
    added = sorted(cur_keys - base_keys)
    removed = sorted(base_keys - cur_keys)
    changed = sorted([k for k in base_keys & cur_keys if baseline[k] != current[k]])

    for hdr, items in (("Added:", added), ("Removed:", removed), ("Changed:", changed)):
        if items:
            print(hdr); [print(i) for i in items]
    sys.exit(1 if fail and (added or removed or changed) else 0)
if __name__ == "__main__":
    main()
