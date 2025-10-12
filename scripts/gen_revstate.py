#!/usr/bin/env python3
import hashlib, json, os, sys
from pathlib import Path

IGNORE_EXT = {".zip", ".7z", ".exe", ".pdf", ".png", ".jpg", ".jpeg"}
REGISTRY_DIR = Path("registry")
REVSTATE_MD = Path("REVSTATE.md")
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
    create = "--create-if-missing" in sys.argv
    if create and not REVSTATE_MD.exists():
        REVSTATE_MD.write_text("# Reverse State\n", encoding="utf-8")
    if create and not BASELINE_JSON.exists():
        hashes = collect_hashes(REGISTRY_DIR)
        with open(BASELINE_JSON, "w", encoding="utf-8") as f:
            json.dump(hashes, f, indent=2, sort_keys=True)
    return 0

if __name__ == "__main__":
    sys.exit(main())
