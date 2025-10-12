#!/usr/bin/env python3
import json, subprocess, pathlib, sys, hashlib, os
BASE_DIR = pathlib.Path(__file__).resolve().parents[1]
HASHES_JSON = BASE_DIR / "phase13-registry_hashes.json"
def run_optional(cmd):
    try:
        subprocess.run(cmd, check=True)
        return True
    except Exception:
        return False
def sha256_file(p: pathlib.Path) -> str:
    h = hashlib.sha256()
    with open(p, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()
def collect_current():
    targets = [BASE_DIR / 'manifest', BASE_DIR / 'registry']
    files = {}
    for root in targets:
        if not root.exists():
            continue
        for d, _, fs in os.walk(root):
            for fn in fs:
                p = pathlib.Path(d) / fn
                if '.git' in p.parts:
                    continue
                rel = p.relative_to(BASE_DIR).as_posix()
                files[rel] = {'size': p.stat().st_size, 'sha256': sha256_file(p)}
    return files
def main():
    run_optional([sys.executable, str(BASE_DIR / 'tools' / 'build_manifest.py')])
    run_optional([sys.executable, str(BASE_DIR / 'tools' / 'build_registry.py')])
    if not HASHES_JSON.exists():
        print('[error] Missing baseline: phase13-registry_hashes.json'); return 2
    baseline = json.loads(HASHES_JSON.read_text(encoding='utf-8'))
    base_map = {e['path']: (e['size'], e['sha256']) for e in baseline.get('files', [])}
    cur_map = collect_current()
    added = sorted(set(cur_map) - set(base_map))
    removed = sorted(set(base_map) - set(cur_map))
    changed = sorted([p for p in set(cur_map) & set(base_map) if base_map[p] != (cur_map[p]['size'], cur_map[p]['sha256'])])
    if not any([added, removed, changed]):
        print('[ok] No drift detected.'); return 0
    print('=== DRIFT DETECTED ===')
    if added: print('Added ('+str(len(added))+'):\n  ' + '\n  '.join(added))
    if removed: print('Removed ('+str(len(removed))+'):\n  ' + '\n  '.join(removed))
    if changed: print('Changed ('+str(len(changed))+'):\n  ' + '\n  '.join(changed))
    return 1
if __name__ == '__main__':
    sys.exit(main())
