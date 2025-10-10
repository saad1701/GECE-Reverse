#!/usr/bin/env python3
# Diff engine for GECE-Reverse manifests: compares JSON across base git ref and head working tree.
# - Loads three manifests: data_map.json, logic_map.json, ui_map.json
# - Produces a machine-readable JSON diff with added/removed/changed per file
# - Applies simple breaking rules from YAML (parsed minimally without PyYAML)
# - Exits 1 if --fail-on-breaking and any breaking changes found
import argparse
import json
import os
import subprocess
import sys
from typing import Any, Dict, List, Tuple

MANIFEST_FILES = [
    "manifest/data_map.json",
    "manifest/logic_map.json",
    "manifest/ui_map.json",
]

def read_file_from_git(ref: str, path: str) -> str:
    try:
        output = subprocess.check_output(["git", "show", ref + ":" + path])
        return output.decode("utf-8")
    except subprocess.CalledProcessError:
        return ""

def read_json_from_git(ref: str, path: str) -> Any:
    content = read_file_from_git(ref, path)
    if not content.strip():
        return None
    try:
        return json.loads(content)
    except Exception:
        return None

def read_json_from_fs(root: str, path: str) -> Any:
    full = os.path.join(root, path)
    if not os.path.exists(full):
        return None
    try:
        with open(full, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None

def diff_values(base: Any, head: Any) -> Dict[str, Any]:
    if base == head:
        return {}
    return {"from": base, "to": head}

def diff_dict(base: Dict[str, Any], head: Dict[str, Any]) -> Dict[str, Any]:
    added, removed, changed = {}, {}, {}
    base_keys, head_keys = set(base.keys()), set(head.keys())
    for k in sorted(head_keys - base_keys): added[k] = head.get(k)
    for k in sorted(base_keys - head_keys): removed[k] = base.get(k)
    for k in sorted(base_keys & head_keys):
        b, h = base.get(k), head.get(k)
        if isinstance(b, dict) and isinstance(h, dict):
            sub = diff_dict(b, h)
            if sub.get("added") or sub.get("removed") or sub.get("changed"): changed[k] = sub
        elif isinstance(b, list) and isinstance(h, list):
            if b != h: changed[k] = {"from": b, "to": h}
        else:
            vdiff = diff_values(b, h)
            if vdiff: changed[k] = vdiff
    return {"added": added, "removed": removed, "changed": changed}

def normalize_none(obj: Any) -> Any:
    return {} if obj is None else obj

def compare_files(base_json: Any, head_json: Any) -> Dict[str, Any]:
    base_n, head_n = normalize_none(base_json), normalize_none(head_json)
    if isinstance(base_n, dict) and isinstance(head_n, dict):
        return diff_dict(base_n, head_n)
    if base_n == head_n:
        return {"added": {}, "removed": {}, "changed": {}}
    return {"added": head_n, "removed": base_n, "changed": {}}

def parse_rules_yaml(yaml_text: str) -> Dict[str, Any]:
    rules = {"breaking": []}
    lines = [ln.rstrip("\n") for ln in yaml_text.splitlines()]
    section, current = None, None
    for ln in lines:
        s = ln.strip()
        if not s or s.startswith("#"): continue
        if s.startswith("breaking:"): section = "breaking"; current=None; continue
        if section == "breaking":
            if s.startswith("-"):
                current = {}; rules["breaking"].append(current)
                s = s[1:].strip()
                if s and ":" in s:
                    k, v = s.split(":", 1); current[k.strip()] = v.strip()
                continue
            if ":" in s and current is not None:
                k, v = s.split(":", 1); current[k.strip()] = v.strip()
    return rules

def collect_changes_for_path(diff_obj: Any, path_expr: str, mode: str):
    parts = path_expr.split("."); results = []
    def walk(node, idx, prefix):
        if idx == len(parts):
            if isinstance(node, dict):
                if mode == "removal" and node.get("removed") not in (None, {}): results.append((prefix, node.get("removed")))
                elif mode == "change" and node.get("changed"): results.append((prefix, node.get("changed")))
            return
        token = parts[idx]; is_list = token.endswith("[]"); key = token[:-2] if is_list else token
        if not isinstance(node, dict): return
        if "changed" in node and key in node["changed"]: walk(node["changed"][key], idx+1, f"{prefix}.{key}")
        if "removed" in node and key in node["removed"] and not is_list and idx == len(parts)-1 and mode == "removal":
            results.append((f"{prefix}.{key}", node["removed"][key]))
    walk(diff_obj, 0, "")
    return results

def evaluate_breaking(diff_map: Dict[str, Any], rules_text: str):
    rules = parse_rules_yaml(rules_text); out = []
    for rule in rules.get("breaking", []):
        file_rule = rule.get("file"); path_expr = (rule.get("path") or "").strip(); rtype = (rule.get("type") or "").strip().lower()
        if not file_rule or not path_expr or rtype not in {"removal","change"}: continue
        file_diff = diff_map.get(file_rule, {})
        matches = collect_changes_for_path(file_diff, path_expr, "removal" if rtype=="removal" else "change")
        for pref, payload in matches:
            out.append({"file": file_rule, "path": path_expr, "rule_type": rtype, "match_prefix": pref, "payload": payload})
    return out

def main():
    parser = argparse.ArgumentParser(description="Diff GECE manifests and evaluate breaking changes.")
    parser.add_argument("--base", default="origin/main"); parser.add_argument("--head", default=".")
    parser.add_argument("--out", default="manifest_diff.json"); parser.add_argument("--rules", default="manifest/rules/breaking_rules.yml")
    parser.add_argument("--fail-on-breaking", action="store_true")
    args = parser.parse_args()

    diff_map = {}
    for rel in MANIFEST_FILES:
        base_json = read_json_from_git(args.base, rel)
        head_json = read_json_from_fs(args.head, rel)
        diff_map[rel] = compare_files(base_json, head_json)

    rules_text = ""
    if os.path.exists(args.rules):
        with open(args.rules, "r", encoding="utf-8") as f: rules_text = f.read()
    breaking = evaluate_breaking(diff_map, rules_text) if rules_text else []

    out_obj = {"base": args.base, "head": args.head, "files": diff_map, "breaking": breaking}
    with open(args.out, "w", encoding="utf-8") as f: json.dump(out_obj, f, indent=2)
    if args.fail_on_breaking and breaking:
        print("Breaking changes detected:", len(breaking)); sys.exit(1)
    print("Diff written to", args.out)

if __name__ == "__main__":
    main()
