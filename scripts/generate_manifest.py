#!/usr/bin/env python3
"""
Generate manifest files by parsing:
- data/sheets/*.json  (sheet schemas and flows)
- data/vba/*.bas      (modules, subs/functions, reads/writes, calls, sheet refs)
- data/vba/forms/*.frm (forms, controls, events)
Writes:
- manifest/data_map.json
- manifest/logic_map.json
- manifest/ui_map.json
- manifest/manifest_readme.md
"""

import os
import re
import io
import json
from typing import List, Tuple, Dict, Any

# --- helpers ---------------------------------------------------------------

def read_text(path: str) -> str:
    with io.open(path, mode="r", encoding="utf-8", errors="ignore") as f:
        return f.read()

def safe_json_load(path: str):
    try:
        with io.open(path, mode="r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None

def ensure_dir(path: str):
    if not os.path.isdir(path):
        os.makedirs(path, exist_ok=True)

# --- parse sheets ----------------------------------------------------------

def parse_sheets(sheets_dir: str) -> Tuple[List[Dict[str, Any]], List[str]]:
    sheets_out: List[Dict[str, Any]] = []
    notes: List[str] = []
    if not os.path.isdir(sheets_dir):
        notes.append("data/sheets missing")
        return sheets_out, notes

    for fn in sorted(os.listdir(sheets_dir)):
        if not fn.lower().endswith(".json"):
            continue
        p = os.path.join(sheets_dir, fn)
        obj = safe_json_load(p)
        if not isinstance(obj, dict):
            continue

        name = obj.get("sheet_name") or obj.get("name") or os.path.splitext(fn)[0]
        fields = obj.get("fields") or []
        inputs = obj.get("inputs") or []
        outputs = obj.get("outputs") or []
        key_formulas = obj.get("key_formulas") or obj.get("formulas") or []

        sheets_out.append({
            "sheet_name": name,
            "description": obj.get("description") or "",
            "fields": fields,
            "inputs": inputs,
            "outputs": outputs,
            "key_formulas": key_formulas,
        })

    return sheets_out, notes

# --- parse VBA modules (.bas) ---------------------------------------------

# Public/Private Sub|Function Name(params)
SIG_RE = re.compile(r'(?im)^\s*(?:Public|Private)?\s*(Sub|Function)\s+([A-Za-z_][A-Za-z0-9_]*)\s*\(([^)]*)\)')
# Call Foo
CALL_RE_1 = re.compile(r'(?im)\bCall\s+([A-Za-z_][A-Za-z0-9_]*)')
# bare Foo(...)   (heuristic)
CALL_RE_2 = re.compile(r'(?im)\b([A-Za-z_][A-Za-z0-9_]*)\s*\(')
# Application.Run "Foo"
CALL_RE_3 = re.compile(r'(?im)Application\.Run\s*"([A-Za-z_][A-Za-z0-9_]*)"')

# Worksheets("Name") / Sheets("Name")
SHEET_REF_RE = re.compile(r'(?im)(?:Worksheets|Sheets)\s*\(\s*"([^"]+)"\s*\)')
# Range("A1") or Range("NamedRange")
RANGE_REF_RE = re.compile(r'(?im)Range\s*\(\s*"([^"]+)"\s*\)')
# Cells(…) usage
CELLS_RE = re.compile(r'(?im)\bCells\s*\(')

# crude read/write heuristics
WRITE_RE_RANGE = re.compile(r'(?im)Range\s*\(.*?\)\s*\.Value\s*=')
READ_RE_RANGE  = re.compile(r'(?im)=\s*Range\s*\(.*?\)\s*\.Value')
WRITE_RE_CELLS = re.compile(r'(?im)Cells\s*\(.*?\)\s*\.Value\s*=')
READ_RE_CELLS  = re.compile(r'(?im)=\s*Cells\s*\(.*?\)\s*\.Value')

RESERVED = set("if while for select case with do loop function sub elseif call and or not then debug else wend end".split())

def _extract_vba_signatures(text: str):
    sigs = []
    for m in SIG_RE.finditer(text):
        kind = m.group(1)
        name = m.group(2)
        params_raw = (m.group(3) or "").strip()
        params = []
        if params_raw:
            parts = [x.strip() for x in params_raw.split(",")]
            for part in parts:
                pname = part.split(" As ")[0].strip()
                if pname:
                    params.append(pname)
        sigs.append({"type": kind, "name": name, "params": params})
    return sigs

def _extract_calls(text: str):
    calls = set()

    for m in CALL_RE_1.finditer(text):
        calls.add(m.group(1))

    for m in CALL_RE_2.finditer(text):
        cand = m.group(1)
        if cand.lower() not in RESERVED:
            calls.add(cand)

    for m in CALL_RE_3.finditer(text):
        calls.add(m.group(1))

    return sorted(calls)

def _extract_sheet_refs(text: str):
    refs = set()
    for m in SHEET_REF_RE.finditer(text):
        refs.add(m.group(1))
    for m in RANGE_REF_RE.finditer(text):
        refs.add(m.group(1))
    if CELLS_RE.search(text):
        refs.add("Cells")
    return sorted(refs)

def _extract_reads_writes(text: str):
    reads, writes = set(), set()
    if WRITE_RE_RANGE.search(text): writes.add("Range.Value")
    if READ_RE_RANGE.search(text):  reads.add("Range.Value")
    if WRITE_RE_CELLS.search(text): writes.add("Cells.Value")
    if READ_RE_CELLS.search(text):  reads.add("Cells.Value")
    return sorted(reads), sorted(writes)

def parse_vba_modules(vba_dir: str) -> Tuple[List[Dict[str, Any]], List[str]]:
    modules = []
    notes: List[str] = []
    if not os.path.isdir(vba_dir):
        notes.append("data/vba missing")
        return modules, notes

    for fn in sorted(os.listdir(vba_dir)):
        if not fn.lower().endswith(".bas"):
            continue
        p = os.path.join(vba_dir, fn)
        txt = read_text(p)

        sigs  = _extract_vba_signatures(txt)
        calls = _extract_calls(txt)
        refs  = _extract_sheet_refs(txt)
        rds, wrs = _extract_reads_writes(txt)

        modules.append({
            "module_name": os.path.splitext(fn)[0],
            "subs":      [s for s in sigs if s["type"].lower() == "sub"],
            "functions": [s for s in sigs if s["type"].lower() == "function"],
            "calls": calls,
            "reads": rds,
            "writes": wrs,
            "sheets_touched": refs,
        })

    return modules, notes

# --- parse forms (.frm) ----------------------------------------------------

CTRL_RE = re.compile(r'(?im)^\s*Begin\s+VB\.([A-Za-z0-9_]+)\s+([A-Za-z0-9_]+)')
EVT_RE  = re.compile(r'(?im)^\s*Private\s+Sub\s+([A-Za-z0-9_]+)_([A-Za-z0-9_]+)\s*\(')

def parse_frm_forms(forms_dir: str) -> Tuple[List[Dict[str, Any]], List[str]]:
    forms = []
    notes: List[str] = []
    if not os.path.isdir(forms_dir):
        notes.append("data/vba/forms missing")
        return forms, notes

    for fn in sorted(os.listdir(forms_dir)):
        if not fn.lower().endswith(".frm"):
            continue
        p = os.path.join(forms_dir, fn)
        txt = read_text(p)

        controls = [{"type": m.group(1), "name": m.group(2)} for m in CTRL_RE.finditer(txt)]
        events   = [{"control": m.group(1), "event": m.group(2)} for m in EVT_RE.finditer(txt)]
        macros_called  = _extract_calls(txt)
        sheets_touched = _extract_sheet_refs(txt)

        forms.append({
            "form_name": os.path.splitext(fn)[0],
            "controls": controls,
            "events": events,
            "macros_called": macros_called,
            "sheets_touched": sheets_touched,
        })

    return forms, notes

# --- main ------------------------------------------------------------------

def main():
    repo_root = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    data_dir   = os.path.join(repo_root, "data")
    sheets_dir = os.path.join(data_dir, "sheets")
    vba_dir    = os.path.join(data_dir, "vba")
    forms_dir  = os.path.join(vba_dir, "forms")
    manifest_dir = os.path.join(repo_root, "manifest")
    ensure_dir(manifest_dir)

    sheets, sheet_notes = parse_sheets(sheets_dir)
    modules, vba_notes  = parse_vba_modules(vba_dir)
    forms, frm_notes    = parse_frm_forms(forms_dir)

    data_map = {
        "version": "0.1.0",
        "source": "GECE-Reverse",
        "notes": ["Generated by scripts/generate_manifest.py"] + sheet_notes,
        "sheets": sheets,
    }
    logic_map = {
        "version": "0.1.0",
        "source": "GECE-Reverse",
        "notes": ["Generated by scripts/generate_manifest.py"] + vba_notes,
        "modules": modules,
    }
    ui_map = {
        "version": "0.1.0",
        "source": "GECE-Reverse",
        "notes": ["Generated by scripts/generate_manifest.py"] + frm_notes,
        "forms": forms,
    }
    readme = (
        "# GECE-Reverse Manifest\n\n"
        "These files were generated by scripts/generate_manifest.py by parsing:\n"
        "- data/sheets/*.json for sheet schemas and flows\n"
        "- data/vba/*.bas for modules, procedures, IO, and calls\n"
        "- data/vba/forms/*.frm for forms, controls, and events\n\n"
        "Update process:\n"
        "1) Modify data sources, then re-run the script.\n"
        "2) Validate references across data_map, logic_map, ui_map.\n"
        "3) Commit regenerated files.\n"
    )

    with io.open(os.path.join(manifest_dir, "data_map.json"), "w", encoding="utf-8") as f:
        json.dump(data_map, f, indent=2, ensure_ascii=False)
    with io.open(os.path.join(manifest_dir, "logic_map.json"), "w", encoding="utf-8") as f:
        json.dump(logic_map, f, indent=2, ensure_ascii=False)
    with io.open(os.path.join(manifest_dir, "ui_map.json"), "w", encoding="utf-8") as f:
        json.dump(ui_map, f, indent=2, ensure_ascii=False)
    with io.open(os.path.join(manifest_dir, "manifest_readme.md"), "w", encoding="utf-8") as f:
        f.write(readme)

    print("✅ Wrote manifest/data_map.json, logic_map.json, ui_map.json, manifest_readme.md")

if __name__ == "__main__":
    main()
