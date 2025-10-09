import json, sys, pathlib
try:
    import jsonschema
except ImportError:
    sys.exit("jsonschema not installed")
root = pathlib.Path(__file__).resolve().parents[1]
schemas = {
  "data_map.json":  root/"manifest/schema/data_map.schema.json",
  "logic_map.json": root/"manifest/schema/logic_map.schema.json",
  "ui_map.json":    root/"manifest/schema/ui_map.schema.json",
}
ok=True
for fn, schema_path in schemas.items():
    doc_path = root/"manifest"/fn
    try:
        schema = json.load(open(schema_path, encoding="utf-8"))
        doc    = json.load(open(doc_path,    encoding="utf-8"))
        jsonschema.validate(doc, schema)
        print(f"[OK] {fn}")
    except Exception as e:
        ok=False; print(f"[FAIL] {fn}: {e}")
sys.exit(0 if ok else 1)
