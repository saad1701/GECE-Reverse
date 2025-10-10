## How these were generated
- `scripts/diff_manifest.py` compares `manifest/*.json` between a base git ref (default `origin/main`) and the current working tree, producing `manifest_diff.json`.
- `scripts/render_report.py` converts `manifest_diff.json` into `manifest_report.html` for CI artifacts.
- Rules in `manifest/rules/breaking_rules.yml` mark critical removals/changes as breaking for CI enforcement.
