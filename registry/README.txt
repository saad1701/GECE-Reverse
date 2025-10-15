
Generator commands (CI-equivalent):
$ python scripts/gen_revstate.py --create-if-missing
$ python scripts/generate_manifest.py --reset --reset manifest --force
$ python scripts/check_drift.py --fail-on-drift

Reverse Validate drift evidence:
Changed: components.json

Baseline content expected by CI:
{
  "components": []
}

Assumptions and formatting:
- UTF-8, LF endings, 2-space indent, trailing newline

Branch and SHAs:
- Target branch: phase-14-data
- Current file SHA on branch: <sha-from-phase-14-data-registry/components.json>
- Output file SHA256: 00070f6a1fd3332c284a60a19ad819d598d084e00fdd6024672a79f958deb215
