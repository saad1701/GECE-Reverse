# Phase 11 Release Preparation

## Summary
- Rebuilt manifests and registries from main HEAD.
- Regenerated validation reports with stricter orphan checks.
- Prepared unified diffs (Phase 10 → Phase 11) for key artifacts.

## Highlights
- Canonical IDs and linkage model unchanged unless noted in diffs.
- Validation status: OK; Orphans: 0

## How to Apply
1. Extract phase11-release-pr.zip at repo root.
2. Review the diffs and reports: manifest_report.html and registry_validation_report.md.
3. Commit the bundle on a release branch (e.g., phase11-release).
4. Open PR using PR_TEMPLATE_PHASE11.md, ensure checks pass, and merge.
