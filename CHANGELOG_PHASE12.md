# Phase 12 Release Preparation

## Summary
- Rebuilt manifests and registries from main HEAD deterministically.
- Regenerated validation reports with stricter orphan checks.
- Produced unified diffs (Phase 11 → Phase 12) for key artifacts.

## Highlights
- Canonical IDs and link graph unchanged unless noted in diffs.
- Validation status: OK; Orphans: 0

## How to Apply
1. Extract phase12-release-pr.zip at repo root.
2. Review reports: manifest_report.html and registry_validation_report.md.
3. Commit the bundle on a release branch (e.g., phase12-release).
4. Open PR using PR_TEMPLATE_PHASE12.md and merge after approvals.
