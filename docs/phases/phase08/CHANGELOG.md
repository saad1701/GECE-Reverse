# Phase 8 Reconciliation
- Canonical IDs normalized to type namespaces (sheet:/form:/module:) with lowercase hyphen slugs.
- Adopted registry/registry_links_merged.json as the single source of truth.
- Orphans removed and conflicts deduplicated, prioritizing manifest-derived entries.
- Refreshed registry_index.json and registry_crossmap.json with normalized keys and tags.
- Regenerated manifest_curr.json; included manifest_diff.json and manifest_curr.diff.
