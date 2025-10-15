
Generator commands (CI-equivalent):
$ python scripts/gen_revstate.py --create-if-missing
$ python scripts/generate_manifest.py --reset --reset manifest --force
$ python scripts/check_drift.py --fail-on-drift

Unified diff hunk from CI (Manifest Guard):

diff --git a/manifest/ui_map.json b/manifest/ui_map.json
@@ -1,40 +1,40 @@
 {
-  "steps": [
-    {
-      "id": 0,
-      "purchase": [
-        0,
-        1,
-        2,
-        3,
-        4,
-        5,
-        6,
-        7,
-        8,
-        9,
-        10,
-        11,
-        12,
-        13
-      ],
-      "source": "v1.0.1",
-      "target": "GECE-Recover",
-      "event": "executed by scripts/generate_manifest.py",
-      "dataAndForms missing"
-    }
-  ],
-  "ui_map": []
+  "steps": [
+    {
+      "id": 0,
+      "purchase": [
+        0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13
+      ],
+      "source": "v1.0.1",
+      "target": "GECE-Recover",
+      "event": "executed by scripts/generate_manifest.py",
+      "dataAndForms.missing": true
+    }
+  ],
+  "ui_map": []
 }


Assumptions and formatting:
- UTF-8, LF endings, 2-space indent, trailing newline
- Keys and order per CI hunk: steps[id, purchase, source, target, event, dataAndForms.missing], ui_map

Branch and SHAs:
- Target branch: phase-14-data
- Current file SHA on branch: <sha-from-phase-14-data-manifest/ui_map.json>
- Output file SHA256: e074a84a81160474e02bbdcb9e9e6af07a20f6dd6ad33a7097dfdca829d355e9
