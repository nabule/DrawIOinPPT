# v1.0.9 Office E2E Test Report

This report records the full Office E2E run before the v1.0.9 release. The run temporarily installs the package, validates Word and PowerPoint loading, editing, and export flows, and compares the BuildId returned by both hosts with `BuildInfo.txt` during installation verification.

## Result

`13/13 PASS`, with `FullE2ECleanupSucceeded=True`.

| Check | Result |
| --- | --- |
| Release package and Markdown-link validation | Passed |
| Nested-package upgrade under the old install root | Passed |
| Temporary installation and Office registration | Passed |
| Real Word/PowerPoint COM loading and version callbacks | Passed |
| URL editor, actual Word-host write-back, and complex metadata | Passed |
| Word/PowerPoint SVG aspect handling, preview, and URL write-back | Passed |
| draw.io Desktop export | Passed |

Generated report: `artifacts\test-reports\full-e2e-v1.0.9.md`.
