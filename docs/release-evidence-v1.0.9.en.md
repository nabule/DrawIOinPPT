# v1.0.9 Release Evidence

This file records the v1.0.9 package build identifier, ZIP SHA-256, full Office E2E result, and standard installation verification.

- Release x64 build passed with `0` warnings and `0` errors.
- `version-display-safety-test.ps1` passed, covering the Word/PowerPoint Ribbon, COM entries, automation version interfaces, and installation verification.
- `install-release-upgrade-safety-test.ps1` passed for same-root and nested-package upgrades.
- Full release-package Office E2E: `13/13 PASS`, with successful cleanup.
- Standard local installation: Word and PowerPoint both completed real COM loading and returned a BuildId matching package `BuildInfo.txt`.

Use `BuildInfo.txt`, `PACKAGE.txt`, and the GitHub Release SHA-256 digest from the final package as the authoritative build identifier and archive hash.
