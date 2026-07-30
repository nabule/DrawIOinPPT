# v1.0.10 Release Evidence

This file records the v1.0.10 package build ID, ZIP SHA-256, full Office E2E, and standard local-install verification.

- Release x64 build passed with `0` warnings and `0` errors.
- `install-release-encoding-safety-test.ps1` passed: every repository and release-package `.ps1` is ASCII-only, and `install.cmd` prints and returns the failed exit code.
- `install-release-upgrade-safety-test.ps1` passed, covering same-root and nested-package upgrades.
- Full release-package Office E2E passed `14/14`, with cleanup completed successfully.
- During temporary release-package installation, both Word and PowerPoint loaded through real COM and returned BuildIds matching the package `BuildInfo.txt`.

The final package BuildId and SHA-256 are defined by the final package `BuildInfo.txt`, `PACKAGE.txt`, and GitHub Release digest.
