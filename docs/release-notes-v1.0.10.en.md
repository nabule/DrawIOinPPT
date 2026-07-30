# DrawioPpt v1.0.10 Release Notes

[English Home](../README.en.md) | [中文](./release-notes-v1.0.10.md)

## Problem Solved First

On some Windows machines, running v1.0.9 `install.cmd` caused Windows PowerShell to decode localized installer text with the wrong encoding. It produced garbled parser errors before any files were copied, preventing upgrades on already-installed machines.

## Resolution

- Every PowerShell script in the release package is now ASCII-only, avoiding the Windows PowerShell 5.1 UTF-8-without-BOM decoding issue.
- `install.cmd` now keeps the original PowerShell error, prints the failed installation exit code, and directs users to close Word and PowerPoint before retrying.
- A new package encoding-safety check verifies every packaged `.ps1` and the generated `install.cmd`.

## How To Use It

1. Download and extract `DrawioPpt-v1.0.10.zip`.
2. Close Word and PowerPoint.
3. Double-click `install.cmd` in the extracted folder.
4. Open Word or PowerPoint and confirm the Draw.io Ribbon info area shows `v1.0.10+<git-short-hash>`.

If installation fails, keep the command window open. Read the original PowerShell error and `DrawioPpt installation failed with exit code ...`, close any process holding files, and rerun `install.cmd` from the new package. Do not continue using a v1.0.9 package that displays garbled text.

## Verification

- Every PowerShell script in the release package passed ASCII-byte and Windows PowerShell file-parse checks.
- With Word intentionally open, real `install.cmd` retained the original PowerShell error and printed `DrawioPpt installation failed with exit code 1`.
- Full Office E2E passed `14/14`, with `FullE2ECleanupSucceeded=True`; actual Word and PowerPoint version text both matched the package `BuildInfo.txt`.

See the [v1.0.10 E2E test report](./e2e-test-report-v1.0.10.en.md) and [v1.0.10 release evidence](./release-evidence-v1.0.10.en.md) for complete evidence.
