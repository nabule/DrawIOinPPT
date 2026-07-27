# v1.0.6 E2E Test Report

[English](./e2e-test-report-v1.0.6.en.md) | [中文](./e2e-test-report-v1.0.6.md) | [English Home](../README.en.md) | [中文首页](../README.md)

## Environment

- Windows with Microsoft Word Desktop and PowerPoint Desktop
- .NET Framework 4.8, `Release|x64`
- Current-user Office COM Add-in registration

## Regression Results

| Check | Result | Evidence |
| --- | --- | --- |
| No-polling red test | Expected failure observed | Pre-change contract detected `_selectionStateTimer` |
| No-polling green test | Pass | `WORD_SELECTION_SYNC_TEST_PASS` |
| Release build | Pass | 0 warnings, 0 errors |
| Word floating-picture COM E2E | Pass | `NoSelectionPolling=True`, `ManagedSelectionDetected=True`, `PlainPictureCanBind=True` |
| Word URL write-back E2E | Pass | `MockHttpReachable=True`; create, reopen-edit, and final reopen all `True` |
| Test-environment restoration | Pass | URL E2E reports `RegisteredWordAddInLoadBehavior=Restored:3` and leaves no Word process |

## Coverage

`word-selection-event-e2e.ps1` creates managed and normal floating `Shape` pictures in real Word, validates live-selection recognition through the production event reader, and reflects over the tested DLL to ensure the old Timer field and Tick handler are absent. Explicit Word commands use the same live-selection path.

`word-url-e2e.ps1` covers URL-mode `configure -> init -> load -> save -> export`, save, reopen, and a second edit. Its mock runs in the same test process, performs an HTTP health check, and records iframe requests. Word must be closed before the test; the original registered add-in `LoadBehavior` is restored afterward.

See the [release evidence and log index](./release-evidence-v1.0.6.en.md) for final package-install, release-DLL, and full Office E2E records.
