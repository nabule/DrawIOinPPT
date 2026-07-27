# v1.0.7 E2E Test Report

[English](./e2e-test-report-v1.0.7.en.md) | [中文](./e2e-test-report-v1.0.7.md) | [English Home](../README.en.md) | [中文首页](../README.md)

## Test Environment

- Windows + Microsoft Word Desktop / PowerPoint Desktop
- .NET Framework 4.8, `Release|x64`
- Current-user Office COM Add-in registration and installed release package

## Regression Set

| Check | Result | Evidence |
| --- | --- | --- |
| WebView2 user-data-folder source contract | PASS | `WEBVIEW2_USER_DATA_FOLDER_SOURCE_TEST_PASS` |
| Word URL-host test safety contract | PASS | `WORD_URL_ADDIN_HOST_E2E_SAFETY_TEST_PASS` |
| Release build | PASS | 0 warnings, 0 errors |
| Word selection-event E2E | PASS | `NoSelectionPolling=True`, `ManagedSelectionDetected=True`, `PlainPictureCanBind=True` |
| Real Word URL-host E2E | PASS | `WordAddInConnect=True`, `ActualHostProcess=WINWORD`, `ActualWordUrlEditorSaved=True`, `WebView2UserDataFolderExists=True` |
| Full Office E2E against installed release DLL | PASS (8/8) | Release package, installation, Office registration/load, URL smoke, real Word URL host, PowerPoint URL E2E, and desktop export are all PASS |

## Coverage

`word-url-addin-host-e2e.ps1` does not initialize WebView2 from a temporary EXE host. It selects a managed picture only after the Word COM Add-in is loaded in real `WINWORD.EXE`. A local mock URL editor completes `configure -> init -> load -> save -> export`; the test verifies updated XML write-back, the existence of `%LOCALAPPDATA%\Greensoft\DrawioPpt\WebView2`, and no URL-initialization `E_ACCESSDENIED` in the new log segment. The script restores the Word add-in, COM class registrations, and plug-in settings.

The full Office E2E report is `artifacts\test-reports\full-e2e-v1.0.7.md`. Its eight passing checks are `ReleasePackage`, `InstallRelease`, `InstalledOfficeAddInsVerify`, `PowerPointAddInLoad`, `InstalledUrlSmoke`, `InstalledWordUrlHostE2E`, `PowerPointUrlE2E`, and `DesktopExporter`. See the [release evidence and log index](./release-evidence-v1.0.7.en.md) for package-content and release-asset verification.
