# v1.0.7 Release Evidence and Log Index

[English](./release-evidence-v1.0.7.en.md) | [中文](./release-evidence-v1.0.7.md) | [English Home](../README.en.md) | [中文首页](../README.md)

## Deliverables

- Version: `v1.0.7`
- Release directory: `artifacts\releases\v1.0.7\package`
- Release archive: `artifacts\releases\v1.0.7\DrawioPpt-v1.0.7.zip`
- New release verification script: `scripts\word-url-addin-host-e2e.ps1`

## Release Gate

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build.ps1 -Configuration Release -Platform x64
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\webview2-user-data-folder-test.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-url-addin-host-e2e.ps1 -Configuration Release -SkipBuild
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\full-e2e-test.ps1 -Version v1.0.7
```

## Final Verification Results

- `Release|x64` build passed: 0 warnings, 0 errors.
- `webview2-user-data-folder-test.ps1` passed with `WEBVIEW2_USER_DATA_FOLDER_SOURCE_TEST_PASS`.
- `word-url-addin-host-e2e-safety-test.ps1` passed with `WORD_URL_ADDIN_HOST_E2E_SAFETY_TEST_PASS`; it enforces writing the URL setting only after mock binding succeeds, retrying available loopback ports on bind failure, helper PID/start-time cleanup written immediately, and isolated restoration steps.
- `word-url-addin-host-e2e.ps1` passed against the source Release DLL: real `WINWORD.EXE` reported `WordAddInConnect=True`, `ActualWordUrlEditorSaved=True`, and `WebView2UserDataFolderExists=True`, and the new log segment had no URL-initialization `E_ACCESSDENIED`.
- `word-selection-event-e2e.ps1` passed: `NoSelectionPolling=True`, `ManagedSelectionDetected=True`, and `PlainPictureCanBind=True`.
- The full Office E2E passed 8/8. `InstalledWordUrlHostE2E` registered the DLL from the temporary installed release package `bin` and completed URL write-back in the real Word host.
- The final package must contain `DrawioPpt.WordAddIn.dll`, `word-url-addin-host-e2e.ps1`, both source-contract scripts, the versioned Chinese and English release notes, E2E reports, and this evidence index. The archive SHA256 is retained as the GitHub Release asset verification record.

## Known Test Preconditions

Close Word before running the real Word URL-host test and do not run another Word automation task in parallel. The script stops only Word processes it created and restores the Word add-in, COM class registrations, and plug-in settings in `finally`.
