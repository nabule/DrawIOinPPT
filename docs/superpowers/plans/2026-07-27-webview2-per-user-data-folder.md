# WebView2 Per-User Data Folder Implementation Plan

**Goal:** Make the URL editor initialize in the real Word COM Add-in host by keeping its WebView2 profile under the current user's writable application-data directory.

**Architecture:** `UrlDiagramEditorForm` explicitly creates `CoreWebView2Environment` before the control is initialized. Its profile directory is `%LOCALAPPDATA%\Greensoft\DrawioPpt\WebView2`, shared by the add-ins for the current user. This prevents WebView2 from choosing `WINWORD.EXE.WebView2` beside the protected Word executable.

**Tech stack:** C#/.NET Framework 4.8, Microsoft.Web.WebView2, PowerShell regression scripts, Word COM automation.

## Completed implementation work

- [x] Added `scripts/webview2-user-data-folder-test.ps1` as a source contract. Before the production change it failed with `URL editor still initializes WebView2 with the host default data folder.`; afterward it passed with `WEBVIEW2_USER_DATA_FOLDER_SOURCE_TEST_PASS`.
- [x] Added `GetWebView2UserDataFolder()` in `UrlDiagramEditorForm`, which creates and returns `%LOCALAPPDATA%\Greensoft\DrawioPpt\WebView2`.
- [x] Changed URL-editor initialization to call `CoreWebView2Environment.CreateAsync(null, GetWebView2UserDataFolder())` and then `EnsureCoreWebView2Async(environment)`.
- [x] Added `scripts/word-url-addin-host-e2e.ps1`. It temporarily registers the requested release DLL, opens a managed document in real `WINWORD.EXE`, selects its managed picture, and drives the local mock editor through `configure -> init -> load -> save -> export`.
- [x] Ran the new real-host test against the Release output. It reported `WordAddInConnect=True`, `ActualHostProcess=WINWORD`, `ActualWordUrlEditorSaved=True`, and `WebView2UserDataFolderExists=True`; the new log segment had no URL initialization `E_ACCESSDENIED`.
- [x] Updated architecture, installation, user-guide, Word-add-in, regression, README, release-note, package, and full Office E2E documentation so the behavior and test boundary are discoverable.

## Release gate status

- [x] Built Release x64 and reran both source contracts plus the real Word-host regression.
- [x] Ran `scripts/full-e2e-test.ps1 -Version v1.0.7`; it executed `word-url-addin-host-e2e.ps1` from the installed release package and reported 8/8 PASS.
- [x] Rebuilt the final package after recording test evidence; package-content validation passed and the packaged Word DLL matched Release output.
- [x] Performed specification and code-quality review. The initial script-safety findings were fixed, re-tested, and cleared by final review.
- [x] Created the Chinese commit; version tag, upstream push, and GitHub Release are the remaining publication actions.
