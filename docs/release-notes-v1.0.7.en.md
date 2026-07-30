# v1.0.7

[English](./release-notes-v1.0.7.en.md) | [中文](./release-notes-v1.0.7.md) | [English Home](../README.en.md) | [中文首页](../README.md)

`v1.0.7` fixes the possible “Cannot initialize URL editor” access-denied failure when Word uses the URL editor. It does not change Draw.io XML, SVG, picture layout, or the document metadata storage format.

## Problems Solved First

- Some machines failed to open the Word URL editor with an access-denied error because WebView2 tried to write beside Office executables.
- Users should not need to change `Program Files` or Office installation-directory permissions just to use the embedded editor.
- URL-mode failures needed real Word-host regression coverage and log evidence to identify whether initialization was still the problem.

## Highlights

- The URL editor now explicitly creates a WebView2 environment and stores its profile at `%LOCALAPPDATA%\Greensoft\DrawioPpt\WebView2`.
- Word and PowerPoint no longer rely on WebView2 creating its default profile beside `WINWORD.EXE` or `POWERPNT.EXE`, avoiding `E_ACCESSDENIED` when the Office installation directory is not writable.
- Adds the `webview2-user-data-folder-test.ps1` source contract to prevent a regression to the host-default profile.
- Adds `word-url-addin-host-e2e.ps1`: it temporarily registers the tested DLL in real `WINWORD.EXE`, selects a managed picture to trigger the URL editor, and verifies URL write-back, the per-user profile, and the log.
- The regression script cleans only Word processes reported by its own helper with a matching start time; Word registration, settings, and temporary-directory restoration are isolated, and any restore failure is summarized and fails the test.
- The full Office E2E now runs the real Word URL-host regression against the installed release DLL; the release packager also requires the versioned Chinese and English release material.

## Usage Note

After installing, fully exit and reopen Word or PowerPoint. URL mode uses the current user's LocalAppData directory and does not require changes to `C:\Program Files\Microsoft Office` permissions. If initialization still fails, confirm that the WebView2 Runtime is installed and inspect the `UrlEditor` records in `%APPDATA%\Greensoft\DrawioPpt\Logs\drawioppt.log`.

See the [v1.0.7 E2E test report](./e2e-test-report-v1.0.7.en.md) and [release evidence and log index](./release-evidence-v1.0.7.en.md) for complete verification.
