# Word Selection Event Sync Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Remove the Word add-in's periodic 500ms selection polling while preserving event-driven selection state, Ribbon refresh, and explicit diagram commands.

**Architecture:** Word `SelectionMonitor` already forwards `Application.WindowSelectionChange` to `AddInHost`. `AddInHost` retains that path and one startup synchronization, but no longer owns a `System.Windows.Forms.Timer`. A structural regression script prevents reintroducing periodic polling; a real Word COM test validates floating-picture live selection for the explicit command path, while the Word URL E2E verifies editing write-back.

**Tech Stack:** C# .NET Framework 4.8, Word COM Interop, PowerShell, MSBuild.

---

### Task 1: Lock the no-polling invariant

**Files:**
- Create: `scripts/word-selection-sync-test.ps1`
- Test: `scripts/word-selection-sync-test.ps1`

- [x] **Step 1: Write the failing regression test**

```powershell
$source = Get-Content -Raw (Join-Path $repoRoot 'src\\DrawioPpt.WordAddIn\\Services\\AddInHost.cs')
if ($source -match 'System\\.Windows\\.Forms\\.Timer|_selectionStateTimer|OnSelectionStateTimerTick') {
    throw 'Word selection synchronization must not use a periodic timer.'
}
```

- [x] **Step 2: Run it and confirm the current host fails**

Run: `powershell.exe -ExecutionPolicy Bypass -File .\\scripts\\word-selection-sync-test.ps1`

Expected: failure stating that periodic timer-based synchronization is forbidden.

- [x] **Step 3: Add event-path guards**

```powershell
if ($source -notmatch '_selectionMonitor\\.Start\\(\\)' -or $source -notmatch 'ApplySelectionContext\\(ReadLiveSelectionContext\\(\\), false\\)') {
    throw 'Word startup must retain one initial event-driven selection synchronization.'
}
```

- [x] **Step 4: Commit the test with the implementation task**

```powershell
git add scripts/word-selection-sync-test.ps1 src/DrawioPpt.WordAddIn/Services/AddInHost.cs
git commit -m "优化：改为事件驱动的 Word 选区同步"
```

### Task 2: Remove periodic selection polling

**Files:**
- Modify: `src/DrawioPpt.WordAddIn/Services/AddInHost.cs:33-65,101-125,425-431`
- Test: `scripts/word-selection-sync-test.ps1`

- [x] **Step 1: Delete timer ownership and lifecycle code**

```csharp
// Delete the Timer field, constructor initialization, Start(), Dispose(),
// and OnSelectionStateTimerTick() callback. Keep _selectionMonitor.Start()
// and ApplySelectionContext(ReadLiveSelectionContext(), false) in Start().
```

- [x] **Step 2: Run the regression test**

Run: `powershell.exe -ExecutionPolicy Bypass -File .\\scripts\\word-selection-sync-test.ps1`

Expected: `Word selection event synchronization checks passed.`

- [x] **Step 3: Build Debug and run the real Word E2E**

Run:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\\scripts\\build.ps1 -Configuration Debug -Platform x64
powershell.exe -ExecutionPolicy Bypass -File .\\scripts\\word-selection-event-e2e.ps1 -Configuration Debug -SkipBuild
```

Expected: the E2E exits 0 with `NoSelectionPolling=True`, `ManagedSelectionDetected=True`, and `PlainPictureCanBind=True` for real Word floating pictures.

### Task 3: Document and release v1.0.6

**Files:**
- Modify: `README.md`, `README.en.md`, `docs/architecture.md`, `docs/architecture.en.md`, `docs/word-addin.md`, `docs/word-addin.en.md`, `docs/regression-checklist.md`, `docs/regression-checklist.en.md`
- Create: `docs/release-notes-v1.0.6.md`, `docs/release-notes-v1.0.6.en.md`, `docs/e2e-test-report-v1.0.6.md`, `docs/e2e-test-report-v1.0.6.en.md`, `docs/release-evidence-v1.0.6.md`, `docs/release-evidence-v1.0.6.en.md`

- [x] **Step 1: Update product, architecture and regression documentation**

Document that Word selection state is synchronized by Word selection events and a startup read, without background polling. Do not claim SVG or metadata behavior changed.

- [x] **Step 2: Build, package, install-verify, and load-check**

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\\scripts\\build.ps1 -Configuration Release -Platform x64
powershell.exe -ExecutionPolicy Bypass -File .\\scripts\\package-release.ps1 -Version v1.0.6 -Configuration Release -Platform x64 -SkipBuild
powershell.exe -ExecutionPolicy Bypass -File .\\artifacts\\releases\\v1.0.6\\package\\scripts\\verify-office-install.ps1
powershell.exe -ExecutionPolicy Bypass -File .\\scripts\\word-addin-load-check.ps1 -Configuration Release
```

- [ ] **Step 3: Commit and publish**

```powershell
git add README.md README.en.md docs scripts
git commit -m "发布：v1.0.6 优化 Word 选区同步性能"
git push origin codex/word-addin
git tag -a v1.0.6 -m "发布 v1.0.6"
git push origin v1.0.6
gh release create v1.0.6 artifacts/releases/v1.0.6/DrawioPpt-v1.0.6.zip --title "DrawioPpt v1.0.6" --notes-file docs/release-notes-v1.0.6.md
```
