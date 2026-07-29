param(
    [string]$Configuration = "Release",
    [string]$AssemblyRoot,
    [switch]$SkipBuild
)

$ErrorActionPreference = "Stop"

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$buildScript = Join-Path $PSScriptRoot "build.ps1"
$tempRoot = Join-Path $env:TEMP ("DrawioPpt\word-svg-aspect-ratio-" + [Guid]::NewGuid().ToString("N"))
$tempRootPrefix = Join-Path $env:TEMP "DrawioPpt\word-svg-aspect-ratio-"
$programPath = Join-Path $tempRoot "word-svg-aspect-ratio.cs"
$exePath = Join-Path $tempRoot "word-svg-aspect-ratio.exe"
$imagePath = Join-Path $tempRoot "wide-source.png"
$verticalImagePath = Join-Path $tempRoot "vertical-source.png"
$missingImagePath = Join-Path $tempRoot "missing-source.png"
$documentPath = Join-Path $tempRoot "word-svg-aspect-ratio.docx"
$processIdentityPath = Join-Path $tempRoot "word-process.identity"
$cscPath = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
$interopWord = "C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Word\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.Word.dll"
$officeCore = "C:\Windows\assembly\GAC_MSIL\office\15.0.0.0__71e9bce111e9429c\OFFICE.DLL"

function Assert-NoRunningWord {
    if (Get-Process -Name WINWORD -ErrorAction SilentlyContinue) {
        throw "请先关闭正在运行的 Word，再执行 SVG 比例 E2E 测试。"
    }
}

function Remove-TestTempRoot {
    if (-not (Test-Path -LiteralPath $tempRoot)) {
        return
    }

    if (-not $tempRoot.StartsWith($tempRootPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "拒绝清理不属于 SVG 比例 E2E 的临时目录：$tempRoot"
    }

    Remove-Item -LiteralPath $tempRoot -Recurse -Force
}

if (-not (Test-Path -LiteralPath $cscPath)) {
    throw "csc.exe not found: $cscPath"
}

if (-not $SkipBuild) {
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $buildScript -Configuration $Configuration -Platform x64
    if ($LASTEXITCODE -ne 0) {
        exit $LASTEXITCODE
    }
}

if ([string]::IsNullOrWhiteSpace($AssemblyRoot)) {
    $AssemblyRoot = Join-Path $repoRoot ("src\DrawioPpt.WordAddIn\bin\x64\" + $Configuration)
}

if (Test-Path -LiteralPath $AssemblyRoot) {
    $AssemblyRoot = (Resolve-Path -LiteralPath $AssemblyRoot).Path
}

$wordAssemblyPath = Join-Path $AssemblyRoot "DrawioPpt.WordAddIn.dll"
$powerPointAssemblyPath = Join-Path $AssemblyRoot "DrawioPpt.PowerPointAddIn.dll"
$coreAssemblyPath = Join-Path $AssemblyRoot "DrawioPpt.Core.dll"
foreach ($requiredPath in @($wordAssemblyPath, $powerPointAssemblyPath, $coreAssemblyPath, $interopWord, $officeCore)) {
    if (-not (Test-Path -LiteralPath $requiredPath)) {
        throw "Required test assembly not found: $requiredPath"
    }
}

Assert-NoRunningWord
New-Item -ItemType Directory -Force -Path $tempRoot | Out-Null

try {
    @'
using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using DrawioPpt.Core.Services;
using DrawioPpt.PowerPointAddIn.Services;
using DrawioPpt.WordAddIn.Services;
using DrawioPpt.WordAddIn.Word;
using WordInterop = Microsoft.Office.Interop.Word;

public static class WordSvgAspectRatioE2E
{
    private static bool UsesFrontWrap(
        WordPictureReference picture)
    {
        WordInterop.WrapFormat wrapFormat = null;
        try
        {
            if (picture == null ||
                picture.IsInline ||
                picture.Shape == null)
            {
                return false;
            }

            wrapFormat = picture.Shape.WrapFormat;
            return wrapFormat.Type ==
                WordInterop.WdWrapType.wdWrapFront;
        }
        finally
        {
            if (wrapFormat != null)
            {
                try { Marshal.FinalReleaseComObject(wrapFormat); }
                catch { }
            }
        }
    }

    private static void ReleasePicture(
        WordPictureReference picture)
    {
        if (picture == null)
        {
            return;
        }

        try
        {
            if (picture.Shape != null)
            {
                Marshal.FinalReleaseComObject(picture.Shape);
            }
            else if (picture.InlineShape != null)
            {
                Marshal.FinalReleaseComObject(picture.InlineShape);
            }
        }
        catch { }
    }

    [STAThread]
    public static int Main(string[] args)
    {
        WordInterop.Application application = null;
        WordInterop.Document document = null;
        WordPictureReference inserted = null;
        WordPictureReference replaced = null;
        WordPictureReference vertical = null;

        try
        {
            application = new WordInterop.Application();
            application.Visible = false;
            Process wordProcess = null;
            int processAttempt;
            for (processAttempt = 0;
                 processAttempt < 50;
                 processAttempt++)
            {
                Process[] wordProcesses =
                    Process.GetProcessesByName("WINWORD");
                if (wordProcesses.Length == 1)
                {
                    wordProcess = wordProcesses[0];
                    break;
                }

                foreach (Process candidate in wordProcesses)
                {
                    candidate.Dispose();
                }

                Thread.Sleep(100);
            }

            if (wordProcess == null)
            {
                throw new InvalidOperationException(
                    "Unable to identify the test-owned Word process.");
            }

            File.WriteAllText(
                args[4],
                wordProcess.Id.ToString() + "|" +
                wordProcess.StartTime.ToUniversalTime().Ticks.ToString());
            wordProcess.Dispose();
            document = application.Documents.Add();

            WordSvgPictureService service = new WordSvgPictureService(
                new SelectedPictureAccessor(new DiagramEnvelopeSerializer()),
                new SvgPowerPointSupportService());

            inserted = service.InsertAtSelection(
                application,
                args[0],
                "Wide PNG preview");
            float insertedRatio = inserted.Width / inserted.Height;
            bool insertUsesFloatingFront =
                UsesFrontWrap(inserted);
            inserted.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
            inserted.Width = 400f;
            inserted.Height = 400f;
            replaced = service.Replace(
                document,
                inserted,
                args[0]);
            float replacedRatio = replaced.Width / replaced.Height;
            bool replacementUsesFloatingFront =
                UsesFrontWrap(replaced);
            bool replacementRetainsWidth = Math.Abs(replaced.Width - 400f) < 0.1f;
            bool replacementCorrectsHeight = Math.Abs(replaced.Height - 200f) < 0.1f;
            bool insertPreservesSourceAspect = Math.Abs(insertedRatio - 2f) < 0.02f;
            bool replacePreservesSourceAspect = Math.Abs(replacedRatio - 2f) < 0.02f;
            int replacementId = replaced.Shape.ID;
            bool failedReplacementThrew = false;
            try
            {
                service.Replace(document, replaced, args[3]);
            }
            catch
            {
                failedReplacementThrew = true;
            }

            bool failedReplacementPreservedOriginal =
                failedReplacementThrew &&
                document.Shapes.Count == 1 &&
                replaced.Shape.ID == replacementId &&
                Math.Abs(replaced.Width - 400f) < 0.1f &&
                Math.Abs(replaced.Height - 200f) < 0.1f;

            vertical = service.InsertAtSelection(
                application,
                args[1],
                "Vertical PNG preview");
            float verticalRatio =
                vertical.Width / vertical.Height;
            WordInterop.PageSetup pageSetup = null;
            float maximumHeight = 420f;
            try
            {
                pageSetup = document.PageSetup;
                maximumHeight =
                    (pageSetup.PageHeight -
                     pageSetup.TopMargin -
                     pageSetup.BottomMargin) * 0.55f;
            }
            finally
            {
                if (pageSetup != null)
                {
                    try
                    {
                        Marshal.FinalReleaseComObject(pageSetup);
                    }
                    catch { }
                }
            }

            bool verticalContained =
                Math.Abs(verticalRatio - 0.25f) < 0.02f &&
                vertical.Height <= maximumHeight + 0.1f;
            bool documentContainsOnlyFloatingPictures =
                document.InlineShapes.Count == 0 &&
                document.Shapes.Count == 2;

            document.SaveAs2(args[2]);

            Console.WriteLine("InsertedRatio=" + insertedRatio.ToString("0.###"));
            Console.WriteLine("ReplacedRatio=" + replacedRatio.ToString("0.###"));
            Console.WriteLine("ReplacementRetainsWidth=" + replacementRetainsWidth);
            Console.WriteLine("ReplacementCorrectsHeight=" + replacementCorrectsHeight);
            Console.WriteLine("InsertPreservesSourceAspect=" + insertPreservesSourceAspect);
            Console.WriteLine("ReplacePreservesSourceAspect=" + replacePreservesSourceAspect);
            Console.WriteLine("InsertUsesFloatingFront=" + insertUsesFloatingFront);
            Console.WriteLine("ReplacementUsesFloatingFront=" + replacementUsesFloatingFront);
            Console.WriteLine("FailedReplacementPreservedOriginal=" + failedReplacementPreservedOriginal);
            Console.WriteLine("VerticalRatio=" + verticalRatio.ToString("0.###"));
            Console.WriteLine("VerticalContained=" + verticalContained);
            Console.WriteLine("DocumentContainsOnlyFloatingPictures=" + documentContainsOnlyFloatingPictures);

            return insertPreservesSourceAspect &&
                   replacePreservesSourceAspect &&
                   replacementRetainsWidth &&
                   replacementCorrectsHeight &&
                   insertUsesFloatingFront &&
                   replacementUsesFloatingFront &&
                   failedReplacementPreservedOriginal &&
                   verticalContained &&
                   documentContainsOnlyFloatingPictures ? 0 : 1;
        }
        finally
        {
            ReleasePicture(vertical);
            ReleasePicture(replaced);
            ReleasePicture(inserted);

            if (document != null)
            {
                try
                {
                    object saveChanges = WordInterop.WdSaveOptions.wdDoNotSaveChanges;
                    object originalFormat = Type.Missing;
                    object routeDocument = Type.Missing;
                    ((WordInterop._Document)document).Close(ref saveChanges, ref originalFormat, ref routeDocument);
                }
                catch { }
                try { Marshal.FinalReleaseComObject(document); } catch { }
            }

            if (application != null)
            {
                try
                {
                    object saveChanges = WordInterop.WdSaveOptions.wdDoNotSaveChanges;
                    object originalFormat = Type.Missing;
                    object routeDocument = Type.Missing;
                    ((WordInterop._Application)application).Quit(ref saveChanges, ref originalFormat, ref routeDocument);
                }
                catch { }
                try { Marshal.FinalReleaseComObject(application); } catch { }
            }
        }
    }
}
'@ | Set-Content -LiteralPath $programPath -Encoding UTF8

    Add-Type -AssemblyName System.Drawing
    $bitmap = New-Object System.Drawing.Bitmap 1200, 600
    try {
        $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
        try {
            $graphics.Clear([System.Drawing.Color]::FromArgb(234, 244, 255))
            $brush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(31, 95, 168))
            try {
                $graphics.FillRectangle($brush, 80, 100, 1040, 400)
            }
            finally {
                $brush.Dispose()
            }
        }
        finally {
            $graphics.Dispose()
        }

        $bitmap.Save($imagePath, [System.Drawing.Imaging.ImageFormat]::Png)
    }
    finally {
        $bitmap.Dispose()
    }

    $verticalBitmap = New-Object System.Drawing.Bitmap 300, 1200
    try {
        $verticalGraphics = [System.Drawing.Graphics]::FromImage($verticalBitmap)
        try {
            $verticalGraphics.Clear([System.Drawing.Color]::FromArgb(240, 248, 240))
            $verticalBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(48, 122, 70))
            try {
                $verticalGraphics.FillRectangle($verticalBrush, 40, 80, 220, 1040)
            }
            finally {
                $verticalBrush.Dispose()
            }
        }
        finally {
            $verticalGraphics.Dispose()
        }

        $verticalBitmap.Save($verticalImagePath, [System.Drawing.Imaging.ImageFormat]::Png)
    }
    finally {
        $verticalBitmap.Dispose()
    }

    Copy-Item -LiteralPath $coreAssemblyPath -Destination $tempRoot -Force
    Copy-Item -LiteralPath $powerPointAssemblyPath -Destination $tempRoot -Force
    Copy-Item -LiteralPath $wordAssemblyPath -Destination $tempRoot -Force

    $compileArguments = @(
        "/nologo",
        "/target:exe",
        "/out:$exePath",
        "/r:$interopWord",
        "/r:$officeCore",
        "/r:$([IO.Path]::Combine($tempRoot, 'DrawioPpt.Core.dll'))",
        "/r:$([IO.Path]::Combine($tempRoot, 'DrawioPpt.PowerPointAddIn.dll'))",
        "/r:$([IO.Path]::Combine($tempRoot, 'DrawioPpt.WordAddIn.dll'))",
        $programPath
    )
    & $cscPath @compileArguments
    if ($LASTEXITCODE -ne 0) {
        throw "Unable to compile Word SVG aspect-ratio E2E helper."
    }

    Push-Location $tempRoot
    try {
        & $exePath `
            $imagePath `
            $verticalImagePath `
            $documentPath `
            $missingImagePath `
            $processIdentityPath
        $testExitCode = [int]$LASTEXITCODE
    }
    finally {
        Pop-Location
    }

    if ($testExitCode -ne 0) {
        throw "Word SVG aspect-ratio E2E failed."
    }
}
finally {
    $cleanupFailures = New-Object System.Collections.Generic.List[string]
    try {
        $ownedProcessId = 0
        $ownedStartTicks = 0L
        if (Test-Path -LiteralPath $processIdentityPath) {
            $identityParts = (
                Get-Content -LiteralPath $processIdentityPath -Raw
            ).Trim().Split("|")
            if ($identityParts.Count -eq 2) {
                $ownedProcessId = [int]$identityParts[0]
                $ownedStartTicks = [long]$identityParts[1]
            }
        }

        if ($ownedProcessId -gt 0) {
            $cleanupDeadline = [DateTime]::UtcNow.AddSeconds(10)
            do {
                $ownedWord = Get-Process -Id $ownedProcessId -ErrorAction SilentlyContinue
                if ($null -eq $ownedWord -or [DateTime]::UtcNow -ge $cleanupDeadline) {
                    break
                }

                Start-Sleep -Milliseconds 100
            } while ($true)

            if ($null -ne $ownedWord) {
                $actualStartTicks = $ownedWord.StartTime.ToUniversalTime().Ticks
                if ($ownedWord.ProcessName -ne "WINWORD" -or
                    $actualStartTicks -ne $ownedStartTicks) {
                    $cleanupFailures.Add(
                        "测试拥有的 Word 进程身份已变化，拒绝终止 PID $ownedProcessId。")
                }
                else {
                    Stop-Process -Id $ownedProcessId -Force -ErrorAction Stop
                    Wait-Process -Id $ownedProcessId -Timeout 10 -ErrorAction SilentlyContinue
                    if (Get-Process -Id $ownedProcessId -ErrorAction SilentlyContinue) {
                        $cleanupFailures.Add(
                            "SVG 比例 E2E 无法终止测试拥有的 WINWORD 进程：$ownedProcessId")
                    }
                }
            }
        }

        $unexpectedWord = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue)
        if ($unexpectedWord.Count -gt 0) {
            $processIds = ($unexpectedWord | ForEach-Object Id) -join ","
            $cleanupFailures.Add(
                "SVG 比例 E2E 后存在非测试身份的 WINWORD 进程：$processIds")
        }
    }
    catch {
        $cleanupFailures.Add($_.Exception.Message)
    }

    try {
        Remove-TestTempRoot
    }
    catch {
        $cleanupFailures.Add($_.Exception.Message)
    }

    if ($cleanupFailures.Count -gt 0) {
        throw ($cleanupFailures -join [Environment]::NewLine)
    }
}

Write-Output "WORD_SVG_ASPECT_RATIO_E2E_PASS"
