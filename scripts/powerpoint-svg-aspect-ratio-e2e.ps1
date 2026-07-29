param(
    [string]$Configuration = "Release",
    [string]$AssemblyRoot,
    [switch]$SkipBuild
)

$ErrorActionPreference = "Stop"

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$buildScript = Join-Path $PSScriptRoot "build.ps1"
$tempRoot = Join-Path $env:TEMP ("DrawioPpt\powerpoint-svg-aspect-ratio-" + [Guid]::NewGuid().ToString("N"))
$tempRootPrefix = Join-Path $env:TEMP "DrawioPpt\powerpoint-svg-aspect-ratio-"
$programPath = Join-Path $tempRoot "powerpoint-svg-aspect-ratio.cs"
$exePath = Join-Path $tempRoot "powerpoint-svg-aspect-ratio.exe"
$svgPath = Join-Path $tempRoot "wide-source.svg"
$sizelessSvgPath = Join-Path $tempRoot "sizeless-source.svg"
$processIdentityPath = Join-Path $tempRoot "powerpoint-process.txt"
$cscPath = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
$interopPowerPoint = "C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.PowerPoint\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.PowerPoint.dll"
$officeCore = "C:\Windows\assembly\GAC_MSIL\office\15.0.0.0__71e9bce111e9429c\OFFICE.DLL"

function Assert-NoRunningPowerPoint {
    if (Get-Process -Name POWERPNT -ErrorAction SilentlyContinue) {
        throw "请先关闭正在运行的 PowerPoint，再执行 SVG 比例 E2E 测试。"
    }
}

function Get-TestOwnedPowerPointProcess {
    if (-not (Test-Path -LiteralPath $processIdentityPath)) {
        return $null
    }

    $identity = (Get-Content -LiteralPath $processIdentityPath -Raw).Trim().Split("|")
    if ($identity.Count -ne 2) {
        throw "测试拥有的 PowerPoint 进程标识格式无效。"
    }

    $processId = [int]$identity[0]
    $expectedStartTimeTicks = [long]$identity[1]
    $process = Get-Process -Id $processId -ErrorAction SilentlyContinue
    if ($null -eq $process) {
        return $null
    }

    if (-not [string]::Equals($process.ProcessName, "POWERPNT", [System.StringComparison]::OrdinalIgnoreCase) -or
        $process.StartTime.ToUniversalTime().Ticks -ne $expectedStartTimeTicks) {
        throw "拒绝清理身份不匹配的 PowerPoint 进程：$processId"
    }

    return $process
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
    $AssemblyRoot = Join-Path $repoRoot ("src\DrawioPpt.PowerPointAddIn\bin\x64\" + $Configuration)
}

if (Test-Path -LiteralPath $AssemblyRoot) {
    $AssemblyRoot = (Resolve-Path -LiteralPath $AssemblyRoot).Path
}

$powerPointAssemblyPath = Join-Path $AssemblyRoot "DrawioPpt.PowerPointAddIn.dll"
$coreAssemblyPath = Join-Path $AssemblyRoot "DrawioPpt.Core.dll"
foreach ($requiredPath in @($powerPointAssemblyPath, $coreAssemblyPath, $interopPowerPoint, $officeCore)) {
    if (-not (Test-Path -LiteralPath $requiredPath)) {
        throw "Required test assembly not found: $requiredPath"
    }
}

Assert-NoRunningPowerPoint
New-Item -ItemType Directory -Force -Path $tempRoot | Out-Null

try {
    @'
using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using DrawioPpt.PowerPointAddIn.PowerPoint;
using DrawioPpt.PowerPointAddIn.Services;
using Microsoft.Office.Core;
using PptInterop = Microsoft.Office.Interop.PowerPoint;

public static class PowerPointSvgAspectRatioE2E
{
    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr windowHandle, out uint processId);

    private static void WritePowerPointProcessIdentity(PptInterop.Application application, string identityPath)
    {
        uint processId;
        GetWindowThreadProcessId(new IntPtr(application.HWND), out processId);
        Process process = Process.GetProcessById((int)processId);
        File.WriteAllText(
            identityPath,
            process.Id + "|" + process.StartTime.ToUniversalTime().Ticks);
    }

    [STAThread]
    public static int Main(string[] args)
    {
        PptInterop.Application application = null;
        PptInterop.Presentation presentation = null;
        PptInterop.Slide slide = null;
        PptInterop.Shape inserted = null;
        PptInterop.Shape replaced = null;
        PptInterop.Shape fallbackInserted = null;
        PptInterop.Shape fallbackReplaced = null;

        try
        {
            Type powerPointType = Type.GetTypeFromProgID("PowerPoint.Application", true);
            application = (PptInterop.Application)Activator.CreateInstance(powerPointType);
            application.Visible = MsoTriState.msoTrue;
            WritePowerPointProcessIdentity(application, args[2]);

            presentation = application.Presentations.Add(MsoTriState.msoTrue);
            slide = presentation.Slides.Add(1, PptInterop.PpSlideLayout.ppLayoutBlank);
            slide.Select();

            PowerPointSvgShapeService service = new PowerPointSvgShapeService(
                new SelectedShapeAccessor(),
                new SvgPowerPointSupportService());

            float slideWidth = presentation.PageSetup.SlideWidth;
            float slideHeight = presentation.PageSetup.SlideHeight;
            inserted = service.InsertOnActiveSlide(application, args[0], "Wide SVG");

            float insertedWidth = inserted.Width;
            float insertedHeight = inserted.Height;
            float insertedLeft = inserted.Left;
            float insertedTop = inserted.Top;
            float insertedRatio = insertedWidth / insertedHeight;
            float maxWidth = slideWidth * 0.5f;
            float maxHeight = slideHeight * 0.38f;
            float expectedScale = Math.Min(maxWidth / 1200f, maxHeight / 600f);
            float expectedWidth = 1200f * expectedScale;
            float expectedHeight = 600f * expectedScale;
            bool insertPreservesSourceAspect = Math.Abs(insertedRatio - 2f) < 0.02f;
            bool insertFitsWidthBoundary = insertedWidth <= maxWidth + 0.1f;
            bool insertFitsHeightBoundary = insertedHeight <= maxHeight + 0.1f;
            bool insertUsesContainSize =
                Math.Abs(insertedWidth - expectedWidth) < 0.1f &&
                Math.Abs(insertedHeight - expectedHeight) < 0.1f;
            bool insertIsHorizontallyCentered = Math.Abs(
                (insertedLeft + insertedWidth / 2f) - (slideWidth / 2f)) < 0.1f;
            bool insertIsVerticallyCentered = Math.Abs(
                (insertedTop + insertedHeight / 2f) - (slideHeight / 2f)) < 0.1f;
            bool insertLocksAspect = inserted.LockAspectRatio == MsoTriState.msoTrue;

            inserted.LockAspectRatio = MsoTriState.msoFalse;
            inserted.Left = 100f;
            inserted.Top = 120f;
            inserted.Width = 400f;
            inserted.Height = 400f;
            float legacyCenterY = inserted.Top + inserted.Height / 2f;

            replaced = service.Replace(inserted, args[0]);
            inserted = null;

            float replacedRatio = replaced.Width / replaced.Height;
            bool replacePreservesSourceAspect = Math.Abs(replacedRatio - 2f) < 0.02f;
            bool replacementRetainsWidth = Math.Abs(replaced.Width - 400f) < 0.1f;
            bool replacementCorrectsHeight = Math.Abs(replaced.Height - 200f) < 0.1f;
            bool replacementKeepsVerticalCenter = Math.Abs(
                (replaced.Top + replaced.Height / 2f) - legacyCenterY) < 0.1f;
            bool replacementKeepsLeft = Math.Abs(replaced.Left - 100f) < 0.1f;
            bool replacementLocksAspect = replaced.LockAspectRatio == MsoTriState.msoTrue;

            fallbackInserted = service.InsertOnActiveSlide(application, args[1], "Sizeless SVG");
            float fallbackInsertedWidth = fallbackInserted.Width;
            float fallbackInsertedHeight = fallbackInserted.Height;
            MsoTriState fallbackInsertedLockAspectRatio = fallbackInserted.LockAspectRatio;
            bool fallbackInsertUsesLegacyGeometry =
                Math.Abs(fallbackInsertedWidth - maxWidth) < 0.1f &&
                Math.Abs(fallbackInsertedHeight - maxHeight) < 0.1f;
            bool fallbackInsertKeepsDefaultLock =
                fallbackInsertedLockAspectRatio == MsoTriState.msoFalse;

            fallbackInserted.LockAspectRatio = MsoTriState.msoFalse;
            fallbackInserted.Left = 100f;
            fallbackInserted.Top = 120f;
            fallbackInserted.Width = 400f;
            fallbackInserted.Height = 400f;

            fallbackReplaced = service.Replace(fallbackInserted, args[1]);
            fallbackInserted = null;

            bool fallbackReplaceKeepsGeometry =
                Math.Abs(fallbackReplaced.Left - 100f) < 0.1f &&
                Math.Abs(fallbackReplaced.Top - 120f) < 0.1f &&
                Math.Abs(fallbackReplaced.Width - 400f) < 0.1f &&
                Math.Abs(fallbackReplaced.Height - 400f) < 0.1f;
            bool fallbackReplaceKeepsLock =
                fallbackReplaced.LockAspectRatio == MsoTriState.msoFalse;

            Console.WriteLine("SlideWidth=" + slideWidth.ToString("0.###"));
            Console.WriteLine("SlideHeight=" + slideHeight.ToString("0.###"));
            Console.WriteLine("MaxWidth=" + maxWidth.ToString("0.###"));
            Console.WriteLine("MaxHeight=" + maxHeight.ToString("0.###"));
            Console.WriteLine("InsertedWidth=" + insertedWidth.ToString("0.###"));
            Console.WriteLine("InsertedHeight=" + insertedHeight.ToString("0.###"));
            Console.WriteLine("InsertedRatio=" + insertedRatio.ToString("0.###"));
            Console.WriteLine("InsertPreservesSourceAspect=" + insertPreservesSourceAspect);
            Console.WriteLine("InsertFitsWidthBoundary=" + insertFitsWidthBoundary);
            Console.WriteLine("InsertFitsHeightBoundary=" + insertFitsHeightBoundary);
            Console.WriteLine("InsertUsesContainSize=" + insertUsesContainSize);
            Console.WriteLine("InsertIsHorizontallyCentered=" + insertIsHorizontallyCentered);
            Console.WriteLine("InsertIsVerticallyCentered=" + insertIsVerticallyCentered);
            Console.WriteLine("InsertLocksAspect=" + insertLocksAspect);
            Console.WriteLine("ReplacedWidth=" + replaced.Width.ToString("0.###"));
            Console.WriteLine("ReplacedHeight=" + replaced.Height.ToString("0.###"));
            Console.WriteLine("ReplacedLeft=" + replaced.Left.ToString("0.###"));
            Console.WriteLine("ReplacedTop=" + replaced.Top.ToString("0.###"));
            Console.WriteLine("ReplacedRatio=" + replacedRatio.ToString("0.###"));
            Console.WriteLine("ReplacePreservesSourceAspect=" + replacePreservesSourceAspect);
            Console.WriteLine("ReplacementRetainsWidth=" + replacementRetainsWidth);
            Console.WriteLine("ReplacementCorrectsHeight=" + replacementCorrectsHeight);
            Console.WriteLine("ReplacementKeepsVerticalCenter=" + replacementKeepsVerticalCenter);
            Console.WriteLine("ReplacementKeepsLeft=" + replacementKeepsLeft);
            Console.WriteLine("ReplacementLocksAspect=" + replacementLocksAspect);
            Console.WriteLine("FallbackInsertedWidth=" + fallbackInsertedWidth.ToString("0.###"));
            Console.WriteLine("FallbackInsertedHeight=" + fallbackInsertedHeight.ToString("0.###"));
            Console.WriteLine("FallbackInsertedLock=" + fallbackInsertedLockAspectRatio);
            Console.WriteLine("FallbackInsertUsesLegacyGeometry=" + fallbackInsertUsesLegacyGeometry);
            Console.WriteLine("FallbackInsertKeepsDefaultLock=" + fallbackInsertKeepsDefaultLock);
            Console.WriteLine("FallbackReplacedLeft=" + fallbackReplaced.Left.ToString("0.###"));
            Console.WriteLine("FallbackReplacedTop=" + fallbackReplaced.Top.ToString("0.###"));
            Console.WriteLine("FallbackReplacedWidth=" + fallbackReplaced.Width.ToString("0.###"));
            Console.WriteLine("FallbackReplacedHeight=" + fallbackReplaced.Height.ToString("0.###"));
            Console.WriteLine("FallbackReplacedLock=" + fallbackReplaced.LockAspectRatio);
            Console.WriteLine("FallbackReplaceKeepsGeometry=" + fallbackReplaceKeepsGeometry);
            Console.WriteLine("FallbackReplaceKeepsLock=" + fallbackReplaceKeepsLock);

            return insertPreservesSourceAspect &&
                   insertFitsWidthBoundary &&
                   insertFitsHeightBoundary &&
                   insertUsesContainSize &&
                   insertIsHorizontallyCentered &&
                   insertIsVerticallyCentered &&
                   insertLocksAspect &&
                   replacePreservesSourceAspect &&
                   replacementRetainsWidth &&
                   replacementCorrectsHeight &&
                   replacementKeepsVerticalCenter &&
                   replacementKeepsLeft &&
                   replacementLocksAspect &&
                   fallbackInsertUsesLegacyGeometry &&
                   fallbackInsertKeepsDefaultLock &&
                   fallbackReplaceKeepsGeometry &&
                   fallbackReplaceKeepsLock ? 0 : 1;
        }
        finally
        {
            if (fallbackReplaced != null)
            {
                try { Marshal.FinalReleaseComObject(fallbackReplaced); } catch { }
            }

            if (fallbackInserted != null)
            {
                try { Marshal.FinalReleaseComObject(fallbackInserted); } catch { }
            }

            if (replaced != null)
            {
                try { Marshal.FinalReleaseComObject(replaced); } catch { }
            }

            if (inserted != null)
            {
                try { Marshal.FinalReleaseComObject(inserted); } catch { }
            }

            if (slide != null)
            {
                try { Marshal.FinalReleaseComObject(slide); } catch { }
            }

            if (presentation != null)
            {
                try
                {
                    presentation.Saved = MsoTriState.msoTrue;
                    presentation.Close();
                }
                catch { }
                try { Marshal.FinalReleaseComObject(presentation); } catch { }
            }

            if (application != null)
            {
                try { application.Quit(); } catch { }
                try { Marshal.FinalReleaseComObject(application); } catch { }
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}
'@ | Set-Content -LiteralPath $programPath -Encoding UTF8

    @'
<svg xmlns="http://www.w3.org/2000/svg" width="1200" height="600" viewBox="0 0 1200 600">
  <rect width="1200" height="600" fill="#eaf4ff"/>
  <rect x="80" y="100" width="1040" height="400" rx="24" fill="#ffffff" stroke="#1f5fa8" stroke-width="12"/>
  <text x="600" y="330" text-anchor="middle" font-size="72" fill="#17324d">2:1 source aspect ratio</text>
</svg>
'@ | Set-Content -LiteralPath $svgPath -Encoding UTF8

    @'
<svg xmlns="http://www.w3.org/2000/svg">
  <rect x="0" y="0" width="1200" height="600" fill="#f4f4f4"/>
  <text x="600" y="330" text-anchor="middle" font-size="72" fill="#333333">no intrinsic root size</text>
</svg>
'@ | Set-Content -LiteralPath $sizelessSvgPath -Encoding UTF8

    Copy-Item -LiteralPath $coreAssemblyPath -Destination $tempRoot -Force
    Copy-Item -LiteralPath $powerPointAssemblyPath -Destination $tempRoot -Force

    $compileArguments = @(
        "/nologo",
        "/target:exe",
        "/platform:x64",
        "/out:$exePath",
        "/r:$interopPowerPoint",
        "/r:$officeCore",
        "/r:$([IO.Path]::Combine($tempRoot, 'DrawioPpt.Core.dll'))",
        "/r:$([IO.Path]::Combine($tempRoot, 'DrawioPpt.PowerPointAddIn.dll'))",
        $programPath
    )
    & $cscPath @compileArguments
    if ($LASTEXITCODE -ne 0) {
        throw "Unable to compile PowerPoint SVG aspect-ratio E2E helper."
    }

    Push-Location $tempRoot
    try {
        & $exePath $svgPath $sizelessSvgPath $processIdentityPath
        $testExitCode = [int]$LASTEXITCODE
    }
    finally {
        Pop-Location
    }

    if ($testExitCode -ne 0) {
        throw "PowerPoint SVG aspect-ratio E2E failed."
    }
}
finally {
    try {
        $cleanupDeadline = [DateTime]::UtcNow.AddSeconds(10)
        do {
            $testOwnedPowerPoint = Get-TestOwnedPowerPointProcess
            if ($null -eq $testOwnedPowerPoint -or [DateTime]::UtcNow -ge $cleanupDeadline) {
                break
            }

            Start-Sleep -Milliseconds 100
        } while ($true)

        if ($null -ne $testOwnedPowerPoint) {
            Stop-Process -Id $testOwnedPowerPoint.Id -Force
            $testOwnedPowerPoint.WaitForExit(5000)
        }
    }
    finally {
        Remove-TestTempRoot
    }

    $residualPowerPoint = @(Get-Process -Name POWERPNT -ErrorAction SilentlyContinue)
    if ($residualPowerPoint.Count -gt 0) {
        $processIds = ($residualPowerPoint | ForEach-Object Id) -join ","
        throw "SVG 比例 E2E 后存在无法确认归属的 POWERPNT 进程：$processIds"
    }
}

Write-Output "POWERPOINT_SVG_ASPECT_RATIO_E2E_PASS"
