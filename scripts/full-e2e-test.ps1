param(
    [string]$Version = "v0.6.1-stable",
    [switch]$SkipBuild,
    [switch]$KeepInstalled
)

$ErrorActionPreference = "Stop"

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$buildScript = Join-Path $PSScriptRoot "build.ps1"
$packageScript = Join-Path $PSScriptRoot "package-release.ps1"
$registerRepoScript = Join-Path $PSScriptRoot "register-addin.ps1"
$releaseRoot = Join-Path $repoRoot ("artifacts\\releases\\" + $Version)
$packageRoot = Join-Path $releaseRoot "package"
$installRoot = Join-Path $env:TEMP ("DrawioPpt\\installed-" + [Guid]::NewGuid().ToString("N"))
$reportRoot = Join-Path $repoRoot "artifacts\\test-reports"
$reportPath = Join-Path $reportRoot ("full-e2e-" + $Version + ".md")
$settingsPath = Join-Path $env:APPDATA "Greensoft\\DrawioPpt\\settings.xml"
$settingsBackupPath = Join-Path $env:TEMP ("DrawioPpt\\settings-backup-" + [Guid]::NewGuid().ToString("N") + ".xml")
$results = New-Object System.Collections.Generic.List[string]
$drawioExe = "C:\\Program Files\\draw.io\\draw.io.exe"

function Assert-NoRunningPowerPoint {
    $runningPowerPoint = Get-Process -Name POWERPNT -ErrorAction SilentlyContinue
    if ($runningPowerPoint) {
        throw "请先关闭正在运行的 PowerPoint，再执行完整 E2E 测试。"
    }
}

function Stop-RunningPowerPointSilently {
    Get-Process -Name POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
}

function Add-Result {
    param(
        [string]$Name,
        [bool]$Passed,
        [string]$Detail
    )

    $status = if ($Passed) { "PASS" } else { "FAIL" }
    $script:results.Add("| $Name | $status | $Detail |") | Out-Null
    if (-not $Passed) {
        throw "$Name failed: $Detail"
    }
}

function Backup-Settings {
    if (Test-Path $settingsPath) {
        $directory = Split-Path -Parent $settingsBackupPath
        if (-not (Test-Path $directory)) {
            New-Item -ItemType Directory -Force -Path $directory | Out-Null
        }

        Copy-Item $settingsPath $settingsBackupPath -Force
    }
}

function Restore-Settings {
    if (Test-Path $settingsBackupPath) {
        $settingsDirectory = Split-Path -Parent $settingsPath
        if (-not (Test-Path $settingsDirectory)) {
            New-Item -ItemType Directory -Force -Path $settingsDirectory | Out-Null
        }

        Copy-Item $settingsBackupPath $settingsPath -Force
        Remove-Item $settingsBackupPath -Force
        return
    }

    if (Test-Path $settingsPath) {
        Remove-Item $settingsPath -Force
    }
}

function Compile-And-Run-PowerPointUrlE2E {
    param(
        [string]$InstalledRoot
    )

    $tempRoot = Join-Path $env:TEMP ("DrawioPpt\\ppt-url-e2e-" + [Guid]::NewGuid().ToString("N"))
    $serverPort = Get-Random -Minimum 8700 -Maximum 8999
    $mockHtmlPath = Join-Path $tempRoot "mock-editor.html"
    $programPath = Join-Path $tempRoot "powerpoint-url-e2e.cs"
    $exePath = Join-Path $tempRoot "powerpoint-url-e2e.exe"
    $cscPath = "C:\\Windows\\Microsoft.NET\\Framework64\\v4.0.30319\\csc.exe"
    $interopPowerPoint = "C:\\Windows\\assembly\\GAC_MSIL\\Microsoft.Office.Interop.PowerPoint\\15.0.0.0__71e9bce111e9429c\\Microsoft.Office.Interop.PowerPoint.dll"
    $officeCore = "C:\\Windows\\assembly\\GAC_MSIL\\office\\15.0.0.0__71e9bce111e9429c\\OFFICE.DLL"

    if (-not (Test-Path $cscPath)) {
        throw "csc.exe not found: $cscPath"
    }

    $python = Get-Command python -ErrorAction SilentlyContinue
    if (-not $python) {
        $python = Get-Command py -ErrorAction SilentlyContinue
    }

    if (-not $python) {
        throw "Python is required to host the mock URL editor."
    }

    New-Item -ItemType Directory -Force -Path $tempRoot | Out-Null

    @'
<!DOCTYPE html>
<html>
<head><meta charset="utf-8" /></head>
<body>
<script>
(function () {
  var waitingForConfigure = window.location.search.indexOf('configure=1') >= 0;
  function send(message) { window.parent.postMessage(message, '*'); }
  window.addEventListener('message', function (event) {
    var data = event.data || {};
    if (data.action === 'configure') {
      waitingForConfigure = false;
      setTimeout(function () { send({ event: 'init' }); }, 50);
      return;
    }
    if (data.action === 'load') {
      window.__xml = data.xml || '';
      setTimeout(function () { send({ event: 'save', xml: window.__xml, exit: 1 }); }, 120);
      return;
    }
    if (data.action === 'export') {
      var svg = '<svg xmlns="http://www.w3.org/2000/svg" content="old"><text x="10" y="20">e2e</text></svg>';
      setTimeout(function () { send({ event: 'export', data: 'data:image/svg+xml;utf8,' + encodeURIComponent(svg) }); }, 120);
    }
  });
  setTimeout(function () {
    if (waitingForConfigure) {
      send({ event: 'configure' });
      return;
    }

    send({ event: 'init' });
  }, 120);
})();
</script>
</body>
</html>
'@ | Set-Content -Path $mockHtmlPath -Encoding UTF8

    @"
using System;
using System.IO;
using System.Windows.Forms;
using DrawioPpt.Core.Models;
using DrawioPpt.Core.Services;
using DrawioPpt.PowerPointAddIn.Services;
using Microsoft.Office.Core;
using PptInterop = Microsoft.Office.Interop.PowerPoint;

public static class PowerPointUrlE2E
{
    [STAThread]
    public static int Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);

        PluginSettings settings = new PluginSettings();
        settings.EditorMode = EditorMode.Url;
        settings.EditorUrl = "http://127.0.0.1:$serverPort/mock-editor.html";
        settings.UseOfficeCompatibleSvgLabels = true;
        settings.AutoOpenOnSelection = false;
        settings.AutoUpdateOnSave = true;
        settings.KeepSidecarFile = true;
        settings.SidecarFolderName = "drawio-e2e";
        new FilePluginSettingsStore().Save(settings);

        PptInterop.Application application = null;
        AddInHost host = null;
        PptInterop.Presentation presentation = null;

        try
        {
            Type pptType = Type.GetTypeFromProgID("PowerPoint.Application", true);
            application = (PptInterop.Application)Activator.CreateInstance(pptType);
            application.Visible = MsoTriState.msoTrue;

            presentation = application.Presentations.Add(MsoTriState.msoTrue);
            string presentationPath = Path.Combine(Path.GetTempPath(), "DrawioPpt", "ppt-url-e2e-" + Guid.NewGuid().ToString("N") + ".pptx");
            string presentationDirectory = Path.GetDirectoryName(presentationPath);
            if (!Directory.Exists(presentationDirectory))
            {
                Directory.CreateDirectory(presentationDirectory);
            }

            presentation.SaveAs(presentationPath, PptInterop.PpSaveAsFileType.ppSaveAsOpenXMLPresentation, MsoTriState.msoFalse);
            if (presentation.Slides.Count == 0)
            {
                presentation.Slides.Add(1, PptInterop.PpSlideLayout.ppLayoutBlank);
            }

            presentation.Slides[1].Select();

            host = new AddInHost(application);
            host.Start();

            host.CreateNewDiagram();
            System.Threading.Thread.Sleep(1200);

            PptInterop.Slide slide = presentation.Slides[1];
            PptInterop.Shape shape = slide.Shapes[slide.Shapes.Count];
            string diagramId = shape.Tags["DRAWIO_PPT_ID"] ?? string.Empty;
            string partId = shape.Tags["DRAWIO_PPT_PART_ID"] ?? string.Empty;
            bool createdManagedShape = !string.IsNullOrWhiteSpace(diagramId) && !string.IsNullOrWhiteSpace(partId);

            shape.Select(MsoTriState.msoTrue);
            host.EditSelectedDiagram();
            System.Threading.Thread.Sleep(1200);

            slide = presentation.Slides[1];
            shape = slide.Shapes[slide.Shapes.Count];
            string editedPartId = shape.Tags["DRAWIO_PPT_PART_ID"] ?? string.Empty;
            string alternativeText = shape.AlternativeText ?? string.Empty;
            bool editedManagedShape = !string.IsNullOrWhiteSpace(editedPartId) && alternativeText.Length > 0;
            bool customXmlStored = presentation.CustomXMLParts.Count > 0;

            Console.WriteLine("CreatedManagedShape=" + createdManagedShape);
            Console.WriteLine("EditedManagedShape=" + editedManagedShape);
            Console.WriteLine("CustomXmlCount=" + presentation.CustomXMLParts.Count);
            Console.WriteLine("ShapeCount=" + slide.Shapes.Count);

            return createdManagedShape && editedManagedShape && customXmlStored ? 0 : 1;
        }
        catch (Exception ex)
        {
            Console.WriteLine("CaughtType=" + ex.GetType().FullName);
            Console.WriteLine("CaughtMessage=" + ex.Message);
            Console.WriteLine("CaughtStack=" + ex.StackTrace);
            if (ex.InnerException != null)
            {
                Console.WriteLine("InnerType=" + ex.InnerException.GetType().FullName);
                Console.WriteLine("InnerMessage=" + ex.InnerException.Message);
                Console.WriteLine("InnerStack=" + ex.InnerException.StackTrace);
            }

            return 99;
        }
        finally
        {
            if (presentation != null)
            {
                try
                {
                    presentation.Saved = MsoTriState.msoTrue;
                    presentation.Close();
                }
                catch
                {
                }
            }

            if (host != null)
            {
                host.Dispose();
            }

            if (application != null)
            {
                try
                {
                    application.Quit();
                }
                catch
                {
                }
            }
        }
    }
}
"@ | Set-Content -Path $programPath -Encoding UTF8

    Copy-Item (Join-Path $InstalledRoot "bin\\DrawioPpt.Core.dll") $tempRoot -Force
    Copy-Item (Join-Path $InstalledRoot "bin\\DrawioPpt.PowerPointAddIn.dll") $tempRoot -Force
    Copy-Item (Join-Path $InstalledRoot "bin\\Microsoft.Web.WebView2.Core.dll") $tempRoot -Force
    Copy-Item (Join-Path $InstalledRoot "bin\\Microsoft.Web.WebView2.WinForms.dll") $tempRoot -Force
    Copy-Item (Join-Path $InstalledRoot "bin\\runtimes\\win-x64\\native\\WebView2Loader.dll") $tempRoot -Force

    $compileArguments = @(
        "/nologo",
        "/t:exe",
        "/platform:x64",
        "/out:$exePath",
        "/r:$([IO.Path]::Combine($tempRoot, 'DrawioPpt.Core.dll'))",
        "/r:$([IO.Path]::Combine($tempRoot, 'DrawioPpt.PowerPointAddIn.dll'))",
        "/r:$([IO.Path]::Combine($tempRoot, 'Microsoft.Web.WebView2.Core.dll'))",
        "/r:$([IO.Path]::Combine($tempRoot, 'Microsoft.Web.WebView2.WinForms.dll'))",
        "/r:$interopPowerPoint",
        "/r:$officeCore",
        "/r:System.dll",
        "/r:System.Core.dll",
        "/r:System.Windows.Forms.dll",
        "/r:System.Drawing.dll",
        "/r:System.Xml.dll",
        "/r:System.Xml.Linq.dll",
        $programPath
    )

    & $cscPath @compileArguments
    if ($LASTEXITCODE -ne 0) {
        exit $LASTEXITCODE
    }

    $server = Start-Process -FilePath $python.Source -ArgumentList @("-m", "http.server", "$serverPort", "--bind", "127.0.0.1") -WorkingDirectory $tempRoot -WindowStyle Hidden -PassThru
    try {
        Start-Sleep -Seconds 2
        Push-Location $tempRoot
        try {
            $runOutput = & $exePath 2>&1
            $exitCode = $LASTEXITCODE
            if ($runOutput) {
                $runOutput | ForEach-Object { Write-Host $_ }
            }

            return [int]$exitCode
        }
        finally {
            Pop-Location
        }
    }
    finally {
        if ($server -and -not $server.HasExited) {
            Stop-Process -Id $server.Id -Force
        }
    }
}

function Test-DesktopExporter {
    param(
        [string]$ExecutablePath
    )

    if (-not (Test-Path $ExecutablePath)) {
        throw "draw.io Desktop not found: $ExecutablePath"
    }

    $tempRoot = Join-Path $env:TEMP ("DrawioPpt\\desktop-e2e-" + [Guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Force -Path $tempRoot | Out-Null
    $drawioPath = Join-Path $tempRoot "sample.drawio"
    $svgPath = Join-Path $tempRoot "sample.svg"

    @'
<mxfile host="DrawioPpt">
  <diagram id="desktop-e2e" name="Desktop E2E">
    <mxGraphModel dx="1000" dy="1000" grid="1" gridSize="10" guides="1" tooltips="1" connect="1" arrows="1" fold="1" page="1" pageScale="1" pageWidth="850" pageHeight="1100" math="0" shadow="0">
      <root>
        <mxCell id="0"/>
        <mxCell id="1" parent="0"/>
        <mxCell id="2" value="Desktop E2E" style="rounded=1;whiteSpace=wrap;html=1;" vertex="1" parent="1">
          <mxGeometry x="120" y="120" width="160" height="60" as="geometry"/>
        </mxCell>
      </root>
    </mxGraphModel>
  </diagram>
</mxfile>
'@ | Set-Content -Path $drawioPath -Encoding UTF8

    & $ExecutablePath --export --format svg --embed-diagram --embed-svg-images --output $svgPath $drawioPath | Out-Null
    if ($LASTEXITCODE -ne 0) {
        return $false
    }

    return (Test-Path $svgPath) -and ((Get-Content $svgPath -Raw) -like "*<svg*")
}

Backup-Settings

try {
    Assert-NoRunningPowerPoint

    if (-not $SkipBuild) {
        & powershell -ExecutionPolicy Bypass -File $buildScript -Configuration Release -Platform x64
        if ($LASTEXITCODE -ne 0) {
            exit $LASTEXITCODE
        }
    }

    & powershell -ExecutionPolicy Bypass -File $packageScript -Version $Version -Configuration Release -Platform x64 -SkipBuild
    if ($LASTEXITCODE -ne 0) {
        exit $LASTEXITCODE
    }
    Add-Result -Name "ReleasePackage" -Passed (Test-Path $packageRoot) -Detail $packageRoot

    $installScript = Join-Path $packageRoot "scripts\\install-release.ps1"
    & powershell -ExecutionPolicy Bypass -File $installScript -InstallRoot $installRoot
    if ($LASTEXITCODE -ne 0) {
        exit $LASTEXITCODE
    }
    Add-Result -Name "InstallRelease" -Passed (Test-Path (Join-Path $installRoot "bin\\DrawioPpt.PowerPointAddIn.dll")) -Detail $installRoot

    $pp = New-Object -ComObject PowerPoint.Application
    try {
        $addin = $pp.COMAddIns.Item("Greensoft.DrawioPptAddIn")
        $addin.Connect = $true
        Add-Result -Name "PowerPointAddInLoad" -Passed ([bool]$addin.Connect) -Detail "Connect=$($addin.Connect)"
    }
    finally {
        $pp.Quit() | Out-Null
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($pp)
    }

    $installedUrlSmoke = Join-Path $installRoot "scripts\\url-editor-smoke.ps1"
    & powershell -ExecutionPolicy Bypass -File $installedUrlSmoke -SkipBuild
    Add-Result -Name "InstalledUrlSmoke" -Passed ($LASTEXITCODE -eq 0) -Detail $installedUrlSmoke

    $urlE2EExitCode = Compile-And-Run-PowerPointUrlE2E -InstalledRoot $installRoot
    Add-Result -Name "PowerPointUrlE2E" -Passed ($urlE2EExitCode -eq 0) -Detail "ExitCode=$urlE2EExitCode"

    $desktopExporterPassed = Test-DesktopExporter -ExecutablePath $drawioExe
    Add-Result -Name "DesktopExporter" -Passed $desktopExporterPassed -Detail $drawioExe
}
finally {
    Stop-RunningPowerPointSilently

    if (-not $KeepInstalled -and (Test-Path $installRoot)) {
        $uninstallScript = Join-Path $installRoot "scripts\\uninstall-release.ps1"
        if (Test-Path $uninstallScript) {
            & powershell -ExecutionPolicy Bypass -File $uninstallScript -InstallRoot $installRoot | Out-Null
        }
    }

    if (Test-Path $registerRepoScript) {
        & powershell -ExecutionPolicy Bypass -File $registerRepoScript | Out-Null
    }

    Restore-Settings
}

if (-not (Test-Path $reportRoot)) {
    New-Item -ItemType Directory -Force -Path $reportRoot | Out-Null
}

@(
    "# Full E2E Test Report ($Version)",
    "",
    "| Check | Result | Detail |",
    "| --- | --- | --- |"
) + $results | Set-Content -Path $reportPath -Encoding UTF8

Write-Host "E2E report written to: $reportPath"
