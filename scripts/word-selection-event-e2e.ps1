param(
    [string]$Configuration = "Debug",
    [string]$AssemblyRoot,
    [switch]$SkipSourceCheck,
    [switch]$SkipBuild
)

$ErrorActionPreference = "Stop"

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$selectionSyncTest = Join-Path $PSScriptRoot "word-selection-sync-test.ps1"
$selectionSyncSource = Join-Path $repoRoot "src\DrawioPpt.WordAddIn\Services\AddInHost.cs"
$buildScript = Join-Path $PSScriptRoot "build.ps1"
$tempRoot = Join-Path $env:TEMP ("DrawioPpt\word-selection-event-e2e-" + [Guid]::NewGuid().ToString("N"))
$programPath = Join-Path $tempRoot "word-selection-event-e2e.cs"
$exePath = Join-Path $tempRoot "word-selection-event-e2e.exe"
$imagePath = Join-Path $tempRoot "selection-test.png"
$documentPath = Join-Path $tempRoot "word-selection-event.docx"
$cscPath = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
$interopWord = "C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Word\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.Word.dll"
$officeCore = "C:\Windows\assembly\GAC_MSIL\office\15.0.0.0__71e9bce111e9429c\OFFICE.DLL"
$wordBin = $AssemblyRoot
$tempRootPrefix = Join-Path $env:TEMP "DrawioPpt\word-selection-event-e2e-"
$settingsPath = Join-Path $env:APPDATA "Greensoft\DrawioPpt\settings.xml"
$settingsBackupPath = Join-Path $env:TEMP ("DrawioPpt\settings-backup-selection-" + [Guid]::NewGuid().ToString("N") + ".xml")
$settingsExistedAtStart = $false
$settingsBackupCompleted = $false
$testWordProcessIdsBefore = @()

function Assert-NoRunningWord {
    $runningWord = Get-Process -Name WINWORD -ErrorAction SilentlyContinue
    if ($runningWord) {
        throw "请先关闭正在运行的 Word，再执行 Word 选区事件 E2E 测试。"
    }
}

function Backup-Settings {
    $script:settingsExistedAtStart = Test-Path $settingsPath
    if ($script:settingsExistedAtStart) {
        $directory = Split-Path -Parent $settingsBackupPath
        if (-not (Test-Path $directory)) {
            New-Item -ItemType Directory -Force -Path $directory | Out-Null
        }

        Copy-Item $settingsPath $settingsBackupPath -Force
        $script:settingsBackupCompleted = $true
    }
}

function Restore-Settings {
    if ($script:settingsBackupCompleted -and (Test-Path $settingsBackupPath)) {
        $settingsDirectory = Split-Path -Parent $settingsPath
        if (-not (Test-Path $settingsDirectory)) {
            New-Item -ItemType Directory -Force -Path $settingsDirectory | Out-Null
        }

        Copy-Item $settingsBackupPath $settingsPath -Force
        Remove-Item $settingsBackupPath -Force
        return
    }

    if (-not $script:settingsExistedAtStart -and (Test-Path $settingsPath)) {
        Remove-Item $settingsPath -Force
    }
}

function Get-WordAutomationProcesses {
    return @(Get-CimInstance Win32_Process -Filter "Name = 'WINWORD.EXE'" | Where-Object {
        -not [string]::IsNullOrWhiteSpace($_.CommandLine) -and $_.CommandLine -match "(?i)/Automation"
    })
}

function Stop-TestWordProcesses {
    $testProcesses = @(Get-WordAutomationProcesses | Where-Object {
        $script:testWordProcessIdsBefore -notcontains $_.ProcessId
    })

    foreach ($process in $testProcesses) {
        Stop-Process -Id $process.ProcessId -Force -ErrorAction SilentlyContinue
    }

    if ($testProcesses.Count -gt 0) {
        Start-Sleep -Milliseconds 500
    }

    $remainingTestProcesses = @(Get-WordAutomationProcesses | Where-Object {
        $script:testWordProcessIdsBefore -notcontains $_.ProcessId
    })
    if ($remainingTestProcesses.Count -gt 0) {
        throw "测试启动的 Word 自动化进程未能退出：PID=$($remainingTestProcesses.ProcessId -join ',')"
    }
}

function Remove-TestTempRoot {
    if (-not (Test-Path $tempRoot)) {
        return
    }

    if (-not $tempRoot.StartsWith($tempRootPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "拒绝清理不属于 Word 选区事件 E2E 的临时目录：$tempRoot"
    }

    Remove-Item -LiteralPath $tempRoot -Recurse -Force
}

if (-not $SkipSourceCheck -and (Test-Path $selectionSyncSource)) {
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $selectionSyncTest
    if ($LASTEXITCODE -ne 0) {
        exit $LASTEXITCODE
    }
}
elseif (-not $SkipSourceCheck) {
    Write-Host "SelectionSourceContractSkipped=True (installed package has no source tree; runtime reflection remains enabled)"
}

if (-not (Test-Path $cscPath)) {
    throw "csc.exe not found: $cscPath"
}

if (-not $SkipBuild) {
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $buildScript -Configuration $Configuration -Platform x64
    if ($LASTEXITCODE -ne 0) {
        exit $LASTEXITCODE
    }
}

if ([string]::IsNullOrWhiteSpace($wordBin)) {
    $wordBin = Join-Path $repoRoot ("src\DrawioPpt.WordAddIn\bin\x64\" + $Configuration)
}

if (Test-Path $wordBin) {
    $wordBin = (Resolve-Path $wordBin).Path
}

if (-not (Test-Path $wordBin)) {
    throw "Word add-in assembly root not found: $wordBin"
}

Assert-NoRunningWord
$testWordProcessIdsBefore = @(Get-WordAutomationProcesses | ForEach-Object { $_.ProcessId })
New-Item -ItemType Directory -Force -Path $tempRoot | Out-Null
try {
Backup-Settings

@"
using System;
using System.IO;
using System.Reflection;
using DrawioPpt.Core.Models;
using DrawioPpt.Core.Services;
using DrawioPpt.WordAddIn.Services;
using DrawioPpt.WordAddIn.Word;
using WordInterop = Microsoft.Office.Interop.Word;

public static class WordSelectionEventE2E
{
    private static bool HasNoPassiveSelectionPath()
    {
        FieldInfo[] fields = typeof(AddInHost).GetFields(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
        foreach (FieldInfo field in fields)
        {
            if (field.FieldType == typeof(System.Windows.Forms.Timer) ||
                string.Equals(field.Name, "_selectionStateTimer", StringComparison.Ordinal))
            {
                return false;
            }
        }

        MethodInfo timerTick = typeof(AddInHost).GetMethod(
            "OnSelectionStateTimerTick",
            BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
        MethodInfo selectionChange = typeof(SelectionMonitor).GetMethod(
            "OnWindowSelectionChange",
            BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
        EventInfo selectionChanged = typeof(SelectionMonitor).GetEvent(
            "SelectionChanged",
            BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
        return timerTick == null && selectionChange == null && selectionChanged == null;
    }

    private static void SaveSettings()
    {
        PluginSettings settings = new PluginSettings();
        settings.AutoOpenOnSelection = false;
        settings.AutoUpdateOnSave = false;
        settings.KeepSidecarFile = false;
        settings.ShowDiagramInfoDialog = false;
        new FilePluginSettingsStore().Save(settings);
    }

    private static WordInterop.InlineShape AddPicture(WordInterop.Document document, string imagePath, int position)
    {
        object linkToFile = false;
        object saveWithDocument = true;
        object range = document.Range(position, position);
        return document.InlineShapes.AddPicture(imagePath, ref linkToFile, ref saveWithDocument, ref range);
    }

    private static void SaveAndCloseDocument(ref WordInterop.Document document)
    {
        if (document == null)
        {
            return;
        }

        object saveChanges = WordInterop.WdSaveOptions.wdDoNotSaveChanges;
        object originalFormat = Type.Missing;
        object routeDocument = Type.Missing;
        ((WordInterop._Document)document).Close(ref saveChanges, ref originalFormat, ref routeDocument);
        document = null;
    }

    [STAThread]
    public static int Main(string[] args)
    {
        WordInterop.Application application = null;
        WordInterop.Document document = null;
        AddInHost host = null;

        try
        {
            if (args == null || args.Length != 2 || string.IsNullOrWhiteSpace(args[0]) || string.IsNullOrWhiteSpace(args[1]))
            {
                throw new ArgumentException("Expected temporary image and document paths.");
            }

            SaveSettings();
            File.WriteAllBytes(args[0], Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wl0el0AAAAASUVORK5CYII="));

            bool noPassiveSelectionPath = HasNoPassiveSelectionPath();
            Console.WriteLine("NoSelectionPolling=" + noPassiveSelectionPath);
            Console.WriteLine("NoPassiveSelectionMetadataPath=" + noPassiveSelectionPath);
            if (!noPassiveSelectionPath)
            {
                return 1;
            }

            Type wordType = Type.GetTypeFromProgID("Word.Application", true);
            application = (WordInterop.Application)Activator.CreateInstance(wordType);
            application.Visible = true;
            document = application.Documents.Add();

            document.SaveAs2(args[1]);
            WordInterop.InlineShape managedInlineShape = AddPicture(document, args[0], 0);
            WordInterop.InlineShape plainInlineShape = AddPicture(document, args[0], document.Content.End - 1);
            WordInterop.Shape managedShape = managedInlineShape.ConvertToShape();
            WordInterop.Shape plainShape = plainInlineShape.ConvertToShape();

            DiagramEnvelope envelope = new DiagramEnvelope();
            envelope.DiagramId = Guid.NewGuid().ToString("N");
            envelope.DiagramName = "Managed selection test";
            envelope.DrawioXml = "<mxfile><diagram id='selection-test'/></mxfile>";
            new WordPictureMetadataService(new DiagramEnvelopeSerializer(), new PresentationSidecarPathBuilder()).Save(
                WordPictureReference.FromShape(managedShape),
                envelope);

            host = new AddInHost(application);
            host.Start();

            managedShape.Select();
            SelectionContext managedSelection = new WordPictureSelectionReader().Read(application.Selection);
            bool managedSelectionDetected = managedSelection.HasSinglePicture && managedSelection.IsManagedPicture;
            Console.WriteLine("ManagedSelectionDetected=" + managedSelectionDetected);
            Console.WriteLine("ExplicitManagedSelectionDetected=" + managedSelectionDetected);

            plainShape.Select();
            SelectionContext plainSelection = new WordPictureSelectionReader().Read(application.Selection);
            bool plainPictureCanBind = plainSelection.HasSinglePicture && !plainSelection.IsManagedPicture;
            Console.WriteLine("PlainPictureCanBind=" + plainPictureCanBind);
            Console.WriteLine("ExplicitPlainPictureCanBind=" + plainPictureCanBind);

            return managedSelectionDetected && plainPictureCanBind ? 0 : 1;
        }
        catch (Exception ex)
        {
            Console.WriteLine("CaughtType=" + ex.GetType().FullName);
            Console.WriteLine("CaughtMessage=" + ex.Message);
            Console.WriteLine("CaughtStack=" + ex.StackTrace);
            return 99;
        }
        finally
        {
            if (host != null)
            {
                host.Dispose();
            }

            if (document != null)
            {
                try
                {
                    SaveAndCloseDocument(ref document);
                }
                catch
                {
                }
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
                catch
                {
                }
            }
        }
    }

}
"@ | Set-Content -Path $programPath -Encoding UTF8

Copy-Item (Join-Path $wordBin "DrawioPpt.Core.dll") $tempRoot -Force
Copy-Item (Join-Path $wordBin "DrawioPpt.PowerPointAddIn.dll") $tempRoot -Force
Copy-Item (Join-Path $wordBin "DrawioPpt.WordAddIn.dll") $tempRoot -Force

$compileArguments = @(
    "/nologo",
    "/t:exe",
    "/platform:x64",
    "/out:$exePath",
    "/r:$([IO.Path]::Combine($tempRoot, 'DrawioPpt.Core.dll'))",
    "/r:$([IO.Path]::Combine($tempRoot, 'DrawioPpt.PowerPointAddIn.dll'))",
    "/r:$([IO.Path]::Combine($tempRoot, 'DrawioPpt.WordAddIn.dll'))",
    "/r:$interopWord",
    "/r:$officeCore",
    "/r:System.dll",
    "/r:System.Core.dll",
    "/r:System.Windows.Forms.dll",
    "/r:System.Xml.dll",
    "/r:System.Xml.Linq.dll",
    $programPath
)

& $cscPath @compileArguments
if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}

Push-Location $tempRoot
try {
    $runOutput = & $exePath $imagePath $documentPath 2>&1
    $exitCode = $LASTEXITCODE
    if ($runOutput) {
        $runOutput | ForEach-Object { Write-Host $_ }
    }

    exit [int]$exitCode
}
finally {
    Pop-Location
}
}
finally {
    try {
        Stop-TestWordProcesses
    }
    finally {
        try {
            Restore-Settings
        }
        finally {
            Remove-TestTempRoot
        }
    }
}
