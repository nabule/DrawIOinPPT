param(
    [switch]$SkipBuild
)

$ErrorActionPreference = "Stop"

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$buildScript = Join-Path $PSScriptRoot "build.ps1"
$tempRoot = Join-Path $env:TEMP ("DrawioPpt\\word-url-e2e-" + [Guid]::NewGuid().ToString("N"))
$serverPort = Get-Random -Minimum 8800 -Maximum 8999
$mockHtmlPath = Join-Path $tempRoot "mock-editor.html"
$programPath = Join-Path $tempRoot "word-url-e2e.cs"
$exePath = Join-Path $tempRoot "word-url-e2e.exe"
$cscPath = "C:\\Windows\\Microsoft.NET\\Framework64\\v4.0.30319\\csc.exe"
$interopWord = "C:\\Windows\\assembly\\GAC_MSIL\\Microsoft.Office.Interop.Word\\15.0.0.0__71e9bce111e9429c\\Microsoft.Office.Interop.Word.dll"
$officeCore = "C:\\Windows\\assembly\\GAC_MSIL\\office\\15.0.0.0__71e9bce111e9429c\\OFFICE.DLL"
$wordBin = Join-Path $repoRoot "src\\DrawioPpt.WordAddIn\\bin\\x64\\Debug"
$settingsPath = Join-Path $env:APPDATA "Greensoft\\DrawioPpt\\settings.xml"
$settingsBackupPath = Join-Path $env:TEMP ("DrawioPpt\\settings-backup-word-" + [Guid]::NewGuid().ToString("N") + ".xml")

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

function Assert-NoRunningWord {
    $runningWord = Get-Process -Name WINWORD -ErrorAction SilentlyContinue
    if ($runningWord) {
        throw "请先关闭正在运行的 Word，再执行 Word E2E 测试。"
    }
}

function Start-MockHttpServer {
    param(
        [int]$Port,
        [string]$HtmlPath
    )

    $job = Start-Job -ArgumentList $Port, $HtmlPath -ScriptBlock {
        param(
            [int]$Port,
            [string]$HtmlPath
        )

        $listener = New-Object System.Net.HttpListener
        $listener.Prefixes.Add("http://127.0.0.1:$Port/")
        $listener.Prefixes.Add("http://localhost:$Port/")
        $listener.Start()

        try {
            while ($true) {
                $context = $listener.GetContext()
                try {
                    $path = $context.Request.Url.AbsolutePath
                    if ($path -eq "/" -or $path -eq "/mock-editor.html") {
                        $bytes = [System.IO.File]::ReadAllBytes($HtmlPath)
                        $context.Response.StatusCode = 200
                        $context.Response.ContentType = "text/html; charset=utf-8"
                        $context.Response.ContentLength64 = $bytes.Length
                        $context.Response.OutputStream.Write($bytes, 0, $bytes.Length)
                    }
                    else {
                        $context.Response.StatusCode = 404
                    }
                }
                finally {
                    $context.Response.Close()
                }
            }
        }
        finally {
            $listener.Stop()
            $listener.Close()
        }
    }

    Start-Sleep -Seconds 2
    return $job
}

if (-not (Test-Path $cscPath)) {
    throw "csc.exe not found: $cscPath"
}

if (-not $SkipBuild) {
    & powershell.exe -ExecutionPolicy Bypass -File $buildScript -Configuration Debug -Platform x64
    if ($LASTEXITCODE -ne 0) {
        exit $LASTEXITCODE
    }
}

if (-not (Test-Path $wordBin)) {
    throw "Word add-in output not found: $wordBin"
}

New-Item -ItemType Directory -Force -Path $tempRoot | Out-Null
Assert-NoRunningWord
Backup-Settings

@'
<!DOCTYPE html>
<html>
<head><meta charset="utf-8" /></head>
<body>
<script>
(function () {
  var waitingForConfigure = window.location.search.indexOf('configure=1') >= 0;
  var currentLabel = 'Round1';
  function send(message) { window.parent.postMessage(message, '*'); }
  function buildXml(label) {
    return '<mxfile host="DrawioWord"><diagram id="word-url-e2e" name="' + label + '"><mxGraphModel dx="1000" dy="1000" grid="1" gridSize="10" guides="1" tooltips="1" connect="1" arrows="1" fold="1" page="1" pageScale="1" pageWidth="850" pageHeight="1100" math="0" shadow="0"><root><mxCell id="0"/><mxCell id="1" parent="0"/><mxCell id="2" value="' + label + '" style="rounded=1;whiteSpace=wrap;html=1;" vertex="1" parent="1"><mxGeometry x="120" y="120" width="220" height="60" as="geometry"/></mxCell></root></mxGraphModel></diagram></mxfile>';
  }
  window.addEventListener('message', function (event) {
    var data = event.data || {};
    if (data.action === 'configure') {
      waitingForConfigure = false;
      setTimeout(function () { send({ event: 'init' }); }, 50);
      return;
    }
    if (data.action === 'load') {
      var xml = data.xml || '';
      currentLabel = xml.indexOf('Round1') >= 0 ? 'Round2' : 'Round1';
      setTimeout(function () { send({ event: 'save', xml: buildXml(currentLabel), exit: 1 }); }, 120);
      return;
    }
    if (data.action === 'export') {
      var svg = '<svg xmlns="http://www.w3.org/2000/svg" content="' + currentLabel + '"><rect x="10" y="10" width="220" height="80" fill="#dff1ff" stroke="#1f5fa8"/><text x="30" y="58">' + currentLabel + '</text></svg>';
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
using System.Threading;
using System.Windows.Forms;
using DrawioPpt.Core.Models;
using DrawioPpt.Core.Services;
using DrawioPpt.WordAddIn.Services;
using WordInterop = Microsoft.Office.Interop.Word;

public static class WordUrlE2E
{
    private static DiagramEnvelope ReadEnvelope(WordInterop.InlineShape shape, DiagramEnvelopeSerializer serializer)
    {
        if (shape == null || serializer == null)
        {
            return null;
        }

        string alternativeText = shape.AlternativeText ?? string.Empty;
        if (!serializer.CanDeserialize(alternativeText))
        {
            return null;
        }

        return serializer.Deserialize(alternativeText);
    }

    private static bool HasDiagramState(WordInterop.Document document, DiagramEnvelopeSerializer serializer, string expectedText)
    {
        if (document == null || document.InlineShapes.Count < 1)
        {
            return false;
        }

        DiagramEnvelope envelope = ReadEnvelope(document.InlineShapes[1], serializer);
        return envelope != null &&
               !string.IsNullOrWhiteSpace(envelope.DiagramId) &&
               !string.IsNullOrWhiteSpace(envelope.DrawioXml) &&
               envelope.DrawioXml.IndexOf(expectedText, StringComparison.OrdinalIgnoreCase) >= 0 &&
               document.CustomXMLParts.Count > 0;
    }

    private static void SaveAndCloseDocument(ref WordInterop.Document document)
    {
        if (document == null)
        {
            return;
        }

        object saveChanges = WordInterop.WdSaveOptions.wdSaveChanges;
        document.Save();
        document.Close(ref saveChanges);
        document = null;
    }

    private static void DisposeHost(ref AddInHost host)
    {
        if (host == null)
        {
            return;
        }

        host.Dispose();
        host = null;
    }

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
        settings.SidecarFolderName = "drawio-word-e2e";
        settings.ShowDiagramInfoDialog = false;
        new FilePluginSettingsStore().Save(settings);

        WordInterop.Application application = null;
        AddInHost host = null;
        WordInterop.Document document = null;
        DiagramEnvelopeSerializer serializer = new DiagramEnvelopeSerializer();

        try
        {
            Type wordType = Type.GetTypeFromProgID("Word.Application", true);
            application = (WordInterop.Application)Activator.CreateInstance(wordType);
            application.Visible = true;

            document = application.Documents.Add();
            string documentPath = Path.Combine(Path.GetTempPath(), "DrawioPpt", "word-url-e2e-" + Guid.NewGuid().ToString("N") + ".docx");
            string documentDirectory = Path.GetDirectoryName(documentPath);
            if (!Directory.Exists(documentDirectory))
            {
                Directory.CreateDirectory(documentDirectory);
            }

            document.SaveAs2(documentPath);
            application.Selection.HomeKey(WordInterop.WdUnits.wdStory);

            host = new AddInHost(application);
            host.Start();

            host.CreateNewDiagram();
            Thread.Sleep(1800);

            bool createdManagedPicture = HasDiagramState(document, serializer, "Round1");
            Console.WriteLine("CreatedManagedPicture=" + createdManagedPicture);
            Console.WriteLine("InlineShapeCount=" + document.InlineShapes.Count);
            Console.WriteLine("CreatedCustomXmlCount=" + document.CustomXMLParts.Count);

            SaveAndCloseDocument(ref document);
            DisposeHost(ref host);
            Thread.Sleep(1200);

            document = application.Documents.Open(documentPath);
            host = new AddInHost(application);
            host.Start();

            document.InlineShapes[1].Select();
            Thread.Sleep(800);
            host.EditSelectedDiagram();
            Thread.Sleep(1800);

            bool reopenedEditApplied = HasDiagramState(document, serializer, "Round2");
            Console.WriteLine("ReopenedEditApplied=" + reopenedEditApplied);
            Console.WriteLine("EditedCustomXmlCount=" + document.CustomXMLParts.Count);

            SaveAndCloseDocument(ref document);
            DisposeHost(ref host);
            Thread.Sleep(1200);

            document = application.Documents.Open(documentPath);
            bool persistedAfterReopen = HasDiagramState(document, serializer, "Round2");
            Console.WriteLine("PersistedAfterReopen=" + persistedAfterReopen);
            Console.WriteLine("FinalCustomXmlCount=" + document.CustomXMLParts.Count);
            Console.WriteLine("FinalInlineShapeCount=" + document.InlineShapes.Count);

            return createdManagedPicture && reopenedEditApplied && persistedAfterReopen ? 0 : 1;
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

            if (host != null)
            {
                DisposeHost(ref host);
            }

            if (application != null)
            {
                try
                {
                    object saveChanges = WordInterop.WdSaveOptions.wdDoNotSaveChanges;
                    application.Quit(ref saveChanges);
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
Copy-Item (Join-Path $wordBin "Microsoft.Web.WebView2.Core.dll") $tempRoot -Force
Copy-Item (Join-Path $wordBin "Microsoft.Web.WebView2.WinForms.dll") $tempRoot -Force
Copy-Item (Join-Path $wordBin "runtimes\\win-x64\\native\\WebView2Loader.dll") $tempRoot -Force

$compileArguments = @(
    "/nologo",
    "/t:exe",
    "/platform:x64",
    "/out:$exePath",
    "/r:$([IO.Path]::Combine($tempRoot, 'DrawioPpt.Core.dll'))",
    "/r:$([IO.Path]::Combine($tempRoot, 'DrawioPpt.PowerPointAddIn.dll'))",
    "/r:$([IO.Path]::Combine($tempRoot, 'DrawioPpt.WordAddIn.dll'))",
    "/r:$([IO.Path]::Combine($tempRoot, 'Microsoft.Web.WebView2.Core.dll'))",
    "/r:$([IO.Path]::Combine($tempRoot, 'Microsoft.Web.WebView2.WinForms.dll'))",
    "/r:$interopWord",
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
    Restore-Settings
    exit $LASTEXITCODE
}

$server = Start-MockHttpServer -Port $serverPort -HtmlPath $mockHtmlPath
try {
    Push-Location $tempRoot
    try {
        $runOutput = & $exePath 2>&1
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
    if ($server) {
        Stop-Job -Job $server -ErrorAction SilentlyContinue | Out-Null
        Remove-Job -Job $server -Force -ErrorAction SilentlyContinue | Out-Null
    }

    Restore-Settings
}
