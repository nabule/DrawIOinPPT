param(
    [string]$Version = "v1.0.8",
    [switch]$SkipBuild,
    [switch]$KeepInstalled
)

$ErrorActionPreference = "Stop"

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$buildScript = Join-Path $PSScriptRoot "build.ps1"
$packageScript = Join-Path $PSScriptRoot "package-release.ps1"
$registerRepoScript = Join-Path $PSScriptRoot "register-office-addins.ps1"
$releaseRoot = Join-Path $repoRoot ("artifacts\\releases\\" + $Version)
$packageRoot = Join-Path $releaseRoot "package"
$installRoot = Join-Path $env:TEMP ("DrawioPpt\\installed-" + [Guid]::NewGuid().ToString("N"))
$reportRoot = Join-Path $repoRoot "artifacts\\test-reports"
$reportPath = Join-Path $reportRoot ("full-e2e-" + $Version + ".md")
$logRoot = Join-Path $repoRoot ("artifacts\\logs\\" + $Version)
$settingsPath = Join-Path $env:APPDATA "Greensoft\\DrawioPpt\\settings.xml"
$settingsBackupPath = Join-Path $env:TEMP ("DrawioPpt\\settings-backup-" + [Guid]::NewGuid().ToString("N") + ".xml")
$pluginLogPath = Join-Path $env:APPDATA "Greensoft\\DrawioPpt\\Logs\\drawioppt.log"
$pluginLogSnapshotPath = Join-Path $logRoot "drawioppt-full-e2e.log"
$transcriptPath = Join-Path $logRoot ("full-e2e-" + $Version + ".transcript.log")
$complexWordIdentityPath = Join-Path $env:TEMP ("DrawioPpt\\word-complex-identity-" + [Guid]::NewGuid().ToString("N") + ".txt")
$results = New-Object System.Collections.Generic.List[string]
$cleanupErrors = New-Object System.Collections.Generic.List[string]
$drawioExe = "C:\\Program Files\\draw.io\\draw.io.exe"
$fatalError = $null
$transcriptStarted = $false

function Assert-NoRunningPowerPoint {
    $runningPowerPoint = Get-Process -Name POWERPNT -ErrorAction SilentlyContinue
    if ($runningPowerPoint) {
        throw "请先关闭正在运行的 PowerPoint，再执行完整 E2E 测试。"
    }
}

function Assert-NoRunningWord {
    $runningWord = Get-Process -Name WINWORD -ErrorAction SilentlyContinue
    if ($runningWord) {
        throw "请先关闭正在运行的 Word，再执行完整 E2E 测试。"
    }
}

function Get-AvailableLoopbackPort {
    $listener = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Loopback, 0)
    try {
        $listener.Start()
        return ([System.Net.IPEndPoint]$listener.LocalEndpoint).Port
    }
    finally {
        $listener.Stop()
    }
}

function Stop-RunningPowerPointSilently {
    Get-Process -Name POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
}

function Stop-TestOwnedProcess {
    param(
        [Parameter(Mandatory = $true)]
        [string]$IdentityPath,
        [Parameter(Mandatory = $true)]
        [string]$ExpectedProcessName,
        [int]$TimeoutSeconds = 10
    )

    if (-not (Test-Path -LiteralPath $IdentityPath)) {
        return
    }

    $identityLines = @(Get-Content -LiteralPath $IdentityPath -ErrorAction Stop)
    $matchedIdentity = $false
    foreach ($line in $identityLines) {
        if ($line -notmatch '^TestWordProcessIdentity=(\d+):(\d+)$') {
            continue
        }

        $matchedIdentity = $true
        $processId = [int]$Matches[1]
        $expectedStartTicks = [Int64]$Matches[2]
        $process = Get-Process -Id $processId -ErrorAction SilentlyContinue
        if ($null -eq $process) {
            Write-Host "TestOwnedProcessAlreadyExited=$processId"
            continue
        }

        $actualStartTicks = $process.StartTime.ToUniversalTime().Ticks
        $processNameMatches = [string]::Equals(
            $process.ProcessName,
            $ExpectedProcessName,
            [System.StringComparison]::OrdinalIgnoreCase)
        if (-not $processNameMatches -or $actualStartTicks -ne $expectedStartTicks) {
            Write-Host "TestOwnedProcessIdentityMismatchProtected=$processId"
            continue
        }

        Stop-Process -Id $processId -Force -ErrorAction Stop
        $deadline = [DateTime]::UtcNow.AddSeconds($TimeoutSeconds)
        do {
            if ($null -eq (Get-Process -Id $processId -ErrorAction SilentlyContinue)) {
                Write-Host "TestOwnedProcessStopped=$processId"
                break
            }

            if ([DateTime]::UtcNow -ge $deadline) {
                throw "Timed out waiting for the test-owned $ExpectedProcessName process $processId to exit."
            }

            Start-Sleep -Milliseconds 100
        } while ($true)
    }

    if (-not $matchedIdentity) {
        throw "No valid test-owned process identity was found in: $IdentityPath"
    }
}

function Invoke-CheckedPowerShellScript {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ScriptPath,
        [string[]]$Arguments = @(),
        [Parameter(Mandatory = $true)]
        [string]$Description
    )

    if (-not (Test-Path -LiteralPath $ScriptPath)) {
        throw "$Description script not found: $ScriptPath"
    }

    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $ScriptPath @Arguments > $null
    $exitCode = [int]$LASTEXITCODE
    if ($exitCode -ne 0) {
        throw "$Description failed with exit code $exitCode."
    }
}

function Get-ErrorDetail {
    param([object]$ErrorValue)

    $detail = if ($ErrorValue -is [System.Management.Automation.ErrorRecord]) {
        $ErrorValue.Exception.Message
    }
    elseif ($ErrorValue -is [System.Exception]) {
        $ErrorValue.Message
    }
    else {
        [string]$ErrorValue
    }

    if ([string]::IsNullOrWhiteSpace($detail)) {
        $detail = [string]$ErrorValue
    }

    return $detail.Replace("|", "/").Replace("`r", " ").Replace("`n", " ").Trim()
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
    $preferredServerPort = Get-AvailableLoopbackPort
    $mockHtmlPath = Join-Path $tempRoot "mock-editor.html"
    $programPath = Join-Path $tempRoot "powerpoint-url-e2e.cs"
    $exePath = Join-Path $tempRoot "powerpoint-url-e2e.exe"
    $cscPath = "C:\\Windows\\Microsoft.NET\\Framework64\\v4.0.30319\\csc.exe"
    $interopPowerPoint = "C:\\Windows\\assembly\\GAC_MSIL\\Microsoft.Office.Interop.PowerPoint\\15.0.0.0__71e9bce111e9429c\\Microsoft.Office.Interop.PowerPoint.dll"
    $officeCore = "C:\\Windows\\assembly\\GAC_MSIL\\office\\15.0.0.0__71e9bce111e9429c\\OFFICE.DLL"

    if (-not (Test-Path $cscPath)) {
        throw "csc.exe not found: $cscPath"
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
  var currentLabel = 'Round1';
  function send(message) { window.parent.postMessage(message, '*'); }
  function buildXml(label) {
    return '<mxfile host="DrawioPpt"><diagram id="powerpoint-url-e2e" name="' + label + '"><mxGraphModel dx="1000" dy="1000" grid="1" gridSize="10" guides="1" tooltips="1" connect="1" arrows="1" fold="1" page="1" pageScale="1" pageWidth="850" pageHeight="1100" math="0" shadow="0"><root><mxCell id="0"/><mxCell id="1" parent="0"/><mxCell id="2" value="' + label + '" style="rounded=1;whiteSpace=wrap;html=1;" vertex="1" parent="1"><mxGeometry x="120" y="120" width="220" height="60" as="geometry"/></mxCell></root></mxGraphModel></diagram></mxfile>';
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
using System.Net;
using System.Threading;
using System.Windows.Forms;
using DrawioPpt.Core.Models;
using DrawioPpt.Core.Services;
using DrawioPpt.PowerPointAddIn.Services;
using Microsoft.Office.Core;
using PptInterop = Microsoft.Office.Interop.PowerPoint;

public sealed class MockEditorServer : IDisposable
{
    private readonly string _htmlPath;
    private HttpListener _listener;
    private Thread _worker;

    public MockEditorServer(string htmlPath)
    {
        _htmlPath = htmlPath;
    }

    public int Port { get; private set; }

    public void StartWithRetry(int preferredPort)
    {
        int candidatePort = preferredPort;
        Exception lastError = null;
        for (int attempt = 0; attempt < 10; attempt++)
        {
            HttpListener listener = new HttpListener();
            listener.Prefixes.Add("http://127.0.0.1:" + candidatePort + "/");
            listener.Prefixes.Add("http://localhost:" + candidatePort + "/");
            try
            {
                listener.Start();
                _listener = listener;
                Port = candidatePort;
                _worker = new Thread(Listen);
                _worker.IsBackground = true;
                _worker.Start();
                Console.WriteLine("MockHttpServerState=Running");
                Console.WriteLine("MockServerPort=" + Port);
                return;
            }
            catch (HttpListenerException ex)
            {
                lastError = ex;
                listener.Close();
                candidatePort = GetAvailableLoopbackPort();
            }
        }

        throw new InvalidOperationException("Unable to bind the mock editor server after retrying available loopback ports.", lastError);
    }

    private static int GetAvailableLoopbackPort()
    {
        System.Net.Sockets.TcpListener listener = new System.Net.Sockets.TcpListener(IPAddress.Loopback, 0);
        try
        {
            listener.Start();
            return ((System.Net.IPEndPoint)listener.LocalEndpoint).Port;
        }
        finally
        {
            listener.Stop();
        }
    }

    public void AssertReachable()
    {
        string requestUrl = "http://127.0.0.1:" + Port + "/mock-editor.html?configure=1";
        using (WebClient client = new WebClient())
        {
            string html = client.DownloadString(requestUrl);
            if (html.IndexOf("window.parent.postMessage", StringComparison.Ordinal) < 0)
            {
                throw new InvalidOperationException("Mock HTTP server returned unexpected editor content.");
            }
        }

        Console.WriteLine("MockHttpReachable=True");
    }

    private void Listen()
    {
        while (_listener.IsListening)
        {
            HttpListenerContext context;
            try
            {
                context = _listener.GetContext();
            }
            catch (HttpListenerException)
            {
                return;
            }
            catch (ObjectDisposedException)
            {
                return;
            }

            try
            {
                string path = context.Request.Url.AbsolutePath;
                Console.WriteLine("MOCK_HTTP=" + context.Request.HttpMethod + " " + context.Request.Url.PathAndQuery);
                if (string.Equals(path, "/", StringComparison.Ordinal) || string.Equals(path, "/mock-editor.html", StringComparison.Ordinal))
                {
                    byte[] bytes = File.ReadAllBytes(_htmlPath);
                    context.Response.StatusCode = 200;
                    context.Response.ContentType = "text/html; charset=utf-8";
                    context.Response.ContentLength64 = bytes.Length;
                    context.Response.OutputStream.Write(bytes, 0, bytes.Length);
                }
                else
                {
                    context.Response.StatusCode = 404;
                }
            }
            finally
            {
                context.Response.Close();
            }
        }
    }

    public void Dispose()
    {
        if (_listener != null)
        {
            _listener.Close();
        }
        if (_worker != null)
        {
            _worker.Join(5000);
        }
    }
}

public static class PowerPointUrlE2E
{
    private static PptInterop.Shape GetLastShape(PptInterop.Presentation presentation)
    {
        PptInterop.Slide slide = presentation.Slides[1];
        return slide.Shapes[slide.Shapes.Count];
    }

    private static DiagramEnvelope ReadEnvelope(PptInterop.Shape shape, DiagramEnvelopeSerializer serializer)
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

    private static bool HasDiagramState(PptInterop.Shape shape, DiagramEnvelopeSerializer serializer, string expectedText)
    {
        DiagramEnvelope envelope = ReadEnvelope(shape, serializer);
        if (envelope == null)
        {
            return false;
        }

        string partId = shape.Tags["DRAWIO_PPT_PART_ID"] ?? string.Empty;
        return !string.IsNullOrWhiteSpace(shape.Tags["DRAWIO_PPT_ID"]) &&
               !string.IsNullOrWhiteSpace(partId) &&
               !string.IsNullOrWhiteSpace(envelope.DrawioXml) &&
               envelope.DrawioXml.IndexOf(expectedText, StringComparison.OrdinalIgnoreCase) >= 0;
    }

    private static void SaveAndClosePresentation(ref PptInterop.Presentation presentation)
    {
        if (presentation == null)
        {
            return;
        }

        presentation.Save();
        presentation.Saved = MsoTriState.msoTrue;
        presentation.Close();
        presentation = null;
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

        PptInterop.Application application = null;
        AddInHost host = null;
        PptInterop.Presentation presentation = null;
        MockEditorServer mockServer = null;
        DiagramEnvelopeSerializer serializer = new DiagramEnvelopeSerializer();

        try
        {
            mockServer = new MockEditorServer(@"$mockHtmlPath");
            mockServer.StartWithRetry($preferredServerPort);
            mockServer.AssertReachable();

            PluginSettings settings = new PluginSettings();
            settings.EditorMode = EditorMode.Url;
            settings.EditorUrl = "http://127.0.0.1:" + mockServer.Port + "/mock-editor.html";
            settings.UseOfficeCompatibleSvgLabels = true;
            settings.AutoOpenOnSelection = false;
            settings.AutoUpdateOnSave = true;
            settings.KeepSidecarFile = true;
            settings.SidecarFolderName = "drawio-e2e-reopen";
            settings.ShowDiagramInfoDialog = false;
            new FilePluginSettingsStore().Save(settings);

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
            Thread.Sleep(1800);

            PptInterop.Shape shape = GetLastShape(presentation);
            bool createdManagedShape = HasDiagramState(shape, serializer, "Round1");
            string diagramId = shape.Tags["DRAWIO_PPT_ID"] ?? string.Empty;
            Console.WriteLine("CreatedManagedShape=" + createdManagedShape);
            Console.WriteLine("CreatedDiagramId=" + diagramId);
            Console.WriteLine("CreatedCustomXmlCount=" + presentation.CustomXMLParts.Count);

            SaveAndClosePresentation(ref presentation);
            DisposeHost(ref host);
            Thread.Sleep(1200);

            presentation = application.Presentations.Open(presentationPath, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoTrue);
            presentation.Slides[1].Select();
            host = new AddInHost(application);
            host.Start();

            shape = GetLastShape(presentation);
            shape.Select(MsoTriState.msoFalse);
            Thread.Sleep(800);
            host.EditSelectedDiagram();
            Thread.Sleep(1800);

            shape = GetLastShape(presentation);
            bool reopenedEditApplied = HasDiagramState(shape, serializer, "Round2");
            Console.WriteLine("ReopenedEditApplied=" + reopenedEditApplied);
            Console.WriteLine("EditedCustomXmlCount=" + presentation.CustomXMLParts.Count);

            SaveAndClosePresentation(ref presentation);
            DisposeHost(ref host);
            Thread.Sleep(1200);

            presentation = application.Presentations.Open(presentationPath, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoTrue);
            shape = GetLastShape(presentation);
            bool persistedAfterReopen = HasDiagramState(shape, serializer, "Round2") && presentation.CustomXMLParts.Count > 0;
            PptInterop.Slide slide = presentation.Slides[1];
            Console.WriteLine("PersistedAfterReopen=" + persistedAfterReopen);
            Console.WriteLine("FinalCustomXmlCount=" + presentation.CustomXMLParts.Count);
            Console.WriteLine("ShapeCount=" + slide.Shapes.Count);

            return createdManagedShape && reopenedEditApplied && persistedAfterReopen ? 0 : 1;
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
            if (mockServer != null)
            {
                mockServer.Dispose();
            }

            if (presentation != null)
            {
                try
                {
                    SaveAndClosePresentation(ref presentation);
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
    if (-not (Test-Path $logRoot)) {
        New-Item -ItemType Directory -Force -Path $logRoot | Out-Null
    }

    if (Test-Path $pluginLogPath) {
        Remove-Item $pluginLogPath -Force -ErrorAction SilentlyContinue
    }

    try {
        Start-Transcript -Path $transcriptPath -Force | Out-Null
        $transcriptStarted = $true
    }
    catch {
    }

    Assert-NoRunningPowerPoint
    Assert-NoRunningWord

    if (-not $SkipBuild) {
        & powershell.exe -ExecutionPolicy Bypass -File $buildScript -Configuration Release -Platform x64
        if ($LASTEXITCODE -ne 0) {
            exit $LASTEXITCODE
        }
    }

    & powershell.exe -ExecutionPolicy Bypass -File $packageScript -Version $Version -Configuration Release -Platform x64 -SkipBuild
    if ($LASTEXITCODE -ne 0) {
        exit $LASTEXITCODE
    }
    Add-Result -Name "ReleasePackage" -Passed (Test-Path $packageRoot) -Detail $packageRoot

    $installScript = Join-Path $packageRoot "scripts\\install-release.ps1"
    & powershell.exe -ExecutionPolicy Bypass -File $installScript -InstallRoot $installRoot
    if ($LASTEXITCODE -ne 0) {
        exit $LASTEXITCODE
    }
    Add-Result -Name "InstallRelease" -Passed (Test-Path (Join-Path $installRoot "bin\\DrawioPpt.PowerPointAddIn.dll")) -Detail $installRoot

    $verifyInstallScript = Join-Path $installRoot "scripts\\verify-office-install.ps1"
    & powershell.exe -ExecutionPolicy Bypass -File $verifyInstallScript -InstallRoot $installRoot
    Add-Result -Name "InstalledOfficeAddInsVerify" -Passed ($LASTEXITCODE -eq 0) -Detail $verifyInstallScript

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
    & powershell.exe -ExecutionPolicy Bypass -File $installedUrlSmoke -SkipBuild
    Add-Result -Name "InstalledUrlSmoke" -Passed ($LASTEXITCODE -eq 0) -Detail $installedUrlSmoke

    $installedWordUrlHostE2E = Join-Path $installRoot "scripts\\word-url-addin-host-e2e.ps1"
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $installedWordUrlHostE2E -Configuration Release -SkipBuild -AssemblyRoot (Join-Path $installRoot "bin")
    Add-Result -Name "InstalledWordUrlHostE2E" -Passed ($LASTEXITCODE -eq 0) -Detail $installedWordUrlHostE2E

    $installedWordComplexMetadataE2E = Join-Path $installRoot "scripts\\word-complex-metadata-e2e.ps1"
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $installedWordComplexMetadataE2E -Configuration Release -AssemblyRoot (Join-Path $installRoot "bin") -ProcessIdentityPath $complexWordIdentityPath -SkipBuild
    Add-Result -Name "InstalledWordComplexMetadataE2E" -Passed ($LASTEXITCODE -eq 0) -Detail $installedWordComplexMetadataE2E

    $urlE2EExitCode = Compile-And-Run-PowerPointUrlE2E -InstalledRoot $installRoot
    Add-Result -Name "PowerPointUrlE2E" -Passed ($urlE2EExitCode -eq 0) -Detail "ExitCode=$urlE2EExitCode"

    $desktopExporterPassed = Test-DesktopExporter -ExecutablePath $drawioExe
    Add-Result -Name "DesktopExporter" -Passed $desktopExporterPassed -Detail $drawioExe
}
catch {
    $fatalError = $_
}
finally {
    Stop-RunningPowerPointSilently

    try {
        Stop-TestOwnedProcess `
            -IdentityPath $complexWordIdentityPath `
            -ExpectedProcessName "WINWORD" `
            -TimeoutSeconds 10
    }
    catch {
        $cleanupErrors.Add("OwnedWordCleanup: $(Get-ErrorDetail $_)") | Out-Null
    }

    if (-not $KeepInstalled -and (Test-Path $installRoot)) {
        $uninstallScript = Join-Path $installRoot "scripts\\uninstall-release.ps1"
        try {
            Invoke-CheckedPowerShellScript `
                -ScriptPath $uninstallScript `
                -Arguments @("-InstallRoot", $installRoot) `
                -Description "UninstallRelease"
        }
        catch {
            $cleanupErrors.Add("UninstallRelease: $(Get-ErrorDetail $_)") | Out-Null
        }
    }

    try {
        Invoke-CheckedPowerShellScript `
            -ScriptPath $registerRepoScript `
            -Arguments @("-Configuration", "Release") `
            -Description "RegisterRepositoryAddIns"
    }
    catch {
        $cleanupErrors.Add("RegisterRepositoryAddIns: $(Get-ErrorDetail $_)") | Out-Null
    }

    try {
        Restore-Settings
    }
    catch {
        $cleanupErrors.Add("RestoreSettings: $(Get-ErrorDetail $_)") | Out-Null
    }

    try {
        if (Test-Path -LiteralPath $complexWordIdentityPath) {
            Remove-Item -LiteralPath $complexWordIdentityPath -Force
        }
    }
    catch {
        $cleanupErrors.Add("RemoveWordIdentity: $(Get-ErrorDetail $_)") | Out-Null
    }
}

$failureDetails = New-Object System.Collections.Generic.List[string]
if ($fatalError -ne $null) {
    $failureDetails.Add((Get-ErrorDetail $fatalError)) | Out-Null
}

foreach ($cleanupError in $cleanupErrors) {
    $failureDetails.Add($cleanupError) | Out-Null
}

if ($failureDetails.Count -gt 0) {
    $results.Add("| FatalError | FAIL | $($failureDetails -join '; ') |") | Out-Null
}
else {
    Write-Host "FullE2ECleanupSucceeded=True"
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

Copy-Item $reportPath (Join-Path $logRoot ("full-e2e-" + $Version + ".report.md")) -Force

if (Test-Path $pluginLogPath) {
    Copy-Item $pluginLogPath $pluginLogSnapshotPath -Force
}

if ($transcriptStarted) {
    try {
        Stop-Transcript | Out-Null
    }
    catch {
    }
}

if ($failureDetails.Count -gt 0) {
    Write-Error ($failureDetails -join "; ")
    exit 1
}

Write-Host "E2E report written to: $reportPath"
