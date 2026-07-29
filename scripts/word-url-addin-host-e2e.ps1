param(
    [string]$Configuration = "Release",
    [string]$AssemblyRoot,
    [switch]$SkipBuild,
    [string]$DiagnosticLogPath
)

$ErrorActionPreference = "Stop"

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$buildScript = Join-Path $PSScriptRoot "build.ps1"
$registerScript = Join-Path $PSScriptRoot "register-word-addin.ps1"
$unregisterScript = Join-Path $PSScriptRoot "unregister-word-addin.ps1"
$tempRoot = Join-Path $env:TEMP ("DrawioPpt\word-url-addin-host-e2e-" + [Guid]::NewGuid().ToString("N"))
$tempRootPrefix = Join-Path $env:TEMP "DrawioPpt\word-url-addin-host-e2e-"
$mockHtmlPath = Join-Path $tempRoot "mock-editor.html"
$programPath = Join-Path $tempRoot "word-url-addin-host-e2e.cs"
$exePath = Join-Path $tempRoot "word-url-addin-host-e2e.exe"
$documentPath = Join-Path $tempRoot "actual-word-host.docx"
$processIdentityPath = Join-Path $tempRoot "word-process-identities.txt"
$serverPort = $null
$cscPath = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
$interopWord = "C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Word\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.Word.dll"
$officeCore = "C:\Windows\assembly\GAC_MSIL\office\15.0.0.0__71e9bce111e9429c\OFFICE.DLL"
$settingsPath = Join-Path $env:APPDATA "Greensoft\DrawioPpt\settings.xml"
$settingsBackupPath = Join-Path $tempRoot "settings-backup.xml"
$pluginLogPath = Join-Path $env:APPDATA "Greensoft\DrawioPpt\Logs\drawioppt.log"
$wordAddInProgId = "Greensoft.DrawioWordAddIn"
$wordAddInRegistryPaths = @(
    "HKCU\Software\Microsoft\Office\Word\Addins\Greensoft.DrawioWordAddIn",
    "HKCU\Software\Classes\Greensoft.DrawioWordAddIn",
    "HKCU\Software\Classes\CLSID\{F10C5C83-0D86-4C81-A0B8-7E8FE9D31D8D}"
)
$settingsExistedAtStart = $false
$registryBackups = @()
$testWordProcessIdentities = @()

function Write-Diagnostic {
    param([string]$Message)

    if ([string]::IsNullOrWhiteSpace($DiagnosticLogPath)) {
        return
    }

    $diagnosticDirectory = Split-Path -Parent $DiagnosticLogPath
    if (-not [string]::IsNullOrWhiteSpace($diagnosticDirectory) -and -not (Test-Path -LiteralPath $diagnosticDirectory)) {
        New-Item -ItemType Directory -Force -Path $diagnosticDirectory | Out-Null
    }

    Add-Content -LiteralPath $DiagnosticLogPath -Value ("{0:O} {1}" -f [DateTime]::UtcNow, $Message) -Encoding UTF8
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

function Add-TestWordProcessIdentities {
    param([object[]]$Output)

    foreach ($line in @($Output)) {
        if ($line -match '^TestWordProcessIdentity=(\d+):(\d+)$') {
            $identity = [PSCustomObject]@{
                ProcessId = [int]$Matches[1]
                StartTicks = [Int64]$Matches[2]
            }

            if (-not ($script:testWordProcessIdentities | Where-Object {
                $_.ProcessId -eq $identity.ProcessId -and $_.StartTicks -eq $identity.StartTicks
            })) {
                $script:testWordProcessIdentities += $identity
            }
        }
    }
}

function Read-TestWordProcessIdentities {
    if (Test-Path -LiteralPath $processIdentityPath) {
        Add-TestWordProcessIdentities -Output (Get-Content -LiteralPath $processIdentityPath)
    }
}

function Wait-TestWordProcessExit {
    param(
        [Parameter(Mandatory = $true)]
        [int]$ProcessId,
        [Parameter(Mandatory = $true)]
        [Int64]$StartTicks,
        [int]$TimeoutSeconds = 10
    )

    $deadline = [DateTime]::UtcNow.AddSeconds([Math]::Max(0, $TimeoutSeconds))
    do {
        $process = Get-Process -Id $ProcessId -ErrorAction SilentlyContinue
        if ($null -eq $process) {
            return $true
        }

        if ($process.ProcessName -ne "WINWORD" -or
            $process.StartTime.ToUniversalTime().Ticks -ne $StartTicks) {
            return $true
        }

        if ([DateTime]::UtcNow -ge $deadline) {
            return $false
        }

        Start-Sleep -Milliseconds 100
    } while ($true)
}

function Assert-NoRunningWord {
    $runningWord = Get-Process -Name WINWORD -ErrorAction SilentlyContinue
    if ($runningWord) {
        throw "请先关闭正在运行的 Word，再执行真实 Word URL 加载项 E2E 测试。"
    }
}

function Backup-Settings {
    $script:settingsExistedAtStart = Test-Path -LiteralPath $settingsPath
    if ($script:settingsExistedAtStart) {
        Copy-Item -LiteralPath $settingsPath -Destination $settingsBackupPath -Force
    }
}

function Restore-Settings {
    if ($script:settingsExistedAtStart -and (Test-Path -LiteralPath $settingsBackupPath)) {
        $settingsDirectory = Split-Path -Parent $settingsPath
        if (-not (Test-Path -LiteralPath $settingsDirectory)) {
            New-Item -ItemType Directory -Force -Path $settingsDirectory | Out-Null
        }

        Copy-Item -LiteralPath $settingsBackupPath -Destination $settingsPath -Force
        return
    }

    if (-not $script:settingsExistedAtStart -and (Test-Path -LiteralPath $settingsPath)) {
        Remove-Item -LiteralPath $settingsPath -Force
    }
}

function Backup-WordAddInRegistration {
    foreach ($registryPath in $wordAddInRegistryPaths) {
        $backupPath = Join-Path $tempRoot ("registry-" + $registryBackups.Count + ".reg")
        $process = Start-Process -FilePath reg.exe -ArgumentList @("export", $registryPath, $backupPath, "/y") -Wait -PassThru -WindowStyle Hidden
        $script:registryBackups += [PSCustomObject]@{
            RegistryPath = $registryPath
            BackupPath = $backupPath
            Existed = ($process.ExitCode -eq 0)
        }
    }
}

function Restore-WordAddInRegistration {
    foreach ($registryPath in $wordAddInRegistryPaths) {
        Start-Process -FilePath reg.exe -ArgumentList @("delete", $registryPath, "/f") -Wait -WindowStyle Hidden
    }

    foreach ($backup in $script:registryBackups) {
        if ($backup.Existed -and (Test-Path -LiteralPath $backup.BackupPath)) {
            $process = Start-Process -FilePath reg.exe -ArgumentList @("import", $backup.BackupPath) -Wait -PassThru -WindowStyle Hidden
            if ($process.ExitCode -ne 0) {
                throw "无法恢复 Word 加载项注册：$($backup.RegistryPath)"
            }
        }
    }
}

function Stop-TestWordProcesses {
    Read-TestWordProcessIdentities

    foreach ($identity in $script:testWordProcessIdentities) {
        $process = Get-Process -Id $identity.ProcessId -ErrorAction SilentlyContinue
        if (-not $process) {
            continue
        }

        $actualStartTicks = $process.StartTime.ToUniversalTime().Ticks
        if ($process.ProcessName -ne "WINWORD" -or
            $actualStartTicks -ne $identity.StartTicks) {
            throw "测试 Word 进程身份已变化，拒绝终止 PID $($identity.ProcessId)。"
        }

        if (Wait-TestWordProcessExit -ProcessId $identity.ProcessId -StartTicks $identity.StartTicks -TimeoutSeconds 20) {
            Write-Host "WordProcessNaturalExitObserved=$($identity.ProcessId)"
            continue
        }

        Stop-Process -Id $identity.ProcessId -Force -ErrorAction Stop
        if (-not (Wait-TestWordProcessExit -ProcessId $identity.ProcessId -StartTicks $identity.StartTicks -TimeoutSeconds 20)) {
            throw "无法终止测试拥有的 WINWORD 进程：$($identity.ProcessId)"
        }

        Write-Host "TestWordProcessStopped=$($identity.ProcessId)"
    }
}

function Remove-TestTempRoot {
    if (Test-Path -LiteralPath $tempRoot) {
        $resolvedRoot = (Resolve-Path -LiteralPath $tempRoot).Path
        $expectedParent = [IO.Path]::GetFullPath(
            (Join-Path $env:TEMP "DrawioPpt"))
        $leafName = Split-Path -Leaf $resolvedRoot
        if (-not [string]::Equals(
                (Split-Path -Parent $resolvedRoot),
                $expectedParent,
                [System.StringComparison]::OrdinalIgnoreCase) -or
            -not $leafName.StartsWith(
                "word-url-addin-host-e2e-",
                [System.StringComparison]::OrdinalIgnoreCase)) {
            throw "拒绝清理不属于真实 Word URL 加载项 E2E 的临时目录：$resolvedRoot"
        }

        Remove-Item -LiteralPath $resolvedRoot -Recurse -Force
    }
}

function Get-UpdatedDiagramXml {
    param(
        [Parameter(Mandatory = $true)]$InlineShape,
        [Parameter(Mandatory = $true)][string]$CoreAssemblyPath
    )

    [System.Reflection.Assembly]::LoadFrom($CoreAssemblyPath) | Out-Null
    $serializerType = [Type]::GetType("DrawioPpt.Core.Services.DiagramEnvelopeSerializer, DrawioPpt.Core", $true)
    $serializer = [Activator]::CreateInstance($serializerType)
    $envelope = $serializer.Deserialize([string]$InlineShape.AlternativeText)
    return [string]$envelope.DrawioXml
}

if (-not $SkipBuild) {
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $buildScript -Configuration $Configuration -Platform x64
    if ($LASTEXITCODE -ne 0) {
        exit $LASTEXITCODE
    }
}

if (-not (Test-Path -LiteralPath $cscPath)) {
    throw "csc.exe not found: $cscPath"
}

if ([string]::IsNullOrWhiteSpace($AssemblyRoot)) {
    $AssemblyRoot = Join-Path $repoRoot ("src\DrawioPpt.WordAddIn\bin\x64\" + $Configuration)
}

$AssemblyRoot = (Resolve-Path -LiteralPath $AssemblyRoot).Path
$wordAssemblyPath = Join-Path $AssemblyRoot "DrawioPpt.WordAddIn.dll"
$coreAssemblyPath = Join-Path $AssemblyRoot "DrawioPpt.Core.dll"
if (-not (Test-Path -LiteralPath $wordAssemblyPath) -or -not (Test-Path -LiteralPath $coreAssemblyPath)) {
    throw "Expected Word add-in output was not found under: $AssemblyRoot"
}

New-Item -ItemType Directory -Force -Path $tempRoot | Out-Null
Assert-NoRunningWord
$serverPort = Get-AvailableLoopbackPort

@'
<!DOCTYPE html>
<html>
<head><meta charset="utf-8" /></head>
<body>
<script>
(function () {
  function send(message) { window.parent.postMessage(message, '*'); }
  function buildXml() {
    return '<mxfile host="DrawioWord"><diagram id="actual-host" name="Round2"><mxGraphModel><root><mxCell id="0"/><mxCell id="1" parent="0"/><mxCell id="2" value="Round2" vertex="1" parent="1"><mxGeometry x="100" y="100" width="220" height="60" as="geometry"/></mxCell></root></mxGraphModel></diagram></mxfile>';
  }
  window.addEventListener('message', function (event) {
    var data = event.data || {};
    if (data.action === 'configure') {
      setTimeout(function () { send({ event: 'init' }); }, 50);
      return;
    }
    if (data.action === 'load') {
      setTimeout(function () { send({ event: 'save', xml: buildXml(), exit: 1 }); }, 100);
      return;
    }
    if (data.action === 'export') {
      var svg = '<svg xmlns="http://www.w3.org/2000/svg"><rect x="10" y="10" width="220" height="80" fill="#dff1ff" stroke="#1f5fa8"/><text x="30" y="58">Round2</text></svg>';
      setTimeout(function () { send({ event: 'export', data: 'data:image/svg+xml;utf8,' + encodeURIComponent(svg) }); }, 100);
    }
  });
  setTimeout(function () { send({ event: 'configure' }); }, 120);
})();
</script>
</body>
</html>
'@ | Set-Content -LiteralPath $mockHtmlPath -Encoding UTF8

@"
using System;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using DrawioPpt.Core.Models;
using DrawioPpt.Core.Services;
using DrawioPpt.WordAddIn;
using DrawioPpt.WordAddIn.Services;
using DrawioPpt.WordAddIn.Word;
using Microsoft.Office.Core;
using WordInterop = Microsoft.Office.Interop.Word;

public static class WordUrlAddInHostE2E
{
    private static class NativeMethods
    {
        [DllImport("user32.dll", SetLastError = true)]
        internal static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
    }

    private static void WriteWordProcessIdentity(WordInterop.Application application, string processIdentityPath)
    {
        uint processId;
        NativeMethods.GetWindowThreadProcessId(new IntPtr(application.ActiveWindow.Hwnd), out processId);
        if (processId == 0)
        {
            throw new InvalidOperationException("Unable to resolve the Word process identity.");
        }

        Process process = Process.GetProcessById((int)processId);
        string identity = "TestWordProcessIdentity=" + processId + ":" + process.StartTime.ToUniversalTime().Ticks;
        File.AppendAllText(processIdentityPath, identity + Environment.NewLine, Encoding.UTF8);
        Console.WriteLine(identity);
    }

    private static DiagramEnvelope ReadStoredEnvelope(
        WordInterop.Document document,
        WordPictureReference picture,
        DiagramEnvelopeSerializer serializer)
    {
        if (document == null || picture == null || serializer == null)
        {
            return null;
        }

        string alternativeText = picture.AlternativeText ?? string.Empty;
        if (!serializer.CanDeserialize(alternativeText))
        {
            return null;
        }

        DiagramEnvelope reference = serializer.Deserialize(alternativeText);
        if (reference == null)
        {
            return null;
        }

        DiagramEnvelope storedEnvelope;
        DocumentDiagramStore store = new DocumentDiagramStore(serializer);
        return store.TryRead(document, reference.DiagramId, out storedEnvelope)
            ? storedEnvelope
            : null;
    }

    private static WordPictureReference GetSinglePicture(
        WordInterop.Document document)
    {
        if (document == null)
        {
            return null;
        }

        if (document.Shapes.Count == 1)
        {
            return WordPictureReference.FromShape(document.Shapes[1]);
        }

        if (document.InlineShapes.Count == 1)
        {
            return WordPictureReference.FromInlineShape(
                document.InlineShapes[1]);
        }

        return null;
    }

    private sealed class MockEditorServer : IDisposable
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
            using (WebClient client = new WebClient())
            {
                string html = client.DownloadString("http://127.0.0.1:" + Port + "/mock-editor.html");
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
                    if (string.Equals(context.Request.Url.AbsolutePath, "/mock-editor.html", StringComparison.OrdinalIgnoreCase))
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

    private static void CloseDocument(ref WordInterop.Document document)
    {
        if (document == null)
        {
            return;
        }

        object saveChanges = WordInterop.WdSaveOptions.wdDoNotSaveChanges;
        ((WordInterop._Document)document).Close(ref saveChanges);
        document = null;
    }

    private static void QuitApplication(ref WordInterop.Application application)
    {
        if (application == null)
        {
            return;
        }

        object saveChanges = WordInterop.WdSaveOptions.wdDoNotSaveChanges;
        ((WordInterop._Application)application).Quit(ref saveChanges);
        application = null;
    }

    private static void WriteUrlTestSettings(string settingsPath, int serverPort)
    {
        string settingsDirectory = Path.GetDirectoryName(settingsPath);
        if (!Directory.Exists(settingsDirectory))
        {
            Directory.CreateDirectory(settingsDirectory);
        }

        string settingsXml = "<pluginSettings>\r\n"
            + "  <editorMode>Url</editorMode>\r\n"
            + "  <desktopEditorPath></desktopEditorPath>\r\n"
            + "  <editorUrl>http://127.0.0.1:" + serverPort + "/mock-editor.html</editorUrl>\r\n"
            + "  <useOfficeCompatibleSvgLabels>true</useOfficeCompatibleSvgLabels>\r\n"
            + "  <autoOpenOnSelection>true</autoOpenOnSelection>\r\n"
            + "  <autoUpdateOnSave>true</autoUpdateOnSave>\r\n"
            + "  <keepSidecarFile>false</keepSidecarFile>\r\n"
            + "  <sidecarFolderName>drawio-word-addin-host-e2e</sidecarFolderName>\r\n"
            + "  <showDiagramInfoDialog>false</showDiagramInfoDialog>\r\n"
            + "</pluginSettings>\r\n";
        File.WriteAllText(settingsPath, settingsXml, new UTF8Encoding(false));
    }

    private static int PrepareDocument(string documentPath, string processIdentityPath)
    {
        WordInterop.Application application = null;
        WordInterop.Document document = null;
        try
        {
            string directory = Path.GetDirectoryName(documentPath);
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            string imagePath = Path.Combine(directory, "managed-picture.png");
            File.WriteAllBytes(imagePath, Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wl0el0AAAAASUVORK5CYII="));

            Type wordType = Type.GetTypeFromProgID("Word.Application", true);
            application = (WordInterop.Application)Activator.CreateInstance(wordType);
            application.Visible = false;
            application.DisplayAlerts = WordInterop.WdAlertLevel.wdAlertsNone;
            document = application.Documents.Add();
            WriteWordProcessIdentity(application, processIdentityPath);

            object linkToFile = false;
            object saveWithDocument = true;
            object range = document.Range(0, 0);
            WordInterop.InlineShape picture = document.InlineShapes.AddPicture(imagePath, ref linkToFile, ref saveWithDocument, ref range);

            DiagramEnvelope envelope = new DiagramEnvelope();
            envelope.DiagramId = Guid.NewGuid().ToString("N");
            envelope.DiagramName = "Actual Word Host E2E";
            envelope.EditorMode = EditorMode.Url;
            envelope.EditorTarget = "http://127.0.0.1";
            envelope.DrawioXml = "<mxfile host='DrawioWord'><diagram id='round1' name='Round1'/></mxfile>";
            envelope.UpdatedUtc = DateTime.UtcNow;
            picture.Title = envelope.DiagramName;
            DiagramEnvelopeSerializer serializer = new DiagramEnvelopeSerializer();
            picture.AlternativeText = serializer.Serialize(envelope);
            bool mainStoreMissingBeforeMigration = ReadStoredEnvelope(
                document,
                WordPictureReference.FromInlineShape(picture),
                serializer) == null;
            Console.WriteLine("PreparedDocumentMainStoreMissing=" + mainStoreMissingBeforeMigration);
            document.SaveAs2(documentPath);
            Console.WriteLine("PreparedManagedDocument=True");
            return mainStoreMissingBeforeMigration ? 0 : 1;
        }
        catch (Exception ex)
        {
            Console.WriteLine("PrepareCaughtType=" + ex.GetType().FullName);
            Console.WriteLine("PrepareCaughtMessage=" + ex.Message);
            Console.WriteLine("PrepareCaughtStack=" + ex.StackTrace);
            return 99;
        }
        finally
        {
            if (document != null)
            {
                try
                {
                    CloseDocument(ref document);
                }
                catch
                {
                }
            }

            if (application != null)
            {
                try
                {
                    QuitApplication(ref application);
                }
                catch
                {
                }
            }
        }
    }

    private static int RunActualHost(string documentPath, int serverPort, string mockHtmlPath, string settingsPath, string processIdentityPath)
    {
        WordInterop.Application application = null;
        WordInterop.Document document = null;
        MockEditorServer mockServer = null;
        try
        {
            mockServer = new MockEditorServer(mockHtmlPath);
            mockServer.StartWithRetry(serverPort);
            WriteUrlTestSettings(settingsPath, mockServer.Port);
            mockServer.AssertReachable();

            Type wordType = Type.GetTypeFromProgID("Word.Application", true);
            application = (WordInterop.Application)Activator.CreateInstance(wordType);
            application.Visible = true;
            application.DisplayAlerts = WordInterop.WdAlertLevel.wdAlertsNone;
            document = application.Documents.Open(documentPath);
            WriteWordProcessIdentity(application, processIdentityPath);
            Thread.Sleep(800);

            COMAddIn addIn = application.COMAddIns.Item("Greensoft.DrawioWordAddIn");
            addIn.Connect = true;
            bool connected = addIn.Connect;
            Console.WriteLine("WordAddInConnect=" + connected);
            Console.WriteLine("ActualHostProcess=WINWORD");
            WordPictureReference selectedPicture = GetSinglePicture(document);
            if (!connected || selectedPicture == null)
            {
                return 1;
            }

            dynamic automation = addIn.Object;
            bool explicitEditAutomationAvailable = automation != null &&
                string.Equals(
                    (string)automation.GetInteractionMode(),
                    "ExplicitSelectionCommand",
                    StringComparison.Ordinal);
            Console.WriteLine("ExplicitEditAutomationAvailable=" + explicitEditAutomationAvailable);
            if (!explicitEditAutomationAvailable)
            {
                return 1;
            }

            selectedPicture.Select();
            automation.EditSelectedDiagram();
            Console.WriteLine("ExplicitEditCommandInvoked=True");
            Thread.Sleep(500);
            bool saved = false;
            WordPictureReference savedPicture = GetSinglePicture(document);
            if (savedPicture != null)
            {
                DiagramEnvelopeSerializer serializer = new DiagramEnvelopeSerializer();
                DiagramEnvelope storedEnvelope = ReadStoredEnvelope(
                    document,
                    savedPicture,
                    serializer);
                saved = storedEnvelope != null &&
                        !string.IsNullOrWhiteSpace(storedEnvelope.DrawioXml) &&
                        storedEnvelope.DrawioXml.IndexOf("Round2", StringComparison.OrdinalIgnoreCase) >= 0;
            }

            Console.WriteLine("ActualWordUrlEditorSaved=" + saved);
            return saved ? 0 : 1;
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
            if (document != null)
            {
                try
                {
                    CloseDocument(ref document);
                }
                catch
                {
                }
            }

            if (application != null)
            {
                try
                {
                    QuitApplication(ref application);
                }
                catch
                {
                }
            }

            if (mockServer != null)
            {
                mockServer.Dispose();
            }
        }
    }

    [STAThread]
    public static int Main(string[] args)
    {
        if (args == null || args.Length < 3)
        {
            Console.WriteLine("Expected arguments: prepare <documentPath> <processIdentityPath> or run <documentPath> <port> <mockHtmlPath> <settingsPath> <processIdentityPath>.");
            return 99;
        }

        if (string.Equals(args[0], "prepare", StringComparison.OrdinalIgnoreCase) && args.Length == 3)
        {
            return PrepareDocument(args[1], args[2]);
        }

        if (string.Equals(args[0], "run", StringComparison.OrdinalIgnoreCase) && args.Length == 6)
        {
            int serverPort;
            if (!int.TryParse(args[2], out serverPort))
            {
                Console.WriteLine("Invalid mock server port: " + args[2]);
                return 99;
            }

            return RunActualHost(args[1], serverPort, args[3], args[4], args[5]);
        }

        Console.WriteLine("Unknown mode: " + args[0]);
        return 99;
    }
}
"@ | Set-Content -LiteralPath $programPath -Encoding UTF8

Copy-Item -LiteralPath $coreAssemblyPath -Destination $tempRoot -Force
Copy-Item -LiteralPath $wordAssemblyPath -Destination $tempRoot -Force
$compileArguments = @(
    "/nologo",
    "/t:exe",
    "/platform:x64",
    "/out:$exePath",
    "/r:$([IO.Path]::Combine($tempRoot, 'DrawioPpt.Core.dll'))",
    "/r:$([IO.Path]::Combine($tempRoot, 'DrawioPpt.WordAddIn.dll'))",
    "/r:$interopWord",
    "/r:$officeCore",
    "/r:System.dll",
    "/r:System.Core.dll",
    "/r:Microsoft.CSharp.dll",
    $programPath
)
& $cscPath @compileArguments
if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}

Backup-Settings
Backup-WordAddInRegistration
$initialLogLength = if (Test-Path -LiteralPath $pluginLogPath) { (Get-Item -LiteralPath $pluginLogPath).Length } else { 0 }
$testExitCode = 1
try {
    Write-Diagnostic "Temporarily unregistering the pre-existing Word add-in before document preparation."
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $unregisterScript -Configuration $Configuration
    if ($LASTEXITCODE -ne 0) {
        throw "Unable to temporarily unregister the pre-existing Word add-in."
    }

    Write-Diagnostic "Preparing a managed Word document without an active add-in."
    Push-Location $tempRoot
    try {
        $prepareOutput = & $exePath prepare $documentPath $processIdentityPath 2>&1
        $prepareExitCode = $LASTEXITCODE
        Write-Diagnostic "Document preparation exit code: $prepareExitCode"
        if ($prepareOutput) {
            $prepareOutput | ForEach-Object {
                Write-Host $_
                Write-Diagnostic ("Preparation output: " + $_.ToString())
            }
            Add-TestWordProcessIdentities -Output $prepareOutput
        }
    }
    finally {
        Pop-Location
    }

    if ($prepareExitCode -ne 0) {
        throw "Unable to prepare the managed Word document for the actual-host test."
    }

    $prepareWordExitDeadline = [DateTime]::UtcNow.AddSeconds(10)
    do {
        $prepareWordProcesses = @(Get-Process WINWORD -ErrorAction SilentlyContinue)
        if ($prepareWordProcesses.Count -eq 0) {
            break
        }

        Start-Sleep -Milliseconds 100
    } while ([DateTime]::UtcNow -lt $prepareWordExitDeadline)
    if ($prepareWordProcesses.Count -gt 0) {
        throw "Document-preparation Word process did not exit before the actual-host phase."
    }

    Start-Sleep -Milliseconds 1500

    Write-Diagnostic "Registering Word add-in from $wordAssemblyPath."
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $registerScript -Configuration $Configuration -AssemblyPath $wordAssemblyPath
    if ($LASTEXITCODE -ne 0) {
        throw "Unable to register the Word add-in under test."
    }

    Write-Diagnostic "Starting actual Word-host test executable."
    Push-Location $tempRoot
    try {
        $runOutput = & $exePath run $documentPath $serverPort $mockHtmlPath $settingsPath $processIdentityPath 2>&1
        $runnerExitCode = $LASTEXITCODE
        Write-Diagnostic "Word-host executable exit code: $runnerExitCode"
        if ($runOutput) {
            $runOutput | ForEach-Object {
                Write-Host $_
                Write-Diagnostic ("Word-host output: " + $_.ToString())
            }
            Add-TestWordProcessIdentities -Output $runOutput
        }
    }
    finally {
        Pop-Location
    }

    $actualWordUrlEditorSaved = $runnerExitCode -eq 0

    $webView2UserDataFolder = Join-Path $env:LOCALAPPDATA "Greensoft\DrawioPpt\WebView2"
    $webView2UserDataFolderExists = Test-Path -LiteralPath $webView2UserDataFolder
    $newLog = ""
    if (Test-Path -LiteralPath $pluginLogPath) {
        $logContent = Get-Content -LiteralPath $pluginLogPath -Raw
        if ($logContent.Length -gt $initialLogLength) {
            $newLog = $logContent.Substring($initialLogLength)
        }
    }

    $accessDeniedDialog = $newLog -match "Failed to initialize URL editor|E_ACCESSDENIED|拒绝访问"
    Write-Host "ActualWordUrlEditorSaved=$actualWordUrlEditorSaved"
    Write-Host "WebView2UserDataFolderExists=$webView2UserDataFolderExists"
    Write-Host "AccessDeniedDialog=$accessDeniedDialog"
    Write-Diagnostic "Assertions: ActualWordUrlEditorSaved=$actualWordUrlEditorSaved; WebView2UserDataFolderExists=$webView2UserDataFolderExists; AccessDeniedDialog=$accessDeniedDialog"

    $testExitCode = if ($actualWordUrlEditorSaved -and $webView2UserDataFolderExists -and -not $accessDeniedDialog) { 0 } else { 1 }
}
catch {
    Write-Diagnostic ("Unhandled test error: " + $_.Exception.GetType().FullName + ": " + $_.Exception.Message)
    throw
}
finally {
    $cleanupFailures = New-Object System.Collections.Generic.List[string]
    $cleanupActions = @(
        @{ Name = "Stop test Word processes"; Action = { Stop-TestWordProcesses } },
        @{ Name = "Restore Word add-in registration"; Action = { Restore-WordAddInRegistration } },
        @{ Name = "Restore plug-in settings"; Action = { Restore-Settings } },
        @{ Name = "Remove test temporary directory"; Action = { Remove-TestTempRoot } }
    )

    foreach ($cleanupAction in $cleanupActions) {
        try {
            & $cleanupAction.Action
        }
        catch {
            $cleanupFailures.Add("$($cleanupAction.Name): $($_.Exception.Message)") | Out-Null
        }
    }

    if ($cleanupFailures.Count -gt 0) {
        $cleanupSummary = "Cleanup failure: " + ($cleanupFailures -join "; ")
        Write-Diagnostic $cleanupSummary
        Write-Error $cleanupSummary
        $testExitCode = 1
    }
}

exit $testExitCode
