param(
    [switch]$SkipBuild
)

$ErrorActionPreference = "Stop"

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$buildScript = Join-Path $PSScriptRoot "build.ps1"
$tempRoot = Join-Path $env:TEMP ("DrawioPpt\\url-editor-smoke-" + [Guid]::NewGuid().ToString("N"))
$mockHtmlPath = Join-Path $tempRoot "mock-editor.html"
$smokeSourcePath = Join-Path $tempRoot "url-editor-smoke.cs"
$smokeExePath = Join-Path $tempRoot "url-editor-smoke.exe"
$serverPort = Get-Random -Minimum 8600 -Maximum 8999
$powerPointReleaseBin = Join-Path $repoRoot "src\\DrawioPpt.PowerPointAddIn\\bin\\x64\\Release"
$powerPointDebugBin = Join-Path $repoRoot "src\\DrawioPpt.PowerPointAddIn\\bin\\x64\\Debug"
$coreReleaseBin = Join-Path $repoRoot "src\\DrawioPpt.Core\\bin\\x64\\Release"
$coreDebugBin = Join-Path $repoRoot "src\\DrawioPpt.Core\\bin\\x64\\Debug"
$packageBin = Join-Path $repoRoot "bin"
$webView2Package = Join-Path $repoRoot "packages\\Microsoft.Web.WebView2.1.0.3856.49"
$cscPath = "C:\\Windows\\Microsoft.NET\\Framework64\\v4.0.30319\\csc.exe"

function Get-SmokeInputPath {
    param(
        [string[]]$Candidates,
        [string]$Label
    )

    foreach ($candidate in $Candidates) {
        if (-not [string]::IsNullOrWhiteSpace($candidate) -and (Test-Path $candidate)) {
            return $candidate
        }
    }

    throw "$Label not found. Checked: $($Candidates -join '; ')"
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

if (-not $SkipBuild -and (Test-Path $buildScript) -and ((Test-Path $powerPointReleaseBin) -or (Test-Path $powerPointDebugBin))) {
    & powershell.exe -ExecutionPolicy Bypass -File $buildScript
    if ($LASTEXITCODE -ne 0) {
        exit $LASTEXITCODE
    }
}

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
      setTimeout(function () { send({ event: 'save', xml: window.__xml, exit: 1 }); }, 100);
      return;
    }
    if (data.action === 'export') {
      var svg = '<svg xmlns="http://www.w3.org/2000/svg" content="old"><text x="10" y="20">smoke</text></svg>';
      setTimeout(function () { send({ event: 'export', data: 'data:image/svg+xml;utf8,' + encodeURIComponent(svg) }); }, 100);
    }
  });
  setTimeout(function () {
    if (waitingForConfigure) {
      send({ event: 'configure' });
      return;
    }

    send({ event: 'init' });
  }, 100);
})();
</script>
</body>
</html>
'@ | Set-Content -Path $mockHtmlPath -Encoding UTF8

@"
using System;
using System.Windows.Forms;
using DrawioPpt.PowerPointAddIn.Services;
using DrawioPpt.PowerPointAddIn.UI;

public static class UrlEditorSmoke
{
    [STAThread]
    public static int Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);

        PluginTraceLog traceLog = new PluginTraceLog();
        string editorUrl = "http://127.0.0.1:$serverPort/mock-editor.html";
        string xml = "<mxfile host=\"Smoke\"><diagram id=\"smoke\" name=\"Smoke\"><mxGraphModel /></diagram></mxfile>";
        bool saved = false;
        string savedSvg = string.Empty;
        string savedXml = string.Empty;

        using (UrlDiagramEditorForm form = new UrlDiagramEditorForm(editorUrl, "URL Smoke", xml, true, traceLog))
        {
            form.DiagramSaved += delegate(object sender, UrlDiagramSavedEventArgs args)
            {
                saved = true;
                savedSvg = args.SvgMarkup ?? string.Empty;
                savedXml = args.Xml ?? string.Empty;
            };

            DialogResult result = form.ShowDialog();
            Console.WriteLine("DialogResult=" + result);
        }

        Console.WriteLine("Saved=" + saved);
        Console.WriteLine("SvgHasSmoke=" + savedSvg.Contains("smoke"));
        Console.WriteLine("XmlHasSmokeId=" + savedXml.Contains("id=\"smoke\""));
        Console.WriteLine("LogPath=" + traceLog.LogPath);
        return saved && savedSvg.Contains("smoke") && savedXml.Contains("id=\"smoke\"") ? 0 : 1;
    }
}
"@ | Set-Content -Path $smokeSourcePath -Encoding UTF8

$powerPointDll = Get-SmokeInputPath -Candidates @(
    (Join-Path $powerPointReleaseBin "DrawioPpt.PowerPointAddIn.dll"),
    (Join-Path $powerPointDebugBin "DrawioPpt.PowerPointAddIn.dll"),
    (Join-Path $packageBin "DrawioPpt.PowerPointAddIn.dll")
) -Label "DrawioPpt.PowerPointAddIn.dll"

$coreDll = Get-SmokeInputPath -Candidates @(
    (Join-Path $coreReleaseBin "DrawioPpt.Core.dll"),
    (Join-Path $coreDebugBin "DrawioPpt.Core.dll"),
    (Join-Path $packageBin "DrawioPpt.Core.dll")
) -Label "DrawioPpt.Core.dll"

$webView2CoreDll = Get-SmokeInputPath -Candidates @(
    (Join-Path $webView2Package "lib\\net462\\Microsoft.Web.WebView2.Core.dll"),
    (Join-Path $packageBin "Microsoft.Web.WebView2.Core.dll")
) -Label "Microsoft.Web.WebView2.Core.dll"

$webView2WinFormsDll = Get-SmokeInputPath -Candidates @(
    (Join-Path $webView2Package "lib\\net462\\Microsoft.Web.WebView2.WinForms.dll"),
    (Join-Path $packageBin "Microsoft.Web.WebView2.WinForms.dll")
) -Label "Microsoft.Web.WebView2.WinForms.dll"

$webView2LoaderDll = Get-SmokeInputPath -Candidates @(
    (Join-Path $webView2Package "runtimes\\win-x64\\native\\WebView2Loader.dll"),
    (Join-Path $packageBin "runtimes\\win-x64\\native\\WebView2Loader.dll")
) -Label "WebView2Loader.dll"

Copy-Item $powerPointDll $tempRoot -Force
Copy-Item $coreDll $tempRoot -Force
Copy-Item $webView2CoreDll $tempRoot -Force
Copy-Item $webView2WinFormsDll $tempRoot -Force
Copy-Item $webView2LoaderDll $tempRoot -Force

$compileArguments = @(
    "/nologo",
    "/t:exe",
    "/platform:x64",
    "/out:$smokeExePath",
    "/r:$([IO.Path]::Combine($tempRoot, 'DrawioPpt.Core.dll'))",
    "/r:$([IO.Path]::Combine($tempRoot, 'DrawioPpt.PowerPointAddIn.dll'))",
    "/r:$([IO.Path]::Combine($tempRoot, 'Microsoft.Web.WebView2.Core.dll'))",
    "/r:$([IO.Path]::Combine($tempRoot, 'Microsoft.Web.WebView2.WinForms.dll'))",
    "/r:System.dll",
    "/r:System.Core.dll",
    "/r:System.Windows.Forms.dll",
    "/r:System.Drawing.dll",
    "/r:System.Web.Extensions.dll",
    $smokeSourcePath
)

& $cscPath @compileArguments

$server = Start-MockHttpServer -Port $serverPort -HtmlPath $mockHtmlPath

try {
    Start-Sleep -Seconds 2
    Push-Location $tempRoot
    try {
        & $smokeExePath
        exit $LASTEXITCODE
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
}
