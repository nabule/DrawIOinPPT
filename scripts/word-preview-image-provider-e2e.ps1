param(
    [string]$Configuration = "Release",
    [string]$CoreAssemblyPath = "",
    [string]$DrawioExecutablePath = "C:\Program Files\draw.io\draw.io.exe",
    [switch]$SkipBuild
)

$ErrorActionPreference = "Stop"

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$buildScript = Join-Path $PSScriptRoot "build.ps1"
$exporterSourcePath = Join-Path $repoRoot "src\DrawioPpt.PowerPointAddIn\Services\DesktopSvgExporter.cs"
$providerSourcePath = Join-Path $repoRoot "src\DrawioPpt.WordAddIn\Services\WordPreviewImageProvider.cs"
$tempParent = Join-Path $env:TEMP "DrawioPpt"
$tempRootPrefix = Join-Path $tempParent "word-preview-provider-e2e-"
$tempRoot = $tempRootPrefix + [Guid]::NewGuid().ToString("N")
$programPath = Join-Path $tempRoot "word-preview-provider-e2e.cs"
$exePath = Join-Path $tempRoot "word-preview-provider-e2e.exe"
$drawioPath = Join-Path $tempRoot "source.drawio"
$fallbackSvgPath = Join-Path $tempRoot "fallback.svg"
$cscPath = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"

function Get-DrawioProcesses {
    return @(
        Get-Process -ErrorAction SilentlyContinue |
            Where-Object {
                $_.ProcessName -in @(
                    "draw.io",
                    "drawio",
                    "diagrams.net")
            })
}

function Remove-TestTempRoot {
    if (-not (Test-Path -LiteralPath $tempRoot)) {
        return
    }

    $resolvedRoot = (Resolve-Path -LiteralPath $tempRoot).Path
    if (-not $resolvedRoot.StartsWith(
            $tempRootPrefix,
            [StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing to remove a path outside the Word preview-provider E2E root: $resolvedRoot"
    }

    Remove-Item -LiteralPath $resolvedRoot -Recurse -Force
}

if (-not (Test-Path -LiteralPath $cscPath)) {
    throw "csc.exe not found: $cscPath"
}

if (-not (Test-Path -LiteralPath $DrawioExecutablePath)) {
    throw "draw.io Desktop not found: $DrawioExecutablePath"
}

if (-not (Test-Path -LiteralPath $exporterSourcePath)) {
    throw "DesktopSvgExporter source not found: $exporterSourcePath"
}

if (-not (Test-Path -LiteralPath $providerSourcePath)) {
    throw "WordPreviewImageProvider source not found: $providerSourcePath"
}

if ((Get-DrawioProcesses).Count -gt 0) {
    throw "Close draw.io Desktop before running the Word preview-provider E2E."
}

if (-not $SkipBuild) {
    & powershell.exe `
        -NoProfile `
        -ExecutionPolicy Bypass `
        -File $buildScript `
        -Configuration $Configuration `
        -Platform x64
    if ($LASTEXITCODE -ne 0) {
        exit $LASTEXITCODE
    }
}

if ([string]::IsNullOrWhiteSpace($CoreAssemblyPath)) {
    $CoreAssemblyPath = Join-Path `
        $repoRoot `
        ("src\DrawioPpt.Core\bin\x64\" +
            $Configuration +
            "\DrawioPpt.Core.dll")
}

if (-not (Test-Path -LiteralPath $CoreAssemblyPath)) {
    throw "DrawioPpt.Core.dll not found: $CoreAssemblyPath"
}

New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null

try {
    $programSource = @'
using System;
using System.Drawing;
using System.IO;
using System.Security.Cryptography;
using DrawioPpt.Core.Models;
using DrawioPpt.PowerPointAddIn.Services;
using DrawioPpt.WordAddIn.Services;

public static class WordPreviewImageProviderE2E
{
    private static string GetSha256(string path)
    {
        using (SHA256 sha256 = SHA256.Create())
        using (FileStream stream = File.OpenRead(path))
        {
            return BitConverter.ToString(
                sha256.ComputeHash(stream)).Replace("-", string.Empty);
        }
    }

    private static bool HasPngSignature(string path)
    {
        byte[] signature = new byte[8];
        using (FileStream stream = File.OpenRead(path))
        {
            return stream.Read(signature, 0, signature.Length) ==
                    signature.Length &&
                signature[0] == 0x89 &&
                signature[1] == 0x50 &&
                signature[2] == 0x4E &&
                signature[3] == 0x47 &&
                signature[4] == 0x0D &&
                signature[5] == 0x0A &&
                signature[6] == 0x1A &&
                signature[7] == 0x0A;
        }
    }

    public static int Main(string[] args)
    {
        string drawioExecutablePath = args[0];
        string drawioSourcePath = args[1];
        string fallbackSvgPath = args[2];
        string fallbackHashBefore = GetSha256(fallbackSvgPath);
        string previewPath = null;
        string previewDirectory = null;

        try
        {
            DiagramEnvelope envelope = new DiagramEnvelope();
            envelope.DiagramId = "../unsafe:diagram?*";
            envelope.DiagramName = "2:1 preview";
            envelope.DrawioXml = File.ReadAllText(drawioSourcePath);

            WordPreviewImageProvider provider =
                new WordPreviewImageProvider(
                    new DesktopSvgExporter());
            previewPath = provider.GetPreviewImagePath(
                envelope,
                drawioExecutablePath,
                fallbackSvgPath);
            previewDirectory = Path.GetDirectoryName(previewPath);

            bool returnedPng =
                string.Equals(
                    Path.GetExtension(previewPath),
                    ".png",
                    StringComparison.OrdinalIgnoreCase) &&
                File.Exists(previewPath);
            bool pngSignature = returnedPng &&
                HasPngSignature(previewPath);
            int pixelWidth = 0;
            int pixelHeight = 0;
            if (returnedPng)
            {
                using (Image image = Image.FromFile(previewPath))
                {
                    pixelWidth = image.Width;
                    pixelHeight = image.Height;
                }
            }

            double aspectRatio = pixelHeight > 0
                ? pixelWidth / (double)pixelHeight
                : 0;
            bool widthPassed = pixelWidth == 620;
            bool aspectPassed =
                Math.Abs(aspectRatio - 2.0) <= 0.02;
            bool notForcedSquare =
                Math.Abs(aspectRatio - 1.0) > 0.20;
            string expectedRoot = Path.GetFullPath(
                Path.Combine(
                    Path.GetTempPath(),
                    "DrawioPpt",
                    "word-preview"));
            bool safeOutputPath =
                Path.GetFullPath(previewPath).StartsWith(
                    expectedRoot,
                    StringComparison.OrdinalIgnoreCase) &&
                previewDirectory.IndexOf(
                    "..",
                    StringComparison.Ordinal) < 0;
            bool sourceRemoved =
                Directory.GetFiles(
                    previewDirectory,
                    "*.drawio",
                    SearchOption.TopDirectoryOnly).Length == 0;

            string fallbackResult =
                provider.GetPreviewImagePath(
                    envelope,
                    Path.Combine(
                        Path.GetTempPath(),
                        "missing-drawio.exe"),
                    fallbackSvgPath);
            string fallbackHashAfter =
                GetSha256(fallbackSvgPath);
            bool fallbackPassed =
                string.Equals(
                    fallbackResult,
                    fallbackSvgPath,
                    StringComparison.Ordinal) &&
                string.Equals(
                    fallbackHashBefore,
                    fallbackHashAfter,
                    StringComparison.Ordinal) &&
                File.Exists(fallbackSvgPath);

            Console.WriteLine(
                "PreviewPathUnderSafeRoot=" +
                safeOutputPath);
            Console.WriteLine(
                "PngSignaturePassed=" +
                pngSignature);
            Console.WriteLine(
                "PixelWidth=" +
                pixelWidth);
            Console.WriteLine(
                "PixelHeight=" +
                pixelHeight);
            Console.WriteLine(
                "AspectRatio=" +
                aspectRatio.ToString("0.######"));
            Console.WriteLine(
                "Width620Passed=" +
                widthPassed);
            Console.WriteLine(
                "Aspect2To1Passed=" +
                aspectPassed);
            Console.WriteLine(
                "NotForcedSquare=" +
                notForcedSquare);
            Console.WriteLine(
                "TemporaryDrawioRemoved=" +
                sourceRemoved);
            Console.WriteLine(
                "FallbackSvgPassed=" +
                fallbackPassed);

            return returnedPng &&
                pngSignature &&
                widthPassed &&
                aspectPassed &&
                notForcedSquare &&
                safeOutputPath &&
                sourceRemoved &&
                fallbackPassed
                ? 0
                : 1;
        }
        finally
        {
            string expectedRoot = Path.GetFullPath(
                Path.Combine(
                    Path.GetTempPath(),
                    "DrawioPpt",
                    "word-preview"));
            if (!string.IsNullOrWhiteSpace(previewDirectory) &&
                Directory.Exists(previewDirectory) &&
                Path.GetFullPath(previewDirectory).StartsWith(
                    expectedRoot,
                    StringComparison.OrdinalIgnoreCase))
            {
                Directory.Delete(previewDirectory, true);
            }
        }
    }
}
'@
    [IO.File]::WriteAllText(
        $programPath,
        $programSource,
        (New-Object Text.UTF8Encoding($false)))

    $drawioSource = @'
<?xml version="1.0" encoding="UTF-8"?>
<mxfile host="DrawioPpt-E2E">
  <diagram id="preview-2to1" name="Preview">
    <mxGraphModel dx="1200" dy="600" grid="0" page="0" pageScale="1" pageWidth="1200" pageHeight="600">
      <root>
        <mxCell id="0"/>
        <mxCell id="1" parent="0"/>
        <mxCell id="2" value="2:1" style="rounded=0;whiteSpace=wrap;html=1;fillColor=#5B9BD5;strokeColor=none;" vertex="1" parent="1">
          <mxGeometry x="0" y="0" width="1200" height="600" as="geometry"/>
        </mxCell>
      </root>
    </mxGraphModel>
  </diagram>
</mxfile>
'@
    [IO.File]::WriteAllText(
        $drawioPath,
        $drawioSource,
        (New-Object Text.UTF8Encoding($false)))

    $fallbackSvg = @'
<svg xmlns="http://www.w3.org/2000/svg" width="1200" height="600" viewBox="0 0 1200 600">
  <rect width="1200" height="600" fill="#5B9BD5"/>
</svg>
'@
    [IO.File]::WriteAllText(
        $fallbackSvgPath,
        $fallbackSvg,
        (New-Object Text.UTF8Encoding($false)))

    $compileArguments = @(
        "/nologo",
        "/target:exe",
        "/out:$exePath",
        "/r:System.Drawing.dll",
        "/r:$CoreAssemblyPath",
        $exporterSourcePath,
        $providerSourcePath,
        $programPath)
    & $cscPath @compileArguments
    if ($LASTEXITCODE -ne 0) {
        throw "Unable to compile the Word preview-provider E2E helper."
    }

    Copy-Item `
        -LiteralPath $CoreAssemblyPath `
        -Destination $tempRoot `
        -Force

    Push-Location $tempRoot
    try {
        & $exePath `
            $DrawioExecutablePath `
            $drawioPath `
            $fallbackSvgPath
        $testExitCode = [int]$LASTEXITCODE
    }
    finally {
        Pop-Location
    }

    if ($testExitCode -ne 0) {
        throw "Word preview-provider E2E failed."
    }
}
finally {
    $deadline = [DateTime]::UtcNow.AddSeconds(10)
    do {
        $residualProcesses = Get-DrawioProcesses
        if ($residualProcesses.Count -eq 0 -or
            [DateTime]::UtcNow -ge $deadline) {
            break
        }

        Start-Sleep -Milliseconds 100
    } while ($true)

    if ($residualProcesses.Count -gt 0) {
        $processIds =
            ($residualProcesses | ForEach-Object Id) -join ","
        throw "Residual draw.io processes after preview E2E: $processIds"
    }

    Remove-TestTempRoot
}

Write-Output "WORD_PREVIEW_IMAGE_PROVIDER_E2E_PASS"
