param(
    [string]$Configuration = "Release",
    [string]$AssemblyRoot,
    [switch]$SkipBuild
)

$ErrorActionPreference = "Stop"

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$buildScript = Join-Path $PSScriptRoot "build.ps1"
$tempRoot = Join-Path $env:TEMP ("DrawioPpt\word-complex-metadata-" + [Guid]::NewGuid().ToString("N"))
$tempRootPrefix = Join-Path $env:TEMP "DrawioPpt\word-complex-metadata-"
$programPath = Join-Path $tempRoot "word-complex-metadata.cs"
$exePath = Join-Path $tempRoot "word-complex-metadata.exe"
$imagePath = Join-Path $tempRoot "picture.png"
$documentPath = Join-Path $tempRoot "word-complex-metadata.docx"
$cscPath = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
$interopWord = "C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Word\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.Word.dll"
$officeCore = "C:\Windows\assembly\GAC_MSIL\office\15.0.0.0__71e9bce111e9429c\OFFICE.DLL"

function Assert-NoRunningWord {
    if (Get-Process -Name WINWORD -ErrorAction SilentlyContinue) {
        throw "请先关闭正在运行的 Word，再执行复杂元数据 E2E 测试。"
    }
}

function Remove-TestTempRoot {
    if (-not (Test-Path -LiteralPath $tempRoot)) {
        return
    }

    if (-not $tempRoot.StartsWith($tempRootPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "拒绝清理不属于复杂元数据 E2E 的临时目录：$tempRoot"
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
    $AssemblyRoot = (Resolve-Path $AssemblyRoot).Path
}

if (-not (Test-Path -LiteralPath $AssemblyRoot)) {
    throw "Word add-in assembly root not found: $AssemblyRoot"
}

Assert-NoRunningWord
New-Item -ItemType Directory -Force -Path $tempRoot | Out-Null

try {
@"
using System;
using System.IO;
using System.Text;
using DrawioPpt.Core.Models;
using DrawioPpt.Core.Services;
using DrawioPpt.WordAddIn.Services;
using DrawioPpt.WordAddIn.Word;
using WordInterop = Microsoft.Office.Interop.Word;

public static class WordComplexMetadataE2E
{
    private static WordInterop.InlineShape AddPicture(WordInterop.Document document, string imagePath)
    {
        object linkToFile = false;
        object saveWithDocument = true;
        object range = document.Range(0, 0);
        return document.InlineShapes.AddPicture(imagePath, ref linkToFile, ref saveWithDocument, ref range);
    }

    private static string BuildComplexDrawioXml()
    {
        StringBuilder builder = new StringBuilder();
        builder.Append("<mxfile><diagram id='complex'><mxGraphModel><root>");
        for (int index = 0; index < 1200; index++)
        {
            builder.Append("<mxCell id='");
            builder.Append(Guid.NewGuid().ToString("N"));
            builder.Append("' value='");
            builder.Append(Guid.NewGuid().ToString("N"));
            builder.Append("' style='rounded=1;whiteSpace=wrap;html=1;' vertex='1' parent='1'><mxGeometry x='");
            builder.Append(index % 40 * 38);
            builder.Append("' y='");
            builder.Append(index / 40 * 30);
            builder.Append("' width='32' height='22' as='geometry'/></mxCell>");
        }

        builder.Append("</root></mxGraphModel></diagram></mxfile>");
        return builder.ToString();
    }

    private static void CloseDocument(ref WordInterop.Document document)
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

        try
        {
            if (args == null || args.Length != 2)
            {
                throw new ArgumentException("Expected image and document paths.");
            }

            File.WriteAllBytes(args[0], Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wl0el0AAAAASUVORK5CYII="));
            application = (WordInterop.Application)Activator.CreateInstance(Type.GetTypeFromProgID("Word.Application", true));
            application.Visible = true;
            document = application.Documents.Add();
            document.SaveAs2(args[1]);

            string drawioXml = BuildComplexDrawioXml();
            DiagramEnvelope envelope = new DiagramEnvelope();
            envelope.DiagramId = Guid.NewGuid().ToString("N");
            envelope.DiagramName = "Complex metadata test";
            envelope.DrawioXml = drawioXml;

            DiagramEnvelopeSerializer serializer = new DiagramEnvelopeSerializer();
            DocumentDiagramStore store = new DocumentDiagramStore(serializer);
            WordPictureMetadataService metadata = new WordPictureMetadataService(serializer, new PresentationSidecarPathBuilder());
            WordPictureReference picture = WordPictureReference.FromInlineShape(AddPicture(document, args[0]));
            store.Upsert(document, envelope);
            metadata.Save(picture, envelope);

            string alternativeText = picture.AlternativeText ?? string.Empty;
            DiagramEnvelope pictureEnvelope = serializer.Deserialize(alternativeText);
            DiagramEnvelope storedEnvelope;
            bool storedPayload = store.TryRead(document, envelope.DiagramId, out storedEnvelope) &&
                storedEnvelope != null && string.Equals(storedEnvelope.DrawioXml, drawioXml, StringComparison.Ordinal);
            bool lightweightAlternativeText = alternativeText.Length <= 2048 &&
                pictureEnvelope != null && string.IsNullOrEmpty(pictureEnvelope.DrawioXml);
            bool sameDiagramId = pictureEnvelope != null &&
                string.Equals(pictureEnvelope.DiagramId, envelope.DiagramId, StringComparison.Ordinal);

            Console.WriteLine("DrawioXmlChars=" + drawioXml.Length);
            Console.WriteLine("AlternativeTextChars=" + alternativeText.Length);
            Console.WriteLine("StoredPayload=" + storedPayload);
            Console.WriteLine("LightweightAlternativeText=" + lightweightAlternativeText);
            Console.WriteLine("SameDiagramId=" + sameDiagramId);
            return storedPayload && lightweightAlternativeText && sameDiagramId ? 0 : 1;
        }
        catch (Exception exception)
        {
            Console.WriteLine("CaughtType=" + exception.GetType().FullName);
            Console.WriteLine("CaughtMessage=" + exception.Message);
            Console.WriteLine("CaughtStack=" + exception.StackTrace);
            return 99;
        }
        finally
        {
            try
            {
                CloseDocument(ref document);
            }
            catch
            {
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

Copy-Item (Join-Path $AssemblyRoot "DrawioPpt.Core.dll") $tempRoot -Force
Copy-Item (Join-Path $AssemblyRoot "DrawioPpt.PowerPointAddIn.dll") $tempRoot -Force
Copy-Item (Join-Path $AssemblyRoot "DrawioPpt.WordAddIn.dll") $tempRoot -Force

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
    & $exePath $imagePath $documentPath
    exit [int]$LASTEXITCODE
}
finally {
    Pop-Location
}
}
finally {
    Remove-TestTempRoot
}
