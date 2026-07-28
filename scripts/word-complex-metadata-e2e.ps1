param(
    [string]$Configuration = "Release",
    [string]$AssemblyRoot,
    [string]$ProcessIdentityPath,
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
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using DrawioPpt.Core.Models;
using DrawioPpt.Core.Services;
using DrawioPpt.WordAddIn.Services;
using DrawioPpt.WordAddIn.Word;
using WordInterop = Microsoft.Office.Interop.Word;

public static class WordComplexMetadataE2E
{
    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr windowHandle, out uint processId);

    private sealed class FailingDocumentDiagramStore : DocumentDiagramStore
    {
        public FailingDocumentDiagramStore(DiagramEnvelopeSerializer serializer)
            : base(serializer)
        {
        }

        public override string Upsert(WordInterop.Document document, DiagramEnvelope envelope)
        {
            return string.Empty;
        }
    }

    private static WordInterop.InlineShape AddPicture(WordInterop.Document document, string imagePath, int position)
    {
        object linkToFile = false;
        object saveWithDocument = true;
        object range = document.Range(position, position);
        return document.InlineShapes.AddPicture(imagePath, ref linkToFile, ref saveWithDocument, ref range);
    }

    private static WordPictureReference FindPictureByDiagramId(
        WordInterop.Document document,
        DiagramEnvelopeSerializer serializer,
        string diagramId)
    {
        int index;
        for (index = 1; index <= document.InlineShapes.Count; index++)
        {
            WordPictureReference candidate = WordPictureReference.FromInlineShape(document.InlineShapes[index]);
            string alternativeText = candidate.AlternativeText ?? string.Empty;
            if (!serializer.CanDeserialize(alternativeText))
            {
                continue;
            }

            DiagramEnvelope candidateEnvelope = serializer.Deserialize(alternativeText);
            if (candidateEnvelope != null &&
                string.Equals(candidateEnvelope.DiagramId, diagramId, StringComparison.Ordinal))
            {
                return candidate;
            }
        }

        return null;
    }

    private static string GetDiagramPartIds(
        WordInterop.Document document,
        DiagramEnvelopeSerializer serializer,
        string diagramId)
    {
        List<string> partIds = new List<string>();
        int index;
        for (index = 1; index <= document.CustomXMLParts.Count; index++)
        {
            Microsoft.Office.Core.CustomXMLPart candidate = document.CustomXMLParts[index];
            string xml = candidate.XML ?? string.Empty;
            if (!serializer.CanDeserialize(xml))
            {
                continue;
            }

            DiagramEnvelope candidateEnvelope = serializer.Deserialize(xml);
            if (candidateEnvelope != null &&
                string.Equals(candidateEnvelope.DiagramId, diagramId, StringComparison.Ordinal))
            {
                partIds.Add(candidate.Id ?? string.Empty);
            }
        }

        partIds.Sort(StringComparer.OrdinalIgnoreCase);
        return string.Join("|", partIds.ToArray());
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

    private static void WriteWordProcessIdentity(WordInterop.Application application, string identityPath)
    {
        if (application == null ||
            string.IsNullOrWhiteSpace(identityPath) ||
            string.Equals(identityPath, "-", StringComparison.Ordinal))
        {
            return;
        }

        WordInterop.Window activeWindow = application.ActiveWindow;
        uint processId;
        GetWindowThreadProcessId(new IntPtr(activeWindow.Hwnd), out processId);
        if (processId == 0)
        {
            throw new InvalidOperationException("Unable to resolve the Word process identity.");
        }

        string identityDirectory = Path.GetDirectoryName(identityPath);
        if (!string.IsNullOrWhiteSpace(identityDirectory) && !Directory.Exists(identityDirectory))
        {
            Directory.CreateDirectory(identityDirectory);
        }

        using (Process process = Process.GetProcessById((int)processId))
        {
            string identity = "TestWordProcessIdentity=" +
                process.Id +
                ":" +
                process.StartTime.ToUniversalTime().Ticks;
            File.WriteAllText(identityPath, identity + Environment.NewLine, new UTF8Encoding(false));
            Console.WriteLine(identity);
        }
    }

    [STAThread]
    public static int Main(string[] args)
    {
        WordInterop.Application application = null;
        WordInterop.Document document = null;
        int exitCode = 99;

        try
        {
            if (args == null || args.Length != 3)
            {
                throw new ArgumentException("Expected image, document, and process identity paths.");
            }

            File.WriteAllBytes(args[0], Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wl0el0AAAAASUVORK5CYII="));
            application = (WordInterop.Application)Activator.CreateInstance(Type.GetTypeFromProgID("Word.Application", true));
            application.Visible = true;
            document = application.Documents.Add();
            WriteWordProcessIdentity(application, args[2]);
            document.SaveAs2(args[1]);

            string drawioXml = BuildComplexDrawioXml();
            DiagramEnvelope envelope = new DiagramEnvelope();
            envelope.FormatVersion = "word-complex-metadata-v1";
            envelope.DiagramId = Guid.NewGuid().ToString("N");
            envelope.DiagramName = "Complex metadata test";
            envelope.EditorMode = EditorMode.Url;
            envelope.EditorTarget = "https://app.diagrams.net";
            envelope.SidecarPath = "sidecars\\complex-metadata.drawio";
            envelope.UpdatedUtc = new DateTime(2026, 7, 28, 3, 0, 0, DateTimeKind.Utc);
            envelope.DrawioXml = drawioXml;

            DiagramEnvelopeSerializer serializer = new DiagramEnvelopeSerializer();
            DocumentDiagramStore store = new DocumentDiagramStore(serializer);
            WordPictureMetadataService metadata = new WordPictureMetadataService(serializer, new PresentationSidecarPathBuilder());
            WordPictureReference initialPicture = WordPictureReference.FromInlineShape(AddPicture(document, args[0], 0));
            store.Upsert(document, envelope);
            metadata.Save(initialPicture, envelope);
            initialPicture = null;

            document.Save();
            CloseDocument(ref document);
            document = application.Documents.Open(args[1]);

            string alternativeText = string.Empty;
            DiagramEnvelope pictureEnvelope = null;
            if (document.InlineShapes.Count > 0)
            {
                WordPictureReference reopenedPicture = WordPictureReference.FromInlineShape(document.InlineShapes[1]);
                alternativeText = reopenedPicture.AlternativeText ?? string.Empty;
                if (serializer.CanDeserialize(alternativeText))
                {
                    pictureEnvelope = serializer.Deserialize(alternativeText);
                }
            }

            DiagramEnvelope storedEnvelope;
            bool storedPayload = store.TryRead(document, envelope.DiagramId, out storedEnvelope) &&
                storedEnvelope != null && string.Equals(storedEnvelope.DrawioXml, drawioXml, StringComparison.Ordinal);
            bool lightweightAlternativeText = alternativeText.Length <= 2048 &&
                pictureEnvelope != null && string.IsNullOrEmpty(pictureEnvelope.DrawioXml);
            bool sameDiagramId = pictureEnvelope != null &&
                string.Equals(pictureEnvelope.DiagramId, envelope.DiagramId, StringComparison.Ordinal);
            bool referenceFieldsPreserved = pictureEnvelope != null &&
                string.Equals(pictureEnvelope.FormatVersion, envelope.FormatVersion, StringComparison.Ordinal) &&
                string.Equals(pictureEnvelope.DiagramId, envelope.DiagramId, StringComparison.Ordinal) &&
                string.Equals(pictureEnvelope.DiagramName, envelope.DiagramName, StringComparison.Ordinal) &&
                pictureEnvelope.EditorMode == envelope.EditorMode &&
                string.Equals(pictureEnvelope.EditorTarget, envelope.EditorTarget, StringComparison.Ordinal) &&
                string.Equals(pictureEnvelope.SidecarPath, envelope.SidecarPath, StringComparison.Ordinal) &&
                pictureEnvelope.UpdatedUtc == envelope.UpdatedUtc;
            bool sourceEnvelopeUntouched = string.Equals(envelope.DrawioXml, drawioXml, StringComparison.Ordinal);

            DiagramEnvelope fallbackEnvelope = new DiagramEnvelope();
            fallbackEnvelope.FormatVersion = "word-complex-metadata-v1";
            fallbackEnvelope.DiagramId = Guid.NewGuid().ToString("N");
            fallbackEnvelope.DiagramName = "Fallback metadata test";
            fallbackEnvelope.EditorMode = EditorMode.Desktop;
            fallbackEnvelope.EditorTarget = "C:\\Program Files\\draw.io\\draw.io.exe";
            fallbackEnvelope.SidecarPath = "sidecars\\fallback-metadata.drawio";
            fallbackEnvelope.UpdatedUtc = new DateTime(2026, 7, 28, 3, 5, 0, DateTimeKind.Utc);
            fallbackEnvelope.DrawioXml = drawioXml + "<!--fallback-->";

            int fallbackPosition = document.Content.End - 1;
            WordPictureReference fallbackPicture = WordPictureReference.FromInlineShape(
                AddPicture(document, args[0], fallbackPosition));
            bool fallbackStoreReplaced = false;
            AddInHost fallbackHost = null;
            try
            {
                fallbackHost = new AddInHost(application);
                FieldInfo documentStoreField = typeof(AddInHost).GetField(
                    "_documentDiagramStore",
                    BindingFlags.Instance | BindingFlags.NonPublic);
                if (documentStoreField == null)
                {
                    throw new MissingFieldException(typeof(AddInHost).FullName, "_documentDiagramStore");
                }

                FailingDocumentDiagramStore failingStore = new FailingDocumentDiagramStore(serializer);
                documentStoreField.SetValue(fallbackHost, failingStore);
                fallbackStoreReplaced = object.ReferenceEquals(
                    documentStoreField.GetValue(fallbackHost),
                    failingStore);
                if (!fallbackStoreReplaced)
                {
                    throw new InvalidOperationException("Failed to replace the document diagram store.");
                }

                MethodInfo saveManagedEnvelope = typeof(AddInHost).GetMethod(
                    "SaveManagedEnvelope",
                    BindingFlags.Instance | BindingFlags.NonPublic);
                if (saveManagedEnvelope == null)
                {
                    throw new MissingMethodException(typeof(AddInHost).FullName, "SaveManagedEnvelope");
                }

                saveManagedEnvelope.Invoke(
                    fallbackHost,
                    new object[] { fallbackPicture, fallbackEnvelope });
            }
            finally
            {
                if (fallbackHost != null)
                {
                    fallbackHost.Dispose();
                }
            }

            string fallbackAlternativeText = fallbackPicture.AlternativeText ?? string.Empty;
            DiagramEnvelope savedFallbackEnvelope = serializer.CanDeserialize(fallbackAlternativeText)
                ? serializer.Deserialize(fallbackAlternativeText)
                : null;
            bool fallbackRetainsPayload = fallbackStoreReplaced &&
                savedFallbackEnvelope != null &&
                string.Equals(savedFallbackEnvelope.DiagramId, fallbackEnvelope.DiagramId, StringComparison.Ordinal) &&
                string.Equals(savedFallbackEnvelope.DrawioXml, fallbackEnvelope.DrawioXml, StringComparison.Ordinal);
            fallbackPicture = null;

            DiagramEnvelope legacyEnvelope = new DiagramEnvelope();
            legacyEnvelope.FormatVersion = "word-complex-metadata-v1";
            legacyEnvelope.DiagramId = Guid.NewGuid().ToString("N");
            legacyEnvelope.DiagramName = "Legacy migration test";
            legacyEnvelope.EditorMode = EditorMode.Url;
            legacyEnvelope.EditorTarget = "https://app.diagrams.net";
            legacyEnvelope.SidecarPath = "sidecars\\legacy-metadata.drawio";
            legacyEnvelope.UpdatedUtc = new DateTime(2026, 7, 28, 3, 10, 0, DateTimeKind.Utc);
            legacyEnvelope.DrawioXml = drawioXml + "<!--legacy-->";

            DiagramEnvelope unexpectedLegacyStoredEnvelope;
            bool legacyInitiallyAbsent = !store.TryRead(
                document,
                legacyEnvelope.DiagramId,
                out unexpectedLegacyStoredEnvelope);
            int legacyPosition = document.Content.End - 1;
            WordPictureReference legacyPicture = WordPictureReference.FromInlineShape(
                AddPicture(document, args[0], legacyPosition));
            metadata.SaveFallback(legacyPicture, legacyEnvelope);

            bool legacyMigrationRead = false;
            bool secondLegacyMigrationRead = false;
            DiagramEnvelope migratedEnvelope = null;
            DiagramEnvelope secondMigratedEnvelope = null;
            int customXmlPartCountBeforeSecondMigration = -1;
            int customXmlPartCountAfterSecondMigration = -1;
            string legacyPartIdsBeforeSecondMigration = string.Empty;
            string legacyPartIdsAfterSecondMigration = string.Empty;
            AddInHost legacyHost = null;
            try
            {
                legacyHost = new AddInHost(application);
                MethodInfo tryReadManagedEnvelope = typeof(AddInHost).GetMethod(
                    "TryReadManagedEnvelope",
                    BindingFlags.Instance | BindingFlags.NonPublic);
                if (tryReadManagedEnvelope == null)
                {
                    throw new MissingMethodException(typeof(AddInHost).FullName, "TryReadManagedEnvelope");
                }

                object[] migrationArguments = new object[] { legacyPicture, null };
                legacyMigrationRead = (bool)tryReadManagedEnvelope.Invoke(legacyHost, migrationArguments);
                migratedEnvelope = migrationArguments[1] as DiagramEnvelope;

                customXmlPartCountBeforeSecondMigration = document.CustomXMLParts.Count;
                legacyPartIdsBeforeSecondMigration = GetDiagramPartIds(
                    document,
                    serializer,
                    legacyEnvelope.DiagramId);
                object[] secondMigrationArguments = new object[] { legacyPicture, null };
                secondLegacyMigrationRead = (bool)tryReadManagedEnvelope.Invoke(
                    legacyHost,
                    secondMigrationArguments);
                secondMigratedEnvelope = secondMigrationArguments[1] as DiagramEnvelope;
                customXmlPartCountAfterSecondMigration = document.CustomXMLParts.Count;
                legacyPartIdsAfterSecondMigration = GetDiagramPartIds(
                    document,
                    serializer,
                    legacyEnvelope.DiagramId);
            }
            finally
            {
                if (legacyHost != null)
                {
                    legacyHost.Dispose();
                }
            }

            DiagramEnvelope storedLegacyEnvelope;
            bool legacyStoredPayload = store.TryRead(
                    document,
                    legacyEnvelope.DiagramId,
                    out storedLegacyEnvelope) &&
                storedLegacyEnvelope != null &&
                string.Equals(storedLegacyEnvelope.DrawioXml, legacyEnvelope.DrawioXml, StringComparison.Ordinal);
            string migratedAlternativeText = legacyPicture.AlternativeText ?? string.Empty;
            DiagramEnvelope migratedPictureEnvelope = serializer.CanDeserialize(migratedAlternativeText)
                ? serializer.Deserialize(migratedAlternativeText)
                : null;
            bool legacyPictureLightweight = migratedAlternativeText.Length <= 2048 &&
                migratedPictureEnvelope != null &&
                string.Equals(migratedPictureEnvelope.DiagramId, legacyEnvelope.DiagramId, StringComparison.Ordinal) &&
                string.IsNullOrEmpty(migratedPictureEnvelope.DrawioXml);
            bool legacyMigrationPartIdsStable =
                !string.IsNullOrWhiteSpace(legacyPartIdsBeforeSecondMigration) &&
                string.Equals(
                    legacyPartIdsBeforeSecondMigration,
                    legacyPartIdsAfterSecondMigration,
                    StringComparison.OrdinalIgnoreCase);
            bool legacyMigrationIdempotent = secondLegacyMigrationRead &&
                secondMigratedEnvelope != null &&
                string.Equals(secondMigratedEnvelope.DrawioXml, legacyEnvelope.DrawioXml, StringComparison.Ordinal) &&
                customXmlPartCountBeforeSecondMigration == customXmlPartCountAfterSecondMigration &&
                legacyMigrationPartIdsStable &&
                legacyStoredPayload &&
                legacyPictureLightweight;
            bool legacyMigrationPassed = legacyInitiallyAbsent &&
                legacyMigrationRead &&
                migratedEnvelope != null &&
                string.Equals(migratedEnvelope.DrawioXml, legacyEnvelope.DrawioXml, StringComparison.Ordinal) &&
                legacyStoredPayload &&
                legacyPictureLightweight;
            legacyPicture = null;

            document.Save();
            CloseDocument(ref document);
            document = application.Documents.Open(args[1]);

            WordPictureReference reopenedFallbackPicture = FindPictureByDiagramId(
                document,
                serializer,
                fallbackEnvelope.DiagramId);
            string reopenedFallbackAlternativeText = reopenedFallbackPicture == null
                ? string.Empty
                : reopenedFallbackPicture.AlternativeText ?? string.Empty;
            DiagramEnvelope reopenedFallbackEnvelope = serializer.CanDeserialize(reopenedFallbackAlternativeText)
                ? serializer.Deserialize(reopenedFallbackAlternativeText)
                : null;
            bool fallbackPersistedAfterReopen = reopenedFallbackEnvelope != null &&
                string.Equals(reopenedFallbackEnvelope.DiagramId, fallbackEnvelope.DiagramId, StringComparison.Ordinal) &&
                string.Equals(reopenedFallbackEnvelope.DrawioXml, fallbackEnvelope.DrawioXml, StringComparison.Ordinal);
            reopenedFallbackPicture = null;

            WordPictureReference reopenedLegacyPicture = FindPictureByDiagramId(
                document,
                serializer,
                legacyEnvelope.DiagramId);
            string reopenedLegacyAlternativeText = reopenedLegacyPicture == null
                ? string.Empty
                : reopenedLegacyPicture.AlternativeText ?? string.Empty;
            DiagramEnvelope reopenedLegacyPictureEnvelope = serializer.CanDeserialize(reopenedLegacyAlternativeText)
                ? serializer.Deserialize(reopenedLegacyAlternativeText)
                : null;
            DiagramEnvelope reopenedStoredLegacyEnvelope;
            bool legacyMigrationPersistedAfterReopen = reopenedLegacyPictureEnvelope != null &&
                string.Equals(
                    reopenedLegacyPictureEnvelope.DiagramId,
                    legacyEnvelope.DiagramId,
                    StringComparison.Ordinal) &&
                string.IsNullOrEmpty(reopenedLegacyPictureEnvelope.DrawioXml) &&
                store.TryRead(document, legacyEnvelope.DiagramId, out reopenedStoredLegacyEnvelope) &&
                reopenedStoredLegacyEnvelope != null &&
                string.Equals(
                    reopenedStoredLegacyEnvelope.DrawioXml,
                    legacyEnvelope.DrawioXml,
                    StringComparison.Ordinal);
            reopenedLegacyPicture = null;

            Console.WriteLine("DrawioXmlChars=" + drawioXml.Length);
            Console.WriteLine("AlternativeTextChars=" + alternativeText.Length);
            Console.WriteLine("StoredPayload=" + storedPayload);
            Console.WriteLine("LightweightAlternativeText=" + lightweightAlternativeText);
            Console.WriteLine("SameDiagramId=" + sameDiagramId);
            Console.WriteLine("ReferenceFieldsPreserved=" + referenceFieldsPreserved);
            Console.WriteLine("SourceEnvelopeUntouched=" + sourceEnvelopeUntouched);
            Console.WriteLine("FallbackRetainsPayload=" + fallbackRetainsPayload);
            Console.WriteLine("LegacyMigrationPassed=" + legacyMigrationPassed);
            Console.WriteLine("LegacyMigrationPartIdsStable=" + legacyMigrationPartIdsStable);
            Console.WriteLine("LegacyMigrationIdempotent=" + legacyMigrationIdempotent);
            Console.WriteLine("FallbackPersistedAfterReopen=" + fallbackPersistedAfterReopen);
            Console.WriteLine("LegacyMigrationPersistedAfterReopen=" + legacyMigrationPersistedAfterReopen);
            exitCode = storedPayload &&
                lightweightAlternativeText &&
                sameDiagramId &&
                referenceFieldsPreserved &&
                sourceEnvelopeUntouched &&
                fallbackRetainsPayload &&
                legacyMigrationPassed &&
                legacyMigrationIdempotent &&
                fallbackPersistedAfterReopen &&
                legacyMigrationPersistedAfterReopen ? 0 : 1;
        }
        catch (Exception exception)
        {
            Console.WriteLine("CaughtType=" + exception.GetType().FullName);
            Console.WriteLine("CaughtMessage=" + exception.Message);
            Console.WriteLine("CaughtStack=" + exception.StackTrace);
            exitCode = 99;
        }
        finally
        {
            bool cleanupSucceeded = true;
            try
            {
                CloseDocument(ref document);
            }
            catch (Exception cleanupException)
            {
                cleanupSucceeded = false;
                Console.WriteLine("CleanupDocumentCloseErrorType=" + cleanupException.GetType().FullName);
                Console.WriteLine("CleanupDocumentCloseErrorMessage=" + cleanupException.Message);
            }

            if (application != null)
            {
                try
                {
                    object saveChanges = WordInterop.WdSaveOptions.wdDoNotSaveChanges;
                    object originalFormat = Type.Missing;
                    object routeDocument = Type.Missing;
                    ((WordInterop._Application)application).Quit(ref saveChanges, ref originalFormat, ref routeDocument);
                    application = null;
                }
                catch (Exception cleanupException)
                {
                    cleanupSucceeded = false;
                    Console.WriteLine("CleanupApplicationQuitErrorType=" + cleanupException.GetType().FullName);
                    Console.WriteLine("CleanupApplicationQuitErrorMessage=" + cleanupException.Message);
                }
            }

            Console.WriteLine("CleanupSucceeded=" + cleanupSucceeded);
            if (!cleanupSucceeded)
            {
                exitCode = 98;
            }
        }

        return exitCode;
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
    $identityArgument = if ([string]::IsNullOrWhiteSpace($ProcessIdentityPath)) {
        "-"
    }
    else {
        [System.IO.Path]::GetFullPath($ProcessIdentityPath)
    }

    if ($identityArgument -ne "-" -and (Test-Path -LiteralPath $identityArgument)) {
        Remove-Item -LiteralPath $identityArgument -Force
    }

    & $exePath $imagePath $documentPath $identityArgument
    $testExitCode = [int]$LASTEXITCODE
}
finally {
    Pop-Location
}

$cleanupTimeoutSeconds = 10
$cleanupDeadline = [DateTime]::UtcNow.AddSeconds($cleanupTimeoutSeconds)
do {
    $runningWordProcesses = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue)
    if ($runningWordProcesses.Count -eq 0) {
        break
    }

    if ([DateTime]::UtcNow -ge $cleanupDeadline) {
        break
    }

    Start-Sleep -Milliseconds 100
} while ($true)

if ($runningWordProcesses.Count -gt 0) {
    Write-Output "CleanupResidualWinWord=True"
    Write-Output ("CleanupResidualWinWordProcessIds=" + (($runningWordProcesses | ForEach-Object Id) -join ","))
    Write-Output ("CleanupResidualWinWordError=Word did not exit within " + $cleanupTimeoutSeconds + " seconds.")
    exit 98
}

Write-Output "CleanupResidualWinWord=False"
exit $testExitCode
}
finally {
    Remove-TestTempRoot
}
