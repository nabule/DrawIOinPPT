param(
    [string]$InstallRoot = (Join-Path $env:LOCALAPPDATA "Greensoft\DrawioPpt"),
    [string]$DrawioSourcePath = "",
    [ValidateSet("Square", "Front")]
    [string]$PictureWrapMode = "Square",
    [ValidateSet("Svg", "Png", "Jpeg", "Placeholder")]
    [string]$PictureRenderFormat = "Svg",
    [int]$PreviewPixelWidth = 0,
    [int]$DurationSeconds = 12,
    [double]$MaximumP95Ratio = 1.25,
    [double]$MaximumPerRoundP95Ratio = 1.25,
    [double]$MaximumP95DeltaMs = 50,
    [double]$MaximumSelectionP95DeltaMs = 50,
    [double]$MaximumAbsoluteSelectionP95Ms = 300,
    [double]$MaximumAbsoluteP95Ms = 300,
    [double]$MaximumAbsoluteResizeP95Ms = 300,
    [int]$MinimumOperationsPerRound = 10,
    [int]$MinimumBodyParagraphs = 100,
    [int]$MinimumBodyCharacters = 10000,
    [int]$MinimumPageCount = 8,
    [int]$MinimumTableCount = 2,
    [int]$MinimumAuxiliaryPictureCount = 3,
    [int]$WarmupSeconds = 4,
    [int]$ResizeOperationsPerRound = 12,
    [int]$SelectionSettleMilliseconds = 75,
    [int]$SelectionEventTimeoutMilliseconds = 500,
    [string]$OutputPath = "",
    [switch]$KeepArtifacts
)

$ErrorActionPreference = "Stop"

if ($DurationSeconds -lt 10) {
    throw "DurationSeconds must be at least 10 seconds for each picture."
}

if ($PreviewPixelWidth -lt 0) {
    throw "PreviewPixelWidth cannot be negative."
}

if ($PictureRenderFormat -notin @("Png", "Jpeg") -and
    $PreviewPixelWidth -gt 0) {
    throw "PreviewPixelWidth can only be used with PictureRenderFormat Png or Jpeg."
}

foreach ($positiveThreshold in @(
        $MaximumP95Ratio,
        $MaximumPerRoundP95Ratio,
        $MaximumP95DeltaMs,
        $MaximumSelectionP95DeltaMs,
        $MaximumAbsoluteSelectionP95Ms,
        $MaximumAbsoluteP95Ms,
        $MaximumAbsoluteResizeP95Ms)) {
    if ($positiveThreshold -le 0) {
        throw "Latency thresholds must be greater than zero."
    }
}

if ($MinimumOperationsPerRound -lt 1) {
    throw "MinimumOperationsPerRound must be at least one."
}

foreach ($minimumDocumentContent in @(
        $MinimumBodyParagraphs,
        $MinimumBodyCharacters,
        $MinimumPageCount,
        $MinimumTableCount,
        $MinimumAuxiliaryPictureCount)) {
    if ($minimumDocumentContent -lt 1) {
        throw "Document-content minimums must be at least one."
    }
}

if ($WarmupSeconds -lt 1) {
    throw "WarmupSeconds must be at least one."
}

if ($ResizeOperationsPerRound -lt 4 -or
    ($ResizeOperationsPerRound % 4) -ne 0) {
    throw "ResizeOperationsPerRound must be a positive multiple of four."
}

if ($SelectionSettleMilliseconds -lt 1) {
    throw "SelectionSettleMilliseconds must be at least one."
}

if ($SelectionEventTimeoutMilliseconds -lt
    $SelectionSettleMilliseconds) {
    throw "SelectionEventTimeoutMilliseconds must be at least SelectionSettleMilliseconds."
}

$installRootResolved = (Resolve-Path -LiteralPath $InstallRoot).Path
$binRoot = Join-Path $installRootResolved "bin"
$coreAssemblyPath = Join-Path $binRoot "DrawioPpt.Core.dll"
$wordAddInAssemblyPath = Join-Path $binRoot "DrawioPpt.WordAddIn.dll"

foreach ($requiredPath in @($coreAssemblyPath, $wordAddInAssemblyPath)) {
    if (-not (Test-Path -LiteralPath $requiredPath)) {
        throw "Required installed file not found: $requiredPath"
    }
}

if (@(Get-Process WINWORD -ErrorAction SilentlyContinue).Count -gt 0) {
    throw "Close Word before running the Word UI-thread stress acceptance."
}

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Join-Path $env:TEMP "DrawioPpt\word-ui-thread-stress-v1.0.8.json"
}

$outputFullPath = [System.IO.Path]::GetFullPath($OutputPath)
$outputDirectory = Split-Path -Parent $outputFullPath
if (-not (Test-Path -LiteralPath $outputDirectory)) {
    New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
}

[void][Reflection.Assembly]::LoadFrom($coreAssemblyPath)
[void][Reflection.Assembly]::LoadFrom($wordAddInAssemblyPath)

if (-not ("DrawioPptWordWindowProcessResolver" -as [type])) {
    Add-Type @'
using System;
using System.Runtime.InteropServices;

public static class DrawioPptWordWindowProcessResolver
{
    [DllImport("user32.dll", SetLastError = true)]
    private static extern uint GetWindowThreadProcessId(
        IntPtr hWnd,
        out uint processId);

    public static int GetProcessId(long windowHandle)
    {
        uint processId;
        GetWindowThreadProcessId(
            new IntPtr(windowHandle),
            out processId);
        return (int)processId;
    }

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool IsWindow(IntPtr windowHandle);
}
'@
}

function Get-WordProcessIdentity {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Application
    )

    $activeWindow = $null
    $process = $null
    try {
        $activeWindow = $Application.ActiveWindow
        $windowHandle = [long]$activeWindow.Hwnd
        $processId = [DrawioPptWordWindowProcessResolver]::GetProcessId(
            $windowHandle)
        if ($processId -le 0) {
            throw "Unable to resolve the Word process from ActiveWindow.Hwnd."
        }

        $process = Get-Process -Id $processId -ErrorAction Stop
        return [pscustomobject]@{
            ProcessId = $process.Id
            StartTimeUtcTicks =
                $process.StartTime.ToUniversalTime().Ticks
            WindowHandle = $windowHandle
        }
    }
    finally {
        if ($null -ne $process) {
            $process.Dispose()
        }

        Release-ComObject -Value $activeWindow
    }
}

function Get-SoleWordProcessIdentity {
    $processes = @(Get-Process WINWORD -ErrorAction SilentlyContinue)
    if ($processes.Count -ne 1) {
        foreach ($process in $processes) {
            $process.Dispose()
        }

        throw "Expected exactly one Word process after creating the isolated COM application, but found $($processes.Count)."
    }

    $process = $processes[0]
    try {
        return [pscustomobject]@{
            ProcessId = $process.Id
            StartTimeUtcTicks =
                $process.StartTime.ToUniversalTime().Ticks
        }
    }
    finally {
        $process.Dispose()
    }
}

function Assert-WordProcessIdentityMatch {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Expected,
        [Parameter(Mandatory = $true)]
        [object]$Actual
    )

    if ($Expected.ProcessId -ne $Actual.ProcessId -or
        $Expected.StartTimeUtcTicks -ne
            $Actual.StartTimeUtcTicks) {
        throw "The Word ActiveWindow process identity does not match the process created for this isolated COM application."
    }

    $Expected | Add-Member `
        -NotePropertyName WindowHandle `
        -NotePropertyValue ([long]$Actual.WindowHandle) `
        -Force
}

function Release-ComObject {
    param(
        [object]$Value
    )

    if ($null -eq $Value -or
        -not [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        return
    }

    try {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject(
            $Value)
    }
    catch {
    }
}

function Get-WordProcessForIdentity {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Identity
    )

    $process = Get-Process `
        -Id $Identity.ProcessId `
        -ErrorAction SilentlyContinue
    if ($null -eq $process) {
        return $null
    }

    try {
        $actualStartTicks =
            $process.StartTime.ToUniversalTime().Ticks
        if ($actualStartTicks -ne $Identity.StartTimeUtcTicks) {
            $process.Dispose()
            return $null
        }

        return $process
    }
    catch {
        $process.Dispose()
        return $null
    }
}

function New-ComplexSvg {
    $builder = New-Object System.Text.StringBuilder
    [void]$builder.Append('<svg xmlns="http://www.w3.org/2000/svg" width="1200" height="800" viewBox="0 0 1200 800">')
    [void]$builder.Append('<rect width="1200" height="800" fill="#f7fbff"/>')
    [void]$builder.Append('<text x="40" y="55" font-family="Segoe UI" font-size="30" fill="#17365d">DrawioPpt v1.0.8 Complex Vector</text>')

    for ($index = 0; $index -lt 120; $index++) {
        $column = $index % 30
        $row = [Math]::Floor($index / 30)
        $x = 22 + ($column * 38)
        $y = 82 + ($row * 34)
        $fill = if (($index % 3) -eq 0) {
            "#5b9bd5"
        }
        elseif (($index % 3) -eq 1) {
            "#70ad47"
        }
        else {
            "#ed7d31"
        }

        [void]$builder.Append(
            ('<g><rect x="{0}" y="{1}" width="30" height="22" rx="4" fill="{2}" stroke="#fff"/>' -f $x, $y, $fill))
        [void]$builder.Append(
            ('<text x="{0}" y="{1}" font-family="Segoe UI" font-size="8" fill="#fff">{2}</text></g>' -f ($x + 4), ($y + 15), $index))
    }

    [void]$builder.Append("</svg>")
    return $builder.ToString()
}

function New-ComplexDrawioXml {
    $builder = New-Object System.Text.StringBuilder
    [void]$builder.Append('<mxfile host="DrawioWord"><diagram id="ui-thread-stress-v108" name="Managed Complex v1.0.8"><mxGraphModel><root><mxCell id="0"/><mxCell id="1" parent="0"/>')

    for ($index = 0; $index -lt 1200; $index++) {
        $x = ($index % 40) * 30
        $y = [Math]::Floor($index / 40) * 25
        [void]$builder.Append(
            ('<mxCell id="n{0}" value="Node {0}" style="rounded=1;whiteSpace=wrap;html=1;fillColor=#dae8fc;strokeColor=#6c8ebf;" vertex="1" parent="1"><mxGeometry x="{1}" y="{2}" width="28" height="20" as="geometry"/></mxCell>' -f $index, $x, $y))
    }

    [void]$builder.Append("</root></mxGraphModel></diagram></mxfile>")
    return $builder.ToString()
}

function New-AuxiliarySvg {
    return @'
<svg xmlns="http://www.w3.org/2000/svg" width="240" height="120" viewBox="0 0 240 120">
  <rect width="240" height="120" rx="12" fill="#eaf2f8"/>
  <circle cx="42" cy="60" r="24" fill="#5b9bd5"/>
  <path d="M82 42h130v14H82zm0 30h92v12H82z" fill="#315b7d"/>
</svg>
'@
}

function Get-Sha256 {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    return (Get-FileHash `
        -LiteralPath $Path `
        -Algorithm SHA256).Hash.ToUpperInvariant()
}

function Get-DrawioSourceStatistics {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [Parameter(Mandatory = $true)]
        [string]$XmlText
    )

    try {
        [xml]$xmlDocument = $XmlText
    }
    catch {
        throw "Draw.io source is not valid XML: $($_.Exception.Message)"
    }

    if ($null -eq $xmlDocument.DocumentElement -or
        -not [string]::Equals(
            $xmlDocument.DocumentElement.LocalName,
            "mxfile",
            [StringComparison]::Ordinal)) {
        throw "Draw.io source root element must be mxfile."
    }

    $diagramNodes = @($xmlDocument.SelectNodes(
        "//*[local-name()='diagram']"))
    $cellNodes = @($xmlDocument.SelectNodes(
        "//*[local-name()='mxCell']"))
    $vertexNodes = @($xmlDocument.SelectNodes(
        "//*[local-name()='mxCell' and @vertex='1']"))
    $edgeNodes = @($xmlDocument.SelectNodes(
        "//*[local-name()='mxCell' and @edge='1']"))
    if ($diagramNodes.Count -lt 1 -or $cellNodes.Count -lt 2) {
        throw "Draw.io source must contain at least one diagram and two mxCell elements."
    }

    return [pscustomobject]@{
        Chars = $XmlText.Length
        Bytes = (Get-Item -LiteralPath $Path).Length
        Sha256 = Get-Sha256 -Path $Path
        DiagramCount = $diagramNodes.Count
        MxCellCount = $cellNodes.Count
        VertexCount = $vertexNodes.Count
        EdgeCount = $edgeNodes.Count
    }
}

function Export-DrawioSourceImage {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourcePath,
        [Parameter(Mandatory = $true)]
        [ValidateSet("Svg", "Png", "Jpeg")]
        [string]$Format,
        [Parameter(Mandatory = $true)]
        [string]$OutputPath,
        [int]$PreviewPixelWidth = 0
    )

    $drawioDesktopPath = "C:\Program Files\draw.io\draw.io.exe"
    if (-not (Test-Path -LiteralPath $drawioDesktopPath)) {
        throw "draw.io Desktop was not found: $drawioDesktopPath"
    }

    $drawioFormat = if ($Format -eq "Jpeg") {
        "jpg"
    }
    else {
        $Format.ToLowerInvariant()
    }
    $drawioArguments = @(
        "--export",
        "--format", $drawioFormat,
        "--page-index", "1",
        "--border", "0")
    if ($Format -in @("Png", "Jpeg") -and
        $PreviewPixelWidth -gt 0) {
        $drawioArguments += @(
            "--width",
            $PreviewPixelWidth.ToString(
                [Globalization.CultureInfo]::InvariantCulture))
    }

    $drawioArguments += @(
        "--output", $OutputPath,
        $SourcePath)
    $process = Start-Process `
        -FilePath $drawioDesktopPath `
        -ArgumentList $drawioArguments `
        -WindowStyle Hidden `
        -Wait `
        -PassThru
    try {
        if ($process.ExitCode -ne 0) {
            throw "draw.io Desktop $Format export failed with exit code $($process.ExitCode)."
        }
    }
    finally {
        $process.Dispose()
    }

    if (-not (Test-Path -LiteralPath $OutputPath) -or
        (Get-Item -LiteralPath $OutputPath).Length -le 0) {
        throw "draw.io Desktop did not create a non-empty $Format image."
    }
}

function Get-RenderedImageStatistics {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("Svg", "Png", "Jpeg", "Placeholder")]
        [string]$Format,
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    [double]$pixelWidth = 0
    [double]$pixelHeight = 0
    if ($Format -eq "Png") {
        $bytes = [System.IO.File]::ReadAllBytes($Path)
        if ($bytes.Length -lt 24 -or
            $bytes[0] -ne 0x89 -or
            $bytes[1] -ne 0x50 -or
            $bytes[2] -ne 0x4E -or
            $bytes[3] -ne 0x47) {
            throw "The rendered PNG does not have a valid PNG signature and IHDR."
        }

        $pixelWidth =
            ([uint32]$bytes[16] -shl 24) -bor
            ([uint32]$bytes[17] -shl 16) -bor
            ([uint32]$bytes[18] -shl 8) -bor
            [uint32]$bytes[19]
        $pixelHeight =
            ([uint32]$bytes[20] -shl 24) -bor
            ([uint32]$bytes[21] -shl 16) -bor
            ([uint32]$bytes[22] -shl 8) -bor
            [uint32]$bytes[23]
    }
    elseif ($Format -eq "Svg") {
        [xml]$svg = [System.IO.File]::ReadAllText($Path)
        $widthMatch = [regex]::Match(
            [string]$svg.DocumentElement.GetAttribute("width"),
            "^[\s]*([0-9]+(?:\.[0-9]+)?)")
        $heightMatch = [regex]::Match(
            [string]$svg.DocumentElement.GetAttribute("height"),
            "^[\s]*([0-9]+(?:\.[0-9]+)?)")
        if ($widthMatch.Success -and $heightMatch.Success) {
            $pixelWidth = [double]::Parse(
                $widthMatch.Groups[1].Value,
                [Globalization.CultureInfo]::InvariantCulture)
            $pixelHeight = [double]::Parse(
                $heightMatch.Groups[1].Value,
                [Globalization.CultureInfo]::InvariantCulture)
        }
        else {
            $viewBoxParts = @(
                ([string]$svg.DocumentElement.GetAttribute("viewBox")) `
                    -split "[,\s]+" |
                    Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
            if ($viewBoxParts.Count -ne 4) {
                throw "The rendered SVG has no readable width, height, or viewBox."
            }

            $pixelWidth = [double]::Parse(
                $viewBoxParts[2],
                [Globalization.CultureInfo]::InvariantCulture)
            $pixelHeight = [double]::Parse(
                $viewBoxParts[3],
                [Globalization.CultureInfo]::InvariantCulture)
        }
    }
    else {
        Add-Type -AssemblyName System.Drawing
        $image = $null
        try {
            $image = [Drawing.Image]::FromFile($Path)
            $pixelWidth = $image.Width
            $pixelHeight = $image.Height
        }
        finally {
            if ($null -ne $image) {
                $image.Dispose()
            }
        }
    }

    if ($pixelWidth -le 0 -or $pixelHeight -le 0) {
        throw "Rendered image dimensions must be greater than zero."
    }

    return [pscustomobject]@{
        Format = $Format
        PixelWidth = $pixelWidth
        PixelHeight = $pixelHeight
        AspectRatio = [Math]::Round(
            $pixelWidth / $pixelHeight,
            6)
        Bytes = (Get-Item -LiteralPath $Path).Length
    }
}

function New-SolidPlaceholderImage {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [Parameter(Mandatory = $true)]
        [double]$SourceAspectRatio
    )

    if ($SourceAspectRatio -le 0) {
        throw "Placeholder source aspect ratio must be greater than zero."
    }

    $placeholderWidth = 0
    $placeholderHeight = 0
    for ($candidateWidth = 16;
        $candidateWidth -le 128;
        $candidateWidth++) {
        $candidateHeight = [Math]::Max(
            1,
            [int][Math]::Round(
                $candidateWidth / $SourceAspectRatio,
                [MidpointRounding]::AwayFromZero))
        $candidateError = [Math]::Abs(
            ($candidateWidth / [double]$candidateHeight) -
                $SourceAspectRatio)
        if ($candidateError -le 0.001) {
            $placeholderWidth = $candidateWidth
            $placeholderHeight = $candidateHeight
            break
        }
    }

    if ($placeholderWidth -le 0 -or
        $placeholderHeight -le 0) {
        throw "Unable to create a small placeholder within the source aspect-ratio tolerance."
    }

    Add-Type -AssemblyName System.Drawing
    $bitmap = $null
    $graphics = $null
    try {
        $bitmap = New-Object Drawing.Bitmap(
            $placeholderWidth,
            $placeholderHeight,
            [Drawing.Imaging.PixelFormat]::Format24bppRgb)
        $graphics = [Drawing.Graphics]::FromImage($bitmap)
        $graphics.Clear(
            [Drawing.Color]::FromArgb(91, 155, 213))
        $bitmap.Save(
            $Path,
            [Drawing.Imaging.ImageFormat]::Png)
    }
    finally {
        if ($null -ne $graphics) {
            $graphics.Dispose()
        }

        if ($null -ne $bitmap) {
            $bitmap.Dispose()
        }
    }
}

function Remove-OwnedTestRoot {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return
    }

    $resolvedPath = (Resolve-Path -LiteralPath $Path).Path
    $expectedParent = [System.IO.Path]::GetFullPath(
        (Join-Path $env:TEMP "DrawioPpt"))
    $actualParent = [System.IO.Path]::GetFullPath(
        (Split-Path -Parent $resolvedPath))
    $leaf = Split-Path -Leaf $resolvedPath
    if (-not [string]::Equals(
            $actualParent.TrimEnd('\'),
            $expectedParent.TrimEnd('\'),
            [StringComparison]::OrdinalIgnoreCase) -or
        -not $leaf.StartsWith(
            "word-real-pointer-",
            [StringComparison]::Ordinal)) {
        throw "Refusing to recursively remove a path outside the dedicated Word test root: $resolvedPath"
    }

    Remove-Item -LiteralPath $resolvedPath -Recurse -Force
}

function Get-ShapeByName {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Document,
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    $shapes = $null
    try {
        $shapes = $Document.Shapes
        for ($index = 1; $index -le $shapes.Count; $index++) {
            $candidate = $shapes.Item($index)
            if ([string]::Equals(
                    [string]$candidate.Name,
                    $Name,
                    [System.StringComparison]::Ordinal)) {
                return $candidate
            }

            Release-ComObject -Value $candidate
        }
    }
    finally {
        Release-ComObject -Value $shapes
    }

    return $null
}

function Get-ShapeComparisonSemantics {
    param(
        [Parameter(Mandatory = $true)]
        [object]$ManagedShape,
        [Parameter(Mandatory = $true)]
        [object]$PlainShape,
        [Parameter(Mandatory = $true)]
        [double]$SourceAspectRatio
    )

    $managedAnchor = $null
    $plainAnchor = $null
    $managedAnchorParagraph = $null
    $plainAnchorParagraph = $null
    $managedAnchorParagraphRange = $null
    $plainAnchorParagraphRange = $null
    $managedWrap = $null
    $plainWrap = $null
    try {
        $managedAnchor = $ManagedShape.Anchor
        $plainAnchor = $PlainShape.Anchor
        $managedAnchorParagraph =
            $managedAnchor.Paragraphs.Item(1)
        $plainAnchorParagraph =
            $plainAnchor.Paragraphs.Item(1)
        $managedAnchorParagraphRange =
            $managedAnchorParagraph.Range
        $plainAnchorParagraphRange =
            $plainAnchorParagraph.Range
        $managedWrap = $ManagedShape.WrapFormat
        $plainWrap = $PlainShape.WrapFormat
        $sameSize =
            [Math]::Abs(
                [single]$ManagedShape.Width -
                    [single]$PlainShape.Width) -lt 0.1 -and
            [Math]::Abs(
                [single]$ManagedShape.Height -
                    [single]$PlainShape.Height) -lt 0.1
        $sameAnchorSemantics =
            [int]$managedAnchor.StoryType -eq
                [int]$plainAnchor.StoryType -and
            [int]$managedAnchorParagraphRange.Start -eq
                [int]$plainAnchorParagraphRange.Start -and
            [int]$ManagedShape.RelativeHorizontalPosition -eq
                [int]$PlainShape.RelativeHorizontalPosition -and
            [int]$ManagedShape.RelativeVerticalPosition -eq
                [int]$PlainShape.RelativeVerticalPosition
        $sameWrap =
            [int]$managedWrap.Type -eq [int]$plainWrap.Type
        $managedDisplayedAspectRatio =
            [double]$ManagedShape.Width /
                [double]$ManagedShape.Height
        $plainDisplayedAspectRatio =
            [double]$PlainShape.Width /
                [double]$PlainShape.Height
        $managedAspectRatioError = [Math]::Abs(
            $managedDisplayedAspectRatio -
                $SourceAspectRatio)
        $plainAspectRatioError = [Math]::Abs(
            $plainDisplayedAspectRatio -
                $SourceAspectRatio)
        $aspectRatioLocked =
            [int]$ManagedShape.LockAspectRatio -ne 0 -and
            [int]$PlainShape.LockAspectRatio -ne 0
        $aspectRatioPassed =
            $managedAspectRatioError -le 0.001 -and
            $plainAspectRatioError -le 0.001 -and
            $aspectRatioLocked
        return [pscustomobject]@{
            SameSize = $sameSize
            SameAnchorSemantics = $sameAnchorSemantics
            SameWrap = $sameWrap
            SourceAspectRatio = $SourceAspectRatio
            ManagedDisplayedAspectRatio =
                [Math]::Round(
                    $managedDisplayedAspectRatio,
                    6)
            PlainDisplayedAspectRatio =
                [Math]::Round(
                    $plainDisplayedAspectRatio,
                    6)
            ManagedAspectRatioError =
                [Math]::Round(
                    $managedAspectRatioError,
                    6)
            PlainAspectRatioError =
                [Math]::Round(
                    $plainAspectRatioError,
                    6)
            AspectRatioLocked = $aspectRatioLocked
            AspectRatioPassed = $aspectRatioPassed
            Passed =
                $sameSize -and
                $sameAnchorSemantics -and
                $sameWrap -and
                $aspectRatioPassed
        }
    }
    finally {
        Release-ComObject -Value $plainWrap
        Release-ComObject -Value $managedWrap
        Release-ComObject -Value $plainAnchorParagraphRange
        Release-ComObject -Value $managedAnchorParagraphRange
        Release-ComObject -Value $plainAnchorParagraph
        Release-ComObject -Value $managedAnchorParagraph
        Release-ComObject -Value $plainAnchor
        Release-ComObject -Value $managedAnchor
    }
}

function Open-WordDocument {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Application,
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $documents = $null
    try {
        $documents = $Application.Documents
        return $documents.Open($Path)
    }
    finally {
        Release-ComObject -Value $documents
    }
}

function Get-WordComAddIn {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Application,
        [Parameter(Mandatory = $true)]
        [string]$ProgId
    )

    $comAddIns = $null
    try {
        $comAddIns = $Application.COMAddIns
        return $comAddIns.Item($ProgId)
    }
    finally {
        Release-ComObject -Value $comAddIns
    }
}

function Get-Percentile {
    param(
        [Parameter(Mandatory = $true)]
        [double[]]$Values,
        [Parameter(Mandatory = $true)]
        [double]$Percentile
    )

    $sorted = @($Values | Sort-Object)
    if ($sorted.Count -eq 0) {
        return 0
    }

    $index = [Math]::Min(
        $sorted.Count - 1,
        [Math]::Floor($sorted.Count * $Percentile))
    return [double]$sorted[$index]
}

function Wait-ShapeSelectionEvent {
    param(
        [Parameter(Mandatory = $true)]
        [object]$SelectionCounter,
        [Parameter(Mandatory = $true)]
        [string]$ShapeName,
        [Parameter(Mandatory = $true)]
        [int]$CountBefore,
        [Parameter(Mandatory = $true)]
        [int]$MinimumWaitMilliseconds,
        [Parameter(Mandatory = $true)]
        [int]$TimeoutMilliseconds
    )

    Start-Sleep -Milliseconds $MinimumWaitMilliseconds
    $watch = [Diagnostics.Stopwatch]::StartNew()
    while ($SelectionCounter.GetShapeSelectionCount(
            $ShapeName) -le $CountBefore -and
        $watch.Elapsed.TotalMilliseconds -lt
            $TimeoutMilliseconds) {
        Start-Sleep -Milliseconds 10
    }

    return $SelectionCounter.GetShapeSelectionCount(
        $ShapeName) -gt $CountBefore
}

function Measure-ShapeMotion {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Document,
        [Parameter(Mandatory = $true)]
        [object]$Shape,
        [Parameter(Mandatory = $true)]
        [object]$OtherShape,
        [Parameter(Mandatory = $true)]
        [string]$Label,
        [Parameter(Mandatory = $true)]
        [int]$Seconds,
        [Parameter(Mandatory = $true)]
        [int]$WarmupSeconds,
        [Parameter(Mandatory = $true)]
        [int]$ResizeOperations,
        [Parameter(Mandatory = $true)]
        [int]$SelectionSettleMilliseconds,
        [Parameter(Mandatory = $true)]
        [int]$SelectionEventTimeoutMilliseconds,
        [Parameter(Mandatory = $true)]
        [object]$SelectionCounter,
        [Parameter(Mandatory = $true)]
        [object]$WordProcessIdentity,
        [Parameter(Mandatory = $true)]
        [double]$SourceAspectRatio,
        [single]$DisplayWidth = 330
    )

    $neutralRange = $null
    try {
        $OtherShape.Visible = 0
        $Shape.Visible = -1
        $Shape.Left = 45
        $Shape.Top = 70
        $Shape.LockAspectRatio = 0
        $Shape.Width = $DisplayWidth
        $Shape.Height =
            [single]($DisplayWidth / $SourceAspectRatio)
        $Shape.LockAspectRatio = -1
        $Shape.Select()
        Start-Sleep -Milliseconds 250

        $neutralRange = $Document.Range(0, 0)
        $warmupWatch = [Diagnostics.Stopwatch]::StartNew()
        $warmupIteration = 0
        while ($warmupWatch.Elapsed.TotalSeconds -lt
            $WarmupSeconds) {
            $warmupIteration++
            $neutralRange.Select()
            $Shape.Select()
            if (($warmupIteration % 2) -eq 0) {
                $Shape.Left = 43
                $Shape.Top = 69
            }
            else {
                $Shape.Left = 47
                $Shape.Top = 71
            }

            Start-Sleep -Milliseconds 20
        }

        $warmupWatch.Stop()
        $Shape.Left = 45
        $Shape.Top = 70
        Start-Sleep -Milliseconds 1000
        $shapeName = [string]$Shape.Name
        $SelectionCounter.ResetShapeSelectionCount(
            $shapeName)
        $shapeSelectionEventsBefore =
            $SelectionCounter.GetShapeSelectionCount(
                $shapeName)
        $baseLeft = [single]$Shape.Left
        $baseTop = [single]$Shape.Top
        $baseWidth = [single]$Shape.Width
        $baseHeight = [single]$Shape.Height
        $samples = New-Object System.Collections.Generic.List[double]
        $selectionSamples =
            New-Object System.Collections.Generic.List[double]
        $selectionFailures = 0
        $confirmedShapeSelectionEvents = 0
        $missingShapeSelectionEvents = 0
        $setFailures = 0
        $transientRetries = 0
        $unresponsiveSamples = 0
        $selectionOperations = 0
        $iteration = 0
        $totalWatch = [Diagnostics.Stopwatch]::StartNew()

        while ($totalWatch.Elapsed.TotalSeconds -lt $Seconds) {
            $iteration++
            $neutralRange.Select()
            Start-Sleep `
                -Milliseconds $SelectionSettleMilliseconds
            $shapeEventCountBeforeSelection =
                $SelectionCounter.GetShapeSelectionCount(
                    $shapeName)
            $selectionWatch =
                [Diagnostics.Stopwatch]::StartNew()
            $selectionSucceeded = $false
            try {
                $Shape.Select()
                $selectionOperations++
                $selectionSucceeded = $true
            }
            catch {
                $selectionFailures++
            }
            finally {
                $selectionWatch.Stop()
                $selectionSamples.Add(
                    $selectionWatch.Elapsed.TotalMilliseconds)
            }

            if ($selectionSucceeded) {
                if (Wait-ShapeSelectionEvent `
                    -SelectionCounter $SelectionCounter `
                    -ShapeName $shapeName `
                    -CountBefore $shapeEventCountBeforeSelection `
                    -MinimumWaitMilliseconds $SelectionSettleMilliseconds `
                    -TimeoutMilliseconds $SelectionEventTimeoutMilliseconds) {
                    $confirmedShapeSelectionEvents++
                }
                else {
                    $missingShapeSelectionEvents++
                }
            }

            if (-not $selectionSucceeded) {
                Start-Sleep -Milliseconds 20
                continue
            }

            $stepWatch =
                [Diagnostics.Stopwatch]::StartNew()
            $setSucceeded = $false
            for ($attempt = 1; $attempt -le 3; $attempt++) {
                try {
                    $phase = $iteration % 4
                    if ($phase -eq 0) {
                        $Shape.Left = $baseLeft - 18
                        $Shape.Top = $baseTop - 8
                    }
                    elseif ($phase -eq 1) {
                        $Shape.Left = $baseLeft + 18
                        $Shape.Top = $baseTop - 8
                    }
                    elseif ($phase -eq 2) {
                        $Shape.Left = $baseLeft + 18
                        $Shape.Top = $baseTop + 8
                    }
                    else {
                        $Shape.Left = $baseLeft - 18
                        $Shape.Top = $baseTop + 8
                    }

                    $setSucceeded = $true
                    break
                }
                catch {
                    if ($attempt -lt 3) {
                        $transientRetries++
                        Start-Sleep -Milliseconds 250
                    }
                }
            }

            if (-not $setSucceeded) {
                $setFailures++
            }

            $stepWatch.Stop()
            $samples.Add($stepWatch.Elapsed.TotalMilliseconds)

            $wordProcess = Get-WordProcessForIdentity `
                -Identity $WordProcessIdentity
            if ($null -eq $wordProcess) {
                throw "The test-owned Word process exited or its identity changed."
            }

            try {
                $wordProcess.Refresh()
                if (-not $wordProcess.Responding) {
                    $unresponsiveSamples++
                }
            }
            finally {
                $wordProcess.Dispose()
            }

            Start-Sleep -Milliseconds 20
        }

        $totalWatch.Stop()
        Start-Sleep -Milliseconds 500

        $resizeSamples = New-Object System.Collections.Generic.List[double]
        $resizeTransientRetries = 0
        $resizeFailures = 0
        for ($resizeIndex = 0;
            $resizeIndex -lt $ResizeOperations;
            $resizeIndex++) {
            $neutralRange.Select()
            Start-Sleep `
                -Milliseconds $SelectionSettleMilliseconds
            $shapeEventCountBeforeResizeSelection =
                $SelectionCounter.GetShapeSelectionCount(
                    $shapeName)
            $resizeSelectionWatch =
                [Diagnostics.Stopwatch]::StartNew()
            $resizeSelectionSucceeded = $false
            try {
                $Shape.Select()
                $selectionOperations++
                $resizeSelectionSucceeded = $true
            }
            catch {
                $selectionFailures++
            }
            finally {
                $resizeSelectionWatch.Stop()
                $selectionSamples.Add(
                    $resizeSelectionWatch.Elapsed.TotalMilliseconds)
            }

            if ($resizeSelectionSucceeded) {
                if (Wait-ShapeSelectionEvent `
                    -SelectionCounter $SelectionCounter `
                    -ShapeName $shapeName `
                    -CountBefore $shapeEventCountBeforeResizeSelection `
                    -MinimumWaitMilliseconds $SelectionSettleMilliseconds `
                    -TimeoutMilliseconds $SelectionEventTimeoutMilliseconds) {
                    $confirmedShapeSelectionEvents++
                }
                else {
                    $missingShapeSelectionEvents++
                }
            }

            if (-not $resizeSelectionSucceeded) {
                $resizeFailures++
                continue
            }

            $resizeWatch = [Diagnostics.Stopwatch]::StartNew()
            $resizeSucceeded = $false
            for ($attempt = 1; $attempt -le 3; $attempt++) {
                try {
                    $resizePhase = $resizeIndex % 4
                    if ($resizePhase -eq 0) {
                        $Shape.Width = [single]($baseWidth + 6)
                    }
                    elseif ($resizePhase -eq 1) {
                        $Shape.Width = [single]$baseWidth
                    }
                    elseif ($resizePhase -eq 2) {
                        $Shape.Width = [single]($baseWidth - 6)
                    }
                    else {
                        $Shape.Width = [single]$baseWidth
                    }

                    $resizeSucceeded = $true
                    break
                }
                catch {
                    if ($attempt -lt 3) {
                        $resizeTransientRetries++
                        Start-Sleep -Milliseconds 250
                    }
                }
            }

            if (-not $resizeSucceeded) {
                $resizeFailures++
            }

            $resizeWatch.Stop()
            $resizeSamples.Add($resizeWatch.Elapsed.TotalMilliseconds)
            Start-Sleep -Milliseconds 250
        }

        Start-Sleep -Milliseconds 500
        $shapeSelectionEventsAfter =
            $SelectionCounter.GetShapeSelectionCount(
                $shapeName)
        $classifiedShapeSelectionEventDelta =
            [Math]::Max(
                0,
                $shapeSelectionEventsAfter -
                    $shapeSelectionEventsBefore)

        $sampleArray = [double[]]$samples.ToArray()
        $selectionSampleArray =
            [double[]]$selectionSamples.ToArray()
        $resizeSampleArray = [double[]]$resizeSamples.ToArray()
        $average = ($sampleArray | Measure-Object -Average).Average
        $maximum = ($sampleArray | Measure-Object -Maximum).Maximum
        $resizeAverage = ($resizeSampleArray | Measure-Object -Average).Average
        $resizeMaximum = ($resizeSampleArray | Measure-Object -Maximum).Maximum
        $selectionAverage =
            ($selectionSampleArray | Measure-Object -Average).Average
        $selectionMaximum =
            ($selectionSampleArray | Measure-Object -Maximum).Maximum
        $startAspectRatio =
            [double]$baseWidth / [double]$baseHeight
        $endWidth = [single]$Shape.Width
        $endHeight = [single]$Shape.Height
        $endAspectRatio =
            [double]$endWidth / [double]$endHeight
        $startAspectRatioError = [Math]::Abs(
            $startAspectRatio - $SourceAspectRatio)
        $endAspectRatioError = [Math]::Abs(
            $endAspectRatio - $SourceAspectRatio)
        $geometryPassed =
            $startAspectRatioError -le 0.001 -and
            $endAspectRatioError -le 0.001 -and
            [int]$Shape.LockAspectRatio -ne 0

        return [pscustomobject]@{
            Label = $Label
            DurationSeconds = [Math]::Round(
                $totalWatch.Elapsed.TotalSeconds,
                3)
            Operations = $sampleArray.Count
            SelectionOperations = $selectionOperations
            SelectionAverageMs = [Math]::Round(
                $selectionAverage,
                3)
            SelectionP95Ms = [Math]::Round(
                (Get-Percentile `
                    -Values $selectionSampleArray `
                    -Percentile 0.95),
                3)
            SelectionMaxMs = [Math]::Round(
                $selectionMaximum,
                3)
            SelectionFailures = $selectionFailures
            ConfirmedShapeSelectionEvents =
                $confirmedShapeSelectionEvents
            MissingShapeSelectionEvents =
                $missingShapeSelectionEvents
            ClassifiedShapeSelectionEventDelta =
                $classifiedShapeSelectionEventDelta
            AverageMs = [Math]::Round($average, 3)
            P95Ms = [Math]::Round(
                (Get-Percentile -Values $sampleArray -Percentile 0.95),
                3)
            P99Ms = [Math]::Round(
                (Get-Percentile -Values $sampleArray -Percentile 0.99),
                3)
            MaxMs = [Math]::Round($maximum, 3)
            Over100Ms = @($sampleArray | Where-Object { $_ -gt 100 }).Count
            Over250Ms = @($sampleArray | Where-Object { $_ -gt 250 }).Count
            TransientRetries = $transientRetries
            SetFailures = $setFailures
            ResizeOperations = $resizeSampleArray.Count
            ResizeAverageMs = [Math]::Round($resizeAverage, 3)
            ResizeP95Ms = [Math]::Round(
                (Get-Percentile -Values $resizeSampleArray -Percentile 0.95),
                3)
            ResizeMaxMs = [Math]::Round($resizeMaximum, 3)
            ResizeTransientRetries = $resizeTransientRetries
            ResizeSetFailures = $resizeFailures
            UnresponsiveSamples = $unresponsiveSamples
            StartLeft = $baseLeft
            StartTop = $baseTop
            StartWidth = $baseWidth
            StartHeight = $baseHeight
            StartAspectRatio =
                [Math]::Round($startAspectRatio, 6)
            StartAspectRatioError =
                [Math]::Round($startAspectRatioError, 6)
            EndWidth = $endWidth
            EndHeight = $endHeight
            EndAspectRatio =
                [Math]::Round($endAspectRatio, 6)
            EndAspectRatioError =
                [Math]::Round($endAspectRatioError, 6)
            AspectRatioLocked =
                [int]$Shape.LockAspectRatio -ne 0
            GeometryPassed = $geometryPassed
            SelectionSamplesMs = $selectionSampleArray
            MotionSamplesMs = $sampleArray
            ResizeSamplesMs = $resizeSampleArray
        }
    }
    finally {
        Release-ComObject -Value $neutralRange
    }
}

function Merge-ShapeMotionResults {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Results,
        [Parameter(Mandatory = $true)]
        [string]$Label
    )

    $sampleArray = [double[]]@(
        $Results | ForEach-Object { $_.MotionSamplesMs })
    $selectionSampleArray = [double[]]@(
        $Results | ForEach-Object { $_.SelectionSamplesMs })
    $resizeSampleArray = [double[]]@(
        $Results | ForEach-Object { $_.ResizeSamplesMs })

    return [pscustomobject]@{
        Label = $Label
        Rounds = $Results.Count
        DurationSeconds = [Math]::Round(
            ($Results | Measure-Object -Property DurationSeconds -Sum).Sum,
            3)
        Operations =
            ($Results | Measure-Object -Property Operations -Sum).Sum
        SelectionOperations =
            ($Results | Measure-Object -Property SelectionOperations -Sum).Sum
        SelectionAverageMs = [Math]::Round(
            ($selectionSampleArray |
                Measure-Object -Average).Average,
            3)
        SelectionP95Ms = [Math]::Round(
            (Get-Percentile `
                -Values $selectionSampleArray `
                -Percentile 0.95),
            3)
        SelectionMaxMs = [Math]::Round(
            ($selectionSampleArray |
                Measure-Object -Maximum).Maximum,
            3)
        SelectionFailures =
            ($Results |
                Measure-Object -Property SelectionFailures -Sum).Sum
        ConfirmedShapeSelectionEvents =
            ($Results |
                Measure-Object `
                    -Property ConfirmedShapeSelectionEvents `
                    -Sum).Sum
        MissingShapeSelectionEvents =
            ($Results |
                Measure-Object `
                    -Property MissingShapeSelectionEvents `
                    -Sum).Sum
        ClassifiedShapeSelectionEventDelta =
            ($Results |
                Measure-Object `
                    -Property ClassifiedShapeSelectionEventDelta `
                    -Sum).Sum
        AverageMs = [Math]::Round(
            ($sampleArray | Measure-Object -Average).Average,
            3)
        P95Ms = [Math]::Round(
            (Get-Percentile -Values $sampleArray -Percentile 0.95),
            3)
        P99Ms = [Math]::Round(
            (Get-Percentile -Values $sampleArray -Percentile 0.99),
            3)
        MaxMs = [Math]::Round(
            ($sampleArray | Measure-Object -Maximum).Maximum,
            3)
        Over100Ms =
            ($Results | Measure-Object -Property Over100Ms -Sum).Sum
        Over250Ms =
            ($Results | Measure-Object -Property Over250Ms -Sum).Sum
        TransientRetries =
            ($Results | Measure-Object -Property TransientRetries -Sum).Sum
        SetFailures =
            ($Results | Measure-Object -Property SetFailures -Sum).Sum
        ResizeOperations =
            ($Results | Measure-Object -Property ResizeOperations -Sum).Sum
        ResizeAverageMs = [Math]::Round(
            ($resizeSampleArray | Measure-Object -Average).Average,
            3)
        ResizeP95Ms = [Math]::Round(
            (Get-Percentile -Values $resizeSampleArray -Percentile 0.95),
            3)
        ResizeMaxMs = [Math]::Round(
            ($resizeSampleArray | Measure-Object -Maximum).Maximum,
            3)
        ResizeTransientRetries =
            ($Results |
                Measure-Object -Property ResizeTransientRetries -Sum).Sum
        ResizeSetFailures =
            ($Results |
                Measure-Object -Property ResizeSetFailures -Sum).Sum
        UnresponsiveSamples =
            ($Results |
                Measure-Object -Property UnresponsiveSamples -Sum).Sum
        StartLeft = [single]$Results[0].StartLeft
        StartTop = [single]$Results[0].StartTop
    }
}

function Wait-WordProcessExit {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Identity,
        [int]$TimeoutSeconds = 15
    )

    $watch = [Diagnostics.Stopwatch]::StartNew()
    while ($watch.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
        $process = Get-WordProcessForIdentity -Identity $Identity
        if ($null -eq $process) {
            return $true
        }

        $process.Dispose()
        Start-Sleep -Milliseconds 250
    }

    return $false
}

function Stop-OwnedWordProcess {
    param(
        [object]$Identity
    )

    if ($null -eq $Identity) {
        return
    }

    $process = Get-WordProcessForIdentity -Identity $Identity
    if ($null -eq $process) {
        return
    }

    try {
        if ($null -eq $Identity.PSObject.Properties["WindowHandle"] -or
            [long]$Identity.WindowHandle -le 0 -or
            -not [DrawioPptWordWindowProcessResolver]::IsWindow(
                [IntPtr][long]$Identity.WindowHandle) -or
            [DrawioPptWordWindowProcessResolver]::GetProcessId(
                [long]$Identity.WindowHandle) -ne
                [int]$Identity.ProcessId) {
            throw "Refusing forced cleanup because the captured Word PID, start time, and HWND identity is no longer exact."
        }

        Stop-Process -Id $Identity.ProcessId -Force -ErrorAction Stop
        if (-not $process.WaitForExit(5000)) {
            throw "The test-owned Word process did not exit after forced cleanup."
        }
    }
    finally {
        $process.Dispose()
    }

    $remainingProcess =
        Get-WordProcessForIdentity -Identity $Identity
    if ($null -ne $remainingProcess) {
        $remainingProcess.Dispose()
        throw "The test-owned Word process identity still exists after forced cleanup."
    }
}

function New-TypedStressDocument {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ImagePath,
        [Parameter(Mandatory = $true)]
        [string]$AuxiliaryImagePath,
        [Parameter(Mandatory = $true)]
        [string]$DocumentPath,
        [Parameter(Mandatory = $true)]
        [string]$DrawioXml,
        [Parameter(Mandatory = $true)]
        [string]$CoreAssemblyPath,
        [Parameter(Mandatory = $true)]
        [string]$WordAddInAssemblyPath,
        [Parameter(Mandatory = $true)]
        [ValidateSet("Square", "Front")]
        [string]$PictureWrapMode,
        [Parameter(Mandatory = $true)]
        [double]$SourceAspectRatio,
        [Parameter(Mandatory = $true)]
        [int]$MinimumBodyParagraphs,
        [Parameter(Mandatory = $true)]
        [int]$MinimumBodyCharacters,
        [Parameter(Mandatory = $true)]
        [int]$MinimumPageCount,
        [Parameter(Mandatory = $true)]
        [int]$MinimumTableCount,
        [Parameter(Mandatory = $true)]
        [int]$MinimumAuxiliaryPictureCount
    )

    $wordInteropPath = (
        Get-ChildItem `
            -LiteralPath "C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Word" `
            -Recurse `
            -Filter "Microsoft.Office.Interop.Word.dll" `
            -ErrorAction Stop |
            Select-Object -First 1).FullName
    $officeInteropPath = (
        Get-ChildItem `
            -LiteralPath "C:\Windows\assembly\GAC_MSIL\Office" `
            -Recurse `
            -Filter "OFFICE.DLL" `
            -ErrorAction Stop |
            Select-Object -First 1).FullName

    if ([string]::IsNullOrWhiteSpace($wordInteropPath) -or
        [string]::IsNullOrWhiteSpace($officeInteropPath)) {
        throw "Microsoft Office interop assemblies were not found in the GAC."
    }

    if (-not ("DrawioPptWordUiStressDocumentBuilder" -as [type])) {
        $source = @'
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using DrawioPpt.Core.Models;
using DrawioPpt.Core.Services;
using DrawioPpt.WordAddIn.Services;
using DrawioPpt.WordAddIn.Word;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;

public sealed class DrawioPptWordSelectionChangeCounter :
    IDisposable
{
    private Word.Application application;

    public int Count { get; private set; }
    public int ManagedShapeSelectionCount { get; private set; }
    public int PlainShapeSelectionCount { get; private set; }

    private static void ReleaseComObject(object value)
    {
        if (value != null && Marshal.IsComObject(value))
        {
            try
            {
                Marshal.FinalReleaseComObject(value);
            }
            catch
            {
            }
        }
    }

    public int GetShapeSelectionCount(string shapeName)
    {
        if (string.Equals(
            shapeName,
            "Managed Complex v1.0.8",
            StringComparison.Ordinal))
        {
            return ManagedShapeSelectionCount;
        }

        if (string.Equals(
            shapeName,
            "Plain Complex Same Visual",
            StringComparison.Ordinal))
        {
            return PlainShapeSelectionCount;
        }

        throw new ArgumentException(
            "Unexpected stress-test shape name.",
            "shapeName");
    }

    public void ResetShapeSelectionCount(string shapeName)
    {
        if (string.Equals(
            shapeName,
            "Managed Complex v1.0.8",
            StringComparison.Ordinal))
        {
            ManagedShapeSelectionCount = 0;
            return;
        }

        if (string.Equals(
            shapeName,
            "Plain Complex Same Visual",
            StringComparison.Ordinal))
        {
            PlainShapeSelectionCount = 0;
            return;
        }

        throw new ArgumentException(
            "Unexpected stress-test shape name.",
            "shapeName");
    }

    public void Attach(Word.Application target)
    {
        if (target == null)
        {
            throw new ArgumentNullException("target");
        }

        Detach();
        application = target;
        application.WindowSelectionChange +=
            OnWindowSelectionChange;
    }

    private void OnWindowSelectionChange(
        Word.Selection selection)
    {
        Count++;
        Word.ShapeRange shapeRange = null;
        Word.Shape shape = null;
        try
        {
            shapeRange = selection.ShapeRange;
            shape = shapeRange[1];
            string shapeName = shape.Name;
            if (string.Equals(
                shapeName,
                "Managed Complex v1.0.8",
                StringComparison.Ordinal))
            {
                ManagedShapeSelectionCount++;
            }
            else if (string.Equals(
                shapeName,
                "Plain Complex Same Visual",
                StringComparison.Ordinal))
            {
                PlainShapeSelectionCount++;
            }
        }
        catch
        {
        }
        finally
        {
            ReleaseComObject(shape);
            ReleaseComObject(shapeRange);
        }
    }

    public void Detach()
    {
        if (application == null)
        {
            return;
        }

        try
        {
            application.WindowSelectionChange -=
                OnWindowSelectionChange;
        }
        finally
        {
            application = null;
        }
    }

    public void Dispose()
    {
        Detach();
    }
}

public static class DrawioPptWordUiStressDocumentBuilder
{
    [DllImport("user32.dll", SetLastError = true)]
    private static extern uint GetWindowThreadProcessId(
        IntPtr hWnd,
        out uint processId);

    private static void ReleaseComObject(object value)
    {
        if (value != null && Marshal.IsComObject(value))
        {
            try
            {
                Marshal.FinalReleaseComObject(value);
            }
            catch
            {
            }
        }
    }

    private static bool IsOwnedProcessRunning(
        uint processId,
        long processStartTimeUtcTicks)
    {
        if (processId == 0 || processStartTimeUtcTicks == 0)
        {
            return false;
        }

        try
        {
            using (Process process = Process.GetProcessById((int)processId))
            {
                return process.StartTime.ToUniversalTime().Ticks ==
                    processStartTimeUtcTicks;
            }
        }
        catch
        {
            return false;
        }
    }

    private static bool IsExactOwnedWindow(
        uint processId,
        long windowHandle)
    {
        if (processId == 0 || windowHandle == 0)
        {
            return false;
        }

        uint windowProcessId;
        GetWindowThreadProcessId(
            new IntPtr(windowHandle),
            out windowProcessId);
        return windowProcessId == processId;
    }

    private static void CaptureSoleWordProcess(
        out uint processId,
        out long processStartTimeUtcTicks)
    {
        Process[] processes =
            Process.GetProcessesByName("WINWORD");
        try
        {
            if (processes.Length != 1)
            {
                throw new InvalidOperationException(
                    "Expected exactly one isolated Word process after creating the document-builder COM application, but found " +
                    processes.Length.ToString() + ".");
            }

            processId = (uint)processes[0].Id;
            processStartTimeUtcTicks =
                processes[0].StartTime.ToUniversalTime().Ticks;
        }
        finally
        {
            foreach (Process process in processes)
            {
                process.Dispose();
            }
        }
    }

    private static void StopOwnedProcessIfNeeded(
        uint processId,
        long processStartTimeUtcTicks,
        long windowHandle)
    {
        Stopwatch watch = Stopwatch.StartNew();
        while (watch.Elapsed.TotalSeconds < 15 &&
            IsOwnedProcessRunning(processId, processStartTimeUtcTicks))
        {
            Thread.Sleep(250);
        }

        if (!IsOwnedProcessRunning(processId, processStartTimeUtcTicks))
        {
            return;
        }

        if (!IsExactOwnedWindow(processId, windowHandle))
        {
            throw new InvalidOperationException(
                "Refusing forced cleanup because the document-builder PID, start time, and HWND identity is no longer exact.");
        }

        using (Process process = Process.GetProcessById((int)processId))
        {
            if (process.StartTime.ToUniversalTime().Ticks !=
                processStartTimeUtcTicks)
            {
                return;
            }

            process.Kill();
            if (!process.WaitForExit(5000) ||
                IsOwnedProcessRunning(
                    processId,
                    processStartTimeUtcTicks))
            {
                throw new InvalidOperationException(
                    "The document-builder Word process remained after forced cleanup.");
            }
        }
    }

    private static Word.Shape AddPicture(
        Word.Document document,
        string imagePath,
        float top,
        string name,
        Word.WdWrapType wrapType,
        double sourceAspectRatio)
    {
        object linkToFile = false;
        object saveWithDocument = true;
        Word.Range anchor = null;
        Word.InlineShapes inlineShapes = null;
        Word.InlineShape inline = null;
        Word.WrapFormat wrapFormat = null;
        try
        {
            anchor = document.Range(0, 0);
            inlineShapes = document.InlineShapes;
            inline = inlineShapes.AddPicture(
                imagePath,
                ref linkToFile,
                ref saveWithDocument,
                anchor);
            Word.Shape shape = inline.ConvertToShape();
            shape.Name = name;
            shape.RelativeHorizontalPosition =
                Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage;
            shape.RelativeVerticalPosition =
                Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionPage;
            shape.Left = 45f;
            shape.Top = top;
            shape.LockAspectRatio =
                Office.MsoTriState.msoFalse;
            shape.Width = 330f;
            shape.Height =
                (float)(330.0 / sourceAspectRatio);
            shape.LockAspectRatio =
                Office.MsoTriState.msoTrue;
            wrapFormat = shape.WrapFormat;
            wrapFormat.Type = wrapType;
            return shape;
        }
        finally
        {
            ReleaseComObject(wrapFormat);
            ReleaseComObject(inline);
            ReleaseComObject(inlineShapes);
            ReleaseComObject(anchor);
        }
    }

    private static void BuildRichBody(
        Word.Application application,
        Word.Document document,
        string auxiliaryImagePath,
        int minimumBodyParagraphs,
        int minimumBodyCharacters,
        int minimumPageCount,
        int minimumTableCount,
        int minimumAuxiliaryPictureCount)
    {
        int paragraphTarget = Math.Max(100, minimumBodyParagraphs);
        int pageTarget = Math.Max(8, minimumPageCount);
        int paragraphCharacters = Math.Max(
            140,
            (minimumBodyCharacters / paragraphTarget) + 40);
        int paragraphsPerPage = Math.Max(
            1,
            (int)Math.Ceiling(
                paragraphTarget / (double)pageTarget));
        Word.Selection selection = null;
        try
        {
            selection = application.Selection;
            selection.SetRange(0, 0);
            for (int index = 0; index < paragraphTarget; index++)
            {
                StringBuilder paragraph = new StringBuilder();
                paragraph.Append("DrawioPpt rich Word acceptance paragraph ");
                paragraph.Append((index + 1).ToString("D3"));
                paragraph.Append(
                    ". This document intentionally contains normal business narrative, controls, evidence, responsibilities, review notes, and implementation details. ");
                while (paragraph.Length < paragraphCharacters)
                {
                    paragraph.Append(
                        "The paragraph remains independent from the managed diagram metadata and represents ordinary Word content. ");
                }

                selection.TypeText(paragraph.ToString());
                selection.TypeParagraph();
                if ((index + 1) % paragraphsPerPage == 0 &&
                    (index + 1) < paragraphTarget)
                {
                    selection.InsertBreak(
                        Word.WdBreakType.wdPageBreak);
                }
            }
        }
        finally
        {
            ReleaseComObject(selection);
        }

        for (int tableIndex = 0;
            tableIndex < Math.Max(2, minimumTableCount);
            tableIndex++)
        {
            Word.Range tableAnchor = null;
            Word.Range afterTable = null;
            Word.Tables tables = null;
            Word.Table table = null;
            try
            {
                tableAnchor = document.Content;
                tableAnchor.Collapse(
                    Word.WdCollapseDirection.wdCollapseEnd);
                tables = document.Tables;
                table = tables.Add(tableAnchor, 4, 4);
                for (int row = 1; row <= 4; row++)
                {
                    for (int column = 1; column <= 4; column++)
                    {
                        Word.Cell cell = null;
                        Word.Range cellRange = null;
                        try
                        {
                            cell = table.Cell(row, column);
                            cellRange = cell.Range;
                            cellRange.Text =
                                "T" + (tableIndex + 1).ToString() +
                                "-R" + row.ToString() +
                                "-C" + column.ToString();
                        }
                        finally
                        {
                            ReleaseComObject(cellRange);
                            ReleaseComObject(cell);
                        }
                    }
                }

                afterTable = table.Range;
                afterTable.Collapse(
                    Word.WdCollapseDirection.wdCollapseEnd);
                afterTable.InsertParagraphAfter();
            }
            finally
            {
                ReleaseComObject(afterTable);
                ReleaseComObject(table);
                ReleaseComObject(tables);
                ReleaseComObject(tableAnchor);
            }
        }

        for (int pictureIndex = 0;
            pictureIndex < Math.Max(3, minimumAuxiliaryPictureCount);
            pictureIndex++)
        {
            Word.Range anchor = null;
            Word.InlineShapes inlineShapes = null;
            Word.InlineShape picture = null;
            object linkToFile = false;
            object saveWithDocument = true;
            try
            {
                int paragraphIndex = Math.Min(
                    document.Paragraphs.Count,
                    5 + (pictureIndex * 12));
                Word.Paragraph paragraph =
                    document.Paragraphs[paragraphIndex];
                try
                {
                    anchor = paragraph.Range.Duplicate;
                }
                finally
                {
                    ReleaseComObject(paragraph);
                }

                anchor.Collapse(
                    Word.WdCollapseDirection.wdCollapseStart);
                inlineShapes = document.InlineShapes;
                picture = inlineShapes.AddPicture(
                    auxiliaryImagePath,
                    ref linkToFile,
                    ref saveWithDocument,
                    anchor);
                picture.Title =
                    "Auxiliary Picture " +
                    (pictureIndex + 1).ToString();
                picture.AlternativeText =
                    "Ordinary auxiliary picture for rich-content acceptance.";
                picture.LockAspectRatio =
                    Office.MsoTriState.msoFalse;
                picture.Width = 72f;
                picture.Height = 36f;
            }
            finally
            {
                ReleaseComObject(picture);
                ReleaseComObject(inlineShapes);
                ReleaseComObject(anchor);
            }
        }

        Word.Sections sections = null;
        Word.Section section = null;
        Word.HeadersFooters headers = null;
        Word.HeadersFooters footers = null;
        Word.HeaderFooter header = null;
        Word.HeaderFooter footer = null;
        Word.Range headerRange = null;
        Word.Range footerRange = null;
        try
        {
            sections = document.Sections;
            section = sections[1];
            headers = section.Headers;
            footers = section.Footers;
            header = headers[
                Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
            footer = footers[
                Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
            headerRange = header.Range;
            footerRange = footer.Range;
            headerRange.Text =
                "DrawioPpt Word rich-content performance acceptance";
            footerRange.Text =
                "Confidential test content - page ";
        }
        finally
        {
            ReleaseComObject(footerRange);
            ReleaseComObject(headerRange);
            ReleaseComObject(footer);
            ReleaseComObject(header);
            ReleaseComObject(footers);
            ReleaseComObject(headers);
            ReleaseComObject(section);
            ReleaseComObject(sections);
        }
    }

    private static string ComputeSha256(string value)
    {
        using (SHA256 sha256 = SHA256.Create())
        {
            byte[] bytes = Encoding.UTF8.GetBytes(value ?? string.Empty);
            byte[] hash = sha256.ComputeHash(bytes);
            StringBuilder result = new StringBuilder(hash.Length * 2);
            foreach (byte item in hash)
            {
                result.Append(item.ToString("X2"));
            }

            return result.ToString();
        }
    }

    public static string Inspect(Word.Document document)
    {
        Word.Range bodyRange = null;
        Word.Paragraphs paragraphs = null;
        Word.Tables tables = null;
        Word.InlineShapes inlineShapes = null;
        Word.Sections sections = null;
        Word.Section section = null;
        Word.HeadersFooters headers = null;
        Word.HeadersFooters footers = null;
        Word.HeaderFooter header = null;
        Word.HeaderFooter footer = null;
        Word.Range headerRange = null;
        Word.Range footerRange = null;
        try
        {
            document.Repaginate();
            bodyRange = document.Content;
            string bodyText = bodyRange.Text ?? string.Empty;
            paragraphs = document.Paragraphs;
            tables = document.Tables;
            inlineShapes = document.InlineShapes;
            int auxiliaryPictures = 0;
            for (int index = 1; index <= inlineShapes.Count; index++)
            {
                Word.InlineShape candidate = null;
                try
                {
                    candidate = inlineShapes[index];
                    string title = candidate.Title ?? string.Empty;
                    if (title.StartsWith(
                        "Auxiliary Picture ",
                        StringComparison.Ordinal))
                    {
                        auxiliaryPictures++;
                    }
                }
                finally
                {
                    ReleaseComObject(candidate);
                }
            }

            sections = document.Sections;
            section = sections[1];
            headers = section.Headers;
            footers = section.Footers;
            header = headers[
                Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
            footer = footers[
                Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
            headerRange = header.Range;
            footerRange = footer.Range;
            bool headerPresent =
                !string.IsNullOrWhiteSpace(
                    (headerRange.Text ?? string.Empty)
                        .Trim('\r', '\a', ' '));
            bool footerPresent =
                !string.IsNullOrWhiteSpace(
                    (footerRange.Text ?? string.Empty)
                        .Trim('\r', '\a', ' '));
            int pageCount = document.ComputeStatistics(
                Word.WdStatistic.wdStatisticPages,
                false);

            return
                paragraphs.Count.ToString() + "|" +
                bodyText.Length.ToString() + "|" +
                pageCount.ToString() + "|" +
                tables.Count.ToString() + "|" +
                auxiliaryPictures.ToString() + "|" +
                headerPresent.ToString() + "|" +
                footerPresent.ToString() + "|" +
                ComputeSha256(bodyText);
        }
        finally
        {
            ReleaseComObject(footerRange);
            ReleaseComObject(headerRange);
            ReleaseComObject(footer);
            ReleaseComObject(header);
            ReleaseComObject(footers);
            ReleaseComObject(headers);
            ReleaseComObject(section);
            ReleaseComObject(sections);
            ReleaseComObject(inlineShapes);
            ReleaseComObject(tables);
            ReleaseComObject(paragraphs);
            ReleaseComObject(bodyRange);
        }
    }

    public static string Build(
        string imagePath,
        string auxiliaryImagePath,
        string documentPath,
        string drawioXml,
        string pictureWrapMode,
        double sourceAspectRatio,
        int minimumBodyParagraphs,
        int minimumBodyCharacters,
        int minimumPageCount,
        int minimumTableCount,
        int minimumAuxiliaryPictureCount)
    {
        Word.Application application = null;
        Word.Documents documents = null;
        Word.Document document = null;
        Word.PageSetup pageSetup = null;
        Word.Shape managed = null;
        Word.Shape plain = null;
        uint processId = 0;
        long processStartTimeUtcTicks = 0;
        long windowHandle = 0;
        try
        {
            if (sourceAspectRatio <= 0)
            {
                throw new ArgumentOutOfRangeException(
                    "sourceAspectRatio");
            }

            application = new Word.Application();
            CaptureSoleWordProcess(
                out processId,
                out processStartTimeUtcTicks);
            application.Visible = false;
            application.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
            documents = application.Documents;
            document = documents.Add();
            ReleaseComObject(documents);
            documents = null;
            object activeWindowObject = null;
            try
            {
                dynamic activeWindow = application.ActiveWindow;
                activeWindowObject = activeWindow;
                windowHandle = (long)activeWindow.Hwnd;
                uint activeWindowProcessId;
                GetWindowThreadProcessId(
                    new IntPtr(windowHandle),
                    out activeWindowProcessId);
                if (activeWindowProcessId != processId)
                {
                    throw new InvalidOperationException(
                        "The document-builder ActiveWindow process does not match the isolated Word process.");
                }
            }
            finally
            {
                ReleaseComObject(activeWindowObject);
            }

            if (processId == 0)
            {
                throw new InvalidOperationException(
                    "Unable to resolve the document-builder Word process.");
            }

            pageSetup = document.PageSetup;
            pageSetup.TopMargin = 30f;
            pageSetup.BottomMargin = 30f;
            pageSetup.LeftMargin = 30f;
            pageSetup.RightMargin = 30f;
            ReleaseComObject(pageSetup);
            pageSetup = null;

            Word.WdWrapType comparisonWrapType =
                string.Equals(
                    pictureWrapMode,
                    "Front",
                    StringComparison.OrdinalIgnoreCase)
                ? Word.WdWrapType.wdWrapFront
                : Word.WdWrapType.wdWrapSquare;
            BuildRichBody(
                application,
                document,
                auxiliaryImagePath,
                minimumBodyParagraphs,
                minimumBodyCharacters,
                minimumPageCount,
                minimumTableCount,
                minimumAuxiliaryPictureCount);
            managed = AddPicture(
                document,
                imagePath,
                70f,
                "Managed Complex v1.0.8",
                comparisonWrapType,
                sourceAspectRatio);
            plain = AddPicture(
                document,
                imagePath,
                390f,
                "Plain Complex Same Visual",
                comparisonWrapType,
                sourceAspectRatio);

            DiagramEnvelope envelope = new DiagramEnvelope();
            envelope.DiagramId = Guid.NewGuid().ToString("N");
            envelope.DiagramName = "Managed Complex v1.0.8";
            envelope.EditorMode = EditorMode.Url;
            envelope.EditorTarget = "https://app.diagrams.net";
            envelope.SidecarPath = string.Empty;
            envelope.UpdatedUtc = DateTime.UtcNow;
            envelope.DrawioXml = drawioXml;

            DiagramEnvelopeSerializer serializer =
                new DiagramEnvelopeSerializer();
            DocumentDiagramStore store =
                new DocumentDiagramStore(serializer);
            string partId = store.Upsert(document, envelope);
            if (string.IsNullOrWhiteSpace(partId))
            {
                throw new InvalidOperationException(
                    "Document.CustomXMLParts upsert failed.");
            }

            WordPictureMetadataService metadata =
                new WordPictureMetadataService(
                    serializer,
                    new PresentationSidecarPathBuilder());
            metadata.Save(
                WordPictureReference.FromShape(managed),
                envelope);
            plain.Title = "Plain Complex Same Visual";
            plain.AlternativeText = string.Empty;

            string alternativeText = managed.AlternativeText ?? string.Empty;
            DiagramEnvelope stored;
            if (alternativeText.Length > 2048 ||
                !store.TryRead(document, envelope.DiagramId, out stored) ||
                stored == null ||
                stored.DrawioXml != drawioXml)
            {
                throw new InvalidOperationException(
                    "Initial metadata verification failed.");
            }

            document.SaveAs2(documentPath);
            string contentInspection = Inspect(document);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            ReleaseComObject(document);
            document = null;
            application.Quit(Word.WdSaveOptions.wdDoNotSaveChanges);
            ReleaseComObject(application);
            application = null;

            return
                processId.ToString() + "|" +
                processStartTimeUtcTicks.ToString() + "|" +
                windowHandle.ToString() + "|" +
                partId + "|" +
                alternativeText.Length.ToString() + "|" +
                envelope.DiagramId + "|" +
                contentInspection;
        }
        finally
        {
            ReleaseComObject(pageSetup);
            ReleaseComObject(documents);
            ReleaseComObject(plain);
            ReleaseComObject(managed);
            if (document != null)
            {
                try
                {
                    document.Close(
                        Word.WdSaveOptions.wdDoNotSaveChanges);
                }
                catch
                {
                }

                ReleaseComObject(document);
            }

            if (application != null)
            {
                try
                {
                    application.Quit(
                        Word.WdSaveOptions.wdDoNotSaveChanges);
                }
                catch
                {
                }

                ReleaseComObject(application);
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            StopOwnedProcessIfNeeded(
                processId,
                processStartTimeUtcTicks,
                windowHandle);
        }
    }
}
'@

        Add-Type `
            -TypeDefinition $source `
            -ReferencedAssemblies @(
                $CoreAssemblyPath,
                $WordAddInAssemblyPath,
                $wordInteropPath,
                $officeInteropPath,
                "System.dll",
                "System.Core.dll",
                "Microsoft.CSharp.dll",
                "System.Xml.dll",
                "System.Xml.Linq.dll") `
            -Language CSharp `
            -IgnoreWarnings
    }

    $buildText = [DrawioPptWordUiStressDocumentBuilder]::Build(
        $ImagePath,
        $AuxiliaryImagePath,
        $DocumentPath,
        $DrawioXml,
        $PictureWrapMode,
        $SourceAspectRatio,
        $MinimumBodyParagraphs,
        $MinimumBodyCharacters,
        $MinimumPageCount,
        $MinimumTableCount,
        $MinimumAuxiliaryPictureCount)
    $parts = $buildText -split "\|"
    if ($parts.Count -ne 14) {
        throw "Unexpected typed document builder result: $buildText"
    }

    return [pscustomobject]@{
        ProcessId = [int]$parts[0]
        ProcessStartTimeUtcTicks = [long]$parts[1]
        WindowHandle = [long]$parts[2]
        CustomXmlPartId = $parts[3]
        AlternativeTextChars = [int]$parts[4]
        DiagramId = $parts[5]
        BodyParagraphs = [int]$parts[6]
        BodyCharacters = [int]$parts[7]
        PageCount = [int]$parts[8]
        TableCount = [int]$parts[9]
        AuxiliaryPictureCount = [int]$parts[10]
        HeaderPresent = [bool]::Parse($parts[11])
        FooterPresent = [bool]::Parse($parts[12])
        BodySha256 = $parts[13]
    }
}

function Get-TypedDocumentContent {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Document
    )

    $inspection = [DrawioPptWordUiStressDocumentBuilder]::Inspect(
        $Document)
    $parts = $inspection -split "\|"
    if ($parts.Count -ne 8) {
        throw "Unexpected typed document inspection result: $inspection"
    }

    return [pscustomobject]@{
        BodyParagraphs = [int]$parts[0]
        BodyCharacters = [int]$parts[1]
        PageCount = [int]$parts[2]
        TableCount = [int]$parts[3]
        AuxiliaryPictureCount = [int]$parts[4]
        HeaderPresent = [bool]::Parse($parts[5])
        FooterPresent = [bool]::Parse($parts[6])
        BodySha256 = $parts[7]
    }
}

$testRoot = Join-Path (
    Join-Path $env:TEMP "DrawioPpt") (
    "word-real-pointer-" + [Guid]::NewGuid().ToString("N"))
try {
    New-Item -ItemType Directory -Path $testRoot -Force | Out-Null
$sourceCopyPath = Join-Path $testRoot "source.drawio"
$pictureExtension = switch ($PictureRenderFormat) {
    "Jpeg" { "jpg" }
    "Placeholder" { "png" }
    default { $PictureRenderFormat.ToLowerInvariant() }
}
$picturePath = Join-Path `
    $testRoot `
    ("source." + $pictureExtension)
$svgPath = if ($PictureRenderFormat -eq "Svg") {
    $picturePath
}
else {
    ""
}
$auxiliarySvgPath = Join-Path $testRoot "auxiliary.svg"
$documentPath = Join-Path $testRoot "word-ui-thread-stress.docx"
[System.IO.File]::WriteAllText(
    $auxiliarySvgPath,
    (New-AuxiliarySvg),
    (New-Object System.Text.UTF8Encoding($false)))
$sourceMode = "Synthetic"
$sourceReportPath = "<SYNTHETIC_DRAWIO_SOURCE>"
$originalSourcePath = $null
$originalSourceSha256 = $null
if ([string]::IsNullOrWhiteSpace($DrawioSourcePath)) {
    $drawioXml = New-ComplexDrawioXml
    [System.IO.File]::WriteAllText(
        $sourceCopyPath,
        $drawioXml,
        (New-Object System.Text.UTF8Encoding($false)))
    if ($PictureRenderFormat -eq "Svg") {
        [System.IO.File]::WriteAllText(
            $picturePath,
            (New-ComplexSvg),
            (New-Object System.Text.UTF8Encoding($false)))
    }
}
else {
    if (-not (Test-Path -LiteralPath $DrawioSourcePath -PathType Leaf)) {
        throw "Draw.io source file was not found: $DrawioSourcePath"
    }

    $sourceMode = "UserProvided"
    $sourceReportPath = "<USER_DRAWIO_SOURCE>"
    $originalSourcePath =
        (Resolve-Path -LiteralPath $DrawioSourcePath).Path
    $originalSourceSha256 =
        Get-Sha256 -Path $originalSourcePath
    $drawioXml =
        [System.IO.File]::ReadAllText($originalSourcePath)
    [void](Get-DrawioSourceStatistics `
        -Path $originalSourcePath `
        -XmlText $drawioXml)
    Copy-Item `
        -LiteralPath $originalSourcePath `
        -Destination $sourceCopyPath
    if ((Get-Sha256 -Path $sourceCopyPath) -ne
        $originalSourceSha256) {
        throw "The copied Draw.io source SHA256 differs from the original."
    }

}

$sourceDisplayAspectRatio = 0
if ($PictureRenderFormat -eq "Placeholder") {
    $ratioProbePath = Join-Path `
        $testRoot `
        "source-ratio-probe.png"
    Export-DrawioSourceImage `
        -SourcePath $sourceCopyPath `
        -Format "Png" `
        -OutputPath $ratioProbePath `
        -PreviewPixelWidth 620
    $ratioProbeStatistics =
        Get-RenderedImageStatistics `
            -Format "Png" `
            -Path $ratioProbePath
    $sourceDisplayAspectRatio =
        $ratioProbeStatistics.AspectRatio
    New-SolidPlaceholderImage `
        -Path $picturePath `
        -SourceAspectRatio $sourceDisplayAspectRatio
    Remove-Item -LiteralPath $ratioProbePath -Force
}
elseif ($sourceMode -eq "UserProvided" -or
    $PictureRenderFormat -in @("Png", "Jpeg")) {
    Export-DrawioSourceImage `
        -SourcePath $sourceCopyPath `
        -Format $PictureRenderFormat `
        -OutputPath $picturePath `
        -PreviewPixelWidth $PreviewPixelWidth
}
$renderedImageStatistics = Get-RenderedImageStatistics `
    -Format $PictureRenderFormat `
    -Path $picturePath
if ($sourceDisplayAspectRatio -le 0) {
    $sourceDisplayAspectRatio =
        $renderedImageStatistics.AspectRatio
}
$renderingAspectRatioError = [Math]::Abs(
    $renderedImageStatistics.AspectRatio -
        $sourceDisplayAspectRatio)
$renderingAspectRatioPassed =
    $renderingAspectRatioError -le 0.001
if (-not $renderingAspectRatioPassed) {
    throw "Rendered image aspect ratio differs from the source aspect ratio."
}
$sourceStatistics = Get-DrawioSourceStatistics `
    -Path $sourceCopyPath `
    -XmlText $drawioXml

$application = $null
$document = $null
$managedShape = $null
$plainShape = $null
$reopenedManaged = $null
$wordAddIn = $null
$measurementWordAddIn = $null
$selectionCounter = $null
$activeIdentity = $null
$reopenIdentity = $null
$builderIdentity = $null
$ownedIdentities =
    New-Object System.Collections.Generic.List[object]
$report = $null

    Write-Host "Stage=BuildDocument"
    $buildResult = New-TypedStressDocument `
        -ImagePath $picturePath `
        -AuxiliaryImagePath $auxiliarySvgPath `
        -DocumentPath $documentPath `
        -DrawioXml $drawioXml `
        -CoreAssemblyPath $coreAssemblyPath `
        -WordAddInAssemblyPath $wordAddInAssemblyPath `
        -PictureWrapMode $PictureWrapMode `
        -SourceAspectRatio $sourceDisplayAspectRatio `
        -MinimumBodyParagraphs $MinimumBodyParagraphs `
        -MinimumBodyCharacters $MinimumBodyCharacters `
        -MinimumPageCount $MinimumPageCount `
        -MinimumTableCount $MinimumTableCount `
        -MinimumAuxiliaryPictureCount $MinimumAuxiliaryPictureCount
    $builderIdentity = [pscustomobject]@{
        ProcessId = $buildResult.ProcessId
        StartTimeUtcTicks = $buildResult.ProcessStartTimeUtcTicks
        WindowHandle = $buildResult.WindowHandle
    }
    $ownedIdentities.Add($builderIdentity)
    if (-not (Wait-WordProcessExit -Identity $builderIdentity)) {
        throw "The typed document-builder Word process did not exit."
    }
    $buildDocumentContentPassed =
        $buildResult.BodyParagraphs -ge $MinimumBodyParagraphs -and
        $buildResult.BodyCharacters -ge $MinimumBodyCharacters -and
        $buildResult.PageCount -ge $MinimumPageCount -and
        $buildResult.TableCount -ge $MinimumTableCount -and
        $buildResult.AuxiliaryPictureCount -ge
            $MinimumAuxiliaryPictureCount -and
        $buildResult.HeaderPresent -and
        $buildResult.FooterPresent
    Write-Host "BuildBodyParagraphs=$($buildResult.BodyParagraphs)"
    Write-Host "BuildBodyCharacters=$($buildResult.BodyCharacters)"
    Write-Host "BuildPageCount=$($buildResult.PageCount)"
    Write-Host "BuildTableCount=$($buildResult.TableCount)"
    Write-Host "BuildAuxiliaryPictureCount=$($buildResult.AuxiliaryPictureCount)"
    Write-Host "BuildHeaderPresent=$($buildResult.HeaderPresent)"
    Write-Host "BuildFooterPresent=$($buildResult.FooterPresent)"
    Write-Host "BuildBodySha256=$($buildResult.BodySha256)"
    if (-not $buildDocumentContentPassed) {
        throw "The typed Word document did not meet the requested rich-content minimums."
    }

    Write-Host "Stage=StartVisibleWord"
    $application = New-Object -ComObject Word.Application
    $activeIdentity = Get-SoleWordProcessIdentity
    $ownedIdentities.Add($activeIdentity)
    $application.Visible = $true
    $application.DisplayAlerts = 0
    $application.ScreenUpdating = $true
    $document = Open-WordDocument `
        -Application $application `
        -Path $documentPath
    $activeWindowIdentity =
        Get-WordProcessIdentity -Application $application
    Assert-WordProcessIdentityMatch `
        -Expected $activeIdentity `
        -Actual $activeWindowIdentity
    $managedShape = Get-ShapeByName `
        -Document $document `
        -Name "Managed Complex v1.0.8"
    $plainShape = Get-ShapeByName `
        -Document $document `
        -Name "Plain Complex Same Visual"
    if ($null -eq $managedShape -or $null -eq $plainShape) {
        throw "The stress-test pictures were not found in the generated document."
    }
    $shapeComparisonSemantics =
        Get-ShapeComparisonSemantics `
            -ManagedShape $managedShape `
            -PlainShape $plainShape `
            -SourceAspectRatio `
                $sourceDisplayAspectRatio
    Write-Host "ComparisonSameSize=$($shapeComparisonSemantics.SameSize)"
    Write-Host "ComparisonSameAnchorSemantics=$($shapeComparisonSemantics.SameAnchorSemantics)"
    Write-Host "ComparisonSameWrap=$($shapeComparisonSemantics.SameWrap)"
    Write-Host "ComparisonSourceAspectRatio=$($shapeComparisonSemantics.SourceAspectRatio)"
    Write-Host "ComparisonManagedDisplayedAspectRatio=$($shapeComparisonSemantics.ManagedDisplayedAspectRatio)"
    Write-Host "ComparisonPlainDisplayedAspectRatio=$($shapeComparisonSemantics.PlainDisplayedAspectRatio)"
    Write-Host "ComparisonAspectRatioPassed=$($shapeComparisonSemantics.AspectRatioPassed)"
    if (-not $shapeComparisonSemantics.Passed) {
        throw "Managed and plain comparison pictures do not have identical size, anchor semantics, and wrapping."
    }
    $serializer = New-Object DrawioPpt.Core.Services.DiagramEnvelopeSerializer
    $measurementWordAddIn = Get-WordComAddIn `
        -Application $application `
        -ProgId "Greensoft.DrawioWordAddIn"
    $measurementWordAddInConnected =
        [bool]$measurementWordAddIn.Connect
    if (-not $measurementWordAddInConnected) {
        throw "The DrawioPpt Word add-in is not connected in the measured Word instance."
    }

    $selectionCounter =
        New-Object DrawioPptWordSelectionChangeCounter
    $selectionCounter.Attach($application)
    $selectionCountBefore = $selectionCounter.Count

    Write-Host "Stage=Round1PlainThenManaged"
    $round1Plain = Measure-ShapeMotion `
        -Document $document `
        -Shape $plainShape `
        -OtherShape $managedShape `
        -Label "Round1PlainComplex" `
        -Seconds $DurationSeconds `
        -WarmupSeconds $WarmupSeconds `
        -ResizeOperations $ResizeOperationsPerRound `
        -SelectionSettleMilliseconds $SelectionSettleMilliseconds `
        -SelectionEventTimeoutMilliseconds $SelectionEventTimeoutMilliseconds `
        -SelectionCounter $selectionCounter `
        -WordProcessIdentity $activeIdentity `
        -SourceAspectRatio $sourceDisplayAspectRatio
    $round1Managed = Measure-ShapeMotion `
        -Document $document `
        -Shape $managedShape `
        -OtherShape $plainShape `
        -Label "Round1ManagedComplex" `
        -Seconds $DurationSeconds `
        -WarmupSeconds $WarmupSeconds `
        -ResizeOperations $ResizeOperationsPerRound `
        -SelectionSettleMilliseconds $SelectionSettleMilliseconds `
        -SelectionEventTimeoutMilliseconds $SelectionEventTimeoutMilliseconds `
        -SelectionCounter $selectionCounter `
        -WordProcessIdentity $activeIdentity `
        -SourceAspectRatio $sourceDisplayAspectRatio

    Write-Host "Stage=Round2ManagedThenPlain"
    $round2Managed = Measure-ShapeMotion `
        -Document $document `
        -Shape $managedShape `
        -OtherShape $plainShape `
        -Label "Round2ManagedComplex" `
        -Seconds $DurationSeconds `
        -WarmupSeconds $WarmupSeconds `
        -ResizeOperations $ResizeOperationsPerRound `
        -SelectionSettleMilliseconds $SelectionSettleMilliseconds `
        -SelectionEventTimeoutMilliseconds $SelectionEventTimeoutMilliseconds `
        -SelectionCounter $selectionCounter `
        -WordProcessIdentity $activeIdentity `
        -SourceAspectRatio $sourceDisplayAspectRatio
    $round2Plain = Measure-ShapeMotion `
        -Document $document `
        -Shape $plainShape `
        -OtherShape $managedShape `
        -Label "Round2PlainComplex" `
        -Seconds $DurationSeconds `
        -WarmupSeconds $WarmupSeconds `
        -ResizeOperations $ResizeOperationsPerRound `
        -SelectionSettleMilliseconds $SelectionSettleMilliseconds `
        -SelectionEventTimeoutMilliseconds $SelectionEventTimeoutMilliseconds `
        -SelectionCounter $selectionCounter `
        -WordProcessIdentity $activeIdentity `
        -SourceAspectRatio $sourceDisplayAspectRatio

    $plainRounds = @($round1Plain, $round2Plain)
    $managedRounds = @($round1Managed, $round2Managed)
    $allRoundResults = @(
        $round1Plain,
        $round1Managed,
        $round2Managed,
        $round2Plain)
    $plainResult = Merge-ShapeMotionResults `
        -Results $plainRounds `
        -Label "PlainComplex"
    $managedResult = Merge-ShapeMotionResults `
        -Results $managedRounds `
        -Label "ManagedComplex"
    $observedSelectionChangeEvents =
        $selectionCounter.Count - $selectionCountBefore
    $expectedSelectionChangeEvents =
        ($allRoundResults |
            Measure-Object -Property SelectionOperations -Sum).Sum
    $confirmedShapeSelectionEvents =
        ($allRoundResults |
            Measure-Object `
                -Property ConfirmedShapeSelectionEvents `
                -Sum).Sum
    $missingShapeSelectionEvents =
        ($allRoundResults |
            Measure-Object `
                -Property MissingShapeSelectionEvents `
                -Sum).Sum
    $selectionChangeEventsPassed =
        $confirmedShapeSelectionEvents -eq
            $expectedSelectionChangeEvents -and
        $missingShapeSelectionEvents -eq 0
    $selectionCounter.Detach()
    $selectionCounter = $null
    Write-Host "Stage=MotionMeasured"

    $managedShape.Left = [single]($managedResult.StartLeft + 24)
    $managedShape.Top = [single]($managedResult.StartTop + 12)
    $managedShape.Width = [single]($managedShape.Width + 10)
    $plainShape.Visible = -1
    $plainShape.Left = 45
    $plainShape.Top = 390
    $plainShape.LockAspectRatio = -1
    $plainShape.Width = 330
    $expectedLeft = [single]$managedShape.Left
    $expectedTop = [single]$managedShape.Top
    $expectedWidth = [single]$managedShape.Width
    $expectedHeight = [single]$managedShape.Height
    $document.Save()
    $document.Close(0)
    Release-ComObject -Value $measurementWordAddIn
    $measurementWordAddIn = $null
    Release-ComObject -Value $plainShape
    $plainShape = $null
    Release-ComObject -Value $managedShape
    $managedShape = $null
    Release-ComObject -Value $document
    $document = $null
    $application.Quit()
    Release-ComObject -Value $application
    $application = $null

    if (-not (Wait-WordProcessExit -Identity $activeIdentity)) {
        throw "The first test-owned Word process did not exit within the timeout."
    }

    Write-Host "Stage=ReopenWord"
    $application = New-Object -ComObject Word.Application
    $reopenIdentity = Get-SoleWordProcessIdentity
    $ownedIdentities.Add($reopenIdentity)
    $application.Visible = $false
    $application.DisplayAlerts = 0
    $document = Open-WordDocument `
        -Application $application `
        -Path $documentPath
    $reopenWindowIdentity =
        Get-WordProcessIdentity -Application $application
    Assert-WordProcessIdentityMatch `
        -Expected $reopenIdentity `
        -Actual $reopenWindowIdentity
    $reopenedDocumentContent =
        Get-TypedDocumentContent -Document $document
    $reopenedDocumentContentPassed =
        $reopenedDocumentContent.BodyParagraphs -ge
            $MinimumBodyParagraphs -and
        $reopenedDocumentContent.BodyCharacters -ge
            $MinimumBodyCharacters -and
        $reopenedDocumentContent.PageCount -ge
            $MinimumPageCount -and
        $reopenedDocumentContent.TableCount -ge
            $MinimumTableCount -and
        $reopenedDocumentContent.AuxiliaryPictureCount -ge
            $MinimumAuxiliaryPictureCount -and
        $reopenedDocumentContent.HeaderPresent -and
        $reopenedDocumentContent.FooterPresent
    $documentContentStable =
        $reopenedDocumentContent.BodyParagraphs -eq
            $buildResult.BodyParagraphs -and
        $reopenedDocumentContent.BodyCharacters -eq
            $buildResult.BodyCharacters -and
        $reopenedDocumentContent.PageCount -eq
            $buildResult.PageCount -and
        $reopenedDocumentContent.TableCount -eq
            $buildResult.TableCount -and
        $reopenedDocumentContent.AuxiliaryPictureCount -eq
            $buildResult.AuxiliaryPictureCount -and
        $reopenedDocumentContent.HeaderPresent -eq
            $buildResult.HeaderPresent -and
        $reopenedDocumentContent.FooterPresent -eq
            $buildResult.FooterPresent -and
        [string]::Equals(
            $reopenedDocumentContent.BodySha256,
            $buildResult.BodySha256,
            [StringComparison]::Ordinal)
    if (-not $reopenedDocumentContentPassed -or
        -not $documentContentStable) {
        throw "Rich Word document content changed or fell below its minimums after save and reopen."
    }
    $reopenedManaged = Get-ShapeByName `
        -Document $document `
        -Name "Managed Complex v1.0.8"
    if ($null -eq $reopenedManaged) {
        throw "Managed shape was not found after reopening the document."
    }

    $referenceSerialized = [string]$reopenedManaged.AlternativeText
    $reopenedReference = $serializer.Deserialize($referenceSerialized)
    $storedEnvelope = $null
    $customXmlParts = $null
    try {
        $customXmlParts = $document.CustomXMLParts
        for ($index = 1; $index -le $customXmlParts.Count; $index++) {
            $candidatePart = $null
            try {
                $candidatePart = $customXmlParts.Item($index)
                $candidateXml = [string]$candidatePart.XML
            }
            finally {
                Release-ComObject -Value $candidatePart
            }

            if (-not $serializer.CanDeserialize($candidateXml)) {
                continue
            }

            $candidateEnvelope = $serializer.Deserialize($candidateXml)
            if ($candidateEnvelope.DiagramId -eq
                $reopenedReference.DiagramId) {
                $storedEnvelope = $candidateEnvelope
                break
            }
        }
    }
    finally {
        Release-ComObject -Value $customXmlParts
    }

    $wordAddIn = Get-WordComAddIn `
        -Application $application `
        -ProgId "Greensoft.DrawioWordAddIn"
    $wordAddInConnected = [bool]$wordAddIn.Connect
    $reopenPersistencePassed =
        [Math]::Abs([single]$reopenedManaged.Left - $expectedLeft) -lt 0.1 -and
        [Math]::Abs([single]$reopenedManaged.Top - $expectedTop) -lt 0.1 -and
        [Math]::Abs([single]$reopenedManaged.Width - $expectedWidth) -lt 0.1 -and
        [Math]::Abs([single]$reopenedManaged.Height - $expectedHeight) -lt 0.1
    $reopenedAspectRatio =
        [double]$reopenedManaged.Width /
            [double]$reopenedManaged.Height
    $reopenedAspectRatioError = [Math]::Abs(
        $reopenedAspectRatio -
            $sourceDisplayAspectRatio)
    $reopenedGeometryPassed =
        $reopenedAspectRatioError -le 0.001 -and
        [int]$reopenedManaged.LockAspectRatio -ne 0
    $storedPayloadPassed =
        $null -ne $storedEnvelope -and
        [string]$storedEnvelope.DrawioXml -eq $drawioXml -and
        [string]$reopenedReference.DrawioXml -eq "" -and
        $storedEnvelope.DiagramId -eq $reopenedReference.DiagramId

    $p95Ratio = if ($plainResult.P95Ms -gt 0) {
        [Math]::Round($managedResult.P95Ms / $plainResult.P95Ms, 3)
    }
    else {
        0
    }
    $p95DeltaMs = [Math]::Round(
        $managedResult.P95Ms - $plainResult.P95Ms,
        3)
    $selectionP95Ratio =
        if ($plainResult.SelectionP95Ms -gt 0) {
            [Math]::Round(
                $managedResult.SelectionP95Ms /
                    $plainResult.SelectionP95Ms,
                3)
        }
        else {
            0
        }
    $selectionP95DeltaMs = [Math]::Round(
        $managedResult.SelectionP95Ms -
            $plainResult.SelectionP95Ms,
        3)
    $resizeP95Ratio = if ($plainResult.ResizeP95Ms -gt 0) {
        [Math]::Round(
            $managedResult.ResizeP95Ms / $plainResult.ResizeP95Ms,
            3)
    }
    else {
        0
    }
    $resizeP95DeltaMs = [Math]::Round(
        $managedResult.ResizeP95Ms -
            $plainResult.ResizeP95Ms,
        3)
    $round1P95Ratio = if ($round1Plain.P95Ms -gt 0) {
        [Math]::Round(
            $round1Managed.P95Ms / $round1Plain.P95Ms,
            3)
    }
    else {
        0
    }
    $round1P95DeltaMs = [Math]::Round(
        $round1Managed.P95Ms - $round1Plain.P95Ms,
        3)
    $round2P95Ratio = if ($round2Plain.P95Ms -gt 0) {
        [Math]::Round(
            $round2Managed.P95Ms / $round2Plain.P95Ms,
            3)
    }
    else {
        0
    }
    $round2P95DeltaMs = [Math]::Round(
        $round2Managed.P95Ms - $round2Plain.P95Ms,
        3)
    $round1SelectionP95Ratio =
        if ($round1Plain.SelectionP95Ms -gt 0) {
            [Math]::Round(
                $round1Managed.SelectionP95Ms /
                    $round1Plain.SelectionP95Ms,
                3)
        }
    else {
        0
    }
    $round1ResizeP95DeltaMs = [Math]::Round(
        $round1Managed.ResizeP95Ms -
            $round1Plain.ResizeP95Ms,
        3)
    $round2SelectionP95Ratio =
        if ($round2Plain.SelectionP95Ms -gt 0) {
            [Math]::Round(
                $round2Managed.SelectionP95Ms /
                    $round2Plain.SelectionP95Ms,
                3)
        }
        else {
            0
        }
    $round1SelectionP95DeltaMs = [Math]::Round(
        $round1Managed.SelectionP95Ms -
            $round1Plain.SelectionP95Ms,
        3)
    $round2SelectionP95DeltaMs = [Math]::Round(
        $round2Managed.SelectionP95Ms -
            $round2Plain.SelectionP95Ms,
        3)
    $round1ResizeP95Ratio =
        if ($round1Plain.ResizeP95Ms -gt 0) {
            [Math]::Round(
                $round1Managed.ResizeP95Ms /
                    $round1Plain.ResizeP95Ms,
                3)
        }
        else {
            0
        }
    $round2ResizeP95Ratio =
        if ($round2Plain.ResizeP95Ms -gt 0) {
            [Math]::Round(
                $round2Managed.ResizeP95Ms /
                    $round2Plain.ResizeP95Ms,
                3)
        }
    else {
        0
    }
    $round2ResizeP95DeltaMs = [Math]::Round(
        $round2Managed.ResizeP95Ms -
            $round2Plain.ResizeP95Ms,
        3)
    $roundP95Ratios = @($round1P95Ratio, $round2P95Ratio)
    $roundP95DeltasMs =
        @($round1P95DeltaMs, $round2P95DeltaMs)
    $roundSelectionP95Ratios =
        @(
            $round1SelectionP95Ratio,
            $round2SelectionP95Ratio)
    $roundSelectionP95DeltasMs =
        @(
            $round1SelectionP95DeltaMs,
            $round2SelectionP95DeltaMs)
    $roundResizeP95Ratios =
        @($round1ResizeP95Ratio, $round2ResizeP95Ratio)
    $roundResizeP95DeltasMs =
        @($round1ResizeP95DeltaMs, $round2ResizeP95DeltaMs)
    $allSetFailures =
        ($allRoundResults |
            Measure-Object -Property SetFailures -Sum).Sum
    $allSelectionFailures =
        ($allRoundResults |
            Measure-Object -Property SelectionFailures -Sum).Sum
    $allResizeSetFailures =
        ($allRoundResults |
            Measure-Object -Property ResizeSetFailures -Sum).Sum
    $allUnresponsiveSamples =
        ($allRoundResults |
            Measure-Object -Property UnresponsiveSamples -Sum).Sum
    $selectionComparisonPassed =
        $selectionP95Ratio -gt 0 -and
        $selectionP95Ratio -le $MaximumP95Ratio -and
        $selectionP95DeltaMs -le
            $MaximumSelectionP95DeltaMs -and
        @(
            for ($selectionRoundIndex = 0;
                $selectionRoundIndex -lt
                    $roundSelectionP95Ratios.Count;
                $selectionRoundIndex++) {
                if (
                    $roundSelectionP95Ratios[
                        $selectionRoundIndex] -le 0 -or
                    $roundSelectionP95Ratios[
                        $selectionRoundIndex] -gt
                            $MaximumPerRoundP95Ratio -or
                        $roundSelectionP95DeltasMs[
                            $selectionRoundIndex] -gt
                                $MaximumSelectionP95DeltaMs
                ) {
                    $false
                }
            }).Count -eq 0
    $comparisonPassed =
        $p95Ratio -gt 0 -and
        $p95Ratio -le $MaximumP95Ratio -and
        $p95DeltaMs -le $MaximumP95DeltaMs -and
        $selectionComparisonPassed -and
        $resizeP95Ratio -gt 0 -and
        $resizeP95Ratio -le $MaximumP95Ratio -and
        $resizeP95DeltaMs -le $MaximumP95DeltaMs -and
        @(for ($roundIndex = 0;
                $roundIndex -lt $roundP95Ratios.Count;
                $roundIndex++) {
                if ($roundP95Ratios[$roundIndex] -le 0 -or
                    $roundP95Ratios[$roundIndex] -gt
                        $MaximumPerRoundP95Ratio -or
                    $roundP95DeltasMs[$roundIndex] -gt
                        $MaximumP95DeltaMs) {
                    $false
                }
            }).Count -eq 0 -and
        @(for ($resizeRoundIndex = 0;
                $resizeRoundIndex -lt
                    $roundResizeP95Ratios.Count;
                $resizeRoundIndex++) {
                if ($roundResizeP95Ratios[$resizeRoundIndex] -le 0 -or
                    $roundResizeP95Ratios[$resizeRoundIndex] -gt
                        $MaximumPerRoundP95Ratio -or
                    $roundResizeP95DeltasMs[$resizeRoundIndex] -gt
                        $MaximumP95DeltaMs) {
                    $false
                }
            }).Count -eq 0 -and
        $allSetFailures -eq 0 -and
        $allSelectionFailures -eq 0 -and
        $allResizeSetFailures -eq 0 -and
        $allUnresponsiveSamples -eq 0
    $absoluteLatencyGatePassed =
        @($allRoundResults |
            Where-Object {
                $_.SelectionP95Ms -gt
                    $MaximumAbsoluteSelectionP95Ms -or
                $_.P95Ms -gt $MaximumAbsoluteP95Ms -or
                $_.ResizeP95Ms -gt
                    $MaximumAbsoluteResizeP95Ms
            }).Count -eq 0
    $minimumSamplesPassed =
        @($allRoundResults |
            Where-Object {
                $_.Operations -lt $MinimumOperationsPerRound -or
                $_.SelectionOperations -lt
                    ($_.Operations + $_.ResizeOperations) -or
                $_.ResizeOperations -lt
                    $ResizeOperationsPerRound
            }).Count -eq 0
    $geometryGatePassed =
        $shapeComparisonSemantics.AspectRatioPassed -and
        $reopenedGeometryPassed -and
        @($allRoundResults |
            Where-Object {
                -not $_.GeometryPassed
            }).Count -eq 0
    $noAdditionalDifferenceObserved =
        $comparisonPassed
    $storagePassed =
        $reopenPersistencePassed -and $storedPayloadPassed
    $copiedSourceSha256 = Get-Sha256 -Path $sourceCopyPath
    $copiedSourceUnchanged =
        [string]::Equals(
            $copiedSourceSha256,
            $sourceStatistics.Sha256,
            [StringComparison]::Ordinal)
    $originalSourceUnchanged = $true
    if ($sourceMode -eq "UserProvided") {
        $originalSourceUnchanged =
            [string]::Equals(
                (Get-Sha256 -Path $originalSourcePath),
                $originalSourceSha256,
                [StringComparison]::Ordinal)
    }
    $sourcePassed =
        $copiedSourceUnchanged -and
        $originalSourceUnchanged -and
        $renderingAspectRatioPassed
    $documentContentPassed =
        $buildDocumentContentPassed -and
        $reopenedDocumentContentPassed -and
        $documentContentStable

    $report = [ordered]@{
        ReportVersion = 9
        TestedVersion = "v1.0.8"
        ExecutedUtc = [DateTime]::UtcNow.ToString("o")
        InstallRoot = $installRootResolved
        TestDocument = $documentPath
        TestWordProcessIdentity =
            "$($activeIdentity.ProcessId):$($activeIdentity.StartTimeUtcTicks):$($activeIdentity.WindowHandle)"
        ReopenWordProcessIdentity =
            "$($reopenIdentity.ProcessId):$($reopenIdentity.StartTimeUtcTicks):$($reopenIdentity.WindowHandle)"
        DurationSecondsPerPicturePerRound = $DurationSeconds
        WarmupSecondsPerPicturePerRound = $WarmupSeconds
        ResizeOperationsPerPicturePerRound =
            $ResizeOperationsPerRound
        SelectionSettleMilliseconds =
            $SelectionSettleMilliseconds
        SelectionEventTimeoutMilliseconds =
            $SelectionEventTimeoutMilliseconds
        MeasurementOrder = @(
            "Round1:PlainThenManaged",
            "Round2:ManagedThenPlain")
        PictureWrapMode = $PictureWrapMode
        PictureRenderFormat = $PictureRenderFormat
        PreviewPixelWidth = $PreviewPixelWidth
        MaximumP95Ratio = $MaximumP95Ratio
        MaximumPerRoundP95Ratio = $MaximumPerRoundP95Ratio
        MaximumP95DeltaMs = $MaximumP95DeltaMs
        MaximumSelectionP95DeltaMs =
            $MaximumSelectionP95DeltaMs
        MaximumAbsoluteSelectionP95Ms =
            $MaximumAbsoluteSelectionP95Ms
        MaximumAbsoluteP95Ms = $MaximumAbsoluteP95Ms
        MaximumAbsoluteResizeP95Ms =
            $MaximumAbsoluteResizeP95Ms
        MinimumOperationsPerRound = $MinimumOperationsPerRound
        Source = [ordered]@{
            Mode = $sourceMode
            Path = $sourceReportPath
            CopiedPath = $sourceCopyPath
            SvgPath = $svgPath
            PicturePath = $picturePath
            Rendering = [ordered]@{
                Format = $renderedImageStatistics.Format
                RequestedPixelWidth = $PreviewPixelWidth
                PixelWidth =
                    $renderedImageStatistics.PixelWidth
                PixelHeight =
                    $renderedImageStatistics.PixelHeight
                AspectRatio =
                    $renderedImageStatistics.AspectRatio
                SourceAspectRatio =
                    $sourceDisplayAspectRatio
                AspectRatioError =
                    [Math]::Round(
                        $renderingAspectRatioError,
                        6)
                Bytes = $renderedImageStatistics.Bytes
                BorderPixels = 0
                PreservesAspectRatio =
                    $renderingAspectRatioPassed
                Passed = $renderingAspectRatioPassed
            }
            Chars = $sourceStatistics.Chars
            Bytes = $sourceStatistics.Bytes
            Sha256 = $sourceStatistics.Sha256
            DiagramCount = $sourceStatistics.DiagramCount
            MxCellCount = $sourceStatistics.MxCellCount
            VertexCount = $sourceStatistics.VertexCount
            EdgeCount = $sourceStatistics.EdgeCount
            CopiedSourceSha256 = $copiedSourceSha256
            CopiedSourceUnchanged = $copiedSourceUnchanged
            OriginalSourceUnchanged = $originalSourceUnchanged
            Passed = $sourcePassed
        }
        DocumentContent = [ordered]@{
            MinimumBodyParagraphs = $MinimumBodyParagraphs
            MinimumBodyCharacters = $MinimumBodyCharacters
            MinimumPageCount = $MinimumPageCount
            MinimumTableCount = $MinimumTableCount
            MinimumAuxiliaryPictureCount =
                $MinimumAuxiliaryPictureCount
            Build = [ordered]@{
                BodyParagraphs = $buildResult.BodyParagraphs
                BodyCharacters = $buildResult.BodyCharacters
                PageCount = $buildResult.PageCount
                TableCount = $buildResult.TableCount
                AuxiliaryPictureCount =
                    $buildResult.AuxiliaryPictureCount
                HeaderPresent = $buildResult.HeaderPresent
                FooterPresent = $buildResult.FooterPresent
                BodySha256 = $buildResult.BodySha256
            }
            Reopened = [ordered]@{
                BodyParagraphs =
                    $reopenedDocumentContent.BodyParagraphs
                BodyCharacters =
                    $reopenedDocumentContent.BodyCharacters
                PageCount = $reopenedDocumentContent.PageCount
                TableCount = $reopenedDocumentContent.TableCount
                AuxiliaryPictureCount =
                    $reopenedDocumentContent.AuxiliaryPictureCount
                HeaderPresent =
                    $reopenedDocumentContent.HeaderPresent
                FooterPresent =
                    $reopenedDocumentContent.FooterPresent
                BodySha256 =
                    $reopenedDocumentContent.BodySha256
            }
            StableAfterReopen = $documentContentStable
            Passed = $documentContentPassed
        }
        ComparisonPictureSemantics = [ordered]@{
            SameSize = $shapeComparisonSemantics.SameSize
            SameAnchorSemantics =
                $shapeComparisonSemantics.SameAnchorSemantics
            SameWrap = $shapeComparisonSemantics.SameWrap
            SourceAspectRatio =
                $shapeComparisonSemantics.SourceAspectRatio
            ManagedDisplayedAspectRatio =
                $shapeComparisonSemantics.ManagedDisplayedAspectRatio
            PlainDisplayedAspectRatio =
                $shapeComparisonSemantics.PlainDisplayedAspectRatio
            ManagedAspectRatioError =
                $shapeComparisonSemantics.ManagedAspectRatioError
            PlainAspectRatioError =
                $shapeComparisonSemantics.PlainAspectRatioError
            AspectRatioLocked =
                $shapeComparisonSemantics.AspectRatioLocked
            AspectRatioPassed =
                $shapeComparisonSemantics.AspectRatioPassed
            Passed = $shapeComparisonSemantics.Passed
        }
        DrawioXmlChars = $drawioXml.Length
        AlternativeTextChars = $referenceSerialized.Length
        Rounds = @(
            [ordered]@{
                Round = 1
                Order = "PlainThenManaged"
                Plain = $round1Plain
                Managed = $round1Managed
                ManagedPlainSelectionP95Ratio =
                    $round1SelectionP95Ratio
                ManagedPlainSelectionP95DeltaMs =
                    $round1SelectionP95DeltaMs
                ManagedPlainP95Ratio = $round1P95Ratio
                ManagedPlainP95DeltaMs =
                    $round1P95DeltaMs
                ManagedPlainResizeP95Ratio =
                    $round1ResizeP95Ratio
                ManagedPlainResizeP95DeltaMs =
                    $round1ResizeP95DeltaMs
            },
            [ordered]@{
                Round = 2
                Order = "ManagedThenPlain"
                Plain = $round2Plain
                Managed = $round2Managed
                ManagedPlainSelectionP95Ratio =
                    $round2SelectionP95Ratio
                ManagedPlainSelectionP95DeltaMs =
                    $round2SelectionP95DeltaMs
                ManagedPlainP95Ratio = $round2P95Ratio
                ManagedPlainP95DeltaMs =
                    $round2P95DeltaMs
                ManagedPlainResizeP95Ratio =
                    $round2ResizeP95Ratio
                ManagedPlainResizeP95DeltaMs =
                    $round2ResizeP95DeltaMs
            })
        Plain = $plainResult
        Managed = $managedResult
        ManagedPlainSelectionP95Ratio =
            $selectionP95Ratio
        ManagedPlainSelectionP95DeltaMs =
            $selectionP95DeltaMs
        ManagedPlainP95Ratio = $p95Ratio
        ManagedPlainP95DeltaMs = $p95DeltaMs
        ManagedPlainResizeP95Ratio = $resizeP95Ratio
        ManagedPlainResizeP95DeltaMs =
            $resizeP95DeltaMs
        SelectionComparisonPassed =
            $selectionComparisonPassed
        ComparisonPassed = $comparisonPassed
        AbsoluteLatencyGatePassed =
            $absoluteLatencyGatePassed
        MinimumSamplesPassed = $minimumSamplesPassed
        MeasurementWordAddInConnect =
            $measurementWordAddInConnected
        ObservedSelectionChangeEvents =
            $observedSelectionChangeEvents
        ExpectedSelectionChangeEvents =
            $expectedSelectionChangeEvents
        ConfirmedShapeSelectionEvents =
            $confirmedShapeSelectionEvents
        MissingShapeSelectionEvents =
            $missingShapeSelectionEvents
        SelectionChangeEventsPassed =
            $selectionChangeEventsPassed
        NoAdditionalMetadataPathDifferenceObserved = $noAdditionalDifferenceObserved
        ReopenPersistencePassed = $reopenPersistencePassed
        ReopenedAspectRatio =
            [Math]::Round($reopenedAspectRatio, 6)
        ReopenedAspectRatioError =
            [Math]::Round($reopenedAspectRatioError, 6)
        ReopenedGeometryPassed = $reopenedGeometryPassed
        GeometryGatePassed = $geometryGatePassed
        StoredPayloadPassed = $storedPayloadPassed
        StoragePassed = $storagePassed
        ReopenedPositionAndSize = [ordered]@{
            Left = [single]$reopenedManaged.Left
            Top = [single]$reopenedManaged.Top
            Width = [single]$reopenedManaged.Width
            Height = [single]$reopenedManaged.Height
        }
        StoredDrawioXmlChars = if ($null -ne $storedEnvelope) {
            ([string]$storedEnvelope.DrawioXml).Length
        }
        else {
            0
        }
        WordAddInConnect = $wordAddInConnected
        PointerInputCovered = $false
        PointerInputBoundary = "This test updates Word Shape properties on the visible Word STA UI thread; it does not inject mouse pointer input."
    }

    $document.Close(0)
    Release-ComObject -Value $wordAddIn
    $wordAddIn = $null
    Release-ComObject -Value $reopenedManaged
    $reopenedManaged = $null
    Release-ComObject -Value $document
    $document = $null
    $application.Quit()
    Release-ComObject -Value $application
    $application = $null

    if (-not (Wait-WordProcessExit -Identity $reopenIdentity)) {
        throw "The reopened test-owned Word process did not exit within the timeout."
    }
    Write-Host "Stage=ReopenVerified"

    $report.CleanupResidualWinWord = @(
        Get-Process WINWORD -ErrorAction SilentlyContinue).Count -gt 0
    $report.OverallPassed =
        $report.ComparisonPassed -and
        $report.AbsoluteLatencyGatePassed -and
        $report.MinimumSamplesPassed -and
        $report.MeasurementWordAddInConnect -and
        $report.SelectionChangeEventsPassed -and
        $report.StoragePassed -and
        $report.WordAddInConnect -and
        $report.Source.Passed -and
        $report.DocumentContent.Passed -and
        $report.ComparisonPictureSemantics.Passed -and
        $report.GeometryGatePassed -and
        -not $report.CleanupResidualWinWord

    $reportJson = $report | ConvertTo-Json -Depth 8
    [System.IO.File]::WriteAllText(
        $outputFullPath,
        $reportJson,
        (New-Object System.Text.UTF8Encoding($false)))

    Write-Host "WordUiThreadStressReport=$outputFullPath"
    Write-Host "DurationSecondsPerPicturePerRound=$DurationSeconds"
    Write-Host "PlainP95Ms=$($plainResult.P95Ms)"
    Write-Host "ManagedP95Ms=$($managedResult.P95Ms)"
    Write-Host "ManagedPlainSelectionP95Ratio=$selectionP95Ratio"
    Write-Host "ManagedPlainSelectionP95DeltaMs=$selectionP95DeltaMs"
    Write-Host "ManagedPlainP95Ratio=$p95Ratio"
    Write-Host "ManagedPlainP95DeltaMs=$p95DeltaMs"
    Write-Host "ManagedPlainResizeP95Ratio=$resizeP95Ratio"
    Write-Host "ManagedPlainResizeP95DeltaMs=$resizeP95DeltaMs"
    Write-Host "ComparisonPassed=$($report.ComparisonPassed)"
    Write-Host "AbsoluteLatencyGatePassed=$($report.AbsoluteLatencyGatePassed)"
    Write-Host "MinimumSamplesPassed=$($report.MinimumSamplesPassed)"
    Write-Host "MeasurementWordAddInConnect=$($report.MeasurementWordAddInConnect)"
    Write-Host "ConfirmedShapeSelectionEvents=$($report.ConfirmedShapeSelectionEvents)"
    Write-Host "MissingShapeSelectionEvents=$($report.MissingShapeSelectionEvents)"
    Write-Host "SelectionChangeEventsPassed=$($report.SelectionChangeEventsPassed)"
    Write-Host "NoAdditionalMetadataPathDifferenceObserved=$($report.NoAdditionalMetadataPathDifferenceObserved)"
    Write-Host "ReopenPersistencePassed=$($report.ReopenPersistencePassed)"
    Write-Host "StoredPayloadPassed=$($report.StoredPayloadPassed)"
    Write-Host "SourcePassed=$($report.Source.Passed)"
    Write-Host "SourceMode=$($report.Source.Mode)"
    Write-Host "PictureRenderFormat=$($report.PictureRenderFormat)"
    Write-Host "PreviewPixelWidth=$($report.PreviewPixelWidth)"
    Write-Host "RenderedPixelWidth=$($report.Source.Rendering.PixelWidth)"
    Write-Host "RenderedPixelHeight=$($report.Source.Rendering.PixelHeight)"
    Write-Host "RenderedAspectRatio=$($report.Source.Rendering.AspectRatio)"
    Write-Host "RenderingSourceAspectRatio=$($report.Source.Rendering.SourceAspectRatio)"
    Write-Host "RenderingAspectRatioError=$($report.Source.Rendering.AspectRatioError)"
    Write-Host "RenderedBorderPixels=$($report.Source.Rendering.BorderPixels)"
    Write-Host "SourceSha256=$($report.Source.Sha256)"
    Write-Host "DrawioXmlChars=$($report.Source.Chars)"
    Write-Host "DrawioXmlBytes=$($report.Source.Bytes)"
    Write-Host "DiagramCount=$($report.Source.DiagramCount)"
    Write-Host "MxCellCount=$($report.Source.MxCellCount)"
    Write-Host "VertexCount=$($report.Source.VertexCount)"
    Write-Host "EdgeCount=$($report.Source.EdgeCount)"
    Write-Host "ReopenedBodyParagraphs=$($report.DocumentContent.Reopened.BodyParagraphs)"
    Write-Host "ReopenedBodyCharacters=$($report.DocumentContent.Reopened.BodyCharacters)"
    Write-Host "ReopenedPageCount=$($report.DocumentContent.Reopened.PageCount)"
    Write-Host "ReopenedTableCount=$($report.DocumentContent.Reopened.TableCount)"
    Write-Host "ReopenedAuxiliaryPictureCount=$($report.DocumentContent.Reopened.AuxiliaryPictureCount)"
    Write-Host "ReopenedHeaderPresent=$($report.DocumentContent.Reopened.HeaderPresent)"
    Write-Host "ReopenedFooterPresent=$($report.DocumentContent.Reopened.FooterPresent)"
    Write-Host "ReopenedBodySha256=$($report.DocumentContent.Reopened.BodySha256)"
    Write-Host "DocumentContentStable=$($report.DocumentContent.StableAfterReopen)"
    Write-Host "DocumentContentPassed=$($report.DocumentContent.Passed)"
    Write-Host "ComparisonPictureSemanticsPassed=$($report.ComparisonPictureSemantics.Passed)"
    Write-Host "GeometryGatePassed=$($report.GeometryGatePassed)"
    Write-Host "WordAddInConnect=$($report.WordAddInConnect)"
    Write-Host "PointerInputCovered=False"
    Write-Host "CleanupResidualWinWord=$($report.CleanupResidualWinWord)"
    Write-Host "OverallPassed=$($report.OverallPassed)"

    if (-not $report.OverallPassed) {
        throw "Word UI-thread stress acceptance failed. See $outputFullPath"
    }
}
finally {
    if ($null -ne $selectionCounter) {
        try {
            $selectionCounter.Dispose()
        }
        catch {
        }
    }

    Release-ComObject -Value $measurementWordAddIn
    Release-ComObject -Value $wordAddIn
    Release-ComObject -Value $reopenedManaged
    Release-ComObject -Value $plainShape
    Release-ComObject -Value $managedShape
    if ($null -ne $document) {
        try {
            $document.Close(0)
        }
        catch {
        }

        Release-ComObject -Value $document
    }

    if ($null -ne $application) {
        try {
            $application.Quit()
        }
        catch {
        }

        Release-ComObject -Value $application
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    foreach ($identity in $ownedIdentities) {
        if (-not (Wait-WordProcessExit `
                -Identity $identity `
                -TimeoutSeconds 2)) {
            Stop-OwnedWordProcess -Identity $identity
        }
    }

    if ($sourceMode -eq "UserProvided" -and
        -not [string]::IsNullOrWhiteSpace($originalSourcePath) -and
        -not [string]::IsNullOrWhiteSpace($originalSourceSha256)) {
        $originalSourceSha256After =
            Get-Sha256 -Path $originalSourcePath
        Write-Host "OriginalSourceSha256After=$originalSourceSha256After"
        if (-not [string]::Equals(
                $originalSourceSha256After,
                $originalSourceSha256,
                [StringComparison]::Ordinal)) {
            throw "The original Draw.io source changed during the test."
        }
    }

    if ($KeepArtifacts) {
        Write-Host "TestRoot=$testRoot"
        Write-Host "DocumentPath=$documentPath"
        Write-Host "SvgPath=$svgPath"
        Write-Host "PicturePath=$picturePath"
    }
    else {
        Remove-OwnedTestRoot -Path $testRoot
    }
}
