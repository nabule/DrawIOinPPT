$ErrorActionPreference = "Stop"

$scriptPath = Join-Path `
    $PSScriptRoot `
    "word-ui-thread-stress-acceptance.ps1"
$tokens = $null
$parseErrors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile(
    $scriptPath,
    [ref]$tokens,
    [ref]$parseErrors)
if ($parseErrors.Count -gt 0) {
    throw "Stress script has PowerShell parse errors."
}

function Get-StressFunctionDefinition {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    $functionAst = $ast.Find(
        {
            param($node)
            $node -is [Management.Automation.Language.FunctionDefinitionAst] -and
                [string]::Equals(
                    $node.Name,
                    $Name,
                    [StringComparison]::Ordinal)
        },
        $true)
    if ($null -eq $functionAst) {
        throw "Stress script must define $Name."
    }

    return [scriptblock]::Create($functionAst.Extent.Text)
}

function Get-MainProbeWiringAssignments {
    param(
        [Parameter(Mandatory = $true)]
        [Management.Automation.Language.ScriptBlockAst]$ScriptAst
    )

    $wiringBranches = @($ScriptAst.FindAll(
            {
                param($node)
                if ($node -isnot
                    [Management.Automation.Language.IfStatementAst] -or
                    $node.Clauses.Count -ne 1 -or
                    $node.Parent -isnot
                        [Management.Automation.Language.StatementBlockAst] -or
                    $node.Parent.Parent -isnot
                        [Management.Automation.Language.TryStatementAst] -or
                    $node.Parent.Parent.Parent -ne $ScriptAst.EndBlock) {
                    return $false
                }

                $conditionPipeline = $node.Clauses[0].Item1
                if ($conditionPipeline.PipelineElements.Count -ne 1 -or
                    $conditionPipeline.PipelineElements[0] -isnot
                        [Management.Automation.Language.CommandExpressionAst]) {
                    return $false
                }

                $conditionExpression =
                    $conditionPipeline.PipelineElements[0].Expression
                if ($conditionExpression -isnot
                    [Management.Automation.Language.BinaryExpressionAst] -or
                    $conditionExpression.Operator -ne
                        [Management.Automation.Language.TokenKind]::Or -or
                    $conditionExpression.Left -isnot
                        [Management.Automation.Language.BinaryExpressionAst] -or
                    $conditionExpression.Right -isnot
                        [Management.Automation.Language.BinaryExpressionAst]) {
                    return $false
                }

                $sourceModeCondition = $conditionExpression.Left
                $renderFormatCondition = $conditionExpression.Right
                if ($sourceModeCondition.Operator -ne
                        [Management.Automation.Language.TokenKind]::Ieq -or
                    $sourceModeCondition.Left -isnot
                        [Management.Automation.Language.VariableExpressionAst] -or
                    $sourceModeCondition.Left.VariablePath.UserPath -ne
                        "sourceMode" -or
                    $sourceModeCondition.Right -isnot
                        [Management.Automation.Language.StringConstantExpressionAst] -or
                    $sourceModeCondition.Right.Value -ne "UserProvided" -or
                    $renderFormatCondition.Operator -ne
                        [Management.Automation.Language.TokenKind]::Iin -or
                    $renderFormatCondition.Left -isnot
                        [Management.Automation.Language.VariableExpressionAst] -or
                    $renderFormatCondition.Left.VariablePath.UserPath -ne
                        "PictureRenderFormat" -or
                    $renderFormatCondition.Right -isnot
                        [Management.Automation.Language.ArrayExpressionAst]) {
                    return $false
                }

                $renderFormats =
                    @($renderFormatCondition.Right.SafeGetValue())
                return $renderFormats.Count -eq 3 -and
                    $renderFormats[0] -eq "Png" -and
                    $renderFormats[1] -eq "Jpeg" -and
                    $renderFormats[2] -eq "Placeholder"
            },
            $true))
    if ($wiringBranches.Count -ne 1) {
        return @()
    }

    $directAssignments =
        @($wiringBranches[0].Clauses[0].Item2.Statements |
            Where-Object {
                $_ -is
                    [Management.Automation.Language.AssignmentStatementAst] -and
                $_.Left -is
                    [Management.Automation.Language.VariableExpressionAst] -and
                $_.Left.VariablePath.UserPath -eq
                    "sourceDisplayAspectRatio"
            })
    if ($directAssignments.Count -ne 1 -or
        $directAssignments[0].Operator -ne
            [Management.Automation.Language.TokenKind]::Equals) {
        return @()
    }

    return @($directAssignments[0] |
        Where-Object {
            if ($_.Right -isnot
                [Management.Automation.Language.PipelineAst] -or
                $_.Right.PipelineElements.Count -ne 1 -or
                $_.Right.PipelineElements[0] -isnot
                    [Management.Automation.Language.CommandAst]) {
                return $false
            }

            $command = $_.Right.PipelineElements[0]
            return $command.GetCommandName() -eq
                    "Get-IndependentSourceDisplayAspectRatio" -and
                $command.Redirections.Count -eq 0 -and
                $command.CommandElements.Count -eq 5 -and
                $command.CommandElements[1] -is
                    [Management.Automation.Language.CommandParameterAst] -and
                $command.CommandElements[1].ParameterName -eq
                    "SourcePath" -and
                $command.CommandElements[2] -is
                    [Management.Automation.Language.VariableExpressionAst] -and
                $command.CommandElements[2].VariablePath.UserPath -eq
                    "sourceCopyPath" -and
                $command.CommandElements[3] -is
                    [Management.Automation.Language.CommandParameterAst] -and
                $command.CommandElements[3].ParameterName -eq
                    "TestRoot" -and
                $command.CommandElements[4] -is
                    [Management.Automation.Language.VariableExpressionAst] -and
                $command.CommandElements[4].VariablePath.UserPath -eq
                    "testRoot"
        })
}

function Test-SourceAspectWriteContract {
    param(
        [Parameter(Mandatory = $true)]
        [Management.Automation.Language.ScriptBlockAst]$ScriptAst
    )

    $assignments = @($ScriptAst.FindAll(
            {
                param($node)
                $node -is
                    [Management.Automation.Language.AssignmentStatementAst] -and
                    $node.Left -is
                        [Management.Automation.Language.VariableExpressionAst] -and
                    $node.Left.VariablePath.UserPath -eq
                        "sourceDisplayAspectRatio"
            },
            $true))
    $probeAssignments =
        @(Get-MainProbeWiringAssignments -ScriptAst $ScriptAst)
    $zeroAssignments = @($assignments |
        Where-Object {
            $_.Operator -eq
                [Management.Automation.Language.TokenKind]::Equals -and
            $_.Right.Extent.Text.Trim() -eq "0"
        })
    $syntheticFallbackAssignments = @($assignments |
        Where-Object {
            $_.Operator -eq
                [Management.Automation.Language.TokenKind]::Equals -and
            ($_.Right.Extent.Text -replace '\s+', '') -eq
                "1200.0/800.0"
        })
    $commandWrites = @($ScriptAst.FindAll(
            {
                param($node)
                if ($node -isnot
                    [Management.Automation.Language.CommandAst] -or
                    $node.GetCommandName() -notin @(
                        "Set-Variable",
                        "Clear-Variable",
                        "Remove-Variable")) {
                    return $false
                }

                return @($node.CommandElements |
                    Where-Object {
                        $_ -is
                            [Management.Automation.Language.StringConstantExpressionAst] -and
                        $_.Value -eq "sourceDisplayAspectRatio"
                    }).Count -gt 0
            },
            $true))
    $unaryWrites = @($ScriptAst.FindAll(
            {
                param($node)
                $node -is
                    [Management.Automation.Language.UnaryExpressionAst] -and
                $node.TokenKind -in @(
                    [Management.Automation.Language.TokenKind]::PlusPlus,
                    [Management.Automation.Language.TokenKind]::PostfixPlusPlus,
                    [Management.Automation.Language.TokenKind]::MinusMinus,
                    [Management.Automation.Language.TokenKind]::PostfixMinusMinus) -and
                $node.Child -is
                    [Management.Automation.Language.VariableExpressionAst] -and
                $node.Child.VariablePath.UserPath -eq
                    "sourceDisplayAspectRatio"
            },
            $true))
    if ($assignments.Count -ne 3 -or
        $probeAssignments.Count -ne 1 -or
        $zeroAssignments.Count -ne 1 -or
        $syntheticFallbackAssignments.Count -ne 1 -or
        $commandWrites.Count -ne 0 -or
        $unaryWrites.Count -ne 0 -or
        @($assignments |
            Where-Object {
                $_.Operator -ne
                    [Management.Automation.Language.TokenKind]::Equals
            }).Count -ne 0) {
        return $false
    }

    $probeAssignment = $probeAssignments[0]
    $probeIf = $probeAssignment.Parent.Parent
    $mainTryBlock = $probeIf.Parent
    $initialAssignment = $zeroAssignments[0]
    if ($probeIf -isnot
            [Management.Automation.Language.IfStatementAst] -or
        $mainTryBlock -isnot
            [Management.Automation.Language.StatementBlockAst] -or
        -not [object]::ReferenceEquals(
            $initialAssignment.Parent,
            $mainTryBlock) -or
        $initialAssignment.Extent.EndOffset -ge
            $probeIf.Extent.StartOffset) {
        return $false
    }

    $fallbackAssignment = $syntheticFallbackAssignments[0]
    $fallbackIf = $fallbackAssignment.Parent.Parent
    if ($fallbackIf -isnot
            [Management.Automation.Language.IfStatementAst] -or
        -not [object]::ReferenceEquals(
            $fallbackIf.Parent,
            $mainTryBlock) -or
        $fallbackIf.Extent.StartOffset -le
            $probeIf.Extent.EndOffset -or
        $fallbackIf.Clauses.Count -ne 1) {
        return $false
    }

    $fallbackPipeline = $fallbackIf.Clauses[0].Item1
    if ($fallbackPipeline.PipelineElements.Count -ne 1 -or
        $fallbackPipeline.PipelineElements[0] -isnot
            [Management.Automation.Language.CommandExpressionAst]) {
        return $false
    }

    $fallbackCondition =
        $fallbackPipeline.PipelineElements[0].Expression
    if ($fallbackCondition -isnot
            [Management.Automation.Language.BinaryExpressionAst] -or
        $fallbackCondition.Operator -ne
            [Management.Automation.Language.TokenKind]::Ile -or
        $fallbackCondition.Left -isnot
            [Management.Automation.Language.VariableExpressionAst] -or
        $fallbackCondition.Left.VariablePath.UserPath -ne
            "sourceDisplayAspectRatio" -or
        $fallbackCondition.Right -isnot
            [Management.Automation.Language.ConstantExpressionAst] -or
        [int]$fallbackCondition.Right.Value -ne 0) {
        return $false
    }

    $renderingContractAssignments =
        @($mainTryBlock.Statements |
            Where-Object {
                $_ -is
                    [Management.Automation.Language.AssignmentStatementAst] -and
                $_.Left -is
                    [Management.Automation.Language.VariableExpressionAst] -and
                $_.Left.VariablePath.UserPath -eq "renderingContract" -and
                $_.Operator -eq
                    [Management.Automation.Language.TokenKind]::Equals -and
                $_.Right -is
                    [Management.Automation.Language.PipelineAst] -and
                $_.Right.PipelineElements.Count -eq 1 -and
                $_.Right.PipelineElements[0] -is
                    [Management.Automation.Language.CommandAst] -and
                $_.Right.PipelineElements[0].GetCommandName() -eq
                    "Test-RenderedImageContract"
            })
    if ($renderingContractAssignments.Count -ne 1 -or
        $renderingContractAssignments[0].Extent.StartOffset -le
            $fallbackIf.Extent.EndOffset) {
        return $false
    }

    $renderingCommand =
        $renderingContractAssignments[0].Right.PipelineElements[0]
    $sourceParameterIndexes =
        @(for ($index = 0;
            $index -lt $renderingCommand.CommandElements.Count - 1;
            $index++) {
            if ($renderingCommand.CommandElements[$index] -is
                    [Management.Automation.Language.CommandParameterAst] -and
                $renderingCommand.CommandElements[$index].ParameterName -eq
                    "SourceAspectRatio" -and
                $renderingCommand.CommandElements[$index + 1] -is
                    [Management.Automation.Language.VariableExpressionAst] -and
                $renderingCommand.CommandElements[
                    $index + 1].VariablePath.UserPath -eq
                    "sourceDisplayAspectRatio") {
                $index
            }
        })
    return $sourceParameterIndexes.Count -eq 1
}

. (Get-StressFunctionDefinition -Name "Test-RasterBlankLine")
. (Get-StressFunctionDefinition -Name "Get-RasterBlankPaddingStatistics")
. (Get-StressFunctionDefinition -Name "Test-RenderedImageContract")
. (Get-StressFunctionDefinition -Name "Get-IndependentSourceDisplayAspectRatio")

Add-Type -AssemblyName System.Drawing
$correctPath = $null
$paddedPath = $null
$nearWhitePath = $null
$opaquePaddedPath = $null
$thinLinePath = $null
$probeRoot = $null
$correctBitmap = $null
$correctGraphics = $null
$paddedBitmap = $null
$paddedGraphics = $null
$nearWhiteBitmap = $null
$nearWhiteGraphics = $null
$opaquePaddedBitmap = $null
$opaquePaddedGraphics = $null
$thinLineBitmap = $null
$thinLineGraphics = $null
try {
    $correctPath = [IO.Path]::GetTempFileName()
    $paddedPath = [IO.Path]::GetTempFileName()
    $nearWhitePath = [IO.Path]::GetTempFileName()
    $opaquePaddedPath = [IO.Path]::GetTempFileName()
    $thinLinePath = [IO.Path]::GetTempFileName()
    try {
        $correctBitmap = New-Object Drawing.Bitmap 120, 80
        $correctGraphics = [Drawing.Graphics]::FromImage($correctBitmap)
        $correctGraphics.Clear([Drawing.Color]::FromArgb(40, 100, 180))
        $correctBitmap.Save(
            $correctPath,
            [Drawing.Imaging.ImageFormat]::Png)

        $paddedBitmap = New-Object Drawing.Bitmap 120, 120
        $paddedGraphics = [Drawing.Graphics]::FromImage($paddedBitmap)
        $paddedGraphics.Clear([Drawing.Color]::Transparent)
        $paddedGraphics.FillRectangle(
            [Drawing.Brushes]::SteelBlue,
            0,
            20,
            120,
            100)
        $paddedBitmap.SetPixel(
            0,
            0,
            [Drawing.Color]::FromArgb(1, 0, 0, 0))
        $paddedBitmap.Save(
            $paddedPath,
            [Drawing.Imaging.ImageFormat]::Png)

        $nearWhiteBitmap = New-Object Drawing.Bitmap 120, 80
        $nearWhiteGraphics =
            [Drawing.Graphics]::FromImage($nearWhiteBitmap)
        $nearWhiteGraphics.Clear(
            [Drawing.Color]::FromArgb(255, 247, 251, 255))
        $nearWhiteBitmap.Save(
            $nearWhitePath,
            [Drawing.Imaging.ImageFormat]::Png)

        $opaquePaddedBitmap = New-Object Drawing.Bitmap 120, 120
        $opaquePaddedGraphics =
            [Drawing.Graphics]::FromImage($opaquePaddedBitmap)
        $opaquePaddedGraphics.Clear([Drawing.Color]::White)
        $opaquePaddedGraphics.FillRectangle(
            [Drawing.Brushes]::SteelBlue,
            0,
            20,
            120,
            100)
        $opaquePaddedBitmap.Save(
            $opaquePaddedPath,
            [Drawing.Imaging.ImageFormat]::Png)

        $thinLineBitmap = New-Object Drawing.Bitmap 120, 120
        $thinLineGraphics =
            [Drawing.Graphics]::FromImage($thinLineBitmap)
        $thinLineGraphics.Clear([Drawing.Color]::Transparent)
        $thinLineGraphics.FillRectangle(
            [Drawing.Brushes]::Black,
            0,
            0,
            1,
            20)
        $thinLineGraphics.FillRectangle(
            [Drawing.Brushes]::SteelBlue,
            0,
            20,
            120,
            100)
        $thinLineBitmap.Save(
            $thinLinePath,
            [Drawing.Imaging.ImageFormat]::Png)
    }
    finally {
        foreach ($graphics in @(
                $correctGraphics,
                $paddedGraphics,
                $nearWhiteGraphics,
                $opaquePaddedGraphics,
                $thinLineGraphics)) {
            if ($null -ne $graphics) {
                $graphics.Dispose()
            }
        }

        foreach ($bitmap in @(
                $correctBitmap,
                $paddedBitmap,
                $nearWhiteBitmap,
                $opaquePaddedBitmap,
                $thinLineBitmap)) {
            if ($null -ne $bitmap) {
                $bitmap.Dispose()
            }
        }
    }
    $correctPadding =
        Get-RasterBlankPaddingStatistics -Path $correctPath
    $paddedPadding =
        Get-RasterBlankPaddingStatistics -Path $paddedPath
    $nearWhitePadding =
        Get-RasterBlankPaddingStatistics -Path $nearWhitePath
    $opaquePaddedPadding =
        Get-RasterBlankPaddingStatistics -Path $opaquePaddedPath
    $thinLinePadding =
        Get-RasterBlankPaddingStatistics -Path $thinLinePath

    $correctStatistics = [pscustomobject]@{
        Format = "Png"
        PixelWidth = 1240
        PixelHeight = 1753
        AspectRatio = [Math]::Round(1240 / 1753, 6)
        Bytes = 1
    }
    $squareStatistics = [pscustomobject]@{
        Format = "Png"
        PixelWidth = 1240
        PixelHeight = 1240
        AspectRatio = 1.0
        Bytes = 1
    }
    $sourceAspectRatio = 1240 / 1754
    $correctContract = Test-RenderedImageContract `
        -RenderedImageStatistics $correctStatistics `
        -SourceAspectRatio $sourceAspectRatio `
        -RequestedPixelWidth 1240 `
        -BlankPaddingStatistics $correctPadding
    $squareContract = Test-RenderedImageContract `
        -RenderedImageStatistics $squareStatistics `
        -SourceAspectRatio $sourceAspectRatio `
        -RequestedPixelWidth 1240 `
        -BlankPaddingStatistics $correctPadding
    $paddedContract = Test-RenderedImageContract `
        -RenderedImageStatistics $correctStatistics `
        -SourceAspectRatio $sourceAspectRatio `
        -RequestedPixelWidth 1240 `
        -BlankPaddingStatistics $paddedPadding
    $nearWhiteContract = Test-RenderedImageContract `
        -RenderedImageStatistics $correctStatistics `
        -SourceAspectRatio $sourceAspectRatio `
        -RequestedPixelWidth 1240 `
        -BlankPaddingStatistics $nearWhitePadding

    function Export-DrawioSourceImage {
        param(
            [string]$SourcePath,
            [string]$Format,
            [string]$OutputPath
        )

        if ($Format -ne "Svg" -or
            -not (Test-Path -LiteralPath $SourcePath)) {
            throw "Independent ratio probe did not request a valid SVG export."
        }

        [IO.File]::WriteAllText($OutputPath, "<svg />")
    }

    function Get-RenderedImageStatistics {
        param(
            [string]$Format,
            [string]$Path
        )

        if ($Format -ne "Svg" -or
            -not (Test-Path -LiteralPath $Path)) {
            throw "Independent ratio probe did not inspect its SVG output."
        }

        return [pscustomobject]@{
            AspectRatio = 1.75
        }
    }

    $probeRoot = Join-Path `
        ([IO.Path]::GetTempPath()) `
        ("DrawioPpt-rendering-validation-" + [Guid]::NewGuid().ToString("N"))
    [void][IO.Directory]::CreateDirectory($probeRoot)
    $probeSourcePath = Join-Path $probeRoot "source.drawio"
    [IO.File]::WriteAllText($probeSourcePath, "<mxfile />")
    $independentProbeRatio =
        Get-IndependentSourceDisplayAspectRatio `
            -SourcePath $probeSourcePath `
            -TestRoot $probeRoot
    $independentProbeReturnedMeasuredRatio =
        [Math]::Abs($independentProbeRatio - 1.75) -le 0.000001
    $independentProbeCleanedUp =
        -not (Test-Path -LiteralPath (
            Join-Path $probeRoot "source-ratio-probe.svg"))

    $independentProbeFunction = $ast.Find(
        {
            param($node)
            $node -is [Management.Automation.Language.FunctionDefinitionAst] -and
                [string]::Equals(
                    $node.Name,
                    "Get-IndependentSourceDisplayAspectRatio",
                    [StringComparison]::Ordinal)
        },
        $true)
    $independentProbeFunctionCommands =
        if ($null -eq $independentProbeFunction) {
            @()
        }
        else {
            @($independentProbeFunction.Body.FindAll(
                    {
                        param($node)
                        $node -is [Management.Automation.Language.CommandAst]
                    },
                    $true) |
                ForEach-Object { $_.GetCommandName() })
        }
    $sourceAspectAssignments = @($ast.FindAll(
            {
                param($node)
                $node -is [Management.Automation.Language.AssignmentStatementAst] -and
                    $node.Left.Extent.Text -eq
                        '$sourceDisplayAspectRatio'
            },
            $true))
    $mainProbeAssignments =
        @(Get-MainProbeWiringAssignments -ScriptAst $ast)
    $mutationTokens = $null
    $mutationErrors = $null
    $unusedFunctionMutationAst =
        [Management.Automation.Language.Parser]::ParseInput(
            @'
function NeverCalled {
    if ($sourceMode -eq "UserProvided" -or
        $PictureRenderFormat -in @("Png", "Jpeg", "Placeholder")) {
        $sourceDisplayAspectRatio =
            Get-IndependentSourceDisplayAspectRatio
    }
}
$sourceDisplayAspectRatio = 0
'@,
            [ref]$mutationTokens,
            [ref]$mutationErrors)
    $mutationTopLevelProbeAssignments =
        @(Get-MainProbeWiringAssignments `
            -ScriptAst $unusedFunctionMutationAst)
    $unusedFunctionCannotSatisfyMainWiring =
        $mutationErrors.Count -eq 0 -and
        $mutationTopLevelProbeAssignments.Count -eq 0
    $anonymousMutationTokens = $null
    $anonymousMutationErrors = $null
    $anonymousScriptBlockMutationAst =
        [Management.Automation.Language.Parser]::ParseInput(
            @'
$neverInvokedAspectProbe = {
    if ($sourceMode -eq "UserProvided" -or
        $PictureRenderFormat -in @("Png", "Jpeg", "Placeholder")) {
        $sourceDisplayAspectRatio =
            Get-IndependentSourceDisplayAspectRatio
    }
}
$sourceDisplayAspectRatio = 0
'@,
            [ref]$anonymousMutationTokens,
            [ref]$anonymousMutationErrors)
    $anonymousMutationProbeAssignments =
        @(Get-MainProbeWiringAssignments `
            -ScriptAst $anonymousScriptBlockMutationAst)
    $anonymousScriptBlockCannotSatisfyMainWiring =
        $anonymousMutationErrors.Count -eq 0 -and
        $anonymousMutationProbeAssignments.Count -eq 0
    $deadBranchMutationTokens = $null
    $deadBranchMutationErrors = $null
    $deadBranchMutationAst =
        [Management.Automation.Language.Parser]::ParseInput(
            @'
if ($false) {
    if ($sourceMode -eq "UserProvided" -or
        $PictureRenderFormat -in @("Png", "Jpeg", "Placeholder")) {
        $sourceDisplayAspectRatio =
            Get-IndependentSourceDisplayAspectRatio
    }
}
$sourceDisplayAspectRatio = 0
'@,
            [ref]$deadBranchMutationTokens,
            [ref]$deadBranchMutationErrors)
    $deadBranchProbeAssignments =
        @(Get-MainProbeWiringAssignments -ScriptAst $deadBranchMutationAst)
    $deadBranchCannotSatisfyMainWiring =
        $deadBranchMutationErrors.Count -eq 0 -and
        $deadBranchProbeAssignments.Count -eq 0
    $falseConditionMutationTokens = $null
    $falseConditionMutationErrors = $null
    $falseConditionMutationAst =
        [Management.Automation.Language.Parser]::ParseInput(
            @'
try {
    if ($false -and ($sourceMode -eq "UserProvided" -or
        $PictureRenderFormat -in @("Png", "Jpeg", "Placeholder"))) {
        $sourceDisplayAspectRatio =
            Get-IndependentSourceDisplayAspectRatio `
                -SourcePath $sourceCopyPath `
                -TestRoot $testRoot
    }
}
finally {}
'@,
            [ref]$falseConditionMutationTokens,
            [ref]$falseConditionMutationErrors)
    $falseConditionProbeAssignments =
        @(Get-MainProbeWiringAssignments `
            -ScriptAst $falseConditionMutationAst)
    $falseConditionCannotSatisfyMainWiring =
        $falseConditionMutationErrors.Count -eq 0 -and
        $falseConditionProbeAssignments.Count -eq 0
    $stringRhsMutationTokens = $null
    $stringRhsMutationErrors = $null
    $stringRhsMutationAst =
        [Management.Automation.Language.Parser]::ParseInput(
            @'
try {
    if ($sourceMode -eq "UserProvided" -or
        $PictureRenderFormat -in @("Png", "Jpeg", "Placeholder")) {
        $sourceDisplayAspectRatio =
            "Get-IndependentSourceDisplayAspectRatio"
    }
}
finally {}
'@,
            [ref]$stringRhsMutationTokens,
            [ref]$stringRhsMutationErrors)
    $stringRhsProbeAssignments =
        @(Get-MainProbeWiringAssignments -ScriptAst $stringRhsMutationAst)
    $stringRhsCannotSatisfyMainWiring =
        $stringRhsMutationErrors.Count -eq 0 -and
        $stringRhsProbeAssignments.Count -eq 0
    $hiddenCommandMutationTokens = $null
    $hiddenCommandMutationErrors = $null
    $hiddenCommandMutationAst =
        [Management.Automation.Language.Parser]::ParseInput(
            @'
try {
    if ($sourceMode -eq "UserProvided" -or
        $PictureRenderFormat -in @("Png", "Jpeg", "Placeholder")) {
        $sourceDisplayAspectRatio = if ($false) {
            Get-IndependentSourceDisplayAspectRatio `
                -SourcePath $sourceCopyPath `
                -TestRoot $testRoot
        }
    }
}
finally {}
'@,
            [ref]$hiddenCommandMutationTokens,
            [ref]$hiddenCommandMutationErrors)
    $hiddenCommandProbeAssignments =
        @(Get-MainProbeWiringAssignments -ScriptAst $hiddenCommandMutationAst)
    $hiddenCommandCannotSatisfyMainWiring =
        $hiddenCommandMutationErrors.Count -eq 0 -and
        $hiddenCommandProbeAssignments.Count -eq 0
    $duplicateAssignmentMutationTokens = $null
    $duplicateAssignmentMutationErrors = $null
    $duplicateAssignmentMutationAst =
        [Management.Automation.Language.Parser]::ParseInput(
            @'
try {
    if ($sourceMode -eq "UserProvided" -or
        $PictureRenderFormat -in @("Png", "Jpeg", "Placeholder")) {
        $sourceDisplayAspectRatio =
            Get-IndependentSourceDisplayAspectRatio `
                -SourcePath $sourceCopyPath `
                -TestRoot $testRoot
        $sourceDisplayAspectRatio = 0
    }
}
finally {}
'@,
            [ref]$duplicateAssignmentMutationTokens,
            [ref]$duplicateAssignmentMutationErrors)
    $duplicateAssignmentProbeAssignments =
        @(Get-MainProbeWiringAssignments `
            -ScriptAst $duplicateAssignmentMutationAst)
    $duplicateAssignmentCannotSatisfyMainWiring =
        $duplicateAssignmentMutationErrors.Count -eq 0 -and
        $duplicateAssignmentProbeAssignments.Count -eq 0
    $nonEqualsMutationTokens = $null
    $nonEqualsMutationErrors = $null
    $nonEqualsMutationAst =
        [Management.Automation.Language.Parser]::ParseInput(
            @'
try {
    if ($sourceMode -eq "UserProvided" -or
        $PictureRenderFormat -in @("Png", "Jpeg", "Placeholder")) {
        $sourceDisplayAspectRatio -=
            Get-IndependentSourceDisplayAspectRatio `
                -SourcePath $sourceCopyPath `
                -TestRoot $testRoot
    }
}
finally {}
'@,
            [ref]$nonEqualsMutationTokens,
            [ref]$nonEqualsMutationErrors)
    $nonEqualsProbeAssignments =
        @(Get-MainProbeWiringAssignments -ScriptAst $nonEqualsMutationAst)
    $nonEqualsCannotSatisfyMainWiring =
        $nonEqualsMutationErrors.Count -eq 0 -and
        $nonEqualsProbeAssignments.Count -eq 0
    $redirectedCommandMutationTokens = $null
    $redirectedCommandMutationErrors = $null
    $redirectedCommandMutationAst =
        [Management.Automation.Language.Parser]::ParseInput(
            @'
try {
    if ($sourceMode -eq "UserProvided" -or
        $PictureRenderFormat -in @("Png", "Jpeg", "Placeholder")) {
        $sourceDisplayAspectRatio =
            Get-IndependentSourceDisplayAspectRatio `
                -SourcePath $sourceCopyPath `
                -TestRoot $testRoot > $null
    }
}
finally {}
'@,
            [ref]$redirectedCommandMutationTokens,
            [ref]$redirectedCommandMutationErrors)
    $redirectedCommandProbeAssignments =
        @(Get-MainProbeWiringAssignments `
            -ScriptAst $redirectedCommandMutationAst)
    $redirectedCommandCannotSatisfyMainWiring =
        $redirectedCommandMutationErrors.Count -eq 0 -and
        $redirectedCommandProbeAssignments.Count -eq 0
    $postBranchOverwriteTokens = $null
    $postBranchOverwriteErrors = $null
    $postBranchOverwriteAst =
        [Management.Automation.Language.Parser]::ParseInput(
            @'
try {
    $sourceDisplayAspectRatio = 0
    if ($sourceMode -eq "UserProvided" -or
        $PictureRenderFormat -in @("Png", "Jpeg", "Placeholder")) {
        $sourceDisplayAspectRatio =
            Get-IndependentSourceDisplayAspectRatio `
                -SourcePath $sourceCopyPath `
                -TestRoot $testRoot
    }
    $sourceDisplayAspectRatio = 0
    if ($sourceDisplayAspectRatio -le 0) {
        $sourceDisplayAspectRatio = 1200.0 / 800.0
    }
}
finally {}
'@,
            [ref]$postBranchOverwriteTokens,
            [ref]$postBranchOverwriteErrors)
    $postBranchOverwriteCannotSatisfyWriteContract =
        $postBranchOverwriteErrors.Count -eq 0 -and
        -not (Test-SourceAspectWriteContract `
            -ScriptAst $postBranchOverwriteAst)
    $movedInitializationTokens = $null
    $movedInitializationErrors = $null
    $movedInitializationAst =
        [Management.Automation.Language.Parser]::ParseInput(
            @'
try {
    if ($sourceMode -eq "UserProvided" -or
        $PictureRenderFormat -in @("Png", "Jpeg", "Placeholder")) {
        $sourceDisplayAspectRatio =
            Get-IndependentSourceDisplayAspectRatio `
                -SourcePath $sourceCopyPath `
                -TestRoot $testRoot
    }
    $sourceDisplayAspectRatio = 0
    if ($sourceDisplayAspectRatio -le 0) {
        $sourceDisplayAspectRatio = 1200.0 / 800.0
    }
    $renderingContract = Test-RenderedImageContract `
        -RenderedImageStatistics $renderedImageStatistics `
        -SourceAspectRatio $sourceDisplayAspectRatio `
        -RequestedPixelWidth $PreviewPixelWidth `
        -BlankPaddingStatistics $blankPaddingStatistics
}
finally {}
'@,
            [ref]$movedInitializationTokens,
            [ref]$movedInitializationErrors)
    $movedInitializationCannotSatisfyWriteContract =
        $movedInitializationErrors.Count -eq 0 -and
        -not (Test-SourceAspectWriteContract `
            -ScriptAst $movedInitializationAst)
    $unconditionalFallbackTokens = $null
    $unconditionalFallbackErrors = $null
    $unconditionalFallbackAst =
        [Management.Automation.Language.Parser]::ParseInput(
            @'
try {
    $sourceDisplayAspectRatio = 0
    if ($sourceMode -eq "UserProvided" -or
        $PictureRenderFormat -in @("Png", "Jpeg", "Placeholder")) {
        $sourceDisplayAspectRatio =
            Get-IndependentSourceDisplayAspectRatio `
                -SourcePath $sourceCopyPath `
                -TestRoot $testRoot
    }
    $sourceDisplayAspectRatio = 1200.0 / 800.0
    $renderingContract = Test-RenderedImageContract `
        -RenderedImageStatistics $renderedImageStatistics `
        -SourceAspectRatio $sourceDisplayAspectRatio `
        -RequestedPixelWidth $PreviewPixelWidth `
        -BlankPaddingStatistics $blankPaddingStatistics
}
finally {}
'@,
            [ref]$unconditionalFallbackTokens,
            [ref]$unconditionalFallbackErrors)
    $unconditionalFallbackCannotSatisfyWriteContract =
        $unconditionalFallbackErrors.Count -eq 0 -and
        -not (Test-SourceAspectWriteContract `
            -ScriptAst $unconditionalFallbackAst)
    $setVariableMutationTokens = $null
    $setVariableMutationErrors = $null
    $setVariableMutationAst =
        [Management.Automation.Language.Parser]::ParseInput(
            @'
try {
    $sourceDisplayAspectRatio = 0
    if ($sourceMode -eq "UserProvided" -or
        $PictureRenderFormat -in @("Png", "Jpeg", "Placeholder")) {
        $sourceDisplayAspectRatio =
            Get-IndependentSourceDisplayAspectRatio `
                -SourcePath $sourceCopyPath `
                -TestRoot $testRoot
    }
    Set-Variable -Name sourceDisplayAspectRatio -Value 1.0
    if ($sourceDisplayAspectRatio -le 0) {
        $sourceDisplayAspectRatio = 1200.0 / 800.0
    }
    $renderingContract = Test-RenderedImageContract `
        -RenderedImageStatistics $renderedImageStatistics `
        -SourceAspectRatio $sourceDisplayAspectRatio `
        -RequestedPixelWidth $PreviewPixelWidth `
        -BlankPaddingStatistics $blankPaddingStatistics
}
finally {}
'@,
            [ref]$setVariableMutationTokens,
            [ref]$setVariableMutationErrors)
    $setVariableCannotSatisfyWriteContract =
        $setVariableMutationErrors.Count -eq 0 -and
        -not (Test-SourceAspectWriteContract `
            -ScriptAst $setVariableMutationAst)
    $unaryMutationTokens = $null
    $unaryMutationErrors = $null
    $unaryMutationAst =
        [Management.Automation.Language.Parser]::ParseInput(
            @'
try {
    $sourceDisplayAspectRatio = 0
    if ($sourceMode -eq "UserProvided" -or
        $PictureRenderFormat -in @("Png", "Jpeg", "Placeholder")) {
        $sourceDisplayAspectRatio =
            Get-IndependentSourceDisplayAspectRatio `
                -SourcePath $sourceCopyPath `
                -TestRoot $testRoot
    }
    $sourceDisplayAspectRatio++
    if ($sourceDisplayAspectRatio -le 0) {
        $sourceDisplayAspectRatio = 1200.0 / 800.0
    }
    $renderingContract = Test-RenderedImageContract `
        -RenderedImageStatistics $renderedImageStatistics `
        -SourceAspectRatio $sourceDisplayAspectRatio `
        -RequestedPixelWidth $PreviewPixelWidth `
        -BlankPaddingStatistics $blankPaddingStatistics
}
finally {}
'@,
            [ref]$unaryMutationTokens,
            [ref]$unaryMutationErrors)
    $unaryWriteCannotSatisfyWriteContract =
        $unaryMutationErrors.Count -eq 0 -and
        -not (Test-SourceAspectWriteContract `
            -ScriptAst $unaryMutationAst)
    $sourceAspectWriteContractPassed =
        Test-SourceAspectWriteContract -ScriptAst $ast
    $usesIndependentSvgProbe =
        $null -ne $independentProbeFunction -and
        $independentProbeFunctionCommands -contains
            "Export-DrawioSourceImage" -and
        $independentProbeFunctionCommands -contains
            "Get-RenderedImageStatistics" -and
        $mainProbeAssignments.Count -eq 1 -and
        $independentProbeReturnedMeasuredRatio -and
        $independentProbeCleanedUp -and
        $unusedFunctionCannotSatisfyMainWiring -and
        $anonymousScriptBlockCannotSatisfyMainWiring -and
        $deadBranchCannotSatisfyMainWiring -and
        $falseConditionCannotSatisfyMainWiring -and
        $stringRhsCannotSatisfyMainWiring -and
        $hiddenCommandCannotSatisfyMainWiring -and
        $duplicateAssignmentCannotSatisfyMainWiring -and
        $nonEqualsCannotSatisfyMainWiring -and
        $redirectedCommandCannotSatisfyMainWiring -and
        $postBranchOverwriteCannotSatisfyWriteContract -and
        $movedInitializationCannotSatisfyWriteContract -and
        $unconditionalFallbackCannotSatisfyWriteContract -and
        $setVariableCannotSatisfyWriteContract -and
        $unaryWriteCannotSatisfyWriteContract -and
        $sourceAspectWriteContractPassed
    $avoidsRenderedSelfComparison =
        @($sourceAspectAssignments |
            Where-Object {
                $_.Right.Extent.Text -match
                    '\$renderedImageStatistics\s*\.\s*AspectRatio'
            }).Count -eq 0
    $correctImagePassed =
        -not $correctPadding.DetectionSupported -and
        $null -eq $correctPadding.Passed -and
        $null -eq $correctContract.BorderPaddingPassed -and
        $correctContract.Passed
    $squareRejected =
        -not $squareContract.PreservesAspectRatio -and
        -not $squareContract.Passed
    $paddingDetected =
        $paddedPadding.DetectionSupported -and
        $paddedPadding.BorderPixels -gt 0 -and
        -not $paddedContract.BorderPaddingPassed -and
        -not $paddedContract.Passed
    $opaqueRasterMarkedUnsupported =
        -not $opaquePaddedPadding.DetectionSupported -and
        $null -eq $opaquePaddedPadding.Passed -and
        -not $nearWhitePadding.DetectionSupported -and
        $null -eq $nearWhitePadding.Passed -and
        $null -eq $nearWhiteContract.BorderPaddingPassed -and
        $nearWhiteContract.Passed
    $thinLinePreserved =
        $thinLinePadding.DetectionSupported -and
        $thinLinePadding.Top -eq 0 -and
        $thinLinePadding.BorderPixels -eq 0 -and
        $thinLinePadding.Passed

    Write-Host "UsesIndependentSvgProbe=$usesIndependentSvgProbe"
    Write-Host "UnusedFunctionCannotSatisfyMainWiring=$unusedFunctionCannotSatisfyMainWiring"
    Write-Host "AnonymousScriptBlockCannotSatisfyMainWiring=$anonymousScriptBlockCannotSatisfyMainWiring"
    Write-Host "DeadBranchCannotSatisfyMainWiring=$deadBranchCannotSatisfyMainWiring"
    Write-Host "FalseConditionCannotSatisfyMainWiring=$falseConditionCannotSatisfyMainWiring"
    Write-Host "StringRhsCannotSatisfyMainWiring=$stringRhsCannotSatisfyMainWiring"
    Write-Host "HiddenCommandCannotSatisfyMainWiring=$hiddenCommandCannotSatisfyMainWiring"
    Write-Host "DuplicateAssignmentCannotSatisfyMainWiring=$duplicateAssignmentCannotSatisfyMainWiring"
    Write-Host "NonEqualsCannotSatisfyMainWiring=$nonEqualsCannotSatisfyMainWiring"
    Write-Host "RedirectedCommandCannotSatisfyMainWiring=$redirectedCommandCannotSatisfyMainWiring"
    Write-Host "PostBranchOverwriteCannotSatisfyWriteContract=$postBranchOverwriteCannotSatisfyWriteContract"
    Write-Host "MovedInitializationCannotSatisfyWriteContract=$movedInitializationCannotSatisfyWriteContract"
    Write-Host "UnconditionalFallbackCannotSatisfyWriteContract=$unconditionalFallbackCannotSatisfyWriteContract"
    Write-Host "SetVariableCannotSatisfyWriteContract=$setVariableCannotSatisfyWriteContract"
    Write-Host "UnaryWriteCannotSatisfyWriteContract=$unaryWriteCannotSatisfyWriteContract"
    Write-Host "SourceAspectWriteContractPassed=$sourceAspectWriteContractPassed"
    Write-Host "AvoidsRenderedSelfComparison=$avoidsRenderedSelfComparison"
    Write-Host "CorrectImagePassed=$correctImagePassed"
    Write-Host "SquareRejected=$squareRejected"
    Write-Host "DetectedPaddingPixels=$($paddedPadding.BorderPixels)"
    Write-Host "PaddingDetected=$paddingDetected"
    Write-Host "OpaqueRasterMarkedUnsupported=$opaqueRasterMarkedUnsupported"
    Write-Host "ThinLinePreserved=$thinLinePreserved"

    if (-not (
            $usesIndependentSvgProbe -and
            $avoidsRenderedSelfComparison -and
            $correctImagePassed -and
            $squareRejected -and
            $paddingDetected -and
            $opaqueRasterMarkedUnsupported -and
            $thinLinePreserved)) {
        throw "Word UI-thread rendering validation contract failed."
    }

    Write-Host "WORD_UI_THREAD_RENDERING_VALIDATION_TEST_PASS"
}
finally {
    foreach ($path in @(
            $correctPath,
            $paddedPath,
            $nearWhitePath,
            $opaquePaddedPath,
            $thinLinePath)) {
        if ($null -ne $path -and
            (Test-Path -LiteralPath $path)) {
            Remove-Item -LiteralPath $path -Force
        }
    }

    if ($null -ne $probeRoot -and
        (Test-Path -LiteralPath $probeRoot)) {
        Remove-Item -LiteralPath $probeRoot -Recurse -Force
    }
}
