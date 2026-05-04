# action: render-edit — apply PPT Edit DSL and rebuild only affected PPTX slides.
param([string]$ArgsJson)
$ErrorActionPreference = 'Stop'
try {
    . "$PSScriptRoot\_bootstrap.ps1"
    $ProjectRoot = Get-WpsProjectRoot
    . (Join-Path $ProjectRoot 'wps-driver\Wps.CommonsFetch.ps1')
    . (Join-Path $ProjectRoot 'wps-driver\Wps.RunSpec.ps1')
    . (Join-Path $ProjectRoot 'wps-driver\Wps.ChartRender.ps1')
    . (Join-Path $ProjectRoot 'wps-driver\Wps.ChartCom.ps1')
    . (Join-Path $ProjectRoot 'wps-driver\Wps.UrlAsset.ps1')
    . (Join-Path $ProjectRoot 'wps-driver\Wps.RendererHelpers.ps1')
    . (Join-Path $ProjectRoot 'wps-driver\Wps.ValidateSpec.ps1')
    . (Join-Path $ProjectRoot 'wps-driver\Wps.Dsl.ps1')
    . (Join-Path $ProjectRoot 'wps-driver\Wps.LayoutEngine.ps1')
    . (Join-Path $ProjectRoot 'wps-driver\Wps.AtomicRenderer.ps1')
    . (Join-Path $ProjectRoot 'wps-driver\Wps.Edit.ps1')
    . (Join-Path $ProjectRoot 'wps-driver\Wps.SlideMap.ps1')
    . (Join-Path $ProjectRoot 'wps-driver\Wps.RenderEdit.ps1')

    $a = Read-OcArgs -ArgsJson $ArgsJson
    $edit = $null
    if ($a.editPath) {
        $ep = [string]$a.editPath
        if (-not [System.IO.Path]::IsPathRooted($ep)) { $ep = Join-Path $ProjectRoot $ep }
        if (-not (Test-Path -LiteralPath $ep)) { throw "edit file not found: $ep" }
        $edit = Get-Content -LiteralPath $ep -Raw -Encoding UTF8 | ConvertFrom-Json
    } elseif ($a.edit) {
        if ($a.edit -is [string]) { $edit = $a.edit | ConvertFrom-Json }
        else { $edit = $a.edit }
    } else {
        throw 'Need editPath or edit'
    }

    $dslPath = if ($a.dslPath) { [string]$a.dslPath } elseif ($edit.target -and $edit.target.dslPath) { [string]$edit.target.dslPath } else { $null }
    if (-not $dslPath) { throw 'Need dslPath or edit.target.dslPath' }
    if (-not [System.IO.Path]::IsPathRooted($dslPath)) { $dslPath = Join-Path $ProjectRoot $dslPath }
    if (-not (Test-Path -LiteralPath $dslPath)) { throw "dsl file not found: $dslPath" }

    $presentationPath = if ($a.presentationPath) { [string]$a.presentationPath } else { $null }
    if (-not $presentationPath) { throw 'Need presentationPath for partial render' }
    if (-not [System.IO.Path]::IsPathRooted($presentationPath)) { $presentationPath = Join-Path $ProjectRoot $presentationPath }
    if (-not (Test-Path -LiteralPath $presentationPath)) { throw "presentation file not found: $presentationPath" }

    $mapPath = if ($a.mapPath) { [string]$a.mapPath } else { Get-PptSlideMapPath -PresentationPath $presentationPath }
    if (-not [System.IO.Path]::IsPathRooted($mapPath)) { $mapPath = Join-Path $ProjectRoot $mapPath }

    $dsl = Get-Content -LiteralPath $dslPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $editResult = Invoke-PptDslEdit -Dsl $dsl -Edit $edit -ValidateAfterEdit
    $editResult.dsl | ConvertTo-Json -Depth 100 | Set-Content -LiteralPath $dslPath -Encoding UTF8

    $renderResult = Invoke-WpsRenderEditedSlides -Dsl $editResult.dsl -Edit $edit -PresentationPath $presentationPath `
        -MapPath $mapPath -ProjectRoot $ProjectRoot -DslPath $dslPath

    Write-OcResult $true @{
        dslPath        = (Resolve-Path -LiteralPath $dslPath).Path
        applied        = $editResult.applied
        presentationPath = $renderResult.presentationPath
        slideMapPath   = $renderResult.slideMapPath
        renderedSlides = $renderResult.renderedSlides
    }
} catch {
    $code = Get-OcErrorCode -Message $_.Exception.Message
    Write-OcResult $false $null $_.Exception.Message $_.ScriptStackTrace $code
}
