# action: run-dsl — compile PPT DSL to RenderPlan, then render via WPS COM
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
    . (Join-Path $ProjectRoot 'wps-driver\Wps.RenderPlan.ps1')
    . (Join-Path $ProjectRoot 'wps-driver\Wps.SlideMap.ps1')

    $a = Read-OcArgs -ArgsJson $ArgsJson
    $dsl = $null
    if ($a.dslPath) {
        $p = [string]$a.dslPath
        if (-not [System.IO.Path]::IsPathRooted($p)) {
            $p = Join-Path $ProjectRoot $p
        }
        if (-not (Test-Path -LiteralPath $p)) { throw "dsl file not found: $p" }
        $dsl = Get-Content -LiteralPath $p -Raw -Encoding UTF8 | ConvertFrom-Json
    } elseif ($a.dsl) {
        if ($a.dsl -is [string]) {
            $dsl = $a.dsl | ConvertFrom-Json
        } else {
            $dsl = $a.dsl
        }
    } else {
        throw 'Need dslPath (file) or dsl (JSON object / string)'
    }

    $verr = Test-PptDslObject -Dsl $dsl
    if ($verr.Count -gt 0) {
        $errLines = @($verr | ForEach-Object { [string]$_ })
        Write-OcResult $false @{ errors = $errLines } ($errLines -join '; ') $null 'VALIDATION_FAILED'
        exit 0
    }

    $renderPlan = Convert-PptDslToRenderPlan -Dsl $dsl

    $result = Invoke-WpsDeckFromRenderPlan -RenderPlan $renderPlan -ProjectRoot $ProjectRoot
    $result.renderPlanVersion = $renderPlan.renderPlanVersion
    $sourceDslPath = $null
    if ($a.dslPath) {
        $sourceDslPath = [string]$a.dslPath
        if (-not [System.IO.Path]::IsPathRooted($sourceDslPath)) {
            $sourceDslPath = Join-Path $ProjectRoot $sourceDslPath
        }
    }
    if ($result.presentationPath) {
        $map = Save-PptSlideMap -Dsl $dsl -PresentationPath ([string]$result.presentationPath) -DslPath $sourceDslPath
        $result.slideMapPath = Get-PptSlideMapPath -PresentationPath ([string]$result.presentationPath)
        $result.slideMap = $map
    }
    Write-OcResult $true $result
} catch {
    $code = Get-OcErrorCode -Message $_.Exception.Message
    Write-OcResult $false $null $_.Exception.Message $_.ScriptStackTrace $code
}
