# action: compile-dsl — compile high-level PPT DSL to RenderPlan or legacy spec (no COM)
param([string]$ArgsJson)
$ErrorActionPreference = 'Stop'
try {
    $toolsDir = $PSScriptRoot
    $ProjectRoot = (Resolve-Path (Join-Path $toolsDir '..')).Path
    . (Join-Path $ProjectRoot 'wps-driver\Wps.Common.ps1')
    . (Join-Path $ProjectRoot 'wps-driver\Wps.Dsl.ps1')
    . (Join-Path $ProjectRoot 'wps-driver\Wps.LayoutEngine.ps1')
    . (Join-Path $ProjectRoot 'wps-driver\Wps.RenderPlan.ps1')

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

    $format = if ($a.format) { ([string]$a.format).ToLowerInvariant() } else { 'renderplan' }
    if ($format -eq 'legacyspec' -or $format -eq 'legacy-spec' -or $format -eq 'legacy') {
        Write-OcResult $true @{
            format = 'legacySpec'
            spec   = (Convert-PptDslToLegacySpec -Dsl $dsl)
        }
    } else {
        Write-OcResult $true @{
            format     = 'renderPlan'
            renderPlan = (Convert-PptDslToRenderPlan -Dsl $dsl)
        }
    }
} catch {
    $code = Get-OcErrorCode -Message $_.Exception.Message
    Write-OcResult $false $null $_.Exception.Message $_.ScriptStackTrace $code
}
