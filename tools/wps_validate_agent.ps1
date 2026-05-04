# action: validate-agent — validate two-stage Agent outputs
# Planner: outlinePath/outline. Writer: dslPath/dsl.
param([string]$ArgsJson)
$ErrorActionPreference = 'Stop'
try {
    $toolsDir = $PSScriptRoot
    $ProjectRoot = (Resolve-Path (Join-Path $toolsDir '..')).Path
    . (Join-Path $ProjectRoot 'wps-driver\Wps.Common.ps1')
    . (Join-Path $ProjectRoot 'wps-driver\Wps.Dsl.ps1')
    . (Join-Path $ProjectRoot 'wps-driver\Wps.AgentValidation.ps1')

    $a = Read-OcArgs -ArgsJson $ArgsJson
    $stage = if ($a.stage) { ([string]$a.stage).Trim().ToLowerInvariant() } else { '' }
    if ($stage -notin @('planner', 'writer')) {
        throw 'Need stage: planner or writer'
    }

    $value = $null
    if ($stage -eq 'planner') {
        if ($a.outlinePath) {
            $p = [string]$a.outlinePath
            if (-not [System.IO.Path]::IsPathRooted($p)) { $p = Join-Path $ProjectRoot $p }
            if (-not (Test-Path -LiteralPath $p)) { throw "outline file not found: $p" }
            $value = Get-Content -LiteralPath $p -Raw -Encoding UTF8 | ConvertFrom-Json
        } elseif ($a.outline) {
            if ($a.outline -is [string]) { $value = $a.outline | ConvertFrom-Json }
            else { $value = $a.outline }
        } else {
            throw 'Need outlinePath or outline for planner stage'
        }
    } else {
        if ($a.dslPath) {
            $p = [string]$a.dslPath
            if (-not [System.IO.Path]::IsPathRooted($p)) { $p = Join-Path $ProjectRoot $p }
            if (-not (Test-Path -LiteralPath $p)) { throw "dsl file not found: $p" }
            $value = Get-Content -LiteralPath $p -Raw -Encoding UTF8 | ConvertFrom-Json
        } elseif ($a.dsl) {
            if ($a.dsl -is [string]) { $value = $a.dsl | ConvertFrom-Json }
            else { $value = $a.dsl }
        } else {
            throw 'Need dslPath or dsl for writer stage'
        }
    }

    $errors = Test-PptAgentStageOutput -Stage $stage -Value $value -ProjectRoot $ProjectRoot
    if ($errors.Count -gt 0) {
        $errLines = @($errors | ForEach-Object { [string]$_ })
        Write-OcResult $false @{ stage = $stage; errors = $errLines } ($errLines -join '; ') $null 'VALIDATION_FAILED'
        exit 0
    }

    Write-OcResult $true @{
        valid = $true
        stage = $stage
    }
} catch {
    $code = Get-OcErrorCode -Message $_.Exception.Message
    Write-OcResult $false $null $_.Exception.Message $_.ScriptStackTrace $code
}
