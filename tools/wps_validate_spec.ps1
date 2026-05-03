# action: validate-spec — check spec JSON structure before run-spec (no COM)
param([string]$ArgsJson)
$ErrorActionPreference = 'Stop'
try {
    $toolsDir = $PSScriptRoot
    $ProjectRoot = (Resolve-Path (Join-Path $toolsDir '..')).Path
    . (Join-Path $ProjectRoot 'wps-driver\Wps.Common.ps1')
    . (Join-Path $ProjectRoot 'wps-driver\Wps.ValidateSpec.ps1')

    $a = Read-OcArgs -ArgsJson $ArgsJson
    $spec = $null
    if ($a.specPath) {
        $p = [string]$a.specPath
        if (-not [System.IO.Path]::IsPathRooted($p)) {
            $p = Join-Path $ProjectRoot $p
        }
        if (-not (Test-Path -LiteralPath $p)) { throw "spec file not found: $p" }
        $spec = Get-Content -LiteralPath $p -Raw -Encoding UTF8 | ConvertFrom-Json
    } elseif ($a.spec) {
        if ($a.spec -is [string]) {
            $spec = $a.spec | ConvertFrom-Json
        } else {
            $spec = $a.spec
        }
    } else {
        throw 'Need specPath (file) or spec (JSON object / string)'
    }

    $verr = Test-PptDeckSpecObject -Spec $spec
    if ($verr.Count -gt 0) {
        $errLines = @($verr | ForEach-Object { [string]$_ })
        Write-OcResult $false @{ errors = $errLines; slideCount = @($spec.slides).Count } ($errLines -join '; ') $null 'VALIDATION_FAILED'
        exit 0
    }
    Write-OcResult $true @{
        valid      = $true
        slideCount = @($spec.slides).Count
    }
} catch {
    $code = Get-OcErrorCode -Message $_.Exception.Message
    Write-OcResult $false $null $_.Exception.Message $_.ScriptStackTrace $code
}
