# action: run-spec — build deck from specs/example-spec.json or inline spec object
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

    $skipVal = $false
    if ($null -ne $a.skipValidation) {
        try { $skipVal = [bool]$a.skipValidation } catch { $skipVal = $false }
    }
    if (-not $skipVal) {
        $verr = Test-PptDeckSpecObject -Spec $spec
        if ($verr.Count -gt 0) {
            $errLines = @($verr | ForEach-Object { [string]$_ })
            Write-OcResult $false @{ errors = $errLines } ($errLines -join '; ') $null 'VALIDATION_FAILED'
            exit 0
        }
    }

    $result = Invoke-WpsDeckFromSpec -Spec $spec -ProjectRoot $ProjectRoot
    Write-OcResult $true $result
} catch {
    $code = Get-OcErrorCode -Message $_.Exception.Message
    Write-OcResult $false $null $_.Exception.Message $_.ScriptStackTrace $code
}
