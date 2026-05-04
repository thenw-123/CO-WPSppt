# action: doctor — environment checks (paths, session, optional live COM)
param([string]$ArgsJson)
$ErrorActionPreference = 'Stop'
try {
    $toolsDir = $PSScriptRoot
    $ProjectRoot = (Resolve-Path (Join-Path $toolsDir '..')).Path
    . (Join-Path $ProjectRoot 'wps-driver\Wps.Common.ps1')
    . (Join-Path $ProjectRoot 'wps-driver\Wps.ChartRender.ps1')

    $a = Read-OcArgs -ArgsJson $ArgsJson
    $comProbe = $false
    if ($null -ne $a.comProbe) {
        try { $comProbe = [bool]$a.comProbe } catch { $comProbe = $false }
    }

    $schemaPath = Join-Path $ProjectRoot 'specs\ppt-spec.schema.json'
    $outDir = Join-Path $ProjectRoot 'output'
    Ensure-Dir $outDir

    $outputWritable = $false
    $testFile = Join-Path $outDir '.wps-ppt-write-test'
    try {
        'ok' | Set-Content -LiteralPath $testFile -Encoding UTF8
        Remove-Item -LiteralPath $testFile -Force -ErrorAction Stop
        $outputWritable = $true
    } catch {
        $outputWritable = $false
    }

    $session = Read-WpsSession
    $progIds = @(
        'Kwpp.Application',
        'kwpp.Application',
        'KWPP.Application',
        'wpp.Application',
        'KWPS.Application'
    )
    $comRegistry = [ordered]@{}
    foreach ($id in $progIds) {
        try {
            $t = [Type]::GetTypeFromProgID($id, $false)
            $comRegistry[$id] = [bool]$t
        } catch {
            $comRegistry[$id] = $false
        }
    }

    $pyCmd = Get-ChartPythonCommand
    $chartPng = [ordered]@{
        pythonAvailable     = [bool]$pyCmd
        matplotlibImportOk  = $false
    }
    if ($pyCmd) {
        $tail = if ($pyCmd.Length -gt 1) { @($pyCmd[1..($pyCmd.Length - 1)]) } else { @() }
        $null = & $pyCmd[0] @($tail + @('-c', 'import matplotlib')) 2>&1
        if ($LASTEXITCODE -eq 0) { $chartPng.matplotlibImportOk = $true }
    }

    $data = [ordered]@{
        projectRoot    = $ProjectRoot
        schemaPath     = $schemaPath
        schemaPresent  = (Test-Path -LiteralPath $schemaPath)
        outputDir      = $outDir
        outputWritable = $outputWritable
        session        = if ($session) { @{ presentationPath = [string]$session.presentationPath } } else { $null }
        comProgIdRegistered = $comRegistry
        comLive        = $null
        comLiveError   = $null
        comAnimationProbe = $null
        comAnimationProbeError = $null
        chartPng       = $chartPng
        assetFetchInsecureHttpAllowed = ($env:OC_ASSET_FETCH_ALLOW_HTTP -eq '1')
    }

    if ($comProbe) {
        . (Join-Path $ProjectRoot 'wps-driver\Wps.Connect.ps1')
        . (Join-Path $ProjectRoot 'wps-driver\Wps.Polish.ps1')
        $app = $null
        try {
            $app = Get-WpsApplication
            $data.comLive = $true
        } catch {
            $data.comLive = $false
            $data.comLiveError = $_.Exception.Message
        }
        if ($null -ne $app) {
            try {
                $data.comAnimationProbe = (Invoke-WpsComAnimationCapabilityProbe -App $app)
            } catch {
                $data.comAnimationProbeError = $_.Exception.Message
            }
        }
    }

    Write-OcResult $true $data
} catch {
    $code = Get-OcErrorCode -Message $_.Exception.Message
    Write-OcResult $false $null $_.Exception.Message $_.ScriptStackTrace $code
}
