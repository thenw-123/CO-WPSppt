# action: new — create blank presentation and save to output
param([string]$ArgsJson)
$ErrorActionPreference = 'Stop'
try {
    . "$PSScriptRoot\_bootstrap.ps1"
    $ProjectRoot = Get-WpsProjectRoot
    $a = Read-OcArgs -ArgsJson $ArgsJson
    Ensure-Dir (Join-Path $ProjectRoot 'output')
    $path = $a.outPath
    if (-not $path) {
        $path = Join-Path $ProjectRoot ('output\deck-' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.pptx')
    } elseif (-not [System.IO.Path]::IsPathRooted([string]$path)) {
        $path = Join-Path $ProjectRoot ([string]$path)
    }
    $app = Get-WpsApplication
    $pres = New-WpsPresentation -App $app -SavePath $path
    Write-OcResult $true @{
        presentationPath = (Resolve-Path -LiteralPath $path).Path
        slideCount         = [int]$pres.Slides.Count
    }
} catch {
    $code = Get-OcErrorCode -Message $_.Exception.Message
    Write-OcResult $false $null $_.Exception.Message $_.ScriptStackTrace $code
}
