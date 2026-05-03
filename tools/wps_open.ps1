# action: open — open existing pptx and bind session
param([string]$ArgsJson)
$ErrorActionPreference = 'Stop'
try {
    . "$PSScriptRoot\_bootstrap.ps1"
    $ProjectRoot = Get-WpsProjectRoot
    $a = Read-OcArgs -ArgsJson $ArgsJson
    $path = [string]$a.path
    if (-not $path) { throw 'Missing required arg: path' }
    if (-not [System.IO.Path]::IsPathRooted($path)) {
        $path = Join-Path $ProjectRoot $path
    }
    $app = Get-WpsApplication
    $pres = Open-WpsPresentation -App $app -Path $path
    Write-OcResult $true @{
        presentationPath = (Resolve-Path -LiteralPath $path).Path
        slideCount         = [int]$pres.Slides.Count
    }
} catch {
    $code = Get-OcErrorCode -Message $_.Exception.Message
    Write-OcResult $false $null $_.Exception.Message $_.ScriptStackTrace $code
}
