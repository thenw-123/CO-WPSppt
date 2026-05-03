# action: save — optional path (SaveAs)
param([string]$ArgsJson)
$ErrorActionPreference = 'Stop'
try {
    . "$PSScriptRoot\_bootstrap.ps1"
    $ProjectRoot = Get-WpsProjectRoot
    $a = Read-OcArgs -ArgsJson $ArgsJson
    $app = Get-WpsApplication
    $pres = Get-WpsPresentationFromSession -App $app
    $path = $a.path
    if ($path) {
        if (-not [System.IO.Path]::IsPathRooted([string]$path)) {
            $path = Join-Path $ProjectRoot ([string]$path)
        }
        $dir = Split-Path -Parent $path
        if ($dir) { Ensure-Dir $dir }
        Save-WpsPresentation -Pres $pres -Path $path
    } else {
        Save-WpsPresentation -Pres $pres
    }
    $sess = Read-WpsSession
    Write-OcResult $true @{ presentationPath = [string]$sess.presentationPath }
} catch {
    $code = Get-OcErrorCode -Message $_.Exception.Message
    Write-OcResult $false $null $_.Exception.Message $_.ScriptStackTrace $code
}
