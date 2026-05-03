# action: close — close active presentation (optional save)
param([string]$ArgsJson)
$ErrorActionPreference = 'Stop'
try {
    . "$PSScriptRoot\_bootstrap.ps1"
    $a = Read-OcArgs -ArgsJson $ArgsJson
    $save = $true
    if ($null -ne $a.save) { $save = [bool]$a.save }
    $app = Get-WpsApplication
    $pres = Get-WpsPresentationFromSession -App $app
    Close-WpsPresentation -Pres $pres -Save $save
    Clear-WpsSession
    Write-OcResult $true @{ closed = $true }
} catch {
    $code = Get-OcErrorCode -Message $_.Exception.Message
    Write-OcResult $false $null $_.Exception.Message $_.ScriptStackTrace $code
}
