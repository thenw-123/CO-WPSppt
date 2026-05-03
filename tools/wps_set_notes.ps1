# action: set-notes — speaker notes on slide (1-based)
param([string]$ArgsJson)
$ErrorActionPreference = 'Stop'
try {
    . "$PSScriptRoot\_bootstrap.ps1"
    $a = Read-OcArgs -ArgsJson $ArgsJson
    $idx = [int]$a.slide
    $notes = [string]$a.notes
    if ($idx -lt 1) { throw 'slide must be >= 1' }
    if ($null -eq $a.notes) { throw 'Missing required arg: notes' }
    $app = Get-WpsApplication
    $pres = Get-WpsPresentationFromSession -App $app
    if ($idx -gt $pres.Slides.Count) { throw "slide $idx out of range" }
    $slide = $pres.Slides.Item($idx)
    Set-WpsSlideNotes -Slide $slide -Text $notes
    Save-WpsPresentation -Pres $pres
    Write-OcResult $true @{ slide = $idx }
} catch {
    $code = Get-OcErrorCode -Message $_.Exception.Message
    Write-OcResult $false $null $_.Exception.Message $_.ScriptStackTrace $code
}
