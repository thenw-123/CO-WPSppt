# action: set-text — title / subtitle / bullets on slide index (1-based)
param([string]$ArgsJson)
$ErrorActionPreference = 'Stop'
try {
    . "$PSScriptRoot\_bootstrap.ps1"
    $a = Read-OcArgs -ArgsJson $ArgsJson
    $idx = [int]$a.slide
    if ($idx -lt 1) { throw 'slide must be >= 1' }
    $app = Get-WpsApplication
    $pres = Get-WpsPresentationFromSession -App $app
    if ($idx -gt $pres.Slides.Count) { throw "slide $idx out of range (max $($pres.Slides.Count))" }
    $slide = $pres.Slides.Item($idx)
    $bullets = @()
    if ($a.bullets) { $bullets = @($a.bullets) }
    Set-WpsSlideContent -Slide $slide -Title ([string]$a.title) -Subtitle ([string]$a.subtitle) -Bullets $bullets
    Save-WpsPresentation -Pres $pres
    Write-OcResult $true @{ slide = $idx }
} catch {
    $code = Get-OcErrorCode -Message $_.Exception.Message
    Write-OcResult $false $null $_.Exception.Message $_.ScriptStackTrace $code
}
