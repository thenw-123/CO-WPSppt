# action: add-slide — append slide with layout
param([string]$ArgsJson)
$ErrorActionPreference = 'Stop'
try {
    . "$PSScriptRoot\_bootstrap.ps1"
    $a = Read-OcArgs -ArgsJson $ArgsJson
    $layoutName = if ($a.layout) { [string]$a.layout } else { 'title-content' }
    $layout = Map-LayoutNameToInt -Name $layoutName
    $app = Get-WpsApplication
    $pres = Get-WpsPresentationFromSession -App $app
    Add-WpsSlide -Presentation $pres -LayoutInt $layout | Out-Null
    Save-WpsPresentation -Pres $pres
    Write-OcResult $true @{
        slideCount = [int]$pres.Slides.Count
        layout     = $layoutName
    }
} catch {
    $code = Get-OcErrorCode -Message $_.Exception.Message
    Write-OcResult $false $null $_.Exception.Message $_.ScriptStackTrace $code
}
