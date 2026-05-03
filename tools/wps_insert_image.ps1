# action: insert-image — picture on slide (1-based)
param([string]$ArgsJson)
$ErrorActionPreference = 'Stop'
try {
    . "$PSScriptRoot\_bootstrap.ps1"
    $ProjectRoot = Get-WpsProjectRoot
    $a = Read-OcArgs -ArgsJson $ArgsJson
    $idx = [int]$a.slide
    $img = [string]$a.path
    if ($idx -lt 1) { throw 'slide must be >= 1' }
    if (-not $img) { throw 'Missing required arg: path (image file)' }
    if (-not [System.IO.Path]::IsPathRooted($img)) {
        $img = Join-Path $ProjectRoot $img
    }
    $app = Get-WpsApplication
    $pres = Get-WpsPresentationFromSession -App $app
    if ($idx -gt $pres.Slides.Count) { throw "slide $idx out of range" }
    $slide = $pres.Slides.Item($idx)
    $left = if ($null -ne $a.left) { [double]$a.left } else { 80 }
    $top = if ($null -ne $a.top) { [double]$a.top } else { 220 }
    $w = if ($null -ne $a.width) { [double]$a.width } else { 480 }
    $h = if ($null -ne $a.height) { [double]$a.height } else { 280 }
    Add-WpsSlidePicture -Slide $slide -ImagePath $img -Left $left -Top $top -Width $w -Height $h | Out-Null
    Save-WpsPresentation -Pres $pres
    Write-OcResult $true @{ slide = $idx; image = (Resolve-Path -LiteralPath $img).Path }
} catch {
    $code = Get-OcErrorCode -Message $_.Exception.Message
    Write-OcResult $false $null $_.Exception.Message $_.ScriptStackTrace $code
}
