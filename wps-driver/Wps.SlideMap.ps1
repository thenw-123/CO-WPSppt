# Sidecar slide map for partial edit rendering.

function Get-PptSlideMapPath {
    param([string]$PresentationPath)
    if (-not $PresentationPath) { throw 'PresentationPath is required for slide map' }
    return "$PresentationPath.map.json"
}

function New-PptSlideMap {
    param(
        [object]$Dsl,
        [string]$PresentationPath,
        [string]$DslPath = $null
    )

    $slides = @()
    $idx = 0
    foreach ($s in @($Dsl.slides)) {
        $idx++
        $sid = if ($s.id) { [string]$s.id } else { 'slide-{0:D3}' -f $idx }
        $title = ''
        try { $title = Get-PptDslTitleFromElements -Slide $s } catch { }
        $slides += [ordered]@{
            slideId = $sid
            index   = $idx
            title   = $title
        }
    }

    return [ordered]@{
        mapVersion       = '1.0'
        presentationPath = $PresentationPath
        dslPath          = $DslPath
        updatedAt        = (Get-Date).ToUniversalTime().ToString('o')
        slides           = $slides
    }
}

function Save-PptSlideMap {
    param(
        [object]$Dsl,
        [string]$PresentationPath,
        [string]$DslPath = $null,
        [string]$MapPath = $null
    )

    if (-not $MapPath) { $MapPath = Get-PptSlideMapPath -PresentationPath $PresentationPath }
    $dir = Split-Path -Parent $MapPath
    if ($dir) { Ensure-Dir $dir }
    $map = New-PptSlideMap -Dsl $Dsl -PresentationPath $PresentationPath -DslPath $DslPath
    $map | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $MapPath -Encoding UTF8
    return $map
}

function Read-PptSlideMap {
    param([string]$MapPath)
    if (-not (Test-Path -LiteralPath $MapPath)) { throw "slide map not found: $MapPath" }
    return Get-Content -LiteralPath $MapPath -Raw -Encoding UTF8 | ConvertFrom-Json
}

function Get-PptSlideMapEntry {
    param(
        [object]$SlideMap,
        [string]$SlideId
    )
    foreach ($s in @($SlideMap.slides)) {
        if ([string]$s.slideId -eq $SlideId) { return $s }
    }
    return $null
}

function Get-PptAffectedSlideIdsFromEdit {
    param([object]$Edit)
    $ids = [System.Collections.Generic.List[string]]::new()
    foreach ($op in @($Edit.operations)) {
        $sid = [string]$op.target.slideId
        if ($sid -and -not $ids.Contains($sid)) { [void]$ids.Add($sid) }
    }
    return @($ids)
}
