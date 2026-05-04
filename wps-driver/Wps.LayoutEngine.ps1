# Pure layout engine: semantic DSL (no coordinates) -> atomic RenderPlan operations.
# Coordinates come only from layout functions below; the atomic renderer executes ops as-is.

function Get-PptLayoutSlideConstants {
    return [ordered]@{
        width  = 960.0
        height = 540.0
        margin = 40.0
    }
}

function Get-PptLayoutContentRect {
    $S = Get-PptLayoutSlideConstants
    $m = [double]$S.margin
    $w = [double]$S.width
    $h = [double]$S.height
    return [ordered]@{
        x = $m
        y = $m
        w = $w - 2.0 * $m
        h = $h - 2.0 * $m
    }
}

function New-PptLayoutBox {
    param([double]$X, [double]$Y, [double]$W, [double]$H)
    return [ordered]@{ x = $X; y = $Y; w = $W; h = $H }
}

function Measure-PptDslSlideTextLength {
    param($Slide)
    $len = 0
    $title = Get-PptDslTitleFromElements -Slide $Slide
    if ($title) { $len += [string]$title.Length }
    foreach ($el in @(Get-PptDslArray $Slide.elements)) {
        if ($null -eq $el -or -not $el.type) { continue }
        $t = ([string]$el.type).ToLowerInvariant()
        if ($t -eq 'text' -and $el.text) { $len += [string]$el.text.Length }
        if ($t -eq 'bullets') {
            foreach ($item in @(Get-PptDslBulletTextsFromElement -Element $el)) {
                $len += [string]$item.Length
            }
        }
    }
    return [int]$len
}

function Test-PptDslSlideHasImageElement {
    param($Slide)
    return (@(Get-PptDslElementsByType -Slide $Slide -Type 'image').Count -gt 0)
}

function Choose-PptDslAutoLayout {
    param($Slide)
    $kind = Get-PptDslSlideKind -Slide $Slide
    if ($kind -eq 'cover') {
        return [ordered]@{ Name = 'title'; GridColumns = $null; GridGap = $null }
    }
    if ($kind -eq 'section') {
        return [ordered]@{ Name = 'title-content'; GridColumns = $null; GridGap = $null }
    }

    $contentLen = Measure-PptDslSlideTextLength -Slide $Slide
    $hasImage = Test-PptDslSlideHasImageElement -Slide $Slide
    $elementCount = @(Get-PptDslArray $Slide.elements).Count

    if ($contentLen -gt 200) {
        return [ordered]@{ Name = 'title-content'; GridColumns = $null; GridGap = $null }
    }
    if ($hasImage) {
        return [ordered]@{ Name = 'image-right'; GridColumns = $null; GridGap = $null }
    }
    if (($contentLen -gt 100) -and ($elementCount -gt 3)) {
        return [ordered]@{ Name = 'grid'; GridColumns = 3; GridGap = 10.0 }
    }
    if ($elementCount -eq 2) {
        return [ordered]@{ Name = 'comparison'; GridColumns = $null; GridGap = $null }
    }
    return [ordered]@{ Name = 'title-content'; GridColumns = $null; GridGap = $null }
}

function Get-PptDslSlideLayoutView {
    param($Slide)
    $merged = [ordered]@{}
    $src = $Slide.layout
    if ($src) {
        foreach ($pr in @($src.PSObject.Properties)) {
            $merged[$pr.Name] = $pr.Value
        }
    }
    $rawName = ''
    if ($merged['name']) { $rawName = ([string]$merged['name']).ToLowerInvariant().Trim() }
    $autoUsed = $false
    if ($rawName -eq '' -or $rawName -eq 'auto') {
        $pick = Choose-PptDslAutoLayout -Slide $Slide
        $merged['name'] = $pick.Name
        if ($null -ne $pick.GridColumns) {
            if (-not ($merged.Contains('columns') -and $null -ne $merged['columns'])) {
                $merged['columns'] = $pick.GridColumns
            }
        }
        if ($null -ne $pick.GridGap) {
            if (-not ($merged.Contains('gap') -and $null -ne $merged['gap'])) {
                $merged['gap'] = $pick.GridGap
            }
        }
        $autoUsed = $true
    }
    $merged['autoLayoutApplied'] = $autoUsed
    return [pscustomobject]$merged
}

function Get-PptLayoutDoubleOrDefault {
    param($Value, [double]$Default)
    if ($null -eq $Value) { return $Default }
    try { return [double]$Value }
    catch { return $Default }
}

function Get-PptLayoutInsetBox {
    param($Box, [double]$Inset)
    if ($Inset -le 0) { return $Box }
    $ix = [double]$Box.x + $Inset
    $iy = [double]$Box.y + $Inset
    $iw = [double]$Box.w - (2.0 * $Inset)
    $ih = [double]$Box.h - (2.0 * $Inset)
    if ($iw -lt 4.0) { $iw = 4.0 }
    if ($ih -lt 4.0) { $ih = 4.0 }
    return New-PptLayoutBox $ix $iy $iw $ih
}

function Merge-PptCardChromeHashtable {
    param($LowPriority, $HighPriority)
    if ($null -eq $LowPriority -and $null -eq $HighPriority) { return $null }
    $o = [ordered]@{}
    foreach ($src in @($LowPriority, $HighPriority)) {
        if ($null -eq $src) { continue }
        foreach ($p in @('surface', 'background', 'borderColor', 'padding', 'cornerRadius')) {
            $prop = $null
            try { $prop = $src.PSObject.Properties[$p] } catch { $prop = $null }
            if ($null -eq $prop -or $null -eq $prop.Value) { continue }
            $o[$p] = $prop.Value
        }
    }
    if ($o.Count -eq 0) { return $null }
    return $o
}

function Test-PptLayoutCardChromeDrawsSurface {
    param($Merged)

    if ($null -eq $Merged) { return $false }
    $sur = ''
    if ($Merged.surface -ne $null) { $sur = ([string]$Merged.surface).ToLowerInvariant().Trim() }
    if (-not [string]::IsNullOrWhiteSpace($Merged.background)) { return $true }
    if (-not [string]::IsNullOrWhiteSpace($Merged.borderColor)) { return $true }
    if ($sur -and $sur -notin @('none', 'off')) { return $true }
    return $false
}

function Resolve-PptLayoutCardChromeAppearance {
    param($Merged)
    if (-not (Test-PptLayoutCardChromeDrawsSurface -Merged $Merged)) { return $null }
    $sur = ''
    if ($Merged.surface -ne $null) { $sur = ([string]$Merged.surface).ToLowerInvariant().Trim() }
    $bg = if ($Merged.background -ne $null) { ([string]$Merged.background).Trim() } else { '' }
    $bd = if ($Merged.borderColor -ne $null) { ([string]$Merged.borderColor).Trim() } else { '' }

    switch ($sur) {
        'muted' {
            if ([string]::IsNullOrWhiteSpace($bg)) { $bg = '#F1F5F9' }
            if ([string]::IsNullOrWhiteSpace($bd)) { $bd = '#CBD5E1' }
        }
        'elevated' {
            if ([string]::IsNullOrWhiteSpace($bg)) { $bg = '#FFFFFF' }
            if ([string]::IsNullOrWhiteSpace($bd)) { $bd = '#E2E8F0' }
        }
        'accent' {
            if ([string]::IsNullOrWhiteSpace($bg)) { $bg = '#EEF2FF' }
            if ([string]::IsNullOrWhiteSpace($bd)) { $bd = '#818CF8' }
        }
    }

    $pad = Get-PptLayoutDoubleOrDefault -Value $Merged.padding -Default 12.0
    $adj = Get-PptLayoutDoubleOrDefault -Value $Merged.cornerRadius -Default 0.12
    $adj = [math]::Min(0.45, [math]::Max(0.02, $adj))

    return [ordered]@{
        fillHex       = $(if (-not [string]::IsNullOrWhiteSpace($bg)) { $bg } else { $null })
        strokeHex     = $(if (-not [string]::IsNullOrWhiteSpace($bd)) { $bd } else { $null })
        paddingInset  = $pad
        cornerAdj     = $adj
        strokePresent = (-not [string]::IsNullOrWhiteSpace($bd))
    }
}

function New-PptAtomicRoundedRectOp {
    param(
        [string]$ElementIdChrome,
        $OuterBox,
        $FillHex,
        $StrokeHex,
        [bool]$StrokePresent,
        [double]$CornerAdj
    )
    return [ordered]@{
        op              = 'addRoundedRect'
        elementId       = $ElementIdChrome
        role            = 'decorative'
        box             = $OuterBox
        fill            = $FillHex
        stroke          = $StrokeHex
        strokeVisible   = $StrokePresent
        cornerAdjustment = $CornerAdj
    }
}

function Get-PptLayoutEngineKey {
    param(
        $Slide,
        $LayoutView = $null
    )
    if ($null -eq $LayoutView) { $LayoutView = Get-PptDslSlideLayoutView -Slide $Slide }
    $layoutName = ''
    if ($LayoutView -and $LayoutView.name) { $layoutName = ([string]$LayoutView.name).ToLowerInvariant() }
    $slideType = Get-PptDslSlideKind -Slide $Slide
    if ($layoutName -in @('three-column', '3-column', 'three-columns')) { return 'three-column' }
    if ($layoutName -in @('grid', 'card-grid', 'cards')) { return 'grid' }
    if ($layoutName -eq 'comparison') { return 'comparison' }
    if ($layoutName -in @('image-right', 'left-image')) { return 'left-right' }
    if ($slideType -eq 'left-image') { return 'left-right' }
    if ($slideType -eq 'cover' -or $layoutName -eq 'title') { return 'title-only' }
    return 'title-content'
}

function Get-PptLayoutRegionsTitleContent {
    param([bool]$HasSubtitle)
    $c = Get-PptLayoutContentRect
    $cx = [double]$c.x
    $cy = [double]$c.y
    $cw = [double]$c.w
    $ch = [double]$c.h
    $gap = 20.0
    if ($HasSubtitle) {
        $titleH = 56.0
        $subH = 44.0
        $title = New-PptLayoutBox $cx $cy $cw $titleH
        $subtitle = New-PptLayoutBox $cx ($cy + $titleH + ($gap / 2.0)) $cw $subH
        $yContent = $cy + $titleH + $subH + $gap
        $hContent = $ch - $titleH - $subH - ($gap * 1.5)
        if ($hContent -lt 40.0) { $hContent = 40.0 }
        $content = New-PptLayoutBox $cx $yContent $cw $hContent
        return [ordered]@{ title = $title; subtitle = $subtitle; content = $content }
    }
    $titleH = 80.0
    $title = New-PptLayoutBox $cx $cy $cw $titleH
    $yContent = $cy + $titleH + $gap
    $hContent = $ch - $titleH - $gap
    if ($hContent -lt 40.0) { $hContent = 40.0 }
    $content = New-PptLayoutBox $cx $yContent $cw $hContent
    return [ordered]@{ title = $title; subtitle = $null; content = $content }
}

function Get-PptLayoutRegionsLeftRight {
    param([bool]$HasTopTitleStrip)
    $c = Get-PptLayoutContentRect
    $cx = [double]$c.x
    $cy = [double]$c.y
    $cw = [double]$c.w
    $ch = [double]$c.h
    $gap = 20.0
    if ($HasTopTitleStrip) {
        $titleH = 80.0
        $title = New-PptLayoutBox $cx $cy $cw $titleH
        $y2 = $cy + $titleH + $gap
        $h2 = $ch - $titleH - $gap
        if ($h2 -lt 40.0) { $h2 = 40.0 }
        $half = ($cw - $gap) / 2.0
        $left = New-PptLayoutBox $cx $y2 $half $h2
        $right = New-PptLayoutBox ($cx + $half + $gap) $y2 $half $h2
        return [ordered]@{ title = $title; left = $left; right = $right }
    }
    $half = ($cw - $gap) / 2.0
    $left = New-PptLayoutBox $cx $cy $half $ch
    $right = New-PptLayoutBox ($cx + $half + $gap) $cy $half $ch
    return [ordered]@{ title = $null; left = $left; right = $right }
}

function Get-PptLayoutRegionsTitleOnly {
    $c = Get-PptLayoutContentRect
    $cx = [double]$c.x
    $cy = [double]$c.y
    $cw = [double]$c.w
    $ch = [double]$c.h
    $gap = 20.0
    $titleH = [math]::Min(120.0, $ch * 0.28)
    $subH = [math]::Min(64.0, $ch * 0.16)
    $title = New-PptLayoutBox $cx ($cy + 40.0) $cw $titleH
    $subtitle = New-PptLayoutBox $cx ($cy + 40.0 + $titleH + $gap) $cw $subH
    $yBody = $cy + 40.0 + $titleH + $subH + 2.0 * $gap
    $hBody = [math]::Max(40.0, $cy + $ch - $yBody)
    $body = New-PptLayoutBox $cx $yBody $cw $hBody
    return [ordered]@{ title = $title; subtitle = $subtitle; content = $body }
}

function New-PptLayoutGridCells {
    param(
        [double]$X,
        [double]$Y,
        [double]$W,
        [double]$H,
        [int]$Columns,
        [int]$Rows,
        [double]$Gap = 20.0
    )
    $cols = [math]::Max(1, [int]$Columns)
    $rows = [math]::Max(1, [int]$Rows)
    $usableW = $W - ([double]($cols - 1) * $Gap)
    $usableH = $H - ([double]($rows - 1) * $Gap)
    $cellW = if ($usableW -gt 0) { $usableW / $cols } else { $W / $cols }
    $cellH = if ($usableH -gt 0) { $usableH / $rows } else { $H / $rows }
    $cells = [System.Collections.Generic.List[object]]::new()
    for ($r = 0; $r -lt $rows; $r++) {
        for ($c = 0; $c -lt $cols; $c++) {
            $left = $X + $c * ($cellW + $Gap)
            $top = $Y + $r * ($cellH + $Gap)
            $cells.Add((New-PptLayoutBox $left $top $cellW $cellH)) | Out-Null
        }
    }
    return @($cells)
}

function Get-PptLayoutRegionsThreeColumn {
    param([int]$BodyElementCount, [bool]$HasTopTitleStrip, [double]$Gap = 20.0)
    $c = Get-PptLayoutContentRect
    $cx = [double]$c.x
    $cy = [double]$c.y
    $cw = [double]$c.w
    $ch = [double]$c.h
    $gap = $Gap
    $title = $null
    if ($HasTopTitleStrip) {
        $title = New-PptLayoutBox $cx $cy $cw 80.0
        $cy = $cy + 100.0
        $ch = $ch - 100.0
    }
    if ($ch -lt 40.0) { $ch = 40.0 }
    $count = [math]::Max(1, [int]$BodyElementCount)
    $rows = [math]::Ceiling($count / 3.0)
    $cells = New-PptLayoutGridCells -X $cx -Y $cy -W $cw -H $ch -Columns 3 -Rows $rows -Gap $gap
    return [ordered]@{ title = $title; subtitle = $null; cells = $cells }
}

function Get-PptLayoutRegionsGrid {
    param([int]$BodyElementCount, [bool]$HasTopTitleStrip, [int]$Columns, [double]$Gap = 20.0)
    $c = Get-PptLayoutContentRect
    $cx = [double]$c.x
    $cy = [double]$c.y
    $cw = [double]$c.w
    $ch = [double]$c.h
    $gap = $Gap
    $title = $null
    if ($HasTopTitleStrip) {
        $title = New-PptLayoutBox $cx $cy $cw 80.0
        $cy = $cy + 100.0
        $ch = $ch - 100.0
    }
    if ($ch -lt 40.0) { $ch = 40.0 }
    $count = [math]::Max(1, [int]$BodyElementCount)
    $cols = [math]::Min(8, [math]::Max(1, [int]$Columns))
    $rows = [math]::Ceiling($count / [double]$cols)
    $cells = New-PptLayoutGridCells -X $cx -Y $cy -W $cw -H $ch -Columns $cols -Rows $rows -Gap $gap
    return [ordered]@{ title = $title; subtitle = $null; cells = $cells }
}

function Read-PptSlideLayoutPlacementOptions {
    param(
        $Slide,
        $LayoutView = $null
    )
    $cols = $null
    $gap = 20.0
    $slideCardSpec = $null
    $layoutObj = if ($null -ne $LayoutView) { $LayoutView } else { $Slide.layout }
    if ($layoutObj) {
        $gap = Get-PptLayoutDoubleOrDefault -Value $layoutObj.gap -Default 20.0
        if ($null -ne $layoutObj.columns) {
            try { $cols = [int]$layoutObj.columns } catch { $cols = $null }
        }
        if ($layoutObj.card) { $slideCardSpec = $layoutObj.card }
    }
    return [ordered]@{
        columnsPreferred = $cols
        gap              = $gap
        slideCardSpec    = $slideCardSpec
    }
}

function Get-PptLayoutRegionsForSlide {
    param(
        $Slide,
        [string]$Key,
        [int]$BodyElementCount = 1,
        $Placement = $null,
        $LayoutView = $null
    )
    if ($null -eq $Placement) { $Placement = Read-PptSlideLayoutPlacementOptions -Slide $Slide -LayoutView $LayoutView }
    $gap = [double]$Placement.gap

    $sub = Get-PptDslSubtitleFromElements -Slide $Slide
    $hasSub = ($sub -ne '')
    switch ($Key) {
        'three-column' {
            $titleMain = Get-PptDslTitleFromElements -Slide $Slide
            $hasStrip = ($titleMain -ne '')
            return Get-PptLayoutRegionsThreeColumn -BodyElementCount $BodyElementCount -HasTopTitleStrip:$hasStrip -Gap $gap
        }
        'grid' {
            $titleMain = Get-PptDslTitleFromElements -Slide $Slide
            $hasStrip = ($titleMain -ne '')
            $cols = 2
            if ($null -ne $Placement.columnsPreferred) {
                try {
                    $cv = [int]$Placement.columnsPreferred
                    $cols = [math]::Min(8, [math]::Max(1, $cv))
                } catch {
                    $cols = 2
                }
            }
            return Get-PptLayoutRegionsGrid -BodyElementCount $BodyElementCount -HasTopTitleStrip:$hasStrip -Columns $cols -Gap $gap
        }
        'left-right' {
            $titleMain = Get-PptDslTitleFromElements -Slide $Slide
            $hasStrip = ($titleMain -ne '')
            return Get-PptLayoutRegionsLeftRight -HasTopTitleStrip:$hasStrip
        }
        'comparison' {
            $titleMain = Get-PptDslTitleFromElements -Slide $Slide
            $hasStrip = ($titleMain -ne '')
            return Get-PptLayoutRegionsLeftRight -HasTopTitleStrip:$hasStrip
        }
        'title-only' {
            return Get-PptLayoutRegionsTitleOnly
        }
        default {
            return Get-PptLayoutRegionsTitleContent -HasSubtitle:$hasSub
        }
    }
}

function New-PptAtomicOpFromElement {
    param(
        $Element,
        $Box,
        [int]$ElementOrdinal
    )

    $type = ([string]$Element.type).ToLowerInvariant()
    $id = if ($Element.id) { [string]$Element.id } else { '{0}-{1}' -f $type, $ElementOrdinal }
    $role = if ($Element.role) { [string]$Element.role } else { 'body' }

    switch ($type) {
        'text' {
            return [ordered]@{
                op        = 'addText'
                elementId = $id
                role      = $role
                text      = [string]$Element.text
                box       = $Box
                style     = $Element.style
            }
        }
        'bullets' {
            return [ordered]@{
                op        = 'addBullets'
                elementId = $id
                role      = $role
                items     = @(Get-PptDslBulletTextsFromElement -Element $Element)
                box       = $Box
                style     = $Element.style
            }
        }
        'image' {
            return [ordered]@{
                op        = 'addImage'
                elementId = $id
                role      = $role
                source    = $Element.source
                alt       = $Element.alt
                box       = $Box
                style     = $Element.style
            }
        }
        default {
            return [ordered]@{
                op         = 'unsupported'
                elementId  = $id
                role       = $role
                sourceType = $type
                box        = $Box
            }
        }
    }
}

function Convert-PptDslSlideToLaidOutRenderSlide {
    param(
        $Slide,
        [int]$SlideIndex
    )

    $titleText = Get-PptDslTitleFromElements -Slide $Slide
    $subText = Get-PptDslSubtitleFromElements -Slide $Slide
    $bodyElements = [System.Collections.Generic.List[object]]::new()
    foreach ($el in @(Get-PptDslArray $Slide.elements)) {
        if ($null -eq $el -or -not $el.type) { continue }
        $elType = ([string]$el.type).ToLowerInvariant()
        $elRole = if ($el.role) { ([string]$el.role).ToLowerInvariant() } else { 'body' }
        if ($elType -eq 'text' -and $elRole -eq 'title' -and $titleText -and ([string]$el.text) -eq $titleText) { continue }
        if ($elType -eq 'text' -and $elRole -eq 'subtitle' -and $subText -and ([string]$el.text) -eq $subText) { continue }
        $bodyElements.Add($el) | Out-Null
    }

    $layoutView = Get-PptDslSlideLayoutView -Slide $Slide
    $key = Get-PptLayoutEngineKey -Slide $Slide -LayoutView $layoutView
    $placementCached = Read-PptSlideLayoutPlacementOptions -Slide $Slide -LayoutView $layoutView
    $regions = Get-PptLayoutRegionsForSlide -Slide $Slide -Key $key -BodyElementCount $bodyElements.Count -Placement $placementCached -LayoutView $layoutView
    $ops = [System.Collections.Generic.List[object]]::new()
    $ordinal = 0
    $cellIndex = 0
    $comparisonIdx = 0

    if ($titleText -and $regions.title) {
        $ordinal++
        $tOp = New-PptAtomicOpFromElement -Element ([ordered]@{
                id = 'layout-title'
                type = 'text'
                role = 'title'
                text = $titleText
            }) -Box $regions.title -ElementOrdinal $ordinal
        $ops.Add($tOp) | Out-Null
    }

    if ($subText -and $regions.subtitle) {
        $ordinal++
        $sOp = New-PptAtomicOpFromElement -Element ([ordered]@{
                id = 'layout-subtitle'
                type = 'text'
                role = 'subtitle'
                text = $subText
            }) -Box $regions.subtitle -ElementOrdinal $ordinal
        $ops.Add($sOp) | Out-Null
    }

    foreach ($element in @($bodyElements)) {
        $elType = ([string]$element.type).ToLowerInvariant()
        $elRole = if ($element.role) { ([string]$element.role).ToLowerInvariant() } else { 'body' }

        $box = $null
        switch ($key) {
            'three-column' {
                if ($regions.cells -and $regions.cells.Count -gt 0) {
                    $pick = [math]::Min($cellIndex, $regions.cells.Count - 1)
                    $box = $regions.cells[$pick]
                    $cellIndex++
                } else {
                    $box = $regions.content
                }
            }
            'grid' {
                if ($regions.cells -and $regions.cells.Count -gt 0) {
                    $pick = [math]::Min($cellIndex, $regions.cells.Count - 1)
                    $box = $regions.cells[$pick]
                    $cellIndex++
                } else {
                    $box = $regions.content
                }
            }
            'left-right' {
                if ($elType -eq 'image') { $box = $regions.right }
                else { $box = $regions.left }
            }
            'comparison' {
                if ($comparisonIdx % 2 -eq 0) { $box = $regions.left }
                else { $box = $regions.right }
                $comparisonIdx++
            }
            'title-only' {
                if ($elRole -eq 'title') { $box = $regions.title }
                elseif ($elRole -eq 'subtitle') { $box = $regions.subtitle }
                else { $box = $regions.content }
            }
            default {
                if ($elRole -in @('title', 'subtitle')) {
                    if ($elRole -eq 'title' -and $regions.title) { $box = $regions.title }
                    elseif ($elRole -eq 'subtitle' -and $regions.subtitle) { $box = $regions.subtitle }
                    else { $box = $regions.content }
                }
                else { $box = $regions.content }
            }
        }

        $elStyleCard = $null
        if ($element.style -and $element.style.card) {
            $elStyleCard = $element.style.card
        }
        $mergedCard = Merge-PptCardChromeHashtable -LowPriority $placementCached.slideCardSpec -HighPriority $elStyleCard
        $chromeApp = Resolve-PptLayoutCardChromeAppearance -Merged $mergedCard

        $useCardChrome = ($null -ne $chromeApp -and $key -in @('grid', 'three-column', 'comparison'))

        $outerBox = $box
        $innerBox = $outerBox
        if ($useCardChrome) {
            $innerBox = Get-PptLayoutInsetBox -Box $outerBox -Inset ([double]$chromeApp.paddingInset)
        }

        if ($useCardChrome) {
            $ordinal++
            $baseId = if ($element.id) { [string]$element.id } else { 'el-{0}' -f $ordinal }
            $chromeOp = New-PptAtomicRoundedRectOp -ElementIdChrome ('{0}-chrome' -f $baseId) -OuterBox $outerBox -FillHex $chromeApp.fillHex -StrokeHex $chromeApp.strokeHex -StrokePresent $chromeApp.strokePresent -CornerAdj $chromeApp.cornerAdj
            $ops.Add($chromeOp) | Out-Null
        }

        $ordinal++
        $op = New-PptAtomicOpFromElement -Element $element -Box $innerBox -ElementOrdinal $ordinal
        if ($op.op -ne 'unsupported') { $ops.Add($op) | Out-Null }
    }

    $layoutNameOut = if ($layoutView -and $layoutView.name) { ([string]$layoutView.name).ToLowerInvariant() } else { $key }
    $layoutMeta = [ordered]@{
        name   = $layoutNameOut
        engine = 'layout-v2'
        key    = $key
    }
    if ($layoutView.autoLayoutApplied) {
        $layoutMeta['autoLayout'] = $true
    }

    return [ordered]@{
        slideId    = if ($Slide.id) { [string]$Slide.id } else { 'slide-{0:D3}' -f $SlideIndex }
        index      = $SlideIndex
        type       = if ($Slide.type) { [string]$Slide.type } else { [string]$Slide.kind }
        layout     = $layoutMeta
        notes      = $Slide.notes
        operations = @($ops)
    }
}

function Convert-PptDslToAtomicRenderPlan {
    param([object]$Dsl)

    $errors = Test-PptDslObject -Dsl $Dsl
    if ($errors.Count -gt 0) { throw "DSL validation failed: $($errors -join '; ')" }

    $slides = @()
    $idx = 0
    foreach ($slide in @($Dsl.slides)) {
        $idx++
        $slides += Convert-PptDslSlideToLaidOutRenderSlide -Slide $slide -SlideIndex $idx
    }

    $S = Get-PptLayoutSlideConstants
    return [ordered]@{
        renderPlanVersion = '1.0'
        renderer          = 'wps-com-atomic'
        source            = [ordered]@{
            dslVersion = [string]$Dsl.dslVersion
            compiledAt = (Get-Date).ToUniversalTime().ToString('o')
        }
        page              = [ordered]@{ width = [double]$S.width; height = [double]$S.height }
        meta              = $Dsl.meta
        theme             = $Dsl.theme
        slides            = @($slides)
    }
}
