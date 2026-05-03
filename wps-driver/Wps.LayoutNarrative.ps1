# Narrative layout variants on blank slides: timeline, comparison, thesis-chain, argument,
# thesis-vertical (一页一论题·竖向三步), SWOT 四象限.
# Requires Wps.Polish.ps1 (Convert-HexToOfficeRgbLong) — load after Polish in bootstrap.

function Get-WpsNarrativeMultilineText {
    param($Val)
    if ($null -eq $Val) { return '' }
    if ($Val -is [string]) { return $Val.Trim() }
    if ($Val -is [System.Collections.IEnumerable]) {
        $lines = [System.Collections.ArrayList]@()
        foreach ($x in @($Val)) {
            if ($null -ne $x -and "$x" -ne '') { [void]$lines.Add([string]$x) }
        }
        return ($lines -join "`r`n")
    }
    return ([string]$Val).Trim()
}

function Apply-WpsShapeFillRgb {
    param($Shape, [nullable[int]]$Rgb)
    if ($null -eq $Rgb) { return }
    try {
        $Shape.Fill.Visible = -1
        $Shape.Fill.Solid()
        $Shape.Fill.ForeColor.RGB = $Rgb
    } catch { }
}

function Apply-WpsShapeLineRgb {
    param($Shape, [nullable[int]]$Rgb, [double]$Weight = 1.0)
    if ($null -eq $Rgb) { return }
    try {
        $Shape.Line.Visible = -1
        $Shape.Line.ForeColor.RGB = $Rgb
        $Shape.Line.Weight = $Weight
    } catch { }
}

function Add-WpsNarrativeTextBox {
    param(
        $Slide,
        [double]$Left,
        [double]$Top,
        [double]$Width,
        [double]$Height,
        [string]$Text,
        [double]$FontSize = 14.0,
        [bool]$Bold = $false,
        [int]$Align = 1
    )
    # msoTextOrientationHorizontal = 1
    $tb = $Slide.Shapes.AddTextbox(1, $Left, $Top, $Width, $Height)
    $tr = $tb.TextFrame.TextRange
    $tr.Text = $Text
    $tr.Font.Size = $FontSize
    $tr.Font.Bold = if ($Bold) { -1 } else { 0 }
    try { $tr.ParagraphFormat.Alignment = $Align } catch { }
    try {
        $tb.TextFrame.MarginLeft = 6
        $tb.TextFrame.MarginRight = 6
        $tb.TextFrame.MarginTop = 4
        $tb.TextFrame.MarginBottom = 4
        $tb.TextFrame.WordWrap = -1
        $tb.TextFrame.VerticalAnchor = 1
    } catch { }
    return $tb
}

function Set-WpsNarrativeShapeColors {
    param(
        $Slide,
        [string]$TitleHex,
        [string]$BodyHex,
        [int[]]$TitleShapeIndices,
        [int[]]$BodyShapeIndices
    )
    $titleRgb = if ($TitleHex) { Convert-HexToOfficeRgbLong -Hex $TitleHex } else { $null }
    $bodyRgb = if ($BodyHex) { Convert-HexToOfficeRgbLong -Hex $BodyHex } else { $null }
    foreach ($j in @($TitleShapeIndices)) {
        if ($j -lt 1) { continue }
        try {
            $sh = $Slide.Shapes.Item($j)
            if ($sh.HasTextFrame -eq -1 -and $null -ne $titleRgb) {
                $sh.TextFrame.TextRange.Font.Color.RGB = $titleRgb
            }
        } catch { }
    }
    foreach ($j in @($BodyShapeIndices)) {
        if ($j -lt 1) { continue }
        try {
            $sh = $Slide.Shapes.Item($j)
            if ($sh.HasTextFrame -eq -1 -and $null -ne $bodyRgb) {
                $sh.TextFrame.TextRange.Font.Color.RGB = $bodyRgb
            }
        } catch { }
    }
}

function Set-WpsNarrativeTimeline {
    param(
        $Slide,
        [object]$Spec,
        [double]$SlideWidth,
        [double]$SlideHeight,
        [double]$ReserveRight,
        [double]$MaxContentBottom,
        [string]$TitleColorHex,
        [string]$BodyColorHex
    )
    $items = @($Spec.timeline)
    if ($items.Count -lt 1) { return }
    $n = [Math]::Min(6, $items.Count)
    $leftM = 40.0
    $usableW = $SlideWidth - $leftM - 36.0 - $ReserveRight
    $title = [string]$Spec.title
    if ($title) {
        Add-WpsNarrativeTextBox -Slide $Slide -Left $leftM -Top 24 -Width $usableW -Height 56 -Text $title -FontSize 28 -Bold $true -Align 1 | Out-Null
    }
    $baseY = 118.0
    if ($MaxContentBottom -lt 520) { $baseY = 98.0 }
    $lineY = $baseY + 44.0
    $nodeH = [Math]::Min(88.0, $MaxContentBottom - $lineY - 52.0)
    if ($nodeH -lt 40) { $nodeH = 40 }
    $seg = $usableW / [double]$n
    $nodeW = [Math]::Min(148.0, $seg - 14.0)
    if ($nodeW -lt 72) { $nodeW = 72 }
    $accentRgb = Convert-HexToOfficeRgbLong -Hex '#CBD5E1'
    if (-not $accentRgb) { $accentRgb = 12632256 }
    try {
        $ln = $Slide.Shapes.AddShape(1, $leftM, $lineY, $usableW, 3)
        Apply-WpsShapeFillRgb -Shape $ln -Rgb $accentRgb
        try { $ln.Line.Visible = 0 } catch { }
    } catch { }

    $shapeIdxTitle = 1
    $bodyIdx = [System.Collections.Generic.List[int]]::new()
    if (-not $title) { $shapeIdxTitle = -1 }
    for ($i = 0; $i -lt $n; $i++) {
        $it = $items[$i]
        $mark = [string]$it.mark
        if (-not $mark) { $mark = [string]$it.year }
        $txt = [string]$it.text
        if (-not $txt) { $txt = [string]$it.caption }
        $cx = $leftM + ($i + 0.5) * $seg - $nodeW / 2.0
        try {
            $card = $Slide.Shapes.AddShape(5, $cx, $lineY - 6, $nodeW, 22)
            Apply-WpsShapeFillRgb -Shape $card -Rgb (Convert-HexToOfficeRgbLong -Hex '#E2E8F0')
            try { $card.Line.Visible = 0 } catch { }
            $card.TextFrame.TextRange.Text = $mark
            $card.TextFrame.TextRange.Font.Size = 11
            $card.TextFrame.TextRange.Font.Bold = -1
            $card.TextFrame.VerticalAnchor = 2
            [void]$bodyIdx.Add([int]$Slide.Shapes.Count)
        } catch { }
        $ny = $lineY + 18.0
        Add-WpsNarrativeTextBox -Slide $Slide -Left $cx -Top $ny -Width $nodeW -Height $nodeH -Text $txt -FontSize 12.5 -Bold $false -Align 1 | Out-Null
        [void]$bodyIdx.Add([int]$Slide.Shapes.Count)
    }
    $titleIdxs = if ($shapeIdxTitle -gt 0) { @(1) } else { @() }
    Set-WpsNarrativeShapeColors -Slide $Slide -TitleHex $TitleColorHex -BodyHex $BodyColorHex -TitleShapeIndices $titleIdxs -BodyShapeIndices ($bodyIdx.ToArray())
}

function Set-WpsNarrativeComparison {
    param(
        $Slide,
        [object]$Spec,
        [double]$SlideWidth,
        [double]$SlideHeight,
        [double]$ReserveRight,
        [double]$MaxContentBottom,
        [string]$TitleColorHex,
        [string]$BodyColorHex
    )
    $cmp = $Spec.comparison
    if (-not $cmp) { return }
    $rows = @($cmp.rows)
    if ($rows.Count -lt 1) { return }
    $leftM = 36.0
    $usableW = $SlideWidth - $leftM - 36.0 - $ReserveRight
    $gutter = 14.0
    $half = ($usableW - $gutter) / 2.0
    $title = [string]$Spec.title
    if ($title) {
        Add-WpsNarrativeTextBox -Slide $Slide -Left $leftM -Top 22 -Width $usableW -Height 52 -Text $title -FontSize 26 -Bold $true | Out-Null
    }
    $hdrY = 84.0
    $hdrH = 34.0
    $lh = [string]$cmp.leftHeader
    $rh = [string]$cmp.rightHeader
    $hdrRgb = Convert-HexToOfficeRgbLong -Hex '#334155'
    $cellBg = Convert-HexToOfficeRgbLong -Hex '#F1F5F9'
    try {
        $L = $Slide.Shapes.AddShape(5, $leftM, $hdrY, $half, $hdrH)
        Apply-WpsShapeFillRgb -Shape $L -Rgb $hdrRgb
        $L.TextFrame.TextRange.Text = $lh
        $L.TextFrame.TextRange.Font.Size = 13
        $L.TextFrame.TextRange.Font.Bold = -1
        $L.TextFrame.TextRange.Font.Color.RGB = 16777215
        $L.TextFrame.VerticalAnchor = 2
        $R = $Slide.Shapes.AddShape(5, $leftM + $half + $gutter, $hdrY, $half, $hdrH)
        Apply-WpsShapeFillRgb -Shape $R -Rgb $hdrRgb
        $R.TextFrame.TextRange.Text = $rh
        $R.TextFrame.TextRange.Font.Size = 13
        $R.TextFrame.TextRange.Font.Bold = -1
        $R.TextFrame.TextRange.Font.Color.RGB = 16777215
        $R.TextFrame.VerticalAnchor = 2
    } catch { }

    $y = $hdrY + $hdrH + 8.0
    $maxRows = [Math]::Min(6, $rows.Count)
    $rowH = [Math]::Min(56.0, ($MaxContentBottom - $y - 16.0) / [Math]::Max(1, $maxRows))
    if ($rowH -lt 36) { $rowH = 36 }
    $bodyIdx = [System.Collections.Generic.List[int]]::new()
    for ($r = 0; $r -lt $maxRows; $r++) {
        $row = $rows[$r]
        $lt = [string]$row.left
        $rt = [string]$row.right
        try {
            $b1 = $Slide.Shapes.AddShape(5, $leftM, $y, $half, $rowH)
            $ix1 = [int]$Slide.Shapes.Count
            Apply-WpsShapeFillRgb -Shape $b1 -Rgb $cellBg
            try { $b1.Line.Visible = 0 } catch { }
            $b1.TextFrame.TextRange.Text = $lt
            $b1.TextFrame.TextRange.Font.Size = 12
            $b1.TextFrame.WordWrap = -1
            $b1.TextFrame.MarginLeft = 8
            $b2 = $Slide.Shapes.AddShape(5, $leftM + $half + $gutter, $y, $half, $rowH)
            $ix2 = [int]$Slide.Shapes.Count
            Apply-WpsShapeFillRgb -Shape $b2 -Rgb $cellBg
            try { $b2.Line.Visible = 0 } catch { }
            $b2.TextFrame.TextRange.Text = $rt
            $b2.TextFrame.TextRange.Font.Size = 12
            $b2.TextFrame.WordWrap = -1
            $b2.TextFrame.MarginLeft = 8
            [void]$bodyIdx.Add($ix1)
            [void]$bodyIdx.Add($ix2)
        } catch { }
        $y += $rowH + 6.0
    }
    $titleIdxs = @()
    foreach ($j in 1..[int]$Slide.Shapes.Count) {
        try {
            $sh = $Slide.Shapes.Item($j)
            if ($sh.HasTextFrame -ne -1) { continue }
            if ([double]$sh.TextFrame.TextRange.Font.Size -ge 24) { $titleIdxs += $j }
        } catch { }
    }
    Set-WpsNarrativeShapeColors -Slide $Slide -TitleHex $TitleColorHex -BodyHex $BodyColorHex -TitleShapeIndices $titleIdxs -BodyShapeIndices ($bodyIdx.ToArray())
}

function Set-WpsNarrativeThesisChain {
    param(
        $Slide,
        [object]$Spec,
        [double]$SlideWidth,
        [double]$SlideHeight,
        [double]$ReserveRight,
        [double]$MaxContentBottom,
        [string]$TitleColorHex,
        [string]$BodyColorHex
    )
    $chain = @($Spec.chain)
    if ($chain.Count -lt 2) { return }
    $n = [Math]::Min(4, $chain.Count)
    $leftM = 36.0
    $usableW = $SlideWidth - $leftM - 36.0 - $ReserveRight
    $title = [string]$Spec.title
    if ($title) {
        Add-WpsNarrativeTextBox -Slide $Slide -Left $leftM -Top 22 -Width $usableW -Height 48 -Text $title -FontSize 26 -Bold $true | Out-Null
    }
    $y = 96.0
    $seg = $usableW / ([double]$n + ($n - 1) * 0.28)
    $cardW = $seg
    $arrowW = $seg * 0.28
    if ($cardW -gt 200) { $cardW = 200; $arrowW = 36 }
    $cardH = [Math]::Min(120.0, $MaxContentBottom - $y - 40.0)
    if ($cardH -lt 64) { $cardH = 64 }
    $fillRgb = Convert-HexToOfficeRgbLong -Hex '#EEF2FF'
    $borderRgb = Convert-HexToOfficeRgbLong -Hex '#6366F1'
    $cx = $leftM
    $bodyIdx = [System.Collections.Generic.List[int]]::new()
    for ($i = 0; $i -lt $n; $i++) {
        $it = $chain[$i]
        $lbl = [string]$it.label
        $txt = [string]$it.text
        try {
            $card = $Slide.Shapes.AddShape(5, $cx, $y, $cardW, $cardH)
            Apply-WpsShapeFillRgb -Shape $card -Rgb $fillRgb
            Apply-WpsShapeLineRgb -Shape $card -Rgb $borderRgb -Weight 1.25
            $card.TextFrame.TextRange.Text = ($lbl + "`r`n`r`n" + $txt)
            $card.TextFrame.TextRange.Font.Size = 12
            try {
                $card.TextFrame.TextRange.Paragraphs(1).Font.Bold = -1
                $card.TextFrame.TextRange.Paragraphs(1).Font.Size = 11
            } catch { }
            $card.TextFrame.MarginLeft = 10
            $card.TextFrame.MarginRight = 10
            $card.TextFrame.WordWrap = -1
            [void]$bodyIdx.Add([int]$Slide.Shapes.Count)
        } catch { }
        $cx += $cardW
        if ($i -lt $n - 1) {
            Add-WpsNarrativeTextBox -Slide $Slide -Left $cx -Top ($y + $cardH / 2 - 14) -Width $arrowW -Height 28 -Text '→' -FontSize 20 -Bold $true -Align 2 | Out-Null
            [void]$bodyIdx.Add([int]$Slide.Shapes.Count)
            $cx += $arrowW
        }
    }
    $titleIdxs = @()
    foreach ($j in 1..[int]$Slide.Shapes.Count) {
        try {
            $sh = $Slide.Shapes.Item($j)
            if ($sh.HasTextFrame -ne -1) { continue }
            if ([double]$sh.TextFrame.TextRange.Font.Size -ge 24) { $titleIdxs += $j }
        } catch { }
    }
    Set-WpsNarrativeShapeColors -Slide $Slide -TitleHex $TitleColorHex -BodyHex $BodyColorHex -TitleShapeIndices $titleIdxs -BodyShapeIndices ($bodyIdx.ToArray())
}

function Set-WpsNarrativeArgument {
    param(
        $Slide,
        [object]$Spec,
        [double]$SlideWidth,
        [double]$SlideHeight,
        [double]$ReserveRight,
        [double]$MaxContentBottom,
        [string]$TitleColorHex,
        [string]$BodyColorHex
    )
    $arg = $Spec.argument
    if (-not $arg) { return }
    $thesis = [string]$arg.thesis
    $therefore = [string]$arg.therefore
    $because = @()
    if ($arg.because) { $because = @($arg.because) }
    $leftM = 40.0
    $usableW = $SlideWidth - $leftM - 44.0 - $ReserveRight
    $title = [string]$Spec.title
    $y = 20.0
    if ($title) {
        Add-WpsNarrativeTextBox -Slide $Slide -Left $leftM -Top $y -Width $usableW -Height 44 -Text $title -FontSize 24 -Bold $true | Out-Null
        $y += 50.0
    }
    $titleIdxs = @()
    foreach ($j in 1..[int]$Slide.Shapes.Count) {
        try {
            $sh = $Slide.Shapes.Item($j)
            if ($sh.HasTextFrame -eq -1 -and [double]$sh.TextFrame.TextRange.Font.Size -ge 22) { $titleIdxs += $j }
        } catch { }
    }

    $boxRgb = Convert-HexToOfficeRgbLong -Hex '#1E293B'
    $boxText = 16777215
    $bodyIdx = [System.Collections.Generic.List[int]]::new()
    if ($thesis) {
        try {
            $null = $Slide.Shapes.AddShape(5, $leftM, $y, $usableW, 48)
            $bx = $Slide.Shapes.Item([int]$Slide.Shapes.Count)
            Apply-WpsShapeFillRgb -Shape $bx -Rgb $boxRgb
            try { $bx.Line.Visible = 0 } catch { }
            $bx.TextFrame.TextRange.Text = $thesis
            $bx.TextFrame.TextRange.Font.Size = 14
            $bx.TextFrame.TextRange.Font.Bold = -1
            $bx.TextFrame.TextRange.Font.Color.RGB = $boxText
            $bx.TextFrame.MarginLeft = 12
            $bx.TextFrame.VerticalAnchor = 2
        } catch { }
        $y += 56.0
    }
    $null = Add-WpsNarrativeTextBox -Slide $Slide -Left $leftM -Top $y -Width 120 -Height 22 -Text '因为' -FontSize 13 -Bold $true
    [void]$bodyIdx.Add([int]$Slide.Shapes.Count)
    $y += 24.0
    $bc = [Math]::Min(4, $because.Count)
    $lines = [System.Collections.ArrayList]@()
    for ($b = 0; $b -lt $bc; $b++) {
        $line = [string]$because[$b]
        if ($line) { [void]$lines.Add(('• ' + $line)) }
    }
    $bodyH = [Math]::Max(36.0, 22.0 * $lines.Count + 8.0)
    $cap = $MaxContentBottom - $y - 120.0
    if ($bodyH -gt $cap -and $cap -gt 36) { $bodyH = $cap }
    if ($lines.Count -gt 0) {
        $null = Add-WpsNarrativeTextBox -Slide $Slide -Left ($leftM + 8) -Top $y -Width ($usableW - 16) -Height $bodyH -Text ($lines -join "`r`n") -FontSize 12.5
        [void]$bodyIdx.Add([int]$Slide.Shapes.Count)
        $y += $bodyH + 10.0
    }
    $null = Add-WpsNarrativeTextBox -Slide $Slide -Left $leftM -Top $y -Width 120 -Height 22 -Text '因此' -FontSize 13 -Bold $true
    [void]$bodyIdx.Add([int]$Slide.Shapes.Count)
    $y += 26.0
    if ($therefore) {
        $concRgb = Convert-HexToOfficeRgbLong -Hex '#DCFCE7'
        $concBorder = Convert-HexToOfficeRgbLong -Hex '#16A34A'
        try {
            $hRem = $MaxContentBottom - $y - 12.0
            $ch = [Math]::Min(72.0, [Math]::Max(44.0, $hRem))
            $cz = $Slide.Shapes.AddShape(5, $leftM, $y, $usableW, $ch)
            Apply-WpsShapeFillRgb -Shape $cz -Rgb $concRgb
            Apply-WpsShapeLineRgb -Shape $cz -Rgb $concBorder -Weight 1.5
            $cz.TextFrame.TextRange.Text = $therefore
            $cz.TextFrame.TextRange.Font.Size = 14
            $cz.TextFrame.TextRange.Font.Bold = -1
            $cz.TextFrame.MarginLeft = 12
            $cz.TextFrame.WordWrap = -1
        } catch { }
    }
    Set-WpsNarrativeShapeColors -Slide $Slide -TitleHex $TitleColorHex -BodyHex $BodyColorHex -TitleShapeIndices $titleIdxs -BodyShapeIndices ($bodyIdx.ToArray())
}

function Set-WpsNarrativeThesisVertical {
    param(
        $Slide,
        [object]$Spec,
        [double]$SlideWidth,
        [double]$SlideHeight,
        [double]$ReserveRight,
        [double]$MaxContentBottom,
        [string]$TitleColorHex,
        [string]$BodyColorHex
    )
    $tv = $Spec.thesisVertical
    if (-not $tv) { return }
    $claim = [string]$tv.claim
    if (-not $claim.Trim()) { $claim = [string]$tv.thesis }
    $steps = @($tv.steps)
    if ($steps.Count -lt 1) { return }
    $leftM = 40.0
    $usableW = $SlideWidth - $leftM - 44.0 - $ReserveRight
    $title = [string]$Spec.title
    $y = 18.0
    $titleIdxs = [System.Collections.Generic.List[int]]::new()
    $bodyIdx = [System.Collections.Generic.List[int]]::new()
    if ($title) {
        Add-WpsNarrativeTextBox -Slide $Slide -Left $leftM -Top $y -Width $usableW -Height 40 -Text $title -FontSize 22 -Bold $true | Out-Null
        [void]$titleIdxs.Add([int]$Slide.Shapes.Count)
        $y += 46.0
    }
    $boxRgb = Convert-HexToOfficeRgbLong -Hex '#0F172A'
    if ($claim.Trim()) {
        try {
            $bx = $Slide.Shapes.AddShape(5, $leftM, $y, $usableW, 42)
            Apply-WpsShapeFillRgb -Shape $bx -Rgb $boxRgb
            try { $bx.Line.Visible = 0 } catch { }
            $bx.TextFrame.TextRange.Text = $claim.Trim()
            $bx.TextFrame.TextRange.Font.Size = 13.5
            $bx.TextFrame.TextRange.Font.Bold = -1
            $bx.TextFrame.TextRange.Font.Color.RGB = 16777215
            $bx.TextFrame.MarginLeft = 12
            $bx.TextFrame.VerticalAnchor = 2
            $bx.TextFrame.WordWrap = -1
        } catch { }
        $y += 48.0
    }
    $stepColors = @(
        @{ fill = '#EEF2FF'; border = '#4F46E5' },
        @{ fill = '#ECFEFF'; border = '#0891B2' },
        @{ fill = '#F0FDF4'; border = '#16A34A' }
    )
    $n = [Math]::Min(3, $steps.Count)
    $avail = $MaxContentBottom - $y - 14.0
    $gap = 6.0
    $arrowSpan = 0.0
    if ($n -gt 1) { $arrowSpan = 18.0 * ($n - 1) }
    $stepH = [Math]::Floor(($avail - ($n - 1) * $gap - $arrowSpan) / [Math]::Max(1, $n))
    if ($stepH -lt 52) { $stepH = 52 }
    for ($si = 0; $si -lt $n; $si++) {
        if ($si -gt 0) {
            Add-WpsNarrativeTextBox -Slide $Slide -Left ($leftM + $usableW / 2 - 14) -Top $y -Width 28 -Height 16 -Text '↓' -FontSize 14 -Bold $true -Align 2 | Out-Null
            [void]$bodyIdx.Add([int]$Slide.Shapes.Count)
            $y += 18.0
        }
        $st = $steps[$si]
        $lbl = [string]$st.label
        $txt = Get-WpsNarrativeMultilineText -Val $st.text
        if (-not $txt -and $lbl) { $txt = $lbl; $lbl = '' }
        $pal = $stepColors[[Math]::Min($si, 2)]
        $fillRgb = Convert-HexToOfficeRgbLong -Hex $pal.fill
        $bdRgb = Convert-HexToOfficeRgbLong -Hex $pal.border
        try {
            $card = $Slide.Shapes.AddShape(5, $leftM, $y, $usableW, $stepH)
            Apply-WpsShapeFillRgb -Shape $card -Rgb $fillRgb
            Apply-WpsShapeLineRgb -Shape $card -Rgb $bdRgb -Weight 1.25
            $block = if ($lbl) { ($lbl.Trim() + "`r`n`r`n" + $txt) } else { $txt }
            $card.TextFrame.TextRange.Text = $block
            $card.TextFrame.TextRange.Font.Size = 12
            try {
                $card.TextFrame.TextRange.Paragraphs(1).Font.Bold = -1
                $card.TextFrame.TextRange.Paragraphs(1).Font.Size = 11
            } catch { }
            $card.TextFrame.MarginLeft = 12
            $card.TextFrame.MarginRight = 10
            $card.TextFrame.WordWrap = -1
            [void]$bodyIdx.Add([int]$Slide.Shapes.Count)
        } catch { }
        $y += $stepH + $gap
    }
    Set-WpsNarrativeShapeColors -Slide $Slide -TitleHex $TitleColorHex -BodyHex $BodyColorHex `
        -TitleShapeIndices ($titleIdxs.ToArray()) -BodyShapeIndices ($bodyIdx.ToArray())
}

function Add-WpsSwotQuadrant {
    param(
        $Slide,
        [double]$X,
        [double]$Y,
        [double]$W,
        [double]$H,
        [string]$Header,
        [string]$BodyText,
        [int]$HdrRgb,
        [int]$BodyFillRgb,
        [int]$BodyTextRgb
    )
    $hdrH = 28.0
    try {
        $hdrBar = $Slide.Shapes.AddShape(5, $X, $Y, $W, $hdrH)
        Apply-WpsShapeFillRgb -Shape $hdrBar -Rgb $HdrRgb
        try { $hdrBar.Line.Visible = 0 } catch { }
        $hdrBar.TextFrame.TextRange.Text = $Header
        $hdrBar.TextFrame.TextRange.Font.Size = 11
        $hdrBar.TextFrame.TextRange.Font.Bold = -1
        $hdrBar.TextFrame.TextRange.Font.Color.RGB = 16777215
        $hdrBar.TextFrame.VerticalAnchor = 2
        $hdrBar.TextFrame.MarginLeft = 8
        $bt = $Y + $hdrH + 4.0
        $bh = [Math]::Max(36.0, $H - $hdrH - 8.0)
        $bod = $Slide.Shapes.AddShape(5, $X, $bt, $W, $bh)
        Apply-WpsShapeFillRgb -Shape $bod -Rgb $BodyFillRgb
        Apply-WpsShapeLineRgb -Shape $bod -Rgb (Convert-HexToOfficeRgbLong -Hex '#CBD5E1') -Weight 0.75
        $bod.TextFrame.TextRange.Text = $BodyText
        $bod.TextFrame.TextRange.Font.Size = 11
        $bod.TextFrame.TextRange.Font.Color.RGB = $BodyTextRgb
        $bod.TextFrame.MarginLeft = 8
        $bod.TextFrame.MarginRight = 6
        $bod.TextFrame.WordWrap = -1
    } catch { }
}

function Set-WpsNarrativeSwot {
    param(
        $Slide,
        [object]$Spec,
        [double]$SlideWidth,
        [double]$SlideHeight,
        [double]$ReserveRight,
        [double]$MaxContentBottom,
        [string]$TitleColorHex,
        [string]$BodyColorHex
    )
    $sw = $Spec.swot
    if (-not $sw) { return }
    $hdrs = $sw.headers
    $hS = 'S 优势'
    $hW = 'W 劣势'
    $hO = 'O 机会'
    $hT = 'T 威胁'
    if ($hdrs) {
        if ([string]$hdrs.strengths) { $hS = [string]$hdrs.strengths }
        if ([string]$hdrs.weaknesses) { $hW = [string]$hdrs.weaknesses }
        if ([string]$hdrs.opportunities) { $hO = [string]$hdrs.opportunities }
        if ([string]$hdrs.threats) { $hT = [string]$hdrs.threats }
    }
    $tS = Get-WpsNarrativeMultilineText -Val $sw.strengths
    $tW = Get-WpsNarrativeMultilineText -Val $sw.weaknesses
    $tO = Get-WpsNarrativeMultilineText -Val $sw.opportunities
    $tT = Get-WpsNarrativeMultilineText -Val $sw.threats
    $leftM = 36.0
    $top0 = 20.0
    $usableW = $SlideWidth - $leftM - 36.0 - $ReserveRight
    $title = [string]$Spec.title
    $y = $top0
    $titleIdxs = [System.Collections.Generic.List[int]]::new()
    if ($title) {
        Add-WpsNarrativeTextBox -Slide $Slide -Left $leftM -Top $y -Width $usableW -Height 44 -Text $title -FontSize 24 -Bold $true | Out-Null
        [void]$titleIdxs.Add([int]$Slide.Shapes.Count)
        $y += 50.0
    }
    $gutter = 12.0
    $gridTop = $y
    $gridBottom = $MaxContentBottom - 12.0
    $gridH = $gridBottom - $gridTop
    $halfW = ($usableW - $gutter) / 2.0
    $halfH = ($gridH - $gutter) / 2.0
    if ($halfH -lt 56) { $halfH = [Math]::Max(48.0, ($gridH - $gutter) / 2.0) }
    $bodyRgb = Convert-HexToOfficeRgbLong -Hex $BodyColorHex
    if (-not $bodyRgb) { $bodyRgb = Convert-HexToOfficeRgbLong -Hex '#334155' }
    $fills = @(
        @{ hdr = '#15803D'; body = '#DCFCE7' },
        @{ hdr = '#BE123C'; body = '#FFE4E6' },
        @{ hdr = '#1D4ED8'; body = '#DBEAFE' },
        @{ hdr = '#CA8A04'; body = '#FEF9C3' }
    )
    $hdrRgb0 = Convert-HexToOfficeRgbLong -Hex $fills[0].hdr
    $hdrRgb1 = Convert-HexToOfficeRgbLong -Hex $fills[1].hdr
    $hdrRgb2 = Convert-HexToOfficeRgbLong -Hex $fills[2].hdr
    $hdrRgb3 = Convert-HexToOfficeRgbLong -Hex $fills[3].hdr
    $bf0 = Convert-HexToOfficeRgbLong -Hex $fills[0].body
    $bf1 = Convert-HexToOfficeRgbLong -Hex $fills[1].body
    $bf2 = Convert-HexToOfficeRgbLong -Hex $fills[2].body
    $bf3 = Convert-HexToOfficeRgbLong -Hex $fills[3].body
    $xL = $leftM
    $xR = $leftM + $halfW + $gutter
    $yT = $gridTop
    $yB = $gridTop + $halfH + $gutter
    $null = Add-WpsSwotQuadrant -Slide $Slide -X $xL -Y $yT -W $halfW -H $halfH -Header $hS -BodyText $tS -HdrRgb $hdrRgb0 -BodyFillRgb $bf0 -BodyTextRgb $bodyRgb
    $null = Add-WpsSwotQuadrant -Slide $Slide -X $xR -Y $yT -W $halfW -H $halfH -Header $hW -BodyText $tW -HdrRgb $hdrRgb1 -BodyFillRgb $bf1 -BodyTextRgb $bodyRgb
    $null = Add-WpsSwotQuadrant -Slide $Slide -X $xL -Y $yB -W $halfW -H $halfH -Header $hO -BodyText $tO -HdrRgb $hdrRgb2 -BodyFillRgb $bf2 -BodyTextRgb $bodyRgb
    $null = Add-WpsSwotQuadrant -Slide $Slide -X $xR -Y $yB -W $halfW -H $halfH -Header $hT -BodyText $tT -HdrRgb $hdrRgb3 -BodyFillRgb $bf3 -BodyTextRgb $bodyRgb
    $bodyIdx = [System.Collections.Generic.List[int]]::new()
    foreach ($j in 1..[int]$Slide.Shapes.Count) {
        try {
            $sh = $Slide.Shapes.Item($j)
            if ($sh.HasTextFrame -ne -1) { continue }
            $fs = 0.0
            try { $fs = [double]$sh.TextFrame.TextRange.Font.Size } catch { }
            if ($fs -ge 20) { continue }
            $rgb = 0
            try { $rgb = [int]$sh.TextFrame.TextRange.Font.Color.RGB } catch { }
            if ($rgb -eq 16777215) { continue }
            [void]$bodyIdx.Add($j)
        } catch { }
    }
    Set-WpsNarrativeShapeColors -Slide $Slide -TitleHex $TitleColorHex -BodyHex $BodyColorHex `
        -TitleShapeIndices ($titleIdxs.ToArray()) -BodyShapeIndices ($bodyIdx.ToArray())
}

function Build-WpsNarrativeSlide {
    param(
        $Slide,
        [object]$Spec,
        [double]$SlideWidth = 960.0,
        [double]$SlideHeight = 540.0,
        [double]$ReserveRight = 0.0,
        [double]$MaxContentBottom = 540.0,
        [string]$TitleColorHex,
        [string]$BodyColorHex
    )
    $kind = if ($Spec.layout) { [string]$Spec.layout } else { '' }
    switch ($kind.ToLowerInvariant()) {
        'timeline' {
            Set-WpsNarrativeTimeline -Slide $Slide -Spec $Spec -SlideWidth $SlideWidth -SlideHeight $SlideHeight `
                -ReserveRight $ReserveRight -MaxContentBottom $MaxContentBottom -TitleColorHex $TitleColorHex -BodyColorHex $BodyColorHex
            break
        }
        'comparison' {
            Set-WpsNarrativeComparison -Slide $Slide -Spec $Spec -SlideWidth $SlideWidth -SlideHeight $SlideHeight `
                -ReserveRight $ReserveRight -MaxContentBottom $MaxContentBottom -TitleColorHex $TitleColorHex -BodyColorHex $BodyColorHex
            break
        }
        'thesis-chain' {
            Set-WpsNarrativeThesisChain -Slide $Slide -Spec $Spec -SlideWidth $SlideWidth -SlideHeight $SlideHeight `
                -ReserveRight $ReserveRight -MaxContentBottom $MaxContentBottom -TitleColorHex $TitleColorHex -BodyColorHex $BodyColorHex
            break
        }
        'argument' {
            Set-WpsNarrativeArgument -Slide $Slide -Spec $Spec -SlideWidth $SlideWidth -SlideHeight $SlideHeight `
                -ReserveRight $ReserveRight -MaxContentBottom $MaxContentBottom -TitleColorHex $TitleColorHex -BodyColorHex $BodyColorHex
            break
        }
        'thesis-vertical' {
            Set-WpsNarrativeThesisVertical -Slide $Slide -Spec $Spec -SlideWidth $SlideWidth -SlideHeight $SlideHeight `
                -ReserveRight $ReserveRight -MaxContentBottom $MaxContentBottom -TitleColorHex $TitleColorHex -BodyColorHex $BodyColorHex
            break
        }
        'swot' {
            Set-WpsNarrativeSwot -Slide $Slide -Spec $Spec -SlideWidth $SlideWidth -SlideHeight $SlideHeight `
                -ReserveRight $ReserveRight -MaxContentBottom $MaxContentBottom -TitleColorHex $TitleColorHex -BodyColorHex $BodyColorHex
            break
        }
        default { }
    }
}
