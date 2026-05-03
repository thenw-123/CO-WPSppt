# Speaker notes, optional theme file, and typography pass (best-effort for WPS COM)

function Convert-HexToOfficeRgbLong {
    param([string]$Hex)
    if (-not $Hex) { return $null }
    $h = $Hex.Trim()
    if ($h.StartsWith('#')) { $h = $h.Substring(1) }
    if ($h.Length -ne 6) { return $null }
    try {
        $r = [Convert]::ToInt32($h.Substring(0, 2), 16)
        $g = [Convert]::ToInt32($h.Substring(2, 2), 16)
        $b = [Convert]::ToInt32($h.Substring(4, 2), 16)
        return [int]($b * 65536 + $g * 256 + $r)
    } catch {
        return $null
    }
}

function Set-WpsSlideBackgroundFromHex {
    param(
        $Slide,
        [string]$Hex
    )
    if (-not $Hex) { return }
    $rgb = Convert-HexToOfficeRgbLong -Hex $Hex
    if ($null -eq $rgb) { return }
    try {
        $Slide.FollowMasterBackground = 0
        $fill = $Slide.Background.Fill
        $fill.Visible = -1
        $fill.Solid()
        $fill.ForeColor.RGB = $rgb
    } catch { }
}

function Set-WpsSlideTextColors {
    param(
        $Slide,
        [string]$TitleHex,
        [string]$BodyHex
    )
    $titleRgb = if ($TitleHex) { Convert-HexToOfficeRgbLong -Hex $TitleHex } else { $null }
    $bodyRgb = if ($BodyHex) { Convert-HexToOfficeRgbLong -Hex $BodyHex } else { $null }
    if ($null -eq $titleRgb -and $null -eq $bodyRgb) { return }
    try {
        if ($null -ne $titleRgb) {
            try { $Slide.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = $titleRgb } catch { }
        }
        if ($null -ne $bodyRgb) {
            foreach ($j in 1..[int]$Slide.Shapes.Count) {
                $sh = $Slide.Shapes.Item($j)
                if ($sh.HasTextFrame -ne -1) { continue }
                try {
                    if ($Slide.Shapes.Title -and $sh.Id -eq $Slide.Shapes.Title.Id) { continue }
                } catch { }
                try { $sh.TextFrame.TextRange.Font.Color.RGB = $bodyRgb } catch { }
            }
        }
    } catch { }
}

function Set-WpsSlideNotes {
    param(
        $Slide,
        [string]$Text
    )
    if (-not $Text) { return }
    try {
        $notes = $Slide.NotesPage
        $placed = $false
        try {
            $notes.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = $Text
            $placed = $true
        } catch { }
        if (-not $placed) {
            foreach ($i in 1..[int]$notes.Shapes.Count) {
                try {
                    $sh = $notes.Shapes.Item($i)
                    if ($sh.HasTextFrame -eq -1) {
                        $sh.TextFrame.TextRange.Text = $Text
                        $placed = $true
                        break
                    }
                } catch { }
            }
        }
    } catch { }
}

function Apply-WpsThemeFromPath {
    param(
        $Presentation,
        [string]$ThemePath
    )
    if (-not $ThemePath) { return }
    if (-not (Test-Path -LiteralPath $ThemePath)) { return }
    try {
        $abs = (Resolve-Path -LiteralPath $ThemePath).Path
        $Presentation.ApplyTheme($abs)
    } catch { }
}

function Set-WpsDeckFonts {
    param(
        $Presentation,
        [string]$TitleFont,
        [string]$BodyFont
    )
    if (-not $TitleFont -and -not $BodyFont) { return }
    foreach ($si in 1..[int]$Presentation.Slides.Count) {
        try {
            $slide = $Presentation.Slides.Item($si)
            if ($TitleFont) {
                try { $slide.Shapes.Title.TextFrame.TextRange.Font.Name = $TitleFont } catch { }
                try {
                    foreach ($j in 1..[int]$slide.Shapes.Count) {
                        $sh = $slide.Shapes.Item($j)
                        if ($sh.HasTextFrame -ne -1) { continue }
                        try {
                            if ($slide.Shapes.Title -and $sh.Id -eq $slide.Shapes.Title.Id) { continue }
                        } catch { }
                        $fs = 0.0
                        try { $fs = [double]$sh.TextFrame.TextRange.Font.Size } catch { }
                        if ($fs -ge 22) { $sh.TextFrame.TextRange.Font.Name = $TitleFont }
                    }
                } catch { }
            }
            if ($BodyFont) {
                try {
                    foreach ($j in 1..[int]$slide.Shapes.Count) {
                        $sh = $slide.Shapes.Item($j)
                        if ($sh.HasTextFrame -ne -1) { continue }
                        try {
                            if ($slide.Shapes.Title -and $sh.Id -eq $slide.Shapes.Title.Id) { continue }
                        } catch { }
                        $fs = 0.0
                        try { $fs = [double]$sh.TextFrame.TextRange.Font.Size } catch { }
                        if ($fs -ge 22) { continue }
                        $sh.TextFrame.TextRange.Font.Name = $BodyFont
                    }
                } catch { }
            }
        } catch { }
    }
}

function Set-WpsAgendaSlidePolish {
    param(
        $Slide,
        [double]$TitlePt,
        [double]$BodyPt,
        [string]$AccentBarHex
    )
    if ($null -eq $Slide) { return }
    try {
        if ($TitlePt -gt 0) {
            $Slide.Shapes.Title.TextFrame.TextRange.Font.Size = $TitlePt
            $Slide.Shapes.Title.TextFrame.TextRange.Font.Bold = -1
        }
    } catch { }
    try {
        foreach ($j in 2..8) {
            $sh = $null
            try { $sh = $Slide.Shapes.Placeholders.Item($j) } catch { break }
            if ($null -eq $sh) { break }
            if ($sh.HasTextFrame -ne -1) { continue }
            if ($BodyPt -gt 0) {
                $sh.TextFrame.TextRange.Font.Size = $BodyPt
            }
            try { $sh.TextFrame.TextRange.ParagraphFormat.SpaceAfter = 4 } catch { }
        }
    } catch { }
    if ($AccentBarHex) {
        $rgb = Convert-HexToOfficeRgbLong -Hex $AccentBarHex
        if ($null -ne $rgb) {
            try {
                $bar = $Slide.Shapes.AddShape(1, 26, 112, 5, 352)
                $bar.Line.Visible = 0
                $bar.Fill.Solid()
                $bar.Fill.ForeColor.RGB = $rgb
            } catch { }
        }
    }
}

function Add-WpsAgendaAfterCover {
    param(
        $Presentation,
        [string[]]$SectionTitles,
        [string]$AgendaTitle,
        [string]$BackgroundHex,
        [string]$TitleColorHex,
        [string]$BodyColorHex,
        [bool]$TwoColumn = $true,
        [bool]$Numbered = $true,
        [string]$AccentBarHex,
        [double]$TitlePt = 40,
        [double]$BodyPt = 21
    )
    if (-not $SectionTitles -or $SectionTitles.Count -eq 0) { return }
    $head = if ($AgendaTitle) { $AgendaTitle } else { '目录' }
    $titles = [string[]]@($SectionTitles)
    if ($Numbered) {
        $lab = [System.Collections.ArrayList]@()
        $ix = 1
        foreach ($t in $titles) {
            [void]$lab.Add(('{0:D2}.  {1}' -f $ix, $t))
            $ix++
        }
        $titles = [string[]]@($lab)
    }

    $ag = $null
    $okTwo = $false
    if ($TwoColumn -and $titles.Count -ge 2) {
        try {
            $Presentation.Slides.Add(2, 4) | Out-Null
            $ag = $Presentation.Slides.Item(2)
            $mid = [int][Math]::Ceiling($titles.Count / 2)
            $left = @($titles[0..($mid - 1)])
            $right = @($titles[$mid..($titles.Count - 1)])
            try { $ag.Shapes.Title.TextFrame.TextRange.Text = $head } catch { }
            $ag.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = ($left -join "`r`n")
            $ag.Shapes.Placeholders.Item(3).TextFrame.TextRange.Text = ($right -join "`r`n")
            $okTwo = $true
        } catch {
            $okTwo = $false
            try { $Presentation.Slides.Item(2).Delete() } catch { }
            $ag = $null
        }
    }
    if (-not $okTwo) {
        try {
            $Presentation.Slides.Add(2, 2) | Out-Null
            $ag = $Presentation.Slides.Item(2)
            Set-WpsSlideContent -Slide $ag -Title $head -Subtitle $null -Bullets $titles
        } catch { return }
    }

    if ($BackgroundHex) { Set-WpsSlideBackgroundFromHex -Slide $ag -Hex $BackgroundHex }
    if ($TitleColorHex -or $BodyColorHex) {
        Set-WpsSlideTextColors -Slide $ag -TitleHex $TitleColorHex -BodyHex $BodyColorHex
    }
    Set-WpsAgendaSlidePolish -Slide $ag -TitlePt $TitlePt -BodyPt $BodyPt -AccentBarHex $AccentBarHex
}

function Set-WpsSlideTransitionFromSpec {
    param(
        $Slide,
        [string]$EffectName,
        [double]$Seconds
    )
    if (-not $EffectName) { return }
    $n = $EffectName.Trim().ToLowerInvariant()
    if ($n -eq '' -or $n -eq 'none') { return }
    # PpEntryEffect (common); best-effort for WPS / PowerPoint-compatible COM
    $id = 3849
    switch ($n) {
        'fade' { $id = 3849 }
        'push' { $id = 3850 }
        'wipe' { $id = 3841 }
        'cut' { $id = 257 }
        'uncover' { $id = 3851 }
        default { $id = 3849 }
    }
    try {
        $tr = $Slide.SlideShowTransition
        $tr.EntryEffect = $id
        if ($Seconds -gt 0.05 -and $Seconds -lt 9.9) {
            $tr.Duration = $Seconds
        }
    } catch { }
}

function Set-WpsSlideAnimateBuild {
    param(
        $Slide,
        [bool]$TwoColumn,
        [bool]$Narrative = $false
    )
    try {
        $seq = $Slide.TimeLine.MainSequence
        if ($Narrative) {
            foreach ($j in 1..[int]$Slide.Shapes.Count) {
                try {
                    $sh = $Slide.Shapes.Item($j)
                    if ($sh.HasTextFrame -ne -1) { continue }
                    $null = $seq.AddEffect($sh, 9, 4, 1)
                } catch { }
            }
            return
        }
        if ($TwoColumn) {
            foreach ($j in 2..3) {
                try {
                    $sh = $Slide.Shapes.Placeholders.Item($j)
                    if ($sh.HasTextFrame -eq -1) {
                        # msoAnimEffectFade = 9 ; msoAnimateByParagraph ≈ 4 for Level in some builds
                        $null = $seq.AddEffect($sh, 9, 4, 1)
                    }
                } catch { }
            }
        } else {
            $shp = Get-WpsSlideMainBodyShape -Slide $Slide
            if ($shp) {
                $null = $seq.AddEffect($shp, 9, 4, 1)
            }
        }
    } catch { }
}
