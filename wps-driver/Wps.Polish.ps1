# Speaker notes, optional theme file, and typography pass (best-effort for WPS COM)

function Format-WpsComException {
    param($Err)
    if ($null -eq $Err) { return 'unknown error' }
    $ex = $Err
    if ($Err -is [System.Management.Automation.ErrorRecord]) {
        $ex = $Err.Exception
    }
    $parts = [System.Collections.Generic.List[string]]::new()
    if ($ex.Message) { [void]$parts.Add($ex.Message.Trim()) }
    if ($ex.InnerException -and $ex.InnerException.Message) {
        [void]$parts.Add(('Inner: ' + $ex.InnerException.Message.Trim()))
    }
    try {
        $hrProp = $ex.GetType().GetProperty('HResult')
        if ($hrProp) {
            $h = $hrProp.GetValue($ex)
            if ($null -ne $h) { [void]$parts.Add(('HResult=0x{0:X8}' -f [int]$h)) }
        }
    } catch { }
    if ($parts.Count -eq 0) { return ($ex.ToString()) }
    return ($parts -join ' | ')
}

function Invoke-WpsComAnimationCapabilityProbe {
    <#
    Creates a temporary unsaved presentation (does not touch wps-session.json),
    probes SlideShowTransition + TimeLine.MainSequence + AddEffect, then closes.
    #>
    param($App)
    $pres = $null
    $out = [ordered]@{
        slideShowTransition   = @{ ok = $false; detail = $null; error = $null }
        timeLineMainSequence  = @{ ok = $false; detail = $null; error = $null }
        mainSequenceAddEffect = @{ ok = $false; detail = $null; error = $null }
    }
    try {
        try {
            $pres = $App.Presentations.Add()
        } catch {
            $pres = $App.Presentations.Add(-1)
        }
        while ([int]$pres.Slides.Count -lt 1) {
            $null = $pres.Slides.Add(1, 1)
        }
        $slide = $pres.Slides.Item(1)

        try {
            $tr = $slide.SlideShowTransition
            $tr.EntryEffect = 3849
            $tr.Duration = 0.55
            $readBack = [int]$tr.EntryEffect
            $out.slideShowTransition.ok = $true
            $out.slideShowTransition.detail = @{
                entryEffectSet = 3849
                entryEffectReadBack = $readBack
                duration          = [double]$tr.Duration
            }
        } catch {
            $out.slideShowTransition.error = (Format-WpsComException $_)
        }

        try {
            $seq = $slide.TimeLine.MainSequence
            $out.timeLineMainSequence.ok = $true
            $out.timeLineMainSequence.detail = @{ effectCount = [int]$seq.Count }
        } catch {
            $out.timeLineMainSequence.error = (Format-WpsComException $_)
        }

        try {
            $sh = $null
            try { $sh = $slide.Shapes.Title } catch { }
            if (-not $sh) { $sh = $slide.Shapes.Item(1) }
            $seq = $slide.TimeLine.MainSequence
            $null = $seq.AddEffect($sh, 9, 4, 1)
            $out.mainSequenceAddEffect.ok = $true
            $out.mainSequenceAddEffect.detail = @{ effectCountAfter = [int]$seq.Count }
        } catch {
            $out.mainSequenceAddEffect.error = (Format-WpsComException $_)
        }
    } catch {
        $out.probeSetupError = (Format-WpsComException $_)
    } finally {
        if ($null -ne $pres) {
            try {
                try { $pres.Saved = -1 } catch { }
                $pres.Close()
            } catch {
                $out.probeCloseError = (Format-WpsComException $_)
            }
        }
    }
    return $out
}

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
        [double]$Seconds,
        [int]$SlideNumber = 0,
        [System.Collections.Generic.List[string]]$ComWarnings = $null
    )
    if (-not $EffectName) { return }
    $n = $EffectName.Trim().ToLowerInvariant()
    if ($n -eq '' -or $n -eq 'none') { return }
    # PpEntryEffect (Office-compatible); WPS may ignore some IDs — still best-effort
    $id = 3849
    switch ($n) {
        'fade' { $id = 3849 }
        'push' { $id = 3850 }
        'wipe' { $id = 3841 }
        'cut' { $id = 257 }
        'uncover' { $id = 3851 }
        'split' { $id = 3585 }       # ppEffectSplitHorizontalOut
        'cover' { $id = 1281 }       # ppEffectCoverLeft
        'random' { $id = 513 }       # ppEffectRandom
        'blinds' { $id = 769 }       # ppEffectBlindsHorizontal
        'dissolve' { $id = 1537 }    # ppEffectDissolve (when supported)
        default { $id = 3849 }
    }
    $lab = if ($SlideNumber -gt 0) { "slide[$SlideNumber]" } else { 'slide' }
    try {
        $tr = $Slide.SlideShowTransition
        $tr.EntryEffect = $id
        if ($Seconds -gt 0.05 -and $Seconds -lt 9.9) {
            $tr.Duration = $Seconds
        }
    } catch {
        $msg = (Format-WpsComException $_)
        $line = "${lab} SlideShowTransition (effect=$n, EntryEffect=$id): $msg"
        if ($null -ne $ComWarnings) { [void]$ComWarnings.Add($line) }
        else { Write-Warning $line }
    }
}

function Resolve-WpsAnimEffectId {
    param([string]$Name)
    if (-not $Name) { return 9 }
    switch ($Name.Trim().ToLowerInvariant()) {
        'fade' { return 9 }        # msoAnimEffectFade
        'appear' { return 1 }      # msoAnimEffectAppear
        'fly' { return 7 }         # msoAnimEffectFly
        'float' { return 5 }       # msoAnimEffectFloat
        'wipe' { return 31 }       # msoAnimEffectWipe
        'zoom' { return 16 }       # msoAnimEffectZoom
        default { return 9 }
    }
}

function Set-WpsSlideAnimateBuild {
    param(
        $Slide,
        [bool]$TwoColumn,
        [bool]$Narrative = $false,
        [string]$EffectName = 'fade',
        [int]$SlideNumber = 0,
        [System.Collections.Generic.List[string]]$ComWarnings = $null
    )
    $effectId = Resolve-WpsAnimEffectId -Name $EffectName
    $lab = if ($SlideNumber -gt 0) { "slide[$SlideNumber]" } else { 'slide' }
    function Add-AnimWarning {
        param([string]$Text)
        if ($null -ne $ComWarnings) { [void]$ComWarnings.Add($Text) }
        else { Write-Warning $Text }
    }
    try {
        $seq = $Slide.TimeLine.MainSequence
        if ($Narrative) {
            # COM shape index order != visual reading order; sort top→left for stable builds
            $list = [System.Collections.Generic.List[object]]::new()
            foreach ($j in 1..[int]$Slide.Shapes.Count) {
                try {
                    $sh = $Slide.Shapes.Item($j)
                    if ($sh.HasTextFrame -ne -1) { continue }
                    $list.Add($sh)
                } catch {
                    Add-AnimWarning ("${lab} animateBuild: enumerate shape #${j}: " + (Format-WpsComException $_))
                }
            }
            $sorted = @($list) | Sort-Object { [double]$_.Top }, { [double]$_.Left }
            $si = 0
            foreach ($sh in $sorted) {
                $si++
                try {
                    $null = $seq.AddEffect($sh, $effectId, 4, 1)
                } catch {
                    Add-AnimWarning ("${lab} animateBuild AddEffect #$si (effectId=$effectId): " + (Format-WpsComException $_))
                }
            }
            return
        }
        if ($TwoColumn) {
            foreach ($j in 2..3) {
                try {
                    $sh = $Slide.Shapes.Placeholders.Item($j)
                    if ($sh.HasTextFrame -eq -1) {
                        $null = $seq.AddEffect($sh, $effectId, 4, 1)
                    }
                } catch {
                    Add-AnimWarning ("${lab} animateBuild two-column placeholder ${j}: " + (Format-WpsComException $_))
                }
            }
        } else {
            $shp = Get-WpsSlideMainBodyShape -Slide $Slide
            if ($shp) {
                try {
                    $null = $seq.AddEffect($shp, $effectId, 4, 1)
                } catch {
                    Add-AnimWarning ("${lab} animateBuild body shape: " + (Format-WpsComException $_))
                }
            }
        }
    } catch {
        Add-AnimWarning ("${lab} TimeLine.MainSequence (animateBuild, effect=$EffectName): " + (Format-WpsComException $_))
    }
}
