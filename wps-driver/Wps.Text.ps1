# Title, subtitle, and body text on a slide

function Get-WpsSlideMainBodyShape {
    param($Slide)
    try {
        foreach ($idx in 1..[int]$Slide.Shapes.Count) {
            $sh = $Slide.Shapes.Item($idx)
            if ($sh.HasTextFrame -ne -1) { continue }
            try {
                if ($Slide.Shapes.Title -and $sh.Id -eq $Slide.Shapes.Title.Id) { continue }
            } catch { }
            return $sh
        }
    } catch { }
    return $null
}

function Set-WpsSlideBodyReservedRight {
    param(
        $Slide,
        [double]$SlideWidth,
        [double]$ReserveRight,
        [double]$LeftMargin = 36
    )
    $shp = Get-WpsSlideMainBodyShape -Slide $Slide
    if ($null -eq $shp) { return }
    try {
        $left = [double]$shp.Left
        if ($left -lt $LeftMargin) { $left = $LeftMargin }
        $newW = $SlideWidth - $ReserveRight - $left
        if ($newW -lt 160) { return }
        $shp.Width = $newW
    } catch { }
}

function Set-WpsSlideTwoColumnContent {
    param(
        $Slide,
        [string]$Title,
        [string]$LeftHeading,
        [string[]]$LeftBullets,
        [string]$RightHeading,
        [string[]]$RightBullets
    )
    if ($Title) {
        try {
            $Slide.Shapes.Title.TextFrame.TextRange.Text = $Title
        } catch {
            try { $Slide.Shapes.Placeholders.Item(1).TextFrame.TextRange.Text = $Title } catch { }
        }
    }
    $leftLines = [System.Collections.ArrayList]@()
    if ($LeftHeading) {
        [void]$leftLines.Add($LeftHeading)
        [void]$leftLines.Add('')
    }
    foreach ($b in @($LeftBullets)) {
        if ($null -ne $b -and "$b" -ne '') { [void]$leftLines.Add([string]$b) }
    }
    $rightLines = [System.Collections.ArrayList]@()
    if ($RightHeading) {
        [void]$rightLines.Add($RightHeading)
        [void]$rightLines.Add('')
    }
    foreach ($b in @($RightBullets)) {
        if ($null -ne $b -and "$b" -ne '') { [void]$rightLines.Add([string]$b) }
    }
    try {
        $Slide.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = ($leftLines -join "`r`n")
    } catch { }
    try {
        $Slide.Shapes.Placeholders.Item(3).TextFrame.TextRange.Text = ($rightLines -join "`r`n")
    } catch { }
}

function Set-WpsSlideTwoColumnBodiesBottomLimit {
    param(
        $Slide,
        [double]$MaxBottom
    )
    foreach ($j in 2..3) {
        try {
            $shp = $Slide.Shapes.Placeholders.Item($j)
            if ($shp.HasTextFrame -ne -1) { continue }
            $top = [double]$shp.Top
            $bt = $top + [double]$shp.Height
            if ($bt -gt $MaxBottom) {
                $shp.Height = [Math]::Max(60, $MaxBottom - $top)
            }
        } catch { break }
    }
}

function Set-WpsSlideTwoColumnBodiesReservedRight {
    param(
        $Slide,
        [double]$SlideWidth,
        [double]$ReserveRight,
        [double]$Gutter = 20.0
    )
    $usable = $SlideWidth - $ReserveRight - 48
    $half = [Math]::Floor(($usable - $Gutter) / 2)
    if ($half -lt 120) { return }
    try {
        $L = $Slide.Shapes.Placeholders.Item(2)
        $L.Left = 36
        $L.Width = $half
    } catch { }
    try {
        $R = $Slide.Shapes.Placeholders.Item(3)
        $R.Left = 36 + $half + $Gutter
        $R.Width = $half
    } catch { }
}

function Set-WpsSlideBodyBottomLimit {
    param(
        $Slide,
        [double]$MaxBottom
    )
    $shp = Get-WpsSlideMainBodyShape -Slide $Slide
    if ($null -eq $shp) { return }
    try {
        $top = [double]$shp.Top
        $bt = $top + [double]$shp.Height
        if ($bt -gt $MaxBottom) {
            $nh = [Math]::Max(72, $MaxBottom - $top)
            $shp.Height = $nh
        }
    } catch { }
}

function Set-WpsSlideContent {
    param(
        $Slide,
        [string]$Title,
        [string]$Subtitle,
        [string[]]$Bullets
    )
    if ($Title) {
        try {
            $Slide.Shapes.Title.TextFrame.TextRange.Text = $Title
        } catch {
            try {
                $Slide.Shapes.Placeholders.Item(1).TextFrame.TextRange.Text = $Title
            } catch { }
        }
    }
    if ($Subtitle) {
        try {
            $Slide.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = $Subtitle
        } catch {
            try {
                foreach ($i in 2..[Math]::Min(5, $Slide.Shapes.Count)) {
                    $sh = $Slide.Shapes.Item($i)
                    if ($sh.HasTextFrame -eq -1) {
                        $sh.TextFrame.TextRange.Text = $Subtitle
                        break
                    }
                }
            } catch { }
        }
    }
    if ($Bullets -and $Bullets.Count -gt 0) {
        $body = $null
        try {
            foreach ($idx in 1..$Slide.Shapes.Count) {
                $sh = $Slide.Shapes.Item($idx)
                if ($sh.HasTextFrame -ne -1) { continue }
                try {
                    if ($Slide.Shapes.Title -and $sh.Id -eq $Slide.Shapes.Title.Id) { continue }
                } catch { }
                $body = $sh
                break
            }
        } catch { }
        if ($null -eq $body) {
            try {
                foreach ($idx in 1..$Slide.Shapes.Count) {
                    $sh = $Slide.Shapes.Item($idx)
                    if ($sh.HasTextFrame -eq -1) { $body = $sh; break }
                }
            } catch { }
        }
        if ($body) {
            $body.TextFrame.TextRange.Text = ($Bullets -join "`r`n")
        }
    }
}
