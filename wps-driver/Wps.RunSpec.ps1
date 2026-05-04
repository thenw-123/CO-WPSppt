# Build full deck from JSON spec. Caller must dot-source Common, Connect, Presentation, Slide, Text, Image, Polish, LayoutNarrative, ChartRender, ChartCom, UrlAsset first.

function Invoke-WpsDeckFromSpec {
    param(
        [object]$Spec,
        [string]$ProjectRoot
    )

    $out = $Spec.savePath
    if (-not $out) {
        $ts = Get-Date -Format 'yyyyMMdd-HHmmss'
        $out = Join-Path $ProjectRoot "output\deck-$ts.pptx"
    } elseif (-not [System.IO.Path]::IsPathRooted($out)) {
        $out = Join-Path $ProjectRoot $out
    }
    $dir = Split-Path -Parent $out
    if ($dir) { Ensure-Dir $dir }

    $deckSlug = [System.IO.Path]::GetFileNameWithoutExtension($out)
    if (-not $deckSlug) { $deckSlug = 'deck' }
    $chartDir = Join-Path $ProjectRoot ("output\charts\{0}" -f $deckSlug)
    Ensure-Dir $chartDir
    $assetOpts = Get-AssetFetchOptions -SpecAssetFetch $Spec.assetFetch
    $assetsDir = Join-Path $ProjectRoot ("output\assets\{0}" -f $deckSlug)
    Ensure-Dir $assetsDir

    $app = Get-WpsApplication
    $pres = New-WpsPresentation -App $app -SavePath $out

    $themeBlock = $Spec.theme
    if ($themeBlock) {
        $thPath = [string]$themeBlock.themePath
        if ($thPath -and -not [System.IO.Path]::IsPathRooted($thPath)) {
            $thPath = Join-Path $ProjectRoot $thPath
        }
        if ($thPath) { Apply-WpsThemeFromPath -Presentation $pres -ThemePath $thPath }
    }
    $themeTitleFont = $null
    $themeBodyFont = $null
    $agendaBg = $null
    $agendaTitleCol = $null
    $agendaBodyCol = $null
    $defSlideBg = $null
    $defTitleCol = $null
    $defBodyCol = $null
    $agTwo = $true
    $agNum = $true
    $agAccentBar = $null
    $agTitlePt = 42.0
    $agBodyPt = 21.0
    if ($themeBlock) {
        $themeTitleFont = [string]$themeBlock.titleFont
        $themeBodyFont = [string]$themeBlock.bodyFont
        if (-not $themeTitleFont -and -not $themeBodyFont) {
            $one = [string]$themeBlock.font
            if ($one) { $themeTitleFont = $one; $themeBodyFont = $one }
        }
        $agendaBg = [string]$themeBlock.agendaBackground
        $agendaTitleCol = [string]$themeBlock.agendaTitleColor
        $agendaBodyCol = [string]$themeBlock.agendaBodyColor
        $defSlideBg = [string]$themeBlock.defaultSlideBackground
        $defTitleCol = [string]$themeBlock.defaultTitleColor
        $defBodyCol = [string]$themeBlock.defaultBodyColor
        if ($null -ne $themeBlock.agendaTwoColumn) {
            try { $agTwo = [bool]$themeBlock.agendaTwoColumn } catch { }
        }
        if ($null -ne $themeBlock.agendaNumbered) {
            try { $agNum = [bool]$themeBlock.agendaNumbered } catch { }
        }
        $agAccentBar = [string]$themeBlock.agendaAccentColor
        if ($null -ne $themeBlock.agendaTitleSize) {
            try { $agTitlePt = [double]$themeBlock.agendaTitleSize } catch { }
        }
        if ($null -ne $themeBlock.agendaBodySize) {
            try { $agBodyPt = [double]$themeBlock.agendaBodySize } catch { }
        }
    }

    $slidesSpec = @($Spec.slides)
    if ($slidesSpec.Count -eq 0) {
        Save-WpsPresentation -Pres $pres -Path $out
        return @{
            path             = (Resolve-Path -LiteralPath $out).Path
            slideCount       = [int]$pres.Slides.Count
            presentationPath = (Resolve-Path -LiteralPath $out).Path
        }
    }

    # Some WPS builds return Slides.Count = 0 right after Add(); ensure at least one slide.
    if ([int]$pres.Slides.Count -lt 1) {
        Add-WpsSlide -Presentation $pres -LayoutInt 1 | Out-Null
    }

    $deckSlideW = 960.0
    try { $deckSlideW = [double]$pres.PageSetup.SlideWidth } catch { }
    $deckSlideH = 540.0
    try { $deckSlideH = [double]$pres.PageSetup.SlideHeight } catch { }
    $narrativeLayouts = @('timeline', 'comparison', 'thesis-chain', 'argument', 'thesis-vertical', 'swot')
    $comWarnings = [System.Collections.Generic.List[string]]::new()

    for ($i = 0; $i -lt $slidesSpec.Count; $i++) {
        $s = $slidesSpec[$i]
        $layoutName = if ($s.layout) { [string]$s.layout } else { 'title-content' }
        if ($s.twoColumns) { $layoutName = 'two-content' }
        $layout = Map-LayoutNameToInt -Name $layoutName
        $isNarrative = $narrativeLayouts -contains $layoutName.ToLowerInvariant()

        if ($i -eq 0) {
            $slide = $pres.Slides.Item(1)
            try { $slide.Layout = $layout } catch { }
        } else {
            Add-WpsSlide -Presentation $pres -LayoutInt $layout | Out-Null
            $slide = $pres.Slides.Item($pres.Slides.Count)
        }

        $bullets = @()
        if ($s.bullets) { $bullets = @($s.bullets) }
        elseif ($s.points) { $bullets = @($s.points) }

        $willImg = ($s.imageCommons -and [string]$s.imageCommons) -or ($s.imageUrl -and [string]$s.imageUrl) -or ($s.image -and [string]$s.image)

        if (-not $s.twoColumns) {
            if ($bullets.Count -eq 0) {
                $ct = if ($s.chart_type) { [string]$s.chart_type } else { '' }
                $ds = if ($s.data_summary) { [string]$s.data_summary } else { '' }
                $ins = if ($s.insight) { [string]$s.insight } else { '' }
                if ($ct -or $ds -or $ins) {
                    $typeLabel = if ($ct) {
                        switch ($ct.ToLowerInvariant()) {
                            'bar' { '柱状图' }
                            'line' { '折线图' }
                            'donut' { '环形图' }
                            'flowchart' { '流程图 / 步骤卡片' }
                            default { $ct }
                        }
                    } else { '' }
                    $lines = [System.Collections.ArrayList]@()
                    if ($typeLabel) { [void]$lines.Add("建议图表：$typeLabel") }
                    if ($ds) { [void]$lines.Add($ds) }
                    if ($ins) { [void]$lines.Add("核心洞察：$ins") }
                    $bullets = @($lines)
                }
            }
        }

        if ($isNarrative) {
            $rrN = 0.0
            $mbN = $deckSlideH
            if ($willImg) {
                $ipl0 = if ($s.imagePlacement) { [string]$s.imagePlacement.Trim().ToLowerInvariant() } else { 'right' }
                if ($ipl0 -eq 'right') {
                    $rrN = 408.0
                    if ($null -ne $s.imageReservedRight) {
                        try { $rrN = [double]$s.imageReservedRight } catch { }
                    }
                } else {
                    $mbN = 276.0
                    if ($null -ne $s.bodyBottomLimit) {
                        try { $mbN = [double]$s.bodyBottomLimit } catch { }
                    }
                }
            }
            $slideTitleColN = [string]$s.titleColor
            $slideBodyColN = [string]$s.bodyColor
            if (-not $slideTitleColN) { $slideTitleColN = $defTitleCol }
            if (-not $slideBodyColN) { $slideBodyColN = $defBodyCol }
            Build-WpsNarrativeSlide -Slide $slide -Spec $s -SlideWidth $deckSlideW -SlideHeight $deckSlideH `
                -ReserveRight $rrN -MaxContentBottom $mbN -TitleColorHex $slideTitleColN -BodyColorHex $slideBodyColN
        }
        elseif ($s.twoColumns) {
            $tc = $s.twoColumns
            $lba = @()
            if ($tc.leftBullets) { $lba = @($tc.leftBullets) }
            $rba = @()
            if ($tc.rightBullets) { $rba = @($tc.rightBullets) }
            Set-WpsSlideTwoColumnContent -Slide $slide -Title ([string]$s.title) `
                -LeftHeading ([string]$tc.leftHeading) -LeftBullets $lba `
                -RightHeading ([string]$tc.rightHeading) -RightBullets $rba
        } else {
            Set-WpsSlideContent -Slide $slide -Title ([string]$s.title) -Subtitle ([string]$s.subtitle) -Bullets $bullets
        }

        $slideBg = [string]$s.background
        if (-not $slideBg) { $slideBg = $defSlideBg }
        if ($slideBg) { Set-WpsSlideBackgroundFromHex -Slide $slide -Hex $slideBg }
        $slideTitleCol = [string]$s.titleColor
        $slideBodyCol = [string]$s.bodyColor
        if (-not $slideTitleCol) { $slideTitleCol = $defTitleCol }
        if (-not $slideBodyCol) { $slideBodyCol = $defBodyCol }
        if (-not $isNarrative -and ($slideTitleCol -or $slideBodyCol)) {
            Set-WpsSlideTextColors -Slide $slide -TitleHex $slideTitleCol -BodyHex $slideBodyCol
        }

        if ($willImg -and ($layoutName -in @('title-content', 'content', 'chart', 'two-content', 'timeline', 'comparison', 'thesis-chain', 'argument', 'thesis-vertical', 'swot'))) {
            $ipl = 'right'
            if ($s.imagePlacement) { $ipl = [string]$s.imagePlacement.Trim().ToLowerInvariant() }
            if ($ipl -eq 'bottom') {
                $mb = 276.0
                if ($null -ne $s.bodyBottomLimit) {
                    try { $mb = [double]$s.bodyBottomLimit } catch { }
                }
                if ($s.twoColumns) {
                    Set-WpsSlideTwoColumnBodiesBottomLimit -Slide $slide -MaxBottom $mb
                } else {
                    Set-WpsSlideBodyBottomLimit -Slide $slide -MaxBottom $mb
                }
            } else {
                $rr = 408.0
                if ($null -ne $s.imageReservedRight) {
                    try { $rr = [double]$s.imageReservedRight } catch { }
                }
                if ($s.twoColumns) {
                    Set-WpsSlideTwoColumnBodiesReservedRight -Slide $slide -SlideWidth $deckSlideW -ReserveRight $rr
                } else {
                    Set-WpsSlideBodyReservedRight -Slide $slide -SlideWidth $deckSlideW -ReserveRight $rr
                }
            }
        }

        $noteText = [string]$s.notes
        if ($noteText) { Set-WpsSlideNotes -Slide $slide -Text $noteText }

        Invoke-WpsSlideChartRender -Slide $slide -SlideSpec $s -SlideIndex ($i + 1) `
            -ProjectRoot $ProjectRoot -ChartDir $chartDir -SlideWidth $deckSlideW

        Invoke-WpsSlideAssetRender -Slide $slide -SlideSpec $s -SlideIndex ($i + 1) `
            -ProjectRoot $ProjectRoot -AssetsDir $assetsDir -AssetFetchOptions $assetOpts -SlideWidth $deckSlideW

        $tcFlag = $null -ne $s.twoColumns
        Invoke-WpsSlideMotionRender -Slide $slide -SlideSpec $s -ThemeBlock $themeBlock `
            -TwoColumn:$tcFlag -Narrative:$isNarrative -SlideNumber ($i + 1) -ComWarnings $comWarnings
    }

    $doAgenda = $false
    if ($null -ne $Spec.autoAgenda) { $doAgenda = [bool]$Spec.autoAgenda }
    if ($doAgenda -and $slidesSpec.Count -gt 1) {
        $agendaTitle = [string]$Spec.agendaTitle
        $toc = @()
        for ($k = 1; $k -lt $slidesSpec.Count; $k++) {
            $t = [string]$slidesSpec[$k].title
            if ($t) { $toc += $t }
        }
        if ($toc.Count -gt 0) {
            Add-WpsAgendaAfterCover -Presentation $pres -SectionTitles $toc -AgendaTitle $agendaTitle `
                -BackgroundHex $agendaBg -TitleColorHex $agendaTitleCol -BodyColorHex $agendaBodyCol `
                -TwoColumn:$agTwo -Numbered:$agNum -AccentBarHex $agAccentBar -TitlePt $agTitlePt -BodyPt $agBodyPt
        }
    }

    if ($themeTitleFont -or $themeBodyFont) {
        Set-WpsDeckFonts -Presentation $pres -TitleFont $themeTitleFont -BodyFont $themeBodyFont
    }

    Save-WpsPresentation -Pres $pres -Path $out
    $ret = @{
        path             = (Resolve-Path -LiteralPath $out).Path
        slideCount       = [int]$pres.Slides.Count
        presentationPath = (Resolve-Path -LiteralPath $out).Path
    }
    if ($comWarnings.Count -gt 0) {
        $ret.comWarnings = @($comWarnings)
    }
    return $ret
}
