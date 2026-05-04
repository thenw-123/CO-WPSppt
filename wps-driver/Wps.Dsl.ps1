# PPT DSL validation and compilation helpers.

$script:PptDslKinds = [string[]]@(
    'cover', 'section', 'content', 'bullets', 'argument', 'comparison', 'timeline', 'datachart', 'data', 'swot', 'agenda', 'custom', 'left-image'
)

function Get-PptDslArray {
    param($Value)
    if ($null -eq $Value) { return @() }
    if ($Value -is [string]) { return @($Value) }
    if ($Value -is [System.Collections.IEnumerable]) { return @($Value) }
    return @($Value)
}

function Get-PptDslString {
    param($Value)
    if ($null -eq $Value) { return '' }
    return [string]$Value
}

function Get-PptDslSlideKind {
    param($Slide)
    if ($Slide.type) { return ([string]$Slide.type).ToLowerInvariant() }
    if ($Slide.kind) { return ([string]$Slide.kind).ToLowerInvariant() }
    return ''
}

function Get-PptDslElementsByType {
    param($Slide, [string]$Type)
    if (-not $Slide.elements) { return @() }
    $wanted = $Type.ToLowerInvariant()
    return @(Get-PptDslArray $Slide.elements | Where-Object {
        $null -ne $_ -and $_.type -and ([string]$_.type).ToLowerInvariant() -eq $wanted
    })
}

function Get-PptDslElementByRole {
    param($Slide, [string]$Role)
    if (-not $Slide.elements) { return $null }
    $wanted = $Role.ToLowerInvariant()
    foreach ($el in @(Get-PptDslArray $Slide.elements)) {
        if ($null -ne $el -and $el.role -and ([string]$el.role).ToLowerInvariant() -eq $wanted) {
            return $el
        }
    }
    return $null
}

function Get-PptDslTitleFromElements {
    param($Slide)
    $titleEl = Get-PptDslElementByRole -Slide $Slide -Role 'title'
    if ($titleEl -and $titleEl.text) { return [string]$titleEl.text }
    if ($Slide.title) { return [string]$Slide.title }
    if ($Slide.message) { return [string]$Slide.message }
    return ''
}

function Get-PptDslSubtitleFromElements {
    param($Slide)
    $subEl = Get-PptDslElementByRole -Slide $Slide -Role 'subtitle'
    if ($subEl -and $subEl.text) { return [string]$subEl.text }
    if ($Slide.subtitle) { return [string]$Slide.subtitle }
    return ''
}

function Get-PptDslBulletTextsFromElement {
    param($Element)
    if ($null -eq $Element) { return @() }
    $raw = $null
    if ($null -ne $Element.items) { $raw = $Element.items }
    elseif ($Element.content -and $null -ne $Element.content.items) { $raw = $Element.content.items }
    if ($null -eq $raw) { return @() }
    return @(Get-PptDslArray $raw | ForEach-Object {
        if ($null -eq $_) { '' }
        elseif ($_ -is [string]) { [string]$_ }
        elseif ($_.text) { [string]$_.text }
        else { [string]$_ }
    } | Where-Object { $_ -ne '' })
}

function Test-PptDslObject {
    param([object]$Dsl)

    $errors = [System.Collections.Generic.List[string]]::new()
    if ($null -eq $Dsl) {
        [void]$errors.Add('DSL is null')
        return $errors
    }
    if (-not $Dsl.dslVersion) {
        [void]$errors.Add('dslVersion: required')
    } elseif ([string]$Dsl.dslVersion -ne '2.0') {
        [void]$errors.Add("dslVersion: unsupported '$($Dsl.dslVersion)' (expected 2.0)")
    }
    if ($null -eq $Dsl.slides) {
        [void]$errors.Add('slides: required')
        return $errors
    }
    if ($Dsl.slides -is [string] -or $Dsl.slides -isnot [System.Collections.IEnumerable]) {
        [void]$errors.Add('slides: must be a JSON array')
        return $errors
    }

    $slides = @($Dsl.slides)
    if ($slides.Count -lt 1) {
        [void]$errors.Add('slides: must contain at least one slide')
        return $errors
    }

    $idx = 0
    foreach ($slide in $slides) {
        $idx++
        $kind = Get-PptDslSlideKind -Slide $slide
        if (-not $kind) {
            [void]$errors.Add("slides[$idx].type: required")
            continue
        }
        if ($script:PptDslKinds -notcontains $kind) {
            [void]$errors.Add("slides[$idx].type: unknown '$kind' (allowed: $($script:PptDslKinds -join ', '))")
        }
        if ($slide.elements -and ($slide.elements -is [string] -or $slide.elements -isnot [System.Collections.IEnumerable])) {
            [void]$errors.Add("slides[$idx].elements: must be a JSON array")
        }
        if ($slide.assets -and ($slide.assets -is [string] -or $slide.assets -isnot [System.Collections.IEnumerable])) {
            [void]$errors.Add("slides[$idx].assets: must be a JSON array")
        }
        if ($kind -eq 'data' -or $kind -eq 'datachart') {
            $content = $slide.content
            $chartEl = @(Get-PptDslElementsByType -Slide $slide -Type 'chart' | Select-Object -First 1)
            if (($null -eq $content -or $null -eq $content.labels -or $null -eq $content.series) -and $chartEl.Count -eq 0) {
                [void]$errors.Add("slides[$idx]: data slide requires chart element or content labels/series")
            }
        }
        if ($kind -eq 'swot') {
            $content = $slide.content
            $hasTable = @(Get-PptDslElementsByType -Slide $slide -Type 'table' | Select-Object -First 1).Count -gt 0
            if (-not $hasTable) {
                foreach ($key in @('strengths', 'weaknesses', 'opportunities', 'threats')) {
                    if ($null -eq $content.$key) {
                        [void]$errors.Add("slides[$idx].content.${key}: required for swot when no table element is provided")
                    }
                }
            }
        }
    }
    return $errors
}

function Convert-PptDslThemeToLegacy {
    param($Theme)
    if ($null -eq $Theme) { return $null }

    $legacy = [ordered]@{}
    $preset = if ($Theme.preset) { [string]$Theme.preset } else { '' }
    switch ($preset.ToLowerInvariant()) {
        'business-blue' {
            $legacy.name = 'business-blue'
            $legacy.defaultSlideBackground = '#F7FAFC'
            $legacy.defaultTitleColor = '#123B73'
            $legacy.defaultBodyColor = '#243447'
            $legacy.agendaAccentColor = '#2B6CB0'
        }
        'dark-executive' {
            $legacy.name = 'dark-executive'
            $legacy.defaultSlideBackground = '#111827'
            $legacy.defaultTitleColor = '#F9FAFB'
            $legacy.defaultBodyColor = '#E5E7EB'
            $legacy.agendaAccentColor = '#60A5FA'
        }
        'clean-light' {
            $legacy.name = 'clean-light'
            $legacy.defaultSlideBackground = '#FFFFFF'
            $legacy.defaultTitleColor = '#111827'
            $legacy.defaultBodyColor = '#374151'
            $legacy.agendaAccentColor = '#6B7280'
        }
    }

    if ($Theme.font) { $legacy.font = [string]$Theme.font }
    if ($Theme.titleFont) { $legacy.titleFont = [string]$Theme.titleFont }
    if ($Theme.bodyFont) { $legacy.bodyFont = [string]$Theme.bodyFont }
    if ($Theme.fonts) {
        if ($Theme.fonts.title) { $legacy.titleFont = [string]$Theme.fonts.title }
        if ($Theme.fonts.body) { $legacy.bodyFont = [string]$Theme.fonts.body }
        if (-not $legacy.titleFont -and -not $legacy.bodyFont -and $Theme.fonts.body) {
            $legacy.font = [string]$Theme.fonts.body
        }
    }
    if ($Theme.colors) {
        if ($Theme.colors.background) { $legacy.defaultSlideBackground = [string]$Theme.colors.background }
        if ($Theme.colors.title) { $legacy.defaultTitleColor = [string]$Theme.colors.title }
        if ($Theme.colors.body) { $legacy.defaultBodyColor = [string]$Theme.colors.body }
        if ($Theme.colors.accent) { $legacy.agendaAccentColor = [string]$Theme.colors.accent }
        $legacy.colors = $Theme.colors
    }
    if ($legacy.Count -eq 0) { return $null }
    return $legacy
}

function Add-PptDslAssetToLegacySlide {
    param(
        [hashtable]$LegacySlide,
        $Slide
    )
    if (-not $Slide.assets) { return }
    $asset = @(Get-PptDslArray $Slide.assets | Where-Object { $null -ne $_ } | Select-Object -First 1)
    if ($asset.Count -eq 0) { return }
    $a = $asset[0]
    $type = if ($a.type) { ([string]$a.type).ToLowerInvariant() } else { '' }
    switch ($type) {
        'localimage' { if ($a.path) { $LegacySlide.image = [string]$a.path } }
        'remoteimage' { if ($a.url) { $LegacySlide.imageUrl = [string]$a.url } }
        'commonsimage' { if ($a.title) { $LegacySlide.imageCommons = [string]$a.title } }
    }
    if ($a.placement) { $LegacySlide.imagePlacement = [string]$a.placement }
}

function Add-PptDslImageElementToLegacySlide {
    param(
        [hashtable]$LegacySlide,
        $Slide
    )
    $imgEl = @(Get-PptDslElementsByType -Slide $Slide -Type 'image' | Select-Object -First 1)
    if ($imgEl.Count -eq 0) { return }
    $el = $imgEl[0]
    if ($null -eq $el.source) { return }
    $src = $el.source
    $kind = if ($src.kind) { ([string]$src.kind).ToLowerInvariant() } else { '' }
    switch ($kind) {
        'local' { if ($src.path) { $LegacySlide.image = [string]$src.path } }
        'remote' { if ($src.url) { $LegacySlide.imageUrl = [string]$src.url } }
        'commons' { if ($src.title) { $LegacySlide.imageCommons = [string]$src.title } }
    }
    if ($el.layoutHint -and $el.layoutHint.placement) {
        $placement = [string]$el.layoutHint.placement
        if ($placement -in @('right', 'bottom')) { $LegacySlide.imagePlacement = $placement }
    }
}

function Add-PptDslVisualAndMotionToLegacySlide {
    param(
        [hashtable]$LegacySlide,
        $Slide
    )
    if ($Slide.layout -and $Slide.layout.name) {
        $layoutName = ([string]$Slide.layout.name).ToLowerInvariant()
        if ($layoutName -eq 'image-right') { $LegacySlide.imagePlacement = 'right' }
        if ($layoutName -eq 'image-bottom') { $LegacySlide.imagePlacement = 'bottom' }
    }
    if ($Slide.style) {
        if ($Slide.style.background) { $LegacySlide.background = [string]$Slide.style.background }
        if ($Slide.style.color) { $LegacySlide.titleColor = [string]$Slide.style.color }
    }
    if ($Slide.visual) {
        if ($Slide.visual.background) { $LegacySlide.background = [string]$Slide.visual.background }
        if ($Slide.visual.titleColor) { $LegacySlide.titleColor = [string]$Slide.visual.titleColor }
        if ($Slide.visual.bodyColor) { $LegacySlide.bodyColor = [string]$Slide.visual.bodyColor }
        if ($Slide.visual.imagePlacement) { $LegacySlide.imagePlacement = [string]$Slide.visual.imagePlacement }
    }
    if ($Slide.motion) {
        if ($Slide.motion.transition) { $LegacySlide.transition = [string]$Slide.motion.transition }
        if ($Slide.motion.duration) { $LegacySlide.transitionDuration = [double]$Slide.motion.duration }
        if ($null -ne $Slide.motion.transitionDuration) { $LegacySlide.transitionDuration = [double]$Slide.motion.transitionDuration }
        if ($Slide.motion.build -and ([string]$Slide.motion.build).ToLowerInvariant() -ne 'none') { $LegacySlide.animateBuild = $true }
        if ($null -ne $Slide.motion.animateBuild) { $LegacySlide.animateBuild = [bool]$Slide.motion.animateBuild }
        if ($Slide.motion.effect) { $LegacySlide.animateEffect = [string]$Slide.motion.effect }
        if ($Slide.motion.animateEffect) { $LegacySlide.animateEffect = [string]$Slide.motion.animateEffect }
    }
}

function Convert-PptDslSlideToLegacy {
    param($Slide)

    $content = $Slide.content
    $kind = Get-PptDslSlideKind -Slide $Slide
    if ($kind -eq 'data') { $kind = 'datachart' }
    if ($kind -eq 'custom') { $kind = 'content' }
    $title = Get-PptDslTitleFromElements -Slide $Slide
    $subtitle = Get-PptDslSubtitleFromElements -Slide $Slide
    $bulletEl = @(Get-PptDslElementsByType -Slide $Slide -Type 'bullets' | Select-Object -First 1)
    $chartEl = @(Get-PptDslElementsByType -Slide $Slide -Type 'chart' | Select-Object -First 1)
    $tableEl = @(Get-PptDslElementsByType -Slide $Slide -Type 'table' | Select-Object -First 1)
    $legacy = [ordered]@{
        layout = 'title-content'
    }
    if ($title) { $legacy.title = $title }
    if ($subtitle) { $legacy.subtitle = $subtitle }
    if ($Slide.notes) { $legacy.notes = [string]$Slide.notes }

    switch ($kind) {
        'cover' {
            $legacy.layout = 'title'
            if (-not $legacy.title -and $content.title) { $legacy.title = [string]$content.title }
            if (-not $legacy.subtitle -and $content.subtitle) { $legacy.subtitle = [string]$content.subtitle }
        }
        'section' {
            $legacy.layout = 'section'
        }
        'agenda' {
            $legacy.layout = 'content'
            if ($bulletEl.Count -gt 0) {
                $legacy.bullets = @(Get-PptDslBulletTextsFromElement -Element $bulletEl[0])
            } else {
                $legacy.bullets = @(Get-PptDslArray $content.items | ForEach-Object { [string]$_ })
            }
        }
        { $_ -in @('content', 'bullets') } {
            $legacy.layout = 'title-content'
            if ($bulletEl.Count -gt 0) {
                $legacy.bullets = @(Get-PptDslBulletTextsFromElement -Element $bulletEl[0])
            } else {
                $points = if ($content.points) { $content.points } elseif ($content.bullets) { $content.bullets } else { @() }
                $legacy.bullets = @(Get-PptDslArray $points | ForEach-Object { [string]$_ })
            }
        }
        'argument' {
            $legacy.layout = 'argument'
            $claimEl = Get-PptDslElementByRole -Slide $Slide -Role 'claim'
            $insightEl = Get-PptDslElementByRole -Slide $Slide -Role 'insight'
            $claim = Get-PptDslString $(if ($claimEl -and $claimEl.text) { $claimEl.text } elseif ($content.claim) { $content.claim } elseif ($Slide.message) { $Slide.message } else { $legacy.title })
            $points = if ($bulletEl.Count -gt 0) { Get-PptDslBulletTextsFromElement -Element $bulletEl[0] } elseif ($content.points) { $content.points } elseif ($content.because) { $content.because } else { @() }
            $conclusion = if ($insightEl -and $insightEl.text) { [string]$insightEl.text } else { Get-PptDslString $content.conclusion }
            $legacy.argument = [ordered]@{
                thesis    = $claim
                because   = @(Get-PptDslArray $points | ForEach-Object { [string]$_ })
                therefore = $conclusion
            }
        }
        'comparison' {
            if ($tableEl.Count -gt 0 -and $tableEl[0].table -and @($tableEl[0].table.columns).Count -ge 2) {
                $legacy.layout = 'comparison'
                $cols = @($tableEl[0].table.columns)
                $legacy.comparison = [ordered]@{
                    leftHeader  = [string]$cols[0]
                    rightHeader = [string]$cols[1]
                    rows        = @(Get-PptDslArray $tableEl[0].table.rows | ForEach-Object {
                        $r = @($_)
                        [ordered]@{ left = (Get-PptDslString $r[0]); right = (Get-PptDslString $r[1]) }
                    })
                }
            } else {
                $legacy.layout = 'comparison'
                $legacy.comparison = [ordered]@{
                    leftHeader  = (Get-PptDslString $content.leftHeader)
                    rightHeader = (Get-PptDslString $content.rightHeader)
                    rows        = @(Get-PptDslArray $content.rows | ForEach-Object {
                        [ordered]@{ left = (Get-PptDslString $_.left); right = (Get-PptDslString $_.right) }
                    })
                }
            }
        }
        'timeline' {
            $legacy.layout = 'timeline'
            $items = if ($content.items) { $content.items } elseif ($content.timeline) { $content.timeline } else { @() }
            $legacy.timeline = @(Get-PptDslArray $items | ForEach-Object {
                [ordered]@{
                    mark = (Get-PptDslString $(if ($_.mark) { $_.mark } else { $_.year }))
                    text = (Get-PptDslString $(if ($_.text) { $_.text } else { $_.caption }))
                }
            })
        }
        'datachart' {
            $legacy.layout = 'chart'
            $chart = if ($chartEl.Count -gt 0) { $chartEl[0].chart } else { $null }
            $chartData = if ($chart -and $chart.data) { $chart.data } else { $content }
            $legacy.chart_type = if ($chart -and $chart.kind) { [string]$chart.kind } elseif ($content.chartType) { [string]$content.chartType } elseif ($content.type) { [string]$content.type } else { 'bar' }
            $legacy.chart_data = [ordered]@{
                labels = @(Get-PptDslArray $chartData.labels | ForEach-Object { [string]$_ })
                series = @(Get-PptDslArray $chartData.series | ForEach-Object {
                    [ordered]@{
                        name   = (Get-PptDslString $_.name)
                        values = @(Get-PptDslArray $_.values | ForEach-Object { [double]$_ })
                    }
                })
            }
            if ($chart -and $chart.render -and $chart.render.engine -and ([string]$chart.render.engine) -ne 'auto') { $legacy.chart_engine = [string]$chart.render.engine }
            elseif ($content.engine) { $legacy.chart_engine = [string]$content.engine }
            if ($chart -and $chart.title) { $legacy.data_summary = [string]$chart.title }
            elseif ($content.summary) { $legacy.data_summary = [string]$content.summary }
            if ($chart -and $chart.insight) { $legacy.insight = [string]$chart.insight }
            elseif ($content.insight) { $legacy.insight = [string]$content.insight }
        }
        'swot' {
            $legacy.layout = 'swot'
            $legacy.swot = [ordered]@{
                strengths     = $content.strengths
                weaknesses    = $content.weaknesses
                opportunities = $content.opportunities
                threats       = $content.threats
            }
            if ($content.headers) { $legacy.swot['headers'] = $content.headers }
        }
    }

    Add-PptDslAssetToLegacySlide -LegacySlide $legacy -Slide $Slide
    Add-PptDslImageElementToLegacySlide -LegacySlide $legacy -Slide $Slide
    Add-PptDslVisualAndMotionToLegacySlide -LegacySlide $legacy -Slide $Slide
    return $legacy
}

function Convert-PptDslToLegacySpec {
    param([object]$Dsl)

    $errors = Test-PptDslObject -Dsl $Dsl
    if ($errors.Count -gt 0) {
        throw "DSL validation failed: $($errors -join '; ')"
    }

    $legacy = [ordered]@{
        title  = (Get-PptDslString $(if ($Dsl.meta -and $Dsl.meta.title) { $Dsl.meta.title } else { $Dsl.deck.title }))
        slides = @()
    }
    if ($Dsl.meta -and $Dsl.meta.savePath) { $legacy.savePath = [string]$Dsl.meta.savePath }
    elseif ($Dsl.deck -and $Dsl.deck.savePath) { $legacy.savePath = [string]$Dsl.deck.savePath }
    if ($Dsl.defaults -and $Dsl.defaults.assetFetch) { $legacy.assetFetch = $Dsl.defaults.assetFetch }
    elseif ($Dsl.assetFetch) { $legacy.assetFetch = $Dsl.assetFetch }
    $theme = Convert-PptDslThemeToLegacy -Theme $Dsl.theme
    if ($theme) { $legacy.theme = $theme }

    $slides = @()
    foreach ($slide in @($Dsl.slides)) {
        $slides += Convert-PptDslSlideToLegacy -Slide $slide
    }
    $legacy.slides = $slides
    return $legacy
}
