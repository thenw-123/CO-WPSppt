# Renderer helper functions extracted from Wps.RunSpec.ps1.

function Invoke-WpsSlideChartRender {
    param(
        $Slide,
        $SlideSpec,
        [int]$SlideIndex,
        [string]$ProjectRoot,
        [string]$ChartDir,
        [double]$SlideWidth
    )

    $hasChartData = $null -ne $SlideSpec.chart_data
    $ctLower = if ($SlideSpec.chart_type) { [string]$SlideSpec.chart_type } else { '' }
    $ctLower = $ctLower.ToLowerInvariant()
    if ($hasChartData -and -not $ctLower) { $ctLower = 'bar' }

    $chartEngine = 'png'
    if ($SlideSpec.chart_engine) {
        $ce = [string]$SlideSpec.chart_engine
        if ($ce.ToLowerInvariant() -eq 'com') { $chartEngine = 'com' }
    }

    $comOk = $false
    $hasOwnImage = $SlideSpec.image -or $SlideSpec.imageUrl -or $SlideSpec.imageCommons
    if ($chartEngine -eq 'com' -and $hasChartData -and ($ctLower -in @('bar', 'line', 'donut')) -and (-not $hasOwnImage)) {
        try {
            Add-WpsEmbeddedChartFromData -Slide $Slide -ChartType $ctLower -ChartData $SlideSpec.chart_data
            $comOk = $true
        } catch {
            Write-Verbose ("COM chart skipped: " + $_.Exception.Message)
        }
    }

    $skipChartPng = $false
    if ($null -ne $SlideSpec.chart_render) {
        $cr = $SlideSpec.chart_render
        if ($cr -eq $false) { $skipChartPng = $true }
        elseif ([string]$cr -match '^(0|false|off|none)$') { $skipChartPng = $true }
    }
    $fallbackPng = $true
    if ($null -ne $SlideSpec.chart_fallback_png -and $SlideSpec.chart_fallback_png -eq $false) { $fallbackPng = $false }

    $canPng = (-not $skipChartPng) -and $hasChartData -and ($ctLower -in @('bar', 'line', 'donut')) -and (-not $hasOwnImage)
    if ($chartEngine -eq 'com' -and $comOk) { $canPng = $false }
    if ($chartEngine -eq 'com' -and -not $comOk -and -not $fallbackPng) { $canPng = $false }

    if ($canPng) {
        $pngPath = Join-Path $ChartDir ('slide-{0}.png' -f $SlideIndex)
        try {
            Invoke-PptChartPngRender -ProjectRoot $ProjectRoot -OutPngPath $pngPath -ChartType $ctLower -ChartData $SlideSpec.chart_data
            Add-WpsSlidePictureFromSpec -Slide $Slide -ImagePath $pngPath -SlideSpec $SlideSpec -SlideWidth $SlideWidth | Out-Null
        } catch {
            Write-Verbose ("Chart PNG skipped: " + $_.Exception.Message)
        }
    }
}

function Invoke-WpsSlideAssetRender {
    param(
        $Slide,
        $SlideSpec,
        [int]$SlideIndex,
        [string]$ProjectRoot,
        [string]$AssetsDir,
        $AssetFetchOptions,
        [double]$SlideWidth
    )

    if ($SlideSpec.imageCommons) {
        $wComm = 500
        if ($null -ne $SlideSpec.commonsThumbWidth) {
            try { $wComm = [int]$SlideSpec.commonsThumbWidth } catch { $wComm = 500 }
        }
        $localComm = Join-Path $AssetsDir ('commons-s{0}-{1}.jpg' -f $SlideIndex, [Guid]::NewGuid().ToString('N').Substring(0, 10))
        Save-WpsCommonsThumbFile -FileTitle ([string]$SlideSpec.imageCommons) -OutPath $localComm -Width $wComm
        Add-WpsSlidePictureFromSpec -Slide $Slide -ImagePath $localComm -SlideSpec $SlideSpec -SlideWidth $SlideWidth | Out-Null
    }

    if ($SlideSpec.imageUrl) {
        try {
            $saved = Save-RemoteImageForSlide -UrlString ([string]$SlideSpec.imageUrl) -DestDir $AssetsDir -FetchOptions $AssetFetchOptions
            Add-WpsSlidePictureFromSpec -Slide $Slide -ImagePath $saved -SlideSpec $SlideSpec -SlideWidth $SlideWidth | Out-Null
        } catch {
            if ($AssetFetchOptions.SoftFail) {
                Write-Verbose ("imageUrl skipped: " + $_.Exception.Message)
            } else {
                throw
            }
        }
    }

    if ($SlideSpec.image) {
        $img = [string]$SlideSpec.image
        if (-not [System.IO.Path]::IsPathRooted($img)) {
            $img = Join-Path $ProjectRoot $img
        }
        if (Test-Path -LiteralPath $img) {
            Add-WpsSlidePictureFromSpec -Slide $Slide -ImagePath $img -SlideSpec $SlideSpec -SlideWidth $SlideWidth | Out-Null
        }
    }
}

function Invoke-WpsSlideMotionRender {
    param(
        $Slide,
        $SlideSpec,
        $ThemeBlock,
        [bool]$TwoColumn,
        [bool]$Narrative,
        [int]$SlideNumber,
        [System.Collections.Generic.List[string]]$ComWarnings
    )

    $tfx = $null
    if ($SlideSpec.transition) { $tfx = [string]$SlideSpec.transition }
    if (-not $tfx -and $ThemeBlock -and $ThemeBlock.defaultTransition) { $tfx = [string]$ThemeBlock.defaultTransition }
    $td = 0.55
    if ($null -ne $SlideSpec.transitionDuration) {
        try { $td = [double]$SlideSpec.transitionDuration } catch { }
    } elseif ($ThemeBlock -and $null -ne $ThemeBlock.defaultTransitionDuration) {
        try { $td = [double]$ThemeBlock.defaultTransitionDuration } catch { }
    }
    if ($tfx) {
        Set-WpsSlideTransitionFromSpec -Slide $Slide -EffectName $tfx -Seconds $td `
            -SlideNumber $SlideNumber -ComWarnings $ComWarnings
    }

    $doBuild = $false
    if ($null -ne $SlideSpec.animateBuild) {
        try { $doBuild = [bool]$SlideSpec.animateBuild } catch { }
    }
    if (-not $doBuild -and $ThemeBlock -and $null -ne $ThemeBlock.defaultAnimateBuild) {
        try { $doBuild = [bool]$ThemeBlock.defaultAnimateBuild } catch { }
    }
    $stepList = @()
    if ($null -ne $SlideSpec.animateSteps -and $SlideSpec.animateSteps -is [System.Collections.IEnumerable] -and $SlideSpec.animateSteps -isnot [string]) {
        $stepList = @($SlideSpec.animateSteps)
    } elseif ($doBuild -and $ThemeBlock -and $null -ne $ThemeBlock.defaultAnimateSteps -and $ThemeBlock.defaultAnimateSteps -is [System.Collections.IEnumerable] -and $ThemeBlock.defaultAnimateSteps -isnot [string]) {
        $stepList = @($ThemeBlock.defaultAnimateSteps)
    }
    $useSteps = $stepList.Count -gt 0

    if ($useSteps -or $doBuild) {
        $animFx = 'fade'
        if ($null -ne $SlideSpec.animateEffect -and [string]$SlideSpec.animateEffect) {
            $animFx = [string]$SlideSpec.animateEffect
        } elseif ($ThemeBlock -and $null -ne $ThemeBlock.defaultAnimateEffect -and [string]$ThemeBlock.defaultAnimateEffect) {
            $animFx = [string]$ThemeBlock.defaultAnimateEffect
        }
        if ($useSteps) {
            Set-WpsSlideAnimateFromSpecSteps -Slide $Slide -Steps $stepList -TwoColumn:$TwoColumn -Narrative:$Narrative `
                -DefaultEffectName $animFx -SlideNumber $SlideNumber -ComWarnings $ComWarnings
        } elseif ($doBuild) {
            Set-WpsSlideAnimateBuild -Slide $Slide -TwoColumn:$TwoColumn -Narrative:$Narrative -EffectName $animFx `
                -SlideNumber $SlideNumber -ComWarnings $ComWarnings
        }
    }
}
