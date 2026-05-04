# Partial PPTX update for edited PPT DSL. Rebuilds affected slides using the
# existing WPS renderer in a temporary one-slide deck, then copies that slide
# into the existing presentation at the sidecar-mapped index.

function Get-PptDslSlideAndIndexById {
    param(
        [object]$Dsl,
        [string]$SlideId
    )
    $idx = 0
    foreach ($s in @($Dsl.slides)) {
        $idx++
        if ([string]$s.id -eq $SlideId) {
            return @{ Slide = $s; Index = $idx }
        }
    }
    return $null
}

function Copy-WpsSlideIntoPresentationAtIndex {
    param(
        $SourcePresentation,
        $TargetPresentation,
        [int]$TargetIndex
    )

    $srcSlide = $SourcePresentation.Slides.Item(1)
    $srcSlide.Copy()

    $count = [int]$TargetPresentation.Slides.Count
    if ($TargetIndex -lt 1) { $TargetIndex = 1 }
    if ($TargetIndex -gt ($count + 1)) { $TargetIndex = $count + 1 }

    try {
        $pasted = $TargetPresentation.Slides.Paste($TargetIndex)
        return $pasted
    } catch {
        # Some WPS builds only support appending via Paste(); move after paste.
        $pasted = $TargetPresentation.Slides.Paste()
        try {
            $newSlide = $pasted.Item(1)
            $newSlide.MoveTo($TargetIndex)
        } catch { }
        return $pasted
    }
}

function Invoke-WpsRenderEditedSlides {
    param(
        [object]$Dsl,
        [object]$Edit,
        [string]$PresentationPath,
        [string]$MapPath,
        [string]$ProjectRoot,
        [string]$DslPath = $null
    )

    $map = Read-PptSlideMap -MapPath $MapPath
    $affected = Get-PptAffectedSlideIdsFromEdit -Edit $Edit
    if ($affected.Count -eq 0) { throw 'No affected slide ids in edit request' }

    $app = Get-WpsApplication
    $target = Open-WpsPresentation -App $app -Path $PresentationPath
    $rendered = @()
    $tempFiles = @()

    try {
        foreach ($slideId in $affected) {
            $entry = Get-PptSlideMapEntry -SlideMap $map -SlideId $slideId
            if ($null -eq $entry) { throw "slide map entry not found: $slideId" }
            $slideInfo = Get-PptDslSlideAndIndexById -Dsl $Dsl -SlideId $slideId
            if ($null -eq $slideInfo) { throw "DSL slide not found after edit: $slideId" }

            $targetIndex = [int]$entry.index
            if ($targetIndex -lt 1 -or $targetIndex -gt [int]$target.Slides.Count) {
                throw "slide index out of range for ${slideId}: $targetIndex"
            }

            $tmpName = 'partial-{0}-{1}.pptx' -f $slideId, ([Guid]::NewGuid().ToString('N').Substring(0, 8))
            $tmpPath = Join-Path $ProjectRoot (Join-Path 'output\partial' $tmpName)
            $tempFiles += $tmpPath

            $singleDsl = [pscustomobject]@{
                dslVersion = '2.0'
                meta       = $Dsl.meta
                theme      = $Dsl.theme
                defaults   = $Dsl.defaults
                slides     = @($slideInfo.Slide)
            }
            $singleDsl.meta | Add-Member -NotePropertyName savePath -NotePropertyValue $tmpPath -Force
            $singlePlan = Convert-PptDslToRenderPlan -Dsl $singleDsl
            $singleResult = Invoke-WpsDeckFromRenderPlan -RenderPlan $singlePlan -ProjectRoot $ProjectRoot

            $tempPres = Open-WpsPresentation -App $app -Path ([string]$singleResult.presentationPath)
            try {
                $target.Slides.Item($targetIndex).Delete()
                Copy-WpsSlideIntoPresentationAtIndex -SourcePresentation $tempPres -TargetPresentation $target -TargetIndex $targetIndex | Out-Null
            } finally {
                try { Close-WpsPresentation -Pres $tempPres -Save:$false } catch { }
            }

            $rendered += @{
                slideId = $slideId
                index   = $targetIndex
                temp    = $tmpPath
            }
        }

        Save-WpsPresentation -Pres $target -Path $PresentationPath
    } finally {
        try { Close-WpsPresentation -Pres $target -Save:$false } catch { }
        foreach ($f in $tempFiles) {
            try { if (Test-Path -LiteralPath $f) { Remove-Item -LiteralPath $f -Force } } catch { }
        }
    }

    $newMap = Save-PptSlideMap -Dsl $Dsl -PresentationPath $PresentationPath -DslPath $DslPath -MapPath $MapPath
    return @{
        presentationPath = (Resolve-Path -LiteralPath $PresentationPath).Path
        slideMapPath     = (Resolve-Path -LiteralPath $MapPath).Path
        renderedSlides   = @($rendered)
        slideMap         = $newMap
    }
}
