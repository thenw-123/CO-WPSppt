# Structural validation for ppt deck spec (aligns with specs/ppt-spec.schema.json)

$script:PptLayoutEnum = [string[]]@(
    'title', 'title-content', 'content', 'chart', 'section', 'two-content', 'blank',
    'timeline', 'comparison', 'thesis-chain', 'argument', 'thesis-vertical', 'swot'
)
$script:NarrativeLayouts = [string[]]@('timeline', 'comparison', 'thesis-chain', 'argument', 'thesis-vertical', 'swot')

function Test-PptOptionalHexColor {
    param([string]$Val)
    if (-not $Val) { return $true }
    return [bool]($Val -match '^#?[0-9A-Fa-f]{6}$')
}

function Test-PptNonEmptyNarrativeText {
    param($Val)
    if ($null -eq $Val) { return $false }
    if ($Val -is [string]) { return $Val.Trim().Length -gt 0 }
    if ($Val -is [System.Collections.IEnumerable]) {
        foreach ($x in @($Val)) {
            if ($null -ne $x -and [string]$x -match '\S') { return $true }
        }
        return $false
    }
    return ([string]$Val).Trim().Length -gt 0
}

function Test-PptSwotQuadrantField {
    param($Val, [int]$SlideIdx, [string]$Key)
    if ($null -eq $Val) { return "slides[$SlideIdx].swot.${Key}: required" }
    if ($Val -is [string]) { return $null }
    if ($Val -is [System.Collections.IEnumerable]) {
        foreach ($x in @($Val)) {
            if ($null -ne $x -and $x -isnot [string]) { return "slides[$SlideIdx].swot.${Key}: all items must be strings" }
        }
        return $null
    }
    return "slides[$SlideIdx].swot.${Key}: must be string or array of strings"
}

function Test-PptDeckSpecObject {
    param([object]$Spec)

    $errors = [System.Collections.Generic.List[string]]::new()

    if ($null -eq $Spec) {
        [void]$errors.Add('Spec is null')
        return $errors
    }

    if ($null -eq $Spec.slides) {
        [void]$errors.Add('slides: required')
        return $errors
    }

    $slides = @($Spec.slides)
    $idx = 0
    foreach ($sl in $slides) {
        $idx++
        if (-not $sl.layout) {
            [void]$errors.Add("slides[$idx].layout: required")
            continue
        }
        $lo = [string]$sl.layout
        if ($script:PptLayoutEnum -notcontains $lo) {
            [void]$errors.Add("slides[$idx].layout: unknown '$lo' (allowed: $($script:PptLayoutEnum -join ', '))")
        }

        foreach ($arrName in @('bullets', 'points')) {
            $arr = $sl.$arrName
            if ($null -eq $arr) { continue }
            if ($arr -is [string]) {
                [void]$errors.Add("slides[$idx].${arrName}: must be a JSON array, not a string")
                continue
            }
            if ($arr -isnot [System.Collections.IEnumerable]) {
                [void]$errors.Add("slides[$idx].${arrName}: must be a JSON array")
                continue
            }
            foreach ($item in $arr) {
                if ($null -ne $item -and $item -isnot [string]) {
                    [void]$errors.Add("slides[$idx].${arrName}: all items must be strings")
                    break
                }
            }
        }

        if ($sl.chart_data) {
            $cd = $sl.chart_data
            $labs = $cd.labels
            if ($null -eq $labs -or $labs -is [string]) {
                [void]$errors.Add("slides[$idx].chart_data.labels: must be a non-empty JSON array")
            } else {
                $nLab = @($labs).Count
                if ($nLab -lt 1) {
                    [void]$errors.Add("slides[$idx].chart_data.labels: must have at least one label")
                }
                $ser = $cd.series
                if ($null -eq $ser -or $ser -is [string]) {
                    [void]$errors.Add("slides[$idx].chart_data.series: must be a JSON array")
                } else {
                    $sj = 0
                    foreach ($row in @($ser)) {
                        $sj++
                        if ($null -eq $row) { continue }
                        $vals = $row.values
                        if ($null -eq $vals) {
                            [void]$errors.Add("slides[$idx].chart_data.series[$sj].values: required")
                            continue
                        }
                        if ($vals -is [string]) {
                            [void]$errors.Add("slides[$idx].chart_data.series[$sj].values: must be array of numbers")
                            continue
                        }
                        if (@($vals).Count -ne $nLab) {
                            [void]$errors.Add("slides[$idx].chart_data.series[$sj]: values length must match labels ($nLab)")
                        }
                        foreach ($v in @($vals)) {
                            try {
                                [void][double]::Parse([string]$v, [System.Globalization.CultureInfo]::InvariantCulture)
                            } catch {
                                [void]$errors.Add("slides[$idx].chart_data.series[$sj].values: all entries must be numbers")
                                break
                            }
                        }
                    }
                }
            }
            $ctp = [string]$sl.chart_type
            if ($ctp -and $ctp.ToLowerInvariant() -eq 'flowchart') {
                [void]$errors.Add("slides[$idx]: flowchart does not support chart_data PNG; omit chart_data or use bar|line|donut")
            }
            elseif ($ctp -and $ctp.ToLowerInvariant() -notin @('bar', 'line', 'donut')) {
                [void]$errors.Add("slides[$idx].chart_type: with chart_data use bar, line, or donut (got '$ctp')")
            }
        }

        $imgCount = 0
        if ($sl.image -and [string]$sl.image) { $imgCount++ }
        if ($sl.imageUrl -and [string]$sl.imageUrl) { $imgCount++ }
        if ($sl.imageCommons -and [string]$sl.imageCommons) { $imgCount++ }
        if ($imgCount -gt 1) {
            [void]$errors.Add("slides[$idx]: use only one of image, imageUrl, imageCommons")
        }
        if ($null -ne $sl.imageCommons -and $sl.imageCommons -isnot [string]) {
            [void]$errors.Add("slides[$idx].imageCommons: must be a string (Commons file name)")
        }
        if ($sl.commonsThumbWidth) {
            try {
                $cw = [int]$sl.commonsThumbWidth
                $steps = @(20, 40, 60, 120, 250, 330, 500, 960, 1280, 1920, 3840)
                if ($steps -notcontains $cw) {
                    [void]$errors.Add("slides[$idx].commonsThumbWidth: must be a Wikimedia standard width ($($steps -join ', '))")
                }
            } catch {
                [void]$errors.Add("slides[$idx].commonsThumbWidth: must be an integer")
            }
        }
        if ($sl.imageUrl) {
            try {
                $iu = [Uri][string]$sl.imageUrl
                if (-not $iu.IsAbsoluteUri) {
                    [void]$errors.Add("slides[$idx].imageUrl: must be an absolute URL")
                }
                elseif ($iu.Scheme -ne 'https' -and $iu.Scheme -ne 'http') {
                    [void]$errors.Add("slides[$idx].imageUrl: only http or https")
                }
            } catch {
                [void]$errors.Add("slides[$idx].imageUrl: invalid URL")
            }
        }
        if ($sl.chart_engine) {
            $ce = [string]$sl.chart_engine
            if ($ce.ToLowerInvariant() -notin @('png', 'com')) {
                [void]$errors.Add("slides[$idx].chart_engine: expected png or com")
            }
            elseif ($ce.ToLowerInvariant() -eq 'com' -and -not $sl.chart_data) {
                [void]$errors.Add("slides[$idx].chart_engine com requires chart_data")
            }
        }

        foreach ($hk in @('background', 'titleColor', 'bodyColor')) {
            $hv = $sl.$hk
            if ($null -eq $hv) { continue }
            if ($hv -isnot [string] -or -not (Test-PptOptionalHexColor -Val ([string]$hv))) {
                [void]$errors.Add("slides[$idx].${hk}: expected #RRGGBB (optional leading #)")
            }
        }

        foreach ($ik in @('imageLeft', 'imageTop', 'imageWidth', 'imageHeight')) {
            $iv = $sl.$ik
            if ($null -eq $iv) { continue }
            try {
                [void][double]::Parse([string]$iv, [System.Globalization.CultureInfo]::InvariantCulture)
            } catch {
                [void]$errors.Add("slides[$idx].${ik}: must be a number if present")
            }
        }
        if ($sl.imagePlacement) {
            $ip = [string]$sl.imagePlacement
            if ($ip.ToLowerInvariant() -notin @('right', 'bottom')) {
                [void]$errors.Add("slides[$idx].imagePlacement: use right or bottom")
            }
        }
        if ($null -ne $sl.imageReservedRight) {
            try {
                [void][double]::Parse([string]$sl.imageReservedRight, [System.Globalization.CultureInfo]::InvariantCulture)
            } catch {
                [void]$errors.Add("slides[$idx].imageReservedRight: must be a number if present")
            }
        }
        if ($null -ne $sl.bodyBottomLimit) {
            try {
                [void][double]::Parse([string]$sl.bodyBottomLimit, [System.Globalization.CultureInfo]::InvariantCulture)
            } catch {
                [void]$errors.Add("slides[$idx].bodyBottomLimit: must be a number if present")
            }
        }
        if ($sl.transition) {
            $tv = [string]$sl.transition
            if ($tv.ToLowerInvariant() -notin @('none', 'fade', 'push', 'wipe', 'cut', 'uncover')) {
                [void]$errors.Add("slides[$idx].transition: unknown effect (fade, push, wipe, cut, uncover, none)")
            }
        }
        $loLower = $lo.ToLowerInvariant()
        if ($script:NarrativeLayouts -contains $loLower) {
            if ($sl.twoColumns) {
                [void]$errors.Add("slides[$idx]: narrative layout '$lo' cannot be combined with twoColumns")
            }
            if ($sl.chart_data) {
                [void]$errors.Add("slides[$idx]: narrative layout '$lo' does not support chart_data (use a separate chart slide)")
            }
            switch ($loLower) {
                'timeline' {
                    $tl = $sl.timeline
                    if ($null -eq $tl -or $tl -is [string] -or @($tl).Count -lt 1) {
                        [void]$errors.Add("slides[$idx].timeline: required non-empty array for layout timeline")
                    } else {
                        $ti = 0
                        foreach ($row in @($tl)) {
                            $ti++
                            if ($null -eq $row) { continue }
                            $mk = [string]$row.mark
                            if (-not $mk) { $mk = [string]$row.year }
                            $tx = [string]$row.text
                            if (-not $tx) { $tx = [string]$row.caption }
                            if (-not $mk -and -not $tx) {
                                [void]$errors.Add("slides[$idx].timeline[$ti]: need mark|year and text|caption")
                            }
                        }
                    }
                }
                'comparison' {
                    $cmp = $sl.comparison
                    if ($null -eq $cmp) {
                        [void]$errors.Add("slides[$idx].comparison: required object for layout comparison")
                    } else {
                        $rw = $cmp.rows
                        if ($null -eq $rw -or $rw -is [string] -or @($rw).Count -lt 1) {
                            [void]$errors.Add("slides[$idx].comparison.rows: required non-empty array")
                        } else {
                            $ri = 0
                            foreach ($row in @($rw)) {
                                $ri++
                                if ($null -eq $row) { continue }
                                if ($null -eq $row.PSObject.Properties['left'] -or $null -eq $row.PSObject.Properties['right']) {
                                    [void]$errors.Add("slides[$idx].comparison.rows[$ri]: each row needs left and right")
                                    break
                                }
                            }
                        }
                    }
                }
                'thesis-chain' {
                    $ch = $sl.chain
                    if ($null -eq $ch -or $ch -is [string] -or @($ch).Count -lt 2) {
                        [void]$errors.Add("slides[$idx].chain: required array of at least 2 items for thesis-chain")
                    } else {
                        $ci = 0
                        foreach ($row in @($ch)) {
                            $ci++
                            if ($null -eq $row) { continue }
                            if (-not ([string]$row.label).Trim() -and -not ([string]$row.text).Trim()) {
                                [void]$errors.Add("slides[$idx].chain[$ci]: need label and/or text")
                            }
                        }
                    }
                }
                'argument' {
                    $ar = $sl.argument
                    if ($null -eq $ar) {
                        [void]$errors.Add("slides[$idx].argument: required object for layout argument")
                    } else {
                        if (-not ([string]$ar.thesis).Trim() -and -not ([string]$ar.therefore).Trim()) {
                            [void]$errors.Add("slides[$idx].argument: need thesis and/or therefore")
                        }
                        $bc = $ar.because
                        if ($null -ne $bc -and $bc -isnot [string]) {
                            if ($bc -isnot [System.Collections.IEnumerable]) {
                                [void]$errors.Add("slides[$idx].argument.because: must be array of strings if present")
                            } else {
                                foreach ($x in @($bc)) {
                                    if ($null -ne $x -and $x -isnot [string]) {
                                        [void]$errors.Add("slides[$idx].argument.because: all items must be strings")
                                        break
                                    }
                                }
                            }
                        }
                    }
                }
                'thesis-vertical' {
                    $tv = $sl.thesisVertical
                    if ($null -eq $tv) {
                        [void]$errors.Add("slides[$idx].thesisVertical: required object for layout thesis-vertical")
                    } else {
                        $cl = [string]$tv.claim
                        if (-not $cl.Trim()) { $cl = [string]$tv.thesis }
                        if (-not $cl.Trim()) {
                            [void]$errors.Add("slides[$idx].thesisVertical: need non-empty claim or thesis")
                        }
                        $st = $tv.steps
                        if ($null -eq $st -or $st -is [string] -or @($st).Count -ne 3) {
                            [void]$errors.Add("slides[$idx].thesisVertical.steps: must be exactly 3 items")
                        } else {
                            $sx = 0
                            foreach ($row in @($st)) {
                                $sx++
                                if ($null -eq $row) {
                                    [void]$errors.Add("slides[$idx].thesisVertical.steps[$sx]: required object")
                                    continue
                                }
                                $ok = Test-PptNonEmptyNarrativeText -Val $row.text
                                if (-not $ok) { $ok = [string]$row.label -match '\S' }
                                if (-not $ok) {
                                    [void]$errors.Add("slides[$idx].thesisVertical.steps[$sx]: need label and/or text")
                                }
                            }
                        }
                    }
                }
                'swot' {
                    $sw = $sl.swot
                    if ($null -eq $sw) {
                        [void]$errors.Add("slides[$idx].swot: required object for layout swot")
                    } else {
                        foreach ($qk in @('strengths', 'weaknesses', 'opportunities', 'threats')) {
                            $em = Test-PptSwotQuadrantField -Val $sw.$qk -SlideIdx $idx -Key $qk
                            if ($em) { [void]$errors.Add($em) }
                        }
                        $anyQ = (Test-PptNonEmptyNarrativeText -Val $sw.strengths) -or (Test-PptNonEmptyNarrativeText -Val $sw.weaknesses) -or `
                            (Test-PptNonEmptyNarrativeText -Val $sw.opportunities) -or (Test-PptNonEmptyNarrativeText -Val $sw.threats)
                        if (-not $anyQ) {
                            [void]$errors.Add("slides[$idx].swot: at least one quadrant must be non-empty")
                        }
                        $hd = $sw.headers
                        if ($null -ne $hd) {
                            if ($hd -isnot [pscustomobject] -and $hd -isnot [hashtable]) {
                                [void]$errors.Add("slides[$idx].swot.headers: must be an object if present")
                            } else {
                                foreach ($hk in @('strengths', 'weaknesses', 'opportunities', 'threats')) {
                                    $hv = $hd.$hk
                                    if ($null -ne $hv -and $hv -isnot [string]) {
                                        [void]$errors.Add("slides[$idx].swot.headers.${hk}: must be string if present")
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
        if ($sl.twoColumns) {
            $tc = $sl.twoColumns
            if ($null -eq $tc) {
                [void]$errors.Add("slides[$idx].twoColumns: must be an object")
            } else {
                foreach ($hk in @('leftHeading', 'rightHeading')) {
                    $hv = $tc.$hk
                    if ($null -ne $hv -and $hv -isnot [string]) {
                        [void]$errors.Add("slides[$idx].twoColumns.${hk}: must be string if present")
                    }
                }
                $lb = $tc.leftBullets
                $rb = $tc.rightBullets
                if ($null -eq $lb -or $lb -is [string] -or @($lb).Count -lt 1) {
                    [void]$errors.Add("slides[$idx].twoColumns.leftBullets: required non-empty array")
                }
                if ($null -eq $rb -or $rb -is [string] -or @($rb).Count -lt 1) {
                    [void]$errors.Add("slides[$idx].twoColumns.rightBullets: required non-empty array")
                }
            }
        }
        if ($null -ne $sl.transitionDuration) {
            try {
                [void][double]::Parse([string]$sl.transitionDuration, [System.Globalization.CultureInfo]::InvariantCulture)
            } catch {
                [void]$errors.Add("slides[$idx].transitionDuration: must be a number if present")
            }
        }
    }

    if ($Spec.assetFetch) {
        $af = $Spec.assetFetch
        if ($af.allowedHosts) {
            if ($af.allowedHosts -is [string]) {
                [void]$errors.Add('assetFetch.allowedHosts: must be a JSON array of host strings')
            }
        }
    }

    if ($Spec.theme) {
        $th = $Spec.theme
        foreach ($k in @(
                'themePath', 'font', 'titleFont', 'bodyFont', 'name',
                'agendaBackground', 'agendaTitleColor', 'agendaBodyColor', 'agendaAccentColor',
                'defaultSlideBackground', 'defaultTitleColor', 'defaultBodyColor'
            )) {
            $v = $th.$k
            if ($null -ne $v -and $v -isnot [string]) {
                [void]$errors.Add("theme.${k}: must be string if present")
            }
            elseif ($null -ne $v -and $k -match 'Color$|Background$' -and -not (Test-PptOptionalHexColor -Val ([string]$v))) {
                [void]$errors.Add("theme.${k}: expected #RRGGBB (optional leading #)")
            }
        }
        foreach ($ak in @('agendaTitleSize', 'agendaBodySize', 'defaultTransitionDuration')) {
            $av = $th.$ak
            if ($null -eq $av) { continue }
            try {
                [void][double]::Parse([string]$av, [System.Globalization.CultureInfo]::InvariantCulture)
            } catch {
                [void]$errors.Add("theme.${ak}: must be a number if present")
            }
        }
        if ($th.defaultTransition) {
            $dv = [string]$th.defaultTransition
            if ($dv.ToLowerInvariant() -notin @('none', 'fade', 'push', 'wipe', 'cut', 'uncover')) {
                [void]$errors.Add('theme.defaultTransition: unknown effect')
            }
        }
    }

    if ($null -ne $Spec.savePath -and $Spec.savePath -isnot [string]) {
        [void]$errors.Add('savePath: must be string if present')
    }

    return $errors
}
