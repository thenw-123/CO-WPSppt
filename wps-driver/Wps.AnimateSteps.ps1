# Optional per-slide / theme.animateSteps — entrance timing & order (best-effort COM; WPS support varies)

function Resolve-WpsAnimTriggerTypeInt {
    param([string]$Name)
    if (-not $Name) { return 0 }
    switch ($Name.Trim().ToLowerInvariant()) {
        'click' { return 1 }           # msoAnimTriggerOnPageClick
        'withprevious' { return 2 }    # msoAnimTriggerWithPrevious
        'afterprevious' { return 3 }   # msoAnimTriggerAfterPrevious
        default { return 0 }
    }
}

function Resolve-WpsAnimLevelInt {
    param([string]$By, [object]$ExplicitLevel)
    if ($null -ne $ExplicitLevel) {
        try { return [int]$ExplicitLevel } catch { return 4 }
    }
    $b = if ($By) { $By.Trim().ToLowerInvariant() } else { 'paragraph' }
    if ($b -eq 'shape') { return 1 }      # msoAnimateTextByAllLevels — one block (best-effort)
    return 4                              # legacy driver default (paragraph-style build)
}

function Get-WpsSortedTextShapesOnSlide {
    param($Slide)
    $list = [System.Collections.Generic.List[object]]::new()
    foreach ($j in 1..[int]$Slide.Shapes.Count) {
        try {
            $sh = $Slide.Shapes.Item($j)
            if ($sh.HasTextFrame -ne -1) { continue }
            $list.Add($sh)
        } catch { }
    }
    return @($list | Sort-Object { [double]$_.Top }, { [double]$_.Left })
}

function Get-WpsAnimateStepTargetShapes {
    param(
        $Slide,
        [object]$Step,
        [bool]$TwoColumn,
        [bool]$Narrative
    )
    $raw = [string]$Step.target
    if (-not $raw) { return @() }
    $t = $raw.Trim().ToLowerInvariant()
    $out = [System.Collections.Generic.List[object]]::new()

    switch ($t) {
        'title' {
            try {
                $th = $Slide.Shapes.Title
                if ($th) { $out.Add($th) }
            } catch { }
            if ($out.Count -eq 0) {
                try { $out.Add($Slide.Shapes.Placeholders.Item(1)) } catch { }
            }
        }
        'subtitle' {
            try { $out.Add($Slide.Shapes.Placeholders.Item(2)) } catch { }
        }
        'body' {
            $b = Get-WpsSlideMainBodyShape -Slide $Slide
            if ($b) { $out.Add($b) }
        }
        'leftcolumn' {
            if ($TwoColumn) {
                try { $out.Add($Slide.Shapes.Placeholders.Item(2)) } catch { }
            }
        }
        'rightcolumn' {
            if ($TwoColumn) {
                try { $out.Add($Slide.Shapes.Placeholders.Item(3)) } catch { }
            }
        }
        'placeholder' {
            $pi = 0
            try { $pi = [int]$Step.placeholder } catch { $pi = 0 }
            if ($pi -ge 1) {
                try { $out.Add($Slide.Shapes.Placeholders.Item($pi)) } catch { }
            }
        }
        'shapeindex' {
            $si = 0
            try { $si = [int]$Step.shapeIndex } catch { $si = 0 }
            if ($si -ge 1) {
                try { $out.Add($Slide.Shapes.Item($si)) } catch { }
            }
        }
        'alltextshapes' {
            foreach ($sh in (Get-WpsSortedTextShapesOnSlide -Slide $Slide)) {
                $out.Add($sh)
            }
        }
        default { }
    }
    return @($out)
}

function Apply-WpsAnimationEffectToShape {
    param(
        $Sequence,
        $Shape,
        [int]$EffectId,
        [int]$Level,
        [double]$Duration,
        [double]$Delay,
        [int]$TriggerType,
        [System.Collections.Generic.List[string]]$ComWarnings,
        [string]$WarnPrefix
    )
    try {
        $eff = $Sequence.AddEffect($Shape, $EffectId, $Level, 1)
        try {
            if ($Duration -gt 0.01 -and $Duration -lt 60) {
                $eff.Timing.Duration = $Duration
            }
        } catch { }
        try {
            if ($Delay -ge 0 -and $Delay -lt 120) {
                $eff.Timing.TriggerDelayTime = $Delay
            }
        } catch { }
        try {
            if ($TriggerType -ge 1 -and $TriggerType -le 4) {
                $eff.Timing.TriggerType = $TriggerType
            }
        } catch { }
    } catch {
        $msg = (Format-WpsComException $_)
        $line = "${WarnPrefix}: $msg"
        if ($null -ne $ComWarnings) { [void]$ComWarnings.Add($line) }
        else { Write-Warning $line }
    }
}

function Set-WpsSlideAnimateFromSpecSteps {
    param(
        $Slide,
        [object[]]$Steps,
        [bool]$TwoColumn,
        [bool]$Narrative,
        [string]$DefaultEffectName,
        [int]$SlideNumber = 0,
        [System.Collections.Generic.List[string]]$ComWarnings = $null
    )
    $lab = if ($SlideNumber -gt 0) { "slide[$SlideNumber]" } else { 'slide' }
    $si = 0
    try {
        $seq = $Slide.TimeLine.MainSequence
        foreach ($st in @($Steps)) {
            if ($null -eq $st) { continue }
            $si++
            $targets = Get-WpsAnimateStepTargetShapes -Slide $Slide -Step $st -TwoColumn:$TwoColumn -Narrative:$Narrative
            if ($targets.Count -eq 0) {
                $tn = [string]$st.target
                $w = "${lab} animateSteps[$si]: no shape resolved for target '$tn'"
                if ($null -ne $ComWarnings) { [void]$ComWarnings.Add($w) }
                else { Write-Warning $w }
                continue
            }
            $fxName = $DefaultEffectName
            if ($null -ne $st.effect -and [string]$st.effect) { $fxName = [string]$st.effect }
            $effectId = Resolve-WpsAnimEffectId -Name $fxName
            $level = Resolve-WpsAnimLevelInt -By ([string]$st.by) -ExplicitLevel $st.level
            $dur = 0.0
            if ($null -ne $st.duration) { try { $dur = [double]$st.duration } catch { $dur = 0 } }
            $del = 0.0
            if ($null -ne $st.delay) { try { $del = [double]$st.delay } catch { $del = 0 } }
            $trig = 0
            if ($null -ne $st.trigger -and [string]$st.trigger) {
                $trig = Resolve-WpsAnimTriggerTypeInt -Name ([string]$st.trigger)
            }
            $ti = 0
            foreach ($sh in $targets) {
                $ti++
                $pfx = "${lab} animateSteps[$si] target#$ti ($fxName)"
                Apply-WpsAnimationEffectToShape -Sequence $seq -Shape $sh -EffectId $effectId -Level $level `
                    -Duration $dur -Delay $del -TriggerType $trig -ComWarnings $ComWarnings -WarnPrefix $pfx
            }
        }
    } catch {
        $msg = (Format-WpsComException $_)
        $line = "${lab} TimeLine.MainSequence (animateSteps): $msg"
        if ($null -ne $ComWarnings) { [void]$ComWarnings.Add($line) }
        else { Write-Warning $line }
    }
}
