# Atomic WPS renderer. It executes RenderPlan operations only; semantic DSL
# decisions must already be resolved by Wps.LayoutEngine.ps1.

function Get-WpsAtomicFontSize {
    param([string]$Role)
    switch (($Role | ForEach-Object { if ($_){ $_.ToLowerInvariant() } else { 'body' } })) {
        'title' { return 34.0 }
        'subtitle' { return 20.0 }
        'claim' { return 28.0 }
        'insight' { return 18.0 }
        default { return 20.0 }
    }
}

function Get-WpsAtomicEstimatedLineCount {
    param(
        [string]$Text,
        [double]$FontSize,
        [double]$BoxWidth
    )
    if (-not $Text) { return 1.0 }
    $charsPerLine = [math]::Max(8.0, [math]::Floor($BoxWidth / [math]::Max(1.0, ($FontSize * 0.55))))
    $lineCount = 0.0
    foreach ($line in @(([string]$Text -split "(\r\n|\n)"))) {
        if ($line -eq "`r`n" -or $line -eq "`n") { continue }
        $len = [double]([string]$line).Length
        if ($len -le 0) { $lineCount += 1.0; continue }
        $lineCount += [math]::Ceiling($len / $charsPerLine)
    }
    return [math]::Max(1.0, $lineCount)
}

function Get-WpsAtomicFittedFontSize {
    param(
        [string]$Text,
        $Box,
        [string]$Role,
        [object]$Style
    )
    $base = if ($Style -and $null -ne $Style.fontSize) { [double]$Style.fontSize } else { Get-WpsAtomicFontSize -Role $Role }
    $min = if ($Role -eq 'title') { 16.0 } else { 12.0 }
    if ($Style -and $null -ne $Style.minFontSize) {
        try { $min = [double]$Style.minFontSize } catch { }
    }
    $fs = $base
    while ($fs -gt $min) {
        $lines = Get-WpsAtomicEstimatedLineCount -Text $Text -FontSize $fs -BoxWidth ([double]$Box.w)
        $needH = $lines * ($fs * 1.35)
        if ($needH -le ([double]$Box.h)) { break }
        $fs -= 1.0
    }
    return [math]::Round([math]::Max($min, $fs), 1)
}

function Convert-WpsAtomicHexToOfficeRgbLong {
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

function Add-WpsAtomicRoundedRect {
    param($Slide, $Operation)
    # msoShapeRoundedRectangle = 5
    $b = $Operation.box
    $shape = $Slide.Shapes.AddShape(5, [double]$b.x, [double]$b.y, [double]$b.w, [double]$b.h)
    $fillRgb = Convert-WpsAtomicHexToOfficeRgbLong -Hex ([string]$Operation.fill)
    try {
        if ($null -ne $fillRgb) {
            $shape.Fill.Visible = -1
            $shape.Fill.Solid()
            $shape.Fill.ForeColor.RGB = $fillRgb
        } else {
            try { $shape.Fill.Visible = 0 } catch { }
        }
    } catch { }

    $strokeRgb = Convert-WpsAtomicHexToOfficeRgbLong -Hex ([string]$Operation.stroke)
    try {
        if (($Operation.strokeVisible -eq $true) -and ($null -ne $strokeRgb)) {
            $shape.Line.Visible = -1
            $shape.Line.ForeColor.RGB = $strokeRgb
            $shape.Line.Weight = 1.0
        } else {
            try { $shape.Line.Visible = 0 } catch { }
        }
    } catch { }

    $adj = 0.12
    if ($null -ne $Operation.cornerAdjustment) {
        try { $adj = [double]$Operation.cornerAdjustment } catch { }
    }
    try { $shape.Adjustments.Item(1) = $adj } catch { }
    return $shape
}

function Get-WpsAtomicContainImageBox {
    param($Box, [double]$ImageWidth, [double]$ImageHeight)
    if ($ImageWidth -le 0 -or $ImageHeight -le 0) { return $Box }
    $boxRatio = [double]$Box.w / [double]$Box.h
    $imgRatio = $ImageWidth / $ImageHeight
    $drawX = [double]$Box.x
    $drawY = [double]$Box.y
    $drawW = [double]$Box.w
    $drawH = [double]$Box.h
    if ($imgRatio -gt $boxRatio) {
        $drawH = $drawW / $imgRatio
        $drawY = [double]$Box.y + (([double]$Box.h - $drawH) / 2.0)
    } else {
        $drawW = $drawH * $imgRatio
        $drawX = [double]$Box.x + (([double]$Box.w - $drawW) / 2.0)
    }
    return [ordered]@{
        x = $drawX
        y = $drawY
        w = $drawW
        h = $drawH
    }
}

function Add-WpsAtomicText {
    param($Slide, $Operation)
    $b = $Operation.box
    $shape = $Slide.Shapes.AddTextbox(1, [double]$b.x, [double]$b.y, [double]$b.w, [double]$b.h)
    $text = [string]$Operation.text
    $shape.TextFrame.TextRange.Text = $text
    $size = Get-WpsAtomicFittedFontSize -Text $text -Box $b -Role ([string]$Operation.role) -Style $Operation.style
    try { $shape.TextFrame.TextRange.Font.Size = $size } catch { }
    try { $shape.TextFrame.WordWrap = -1 } catch { }
    try {
        if ($Operation.role -in @('title', 'claim')) { $shape.TextFrame.TextRange.Font.Bold = -1 }
    } catch { }
    return $shape
}

function Add-WpsAtomicBullets {
    param($Slide, $Operation)
    $b = $Operation.box
    $shape = $Slide.Shapes.AddTextbox(1, [double]$b.x, [double]$b.y, [double]$b.w, [double]$b.h)
    $lines = @($Operation.items | ForEach-Object { [string]$_ })
    $text = ($lines -join "`r`n")
    $shape.TextFrame.TextRange.Text = $text
    $size = Get-WpsAtomicFittedFontSize -Text $text -Box $b -Role ([string]$Operation.role) -Style $Operation.style
    try { $shape.TextFrame.TextRange.Font.Size = $size } catch { }
    try { $shape.TextFrame.WordWrap = -1 } catch { }
    try { $shape.TextFrame.TextRange.ParagraphFormat.Bullet.Visible = -1 } catch { }
    return $shape
}

function Resolve-WpsAtomicImagePath {
    param(
        $Operation,
        [string]$ProjectRoot,
        [string]$AssetsDir,
        $AssetFetchOptions
    )
    $src = $Operation.source
    if ($null -eq $src) { throw "image operation missing source: $($Operation.elementId)" }
    $kind = if ($src.kind) { ([string]$src.kind).ToLowerInvariant() } else { 'local' }
    switch ($kind) {
        'local' {
            $p = [string]$src.path
            if (-not [System.IO.Path]::IsPathRooted($p)) { $p = Join-Path $ProjectRoot $p }
            if (-not (Test-Path -LiteralPath $p)) { throw "Image not found: $p" }
            return (Resolve-Path -LiteralPath $p).Path
        }
        'remote' {
            return Save-RemoteImageForSlide -UrlString ([string]$src.url) -DestDir $AssetsDir -FetchOptions $AssetFetchOptions
        }
        'commons' {
            $out = Join-Path $AssetsDir ('commons-{0}.jpg' -f ([Guid]::NewGuid().ToString('N').Substring(0, 10)))
            Save-WpsCommonsThumbFile -FileTitle ([string]$src.title) -OutPath $out -Width 500
            return $out
        }
        default {
            throw "Unsupported image source kind for atomic renderer: $kind"
        }
    }
}

function Add-WpsAtomicImage {
    param(
        $Slide,
        $Operation,
        [string]$ProjectRoot,
        [string]$AssetsDir,
        $AssetFetchOptions
    )
    $path = Resolve-WpsAtomicImagePath -Operation $Operation -ProjectRoot $ProjectRoot -AssetsDir $AssetsDir -AssetFetchOptions $AssetFetchOptions
    $b = $Operation.box
    $style = $Operation.style
    $fitMode = 'contain'
    if ($style -and $style.imageFit) { $fitMode = ([string]$style.imageFit).ToLowerInvariant() }
    if ($fitMode -eq 'fill') {
        return Add-WpsSlidePicture -Slide $Slide -ImagePath $path -Left ([double]$b.x) -Top ([double]$b.y) -Width ([double]$b.w) -Height ([double]$b.h)
    }
    try {
        Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue | Out-Null
        $img = [System.Drawing.Image]::FromFile($path)
        try {
            $iw = [double]$img.Width
            $ih = [double]$img.Height
        } finally {
            $img.Dispose()
        }
        if ($iw -gt 0 -and $ih -gt 0 -and [double]$b.w -gt 0 -and [double]$b.h -gt 0) {
            $fit = Get-WpsAtomicContainImageBox -Box $b -ImageWidth $iw -ImageHeight $ih
            return Add-WpsSlidePicture -Slide $Slide -ImagePath $path -Left ([double]$fit.x) -Top ([double]$fit.y) -Width ([double]$fit.w) -Height ([double]$fit.h)
        }
    } catch { }
    return Add-WpsSlidePicture -Slide $Slide -ImagePath $path -Left ([double]$b.x) -Top ([double]$b.y) -Width ([double]$b.w) -Height ([double]$b.h)
}

function Invoke-WpsAtomicOperation {
    param(
        $Slide,
        $Operation,
        [string]$ProjectRoot,
        [string]$AssetsDir,
        $AssetFetchOptions
    )
    switch ([string]$Operation.op) {
        'addText' { Add-WpsAtomicText -Slide $Slide -Operation $Operation | Out-Null }
        'addBullets' { Add-WpsAtomicBullets -Slide $Slide -Operation $Operation | Out-Null }
        'addImage' { Add-WpsAtomicImage -Slide $Slide -Operation $Operation -ProjectRoot $ProjectRoot -AssetsDir $AssetsDir -AssetFetchOptions $AssetFetchOptions | Out-Null }
        'addRoundedRect' { Add-WpsAtomicRoundedRect -Slide $Slide -Operation $Operation | Out-Null }
        default { throw "Unknown atomic render operation: $($Operation.op)" }
    }
}

function Invoke-WpsDeckFromAtomicRenderPlan {
    param(
        [object]$RenderPlan,
        [string]$ProjectRoot
    )

    $out = $RenderPlan.meta.savePath
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
    $assetsDir = Join-Path $ProjectRoot ("output\assets\{0}" -f $deckSlug)
    Ensure-Dir $assetsDir
    $assetOpts = Get-AssetFetchOptions -SpecAssetFetch $null

    $app = Get-WpsApplication
    $pres = New-WpsPresentation -App $app -SavePath $out
    while ([int]$pres.Slides.Count -lt 1) { Add-WpsSlide -Presentation $pres -LayoutInt 12 | Out-Null }

    $idx = 0
    foreach ($slidePlan in @($RenderPlan.slides)) {
        $idx++
        if ($idx -eq 1) {
            $slide = $pres.Slides.Item(1)
            try { $slide.Layout = 12 } catch { }
        } else {
            Add-WpsSlide -Presentation $pres -LayoutInt 12 | Out-Null
            $slide = $pres.Slides.Item($pres.Slides.Count)
        }
        foreach ($op in @($slidePlan.operations)) {
            Invoke-WpsAtomicOperation -Slide $slide -Operation $op -ProjectRoot $ProjectRoot -AssetsDir $assetsDir -AssetFetchOptions $assetOpts
        }
        if ($slidePlan.notes) { Set-WpsSlideNotes -Slide $slide -Text ([string]$slidePlan.notes) }
    }

    Save-WpsPresentation -Pres $pres -Path $out
    return @{
        path             = (Resolve-Path -LiteralPath $out).Path
        slideCount       = [int]$pres.Slides.Count
        presentationPath = (Resolve-Path -LiteralPath $out).Path
    }
}
