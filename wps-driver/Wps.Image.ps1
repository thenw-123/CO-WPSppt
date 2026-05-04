# Insert picture on slide (local file path)

function Resolve-WpsSlideImagePlacement {
    param(
        [object]$SlideSpec,
        [double]$SlideWidth = 960.0
    )
    $hasCustom = $false
    if ($null -ne $SlideSpec) {
        foreach ($k in @('imageLeft', 'imageTop', 'imageWidth', 'imageHeight')) {
            $v = $SlideSpec.$k
            if ($null -ne $v) { $hasCustom = $true; break }
        }
    }

    $mode = 'right'
    if ($SlideSpec -and $SlideSpec.imagePlacement) {
        $mp = [string]$SlideSpec.imagePlacement
        if ($mp.ToLowerInvariant() -eq 'bottom') { $mode = 'bottom' }
    }

    $rr = 408.0
    if ($SlideSpec -and $null -ne $SlideSpec.imageReservedRight) {
        try { $rr = [double]$SlideSpec.imageReservedRight } catch { }
    }

    if (-not $hasCustom) {
        if ($mode -eq 'bottom') {
            $m = 48.0
            $L = $m
            $T = [Math]::Max(292.0, $SlideWidth * 0.28)
            $W = $SlideWidth - (2 * $m)
            $H = 198.0
            return @{ Left = $L; Top = $T; Width = $W; Height = $H }
        }
        $imgW = [Math]::Min(372.0, $rr - 28)
        $L = $SlideWidth - $imgW - 42.0
        $T = 108.0
        $W = $imgW
        $H = 312.0
        return @{ Left = $L; Top = $T; Width = $W; Height = $H }
    }

    $L = 492.0; $T = 118.0; $W = 430.0; $H = 375.0
    if ($null -eq $SlideSpec) {
        return @{ Left = $L; Top = $T; Width = $W; Height = $H }
    }
    try { if ($null -ne $SlideSpec.imageLeft) { $L = [double]$SlideSpec.imageLeft } } catch { }
    try { if ($null -ne $SlideSpec.imageTop) { $T = [double]$SlideSpec.imageTop } } catch { }
    try { if ($null -ne $SlideSpec.imageWidth) { $W = [double]$SlideSpec.imageWidth } } catch { }
    try { if ($null -ne $SlideSpec.imageHeight) { $H = [double]$SlideSpec.imageHeight } } catch { }
    return @{ Left = $L; Top = $T; Width = $W; Height = $H }
}

function Add-WpsSlidePicture {
    param(
        $Slide,
        [string]$ImagePath,
        [double]$Left = 80,
        [double]$Top = 220,
        [double]$Width = 480,
        [double]$Height = 280
    )
    if (-not (Test-Path -LiteralPath $ImagePath)) {
        throw "Image not found: $ImagePath"
    }
    $abs = (Resolve-Path -LiteralPath $ImagePath).Path
    # msoFalse = 0, msoTrue = -1
    return $Slide.Shapes.AddPicture($abs, $false, -1, $Left, $Top, $Width, $Height)
}

function Add-WpsSlidePictureFromSpec {
    param(
        $Slide,
        [string]$ImagePath,
        [object]$SlideSpec,
        [double]$SlideWidth = 960.0
    )
    $g = Resolve-WpsSlideImagePlacement -SlideSpec $SlideSpec -SlideWidth $SlideWidth
    $pic = Add-WpsSlidePicture -Slide $Slide -ImagePath $ImagePath -Left $g.Left -Top $g.Top -Width $g.Width -Height $g.Height
    $z = 'back'
    if ($SlideSpec -and $null -ne $SlideSpec.imageZOrder) {
        $iz = [string]$SlideSpec.imageZOrder
        if ($iz.Trim().ToLowerInvariant() -eq 'front') { $z = 'front' }
    }
    if ($z -eq 'back') {
        try {
            # msoSendToBack = 1 — keeps text/placeholders above the photo when shapes overlap
            $pic.ZOrder(1)
        } catch { }
    }
    return $pic
}
