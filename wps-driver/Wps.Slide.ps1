# Add slides with layout constants (PowerPoint-compatible enums)

function Add-WpsSlide {
    param($Presentation, [int]$LayoutInt)
    $next = [int]$Presentation.Slides.Count + 1
    return $Presentation.Slides.Add($next, $LayoutInt)
}
