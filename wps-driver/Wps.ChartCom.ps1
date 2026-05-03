<#
.SYNOPSIS
  Optional embedded chart from chart_data via WPS COM (best-effort).
.NOTES
  Requires a build that exposes Chart.ChartData / Excel workbook behind the chart.
  On failure, caller should fall back to PNG rendering.
#>

function Add-WpsEmbeddedChartFromData {
    param(
        $Slide,
        [string]$ChartType,
        [object]$ChartData
    )
    if (-not $Slide) { throw 'Slide is null' }
    if (-not $ChartData) { throw 'ChartData is null' }

    $labels = @($ChartData.labels | ForEach-Object { [string]$_ })
    $seriesIn = @($ChartData.series)
    if ($labels.Count -lt 1) { throw 'chart_data.labels required' }
    if ($seriesIn.Count -lt 1) { throw 'chart_data.series required' }

    foreach ($row in $seriesIn) {
        if (@($row.values).Count -ne $labels.Count) {
            throw 'each series values length must match labels'
        }
    }

    $typeKey = $ChartType.ToLowerInvariant()
    $xl = @{
        'bar'   = 51
        'line'  = 4
        'donut' = -4120
    }
    if (-not $xl.ContainsKey($typeKey)) {
        throw "COM chart supports bar|line|donut, not $ChartType"
    }
    $chartTypeEnum = $xl[$typeKey]

    $shape = $null
    try {
        $shape = $Slide.Shapes.AddChart()
    } catch {
        try {
            $shape = $Slide.Shapes.AddChart2(201, $chartTypeEnum)
        } catch {
            throw "AddChart/AddChart2 failed: $($_.Exception.Message)"
        }
    }

    $chart = $shape.Chart
    if (-not $chart) { throw 'Chart object missing' }
    $chart.ChartType = $chartTypeEnum

    $cd = $chart.ChartData
    if (-not $cd) { throw 'ChartData not available on this WPS build' }
    $cd.Activate() | Out-Null
    $wb = $cd.Workbook
    $ws = $wb.Worksheets.Item(1)

    $ws.Cells.Item(1, 1) = ''
    $numSeries = $seriesIn.Count
    for ($si = 0; $si -lt $numSeries; $si++) {
        $row = $seriesIn[$si]
        $nm = if ($row.name) { [string]$row.name } else { "Series $($si + 1)" }
        $ws.Cells.Item(1, 2 + $si) = $nm
    }
    for ($r = 0; $r -lt $labels.Count; $r++) {
        $ws.Cells.Item(2 + $r, 1) = $labels[$r]
        for ($si = 0; $si -lt $numSeries; $si++) {
            $vals = @($seriesIn[$si].values)
            $ws.Cells.Item(2 + $r, 2 + $si) = $vals[$r]
        }
    }

    try { $chart.Refresh() } catch { }
    return $shape
}
