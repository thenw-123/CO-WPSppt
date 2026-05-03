<#
.SYNOPSIS
  Invoke matplotlib (headless) to produce chart PNGs for slides.
#>

function Get-ChartPythonCommand {
    if (Get-Command python -ErrorAction SilentlyContinue) {
        return , @('python')
    }
    if (Get-Command py -ErrorAction SilentlyContinue) {
        return , @('py', '-3')
    }
    if ($env:PYTHON -and (Test-Path -LiteralPath $env:PYTHON)) {
        return , @([string]$env:PYTHON)
    }
    return $null
}

function Invoke-PptChartPngRender {
    param(
        [string]$ProjectRoot,
        [string]$OutPngPath,
        [string]$ChartType,
        [object]$ChartData
    )
    $scriptPath = Join-Path $ProjectRoot 'tools\chart_to_png.py'
    if (-not (Test-Path -LiteralPath $scriptPath)) {
        throw "chart_to_png.py not found: $scriptPath"
    }

    $cmd = Get-ChartPythonCommand
    if (-not $cmd) {
        throw 'Python not found on PATH (install Python 3 + pip install -r requirements-charts.txt)'
    }

    $cfg = [ordered]@{
        chart_type  = $ChartType
        chart_data  = $ChartData
    }
    $tmp = [System.IO.Path]::GetTempFileName() + '.json'
    try {
        ($cfg | ConvertTo-Json -Depth 8 -Compress) | Set-Content -LiteralPath $tmp -Encoding UTF8
        $dir = Split-Path -Parent $OutPngPath
        if ($dir) { Ensure-Dir $dir }

        $tail = if ($cmd.Length -gt 1) { @($cmd[1..($cmd.Length - 1)]) } else { @() }
        & $cmd[0] @($tail + @($scriptPath, $OutPngPath, $tmp))
        if ($LASTEXITCODE -ne 0) {
            throw "chart_to_png.py exited with code $LASTEXITCODE"
        }
        if (-not (Test-Path -LiteralPath $OutPngPath)) {
            throw "Chart PNG was not created: $OutPngPath"
        }
    } finally {
        if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue }
    }
}
