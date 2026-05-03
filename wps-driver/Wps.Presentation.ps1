# Presentation lifecycle + session binding

function Normalize-FsPath {
    param([string]$Path)
    try {
        return [System.IO.Path]::GetFullPath($Path).TrimEnd('\').ToLowerInvariant()
    } catch {
        return $Path.TrimEnd('\').ToLowerInvariant()
    }
}

function Get-WpsPresentationFromSession {
    param($App)
    $sess = Read-WpsSession
    if (-not $sess -or -not $sess.presentationPath) {
        throw 'No active presentation in session. Run action "new" or "open" first.'
    }
    $path = [string]$sess.presentationPath
    if (-not (Test-Path -LiteralPath $path)) {
        throw "Presentation file missing: $path"
    }
    $abs = (Resolve-Path -LiteralPath $path).Path
    $want = Normalize-FsPath $abs
    $count = [int]$App.Presentations.Count
    for ($i = 1; $i -le $count; $i++) {
        try {
            $p = $App.Presentations.Item($i)
            $full = $null
            if ($p.Path -and $p.Name) {
                $full = Join-Path $p.Path $p.Name
            } elseif ($p.FullName) {
                $full = [string]$p.FullName
            }
            if ($full -and (Test-Path -LiteralPath $full)) {
                $got = Normalize-FsPath ((Resolve-Path -LiteralPath $full).Path)
                if ($got -eq $want) { return $p }
            }
        } catch { }
    }
    return $App.Presentations.Open($abs, $false, $false, $true)
}

function New-WpsPresentation {
    param($App, [string]$SavePath)
    $pres = $App.Presentations.Add()
    if ($SavePath) {
        $dir = Split-Path -Parent $SavePath
        if ($dir) { Ensure-Dir $dir }
        $pres.SaveAs($SavePath)
        Write-WpsSession @{ presentationPath = (Resolve-Path -LiteralPath $SavePath).Path }
    }
    return $pres
}

function Open-WpsPresentation {
    param($App, [string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { throw "File not found: $Path" }
    $abs = (Resolve-Path -LiteralPath $Path).Path
    $pres = $App.Presentations.Open($abs, $false, $false, $true)
    Write-WpsSession @{ presentationPath = $abs }
    return $pres
}

function Save-WpsPresentation {
    param($Pres, [string]$Path)
    if ($Path) {
        $dir = Split-Path -Parent $Path
        if ($dir) { Ensure-Dir $dir }
        $Pres.SaveAs($Path)
        Write-WpsSession @{ presentationPath = (Resolve-Path -LiteralPath $Path).Path }
    } else {
        $Pres.Save()
    }
}

function Close-WpsPresentation {
    param($Pres, [bool]$Save)
    if ($Save) { $Pres.Save() }
    $Pres.Close()
}
