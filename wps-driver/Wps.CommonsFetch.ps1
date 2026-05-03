# Download Wikimedia Commons thumbnails via the MediaWiki API (standard iiurlwidth steps; see https://w.wiki/GHai).

$script:CommonsApiUserAgent = 'OpenClawWpsPpt/1.4 (educational deck; +https://www.mediawiki.org/wiki/API:Etiquette)'

function Normalize-WpsCommonsFileTitle {
    param([string]$Raw)
    if (-not $Raw) { return $null }
    $t = $Raw.Trim()
    if (-not $t.StartsWith('File:', [StringComparison]::OrdinalIgnoreCase)) {
        $t = 'File:' + $t
    }
    return $t
}

function Resolve-WpsCommonsThumbWidth {
    param([int]$Requested)
    $steps = @(20, 40, 60, 120, 250, 330, 500, 960, 1280, 1920, 3840)
    foreach ($s in $steps) {
        if ($s -ge $Requested) { return $s }
    }
    return 3840
}

function Get-WpsCommonsThumbUrl {
    param(
        [string]$FileTitle,
        [int]$Width = 500
    )
    $title = Normalize-WpsCommonsFileTitle -Raw $FileTitle
    if (-not $title) { throw 'imageCommons: empty file title' }
    $w = Resolve-WpsCommonsThumbWidth -Requested $Width
    $encoded = [Uri]::EscapeDataString($title)
    $api = "https://commons.wikimedia.org/w/api.php?action=query&titles=$encoded&prop=imageinfo&iiprop=url&iiurlwidth=$w&format=json"
    $resp = Invoke-RestMethod -Uri $api -Headers @{ 'User-Agent' = $script:CommonsApiUserAgent } -Method Get
    if (-not $resp.query -or -not $resp.query.pages) {
        throw 'Commons API: empty query'
    }
    foreach ($prop in $resp.query.pages.PSObject.Properties) {
        $page = $prop.Value
        $ii = $page.imageinfo
        if (-not $ii -or $ii.Count -lt 1) {
            throw "Commons: file not found or not an image ($title)"
        }
        $tu = [string]$ii[0].thumburl
        if (-not $tu) { throw "Commons: no thumburl for $title" }
        return $tu.Split('?')[0]
    }
    throw 'Commons: no pages in response'
}

function Save-WpsCommonsThumbFile {
    param(
        [string]$FileTitle,
        [string]$OutPath,
        [int]$Width = 500
    )
    $url = Get-WpsCommonsThumbUrl -FileTitle $FileTitle -Width $Width
    $uri = [Uri]::new($url)
    if ($uri.Host -ne 'upload.wikimedia.org') {
        throw "Commons: unexpected download host $($uri.Host)"
    }
    $dir = Split-Path -Parent $OutPath
    if ($dir) { Ensure-Dir $dir }

    $req = [System.Net.HttpWebRequest]::Create($url)
    $req.Method = 'GET'
    $req.UserAgent = $script:CommonsApiUserAgent
    $req.Timeout = 90000
    try {
        $resp = $req.GetResponse()
    } catch {
        throw "Commons download failed: $($_.Exception.Message)"
    }
    try {
        $stream = $resp.GetResponseStream()
        $fs = [System.IO.File]::Create($OutPath)
        try {
            $stream.CopyTo($fs)
        } finally {
            $fs.Close()
            $stream.Close()
        }
    } finally {
        $resp.Close()
    }
}
