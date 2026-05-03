<#
.SYNOPSIS
  Secure download of remote images for slide assets (SSRF-hardened).
.DESCRIPTION
  - HTTPS only by default (opt-in HTTP via env OC_ASSET_FETCH_ALLOW_HTTP=1, dev only).
  - Resolves DNS and rejects private/link-local/CGNAT/metadata-range targets.
  - Optional spec assetFetch.allowedHosts: exact hostname allow-list (case-insensitive).
  - Manual redirect handling with re-validation per hop.
  - Size cap and image/* content-type allow-list.
#>

$script:AssetFetchDefaults = @{
    MaxBytesDefault     = 5242880   # 5 MiB
    TimeoutSecDefault   = 30
    MaxRedirectsDefault = 5
}

$script:AllowedImageMimeToExt = [ordered]@{
    'image/png'       = '.png'
    'image/jpeg'      = '.jpg'
    'image/jpg'       = '.jpg'
    'image/gif'       = '.gif'
    'image/webp'      = '.webp'
}

function Get-AssetFetchOptions {
    param([object]$SpecAssetFetch)
    $af = $SpecAssetFetch
    $max = $script:AssetFetchDefaults.MaxBytesDefault
    $to = $script:AssetFetchDefaults.TimeoutSecDefault
    $redir = $script:AssetFetchDefaults.MaxRedirectsDefault
    $hosts = [string[]]@()
    $softFail = $false
    if ($af) {
        if ($null -ne $af.maxBytes) {
            try { $max = [int]$af.maxBytes } catch { $max = $script:AssetFetchDefaults.MaxBytesDefault }
            if ($max -lt 1) { $max = 1 }
            if ($max -gt 25MB) { $max = 25MB }
        }
        if ($null -ne $af.timeoutSec) {
            try { $to = [int]$af.timeoutSec } catch { $to = $script:AssetFetchDefaults.TimeoutSecDefault }
            if ($to -lt 1) { $to = 1 }
            if ($to -gt 120) { $to = 120 }
        }
        if ($null -ne $af.maxRedirects) {
            try { $redir = [int]$af.maxRedirects } catch { $redir = $script:AssetFetchDefaults.MaxRedirectsDefault }
            if ($redir -lt 0) { $redir = 0 }
            if ($redir -gt 10) { $redir = 10 }
        }
        if ($af.allowedHosts) {
            $hosts = @($af.allowedHosts | ForEach-Object { [string]$_ }) | Where-Object { $_ }
        }
        if ($null -ne $af.softFail) {
            try { $softFail = [bool]$af.softFail } catch { $softFail = $false }
        }
    }
    return @{
        MaxBytes      = $max
        TimeoutSec    = $to
        MaxRedirects  = $redir
        AllowedHosts  = $hosts
        AllowInsecure = ($env:OC_ASSET_FETCH_ALLOW_HTTP -eq '1')
        SoftFail      = $softFail
    }
}

function Test-IpEndpointRestricted {
    param([System.Net.IPAddress]$Addr)
    if ($null -eq $Addr) { return $true }
    if ([System.Net.IPAddress]::IsLoopback($Addr)) { return $true }
    if ($Addr.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork) {
        $b = $Addr.GetAddressBytes()
        if ($b[0] -eq 0) { return $true }
        if ($b[0] -eq 10) { return $true }
        if ($b[0] -eq 127) { return $true }
        if ($b[0] -eq 169 -and $b[1] -eq 254) { return $true }
        if ($b[0] -eq 172 -and $b[1] -ge 16 -and $b[1] -le 31) { return $true }
        if ($b[0] -eq 192 -and $b[1] -eq 168) { return $true }
        if ($b[0] -eq 100 -and $b[1] -ge 64 -and $b[1] -le 127) { return $true }
        return $false
    }
    if ($Addr.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetworkV6) {
        if ($Addr.IsIPv6LinkLocal) { return $true }
        if ($Addr.IsIPv6SiteLocal) { return $true }
        $x = $Addr.GetAddressBytes()
        if ($x[0] -eq 0xfc -or $x[0] -eq 0xfd) { return $true }
        if ($x[0] -eq 0xfe -and ($x[1] -band 0xc0) -eq 0x80) { return $true }
        return $false
    }
    return $true
}

function Test-UriSafeForAssetFetch {
    param(
        [System.Uri]$Uri,
        [string[]]$AllowedHosts,
        [bool]$AllowInsecureHttp
    )
    if ($null -eq $Uri) { return @{ Ok = $false; Reason = 'uri is null' } }
    if ($Uri.Scheme -eq 'https') { }
    elseif ($Uri.Scheme -eq 'http') {
        if (-not $AllowInsecureHttp) {
            return @{ Ok = $false; Reason = 'http blocked; use https or set OC_ASSET_FETCH_ALLOW_HTTP=1 (dev only)' }
        }
    }
    else {
        return @{ Ok = $false; Reason = 'only http and https schemes are allowed' }
    }
    if (-not [string]::IsNullOrEmpty($Uri.UserInfo)) {
        return @{ Ok = $false; Reason = 'credentials in URL are not allowed' }
    }
    $hostName = $Uri.IdnHost
    if ([string]::IsNullOrWhiteSpace($hostName)) {
        return @{ Ok = $false; Reason = 'missing host' }
    }
    $blocked = @('localhost', '127.0.0.1', '::1', 'metadata.google.internal', 'metadata', 'wpad')
    foreach ($b in $blocked) {
        if ($hostName.Equals($b, [StringComparison]::OrdinalIgnoreCase)) {
            return @{ Ok = $false; Reason = "host '$hostName' is blocked" }
        }
    }
    if ($AllowedHosts -and $AllowedHosts.Count -gt 0) {
        $match = $false
        foreach ($h in $AllowedHosts) {
            if ($hostName.Equals([string]$h, [StringComparison]::OrdinalIgnoreCase)) {
                $match = $true
                break
            }
        }
        if (-not $match) {
            return @{ Ok = $false; Reason = "host not in assetFetch.allowedHosts: $hostName" }
        }
    }
    try {
        $addrs = [System.Net.Dns]::GetHostAddresses($hostName)
    } catch {
        return @{ Ok = $false; Reason = ('dns resolution failed: ' + $_.Exception.Message) }
    }
    if (-not $addrs -or $addrs.Count -lt 1) {
        return @{ Ok = $false; Reason = 'no addresses resolved' }
    }
    foreach ($a in $addrs) {
        if (Test-IpEndpointRestricted -Addr $a) {
            return @{ Ok = $false; Reason = ('host resolves to restricted address: ' + $a.ToString()) }
        }
    }
    return @{ Ok = $true; Reason = $null }
}

function Get-ExtensionFromContentType {
    param([string]$ContentType)
    if (-not $ContentType) { return $null }
    $ct = $ContentType.Split(';')[0].Trim().ToLowerInvariant()
    if ($script:AllowedImageMimeToExt.Contains($ct)) {
        return [string]$script:AllowedImageMimeToExt[$ct]
    }
    return $null
}

function Save-RemoteImageForSlide {
    param(
        [string]$UrlString,
        [string]$DestDir,
        [hashtable]$FetchOptions
    )
    if (-not (Test-Path -LiteralPath $DestDir)) {
        New-Item -ItemType Directory -Path $DestDir -Force | Out-Null
    }
    $current = $null
    try {
        $current = [Uri]::new($UrlString)
    } catch {
        throw "Invalid imageUrl: $($_.Exception.Message)"
    }

    $hop = 0
    $maxHops = [int]$FetchOptions.MaxRedirects + 1
    while ($hop -lt $maxHops) {
        $hop++
        $chk = Test-UriSafeForAssetFetch -Uri $current -AllowedHosts $FetchOptions.AllowedHosts -AllowInsecureHttp $FetchOptions.AllowInsecure
        if (-not $chk.Ok) {
            throw "URL blocked (SSRF policy): $($chk.Reason)"
        }

        $req = [System.Net.HttpWebRequest]::Create($current)
        $req.Method = 'GET'
        $req.Timeout = [int]$FetchOptions.TimeoutSec * 1000
        $req.AllowAutoRedirect = $false
        $req.UserAgent = 'cursor-openclaw-wps-ppt/1.4 asset-fetch'
        try {
            $resp = $req.GetResponse()
        } catch {
            throw "HTTP request failed: $($_.Exception.Message)"
        }
        try {
            $code = [int]$resp.StatusCode
            if ($code -ge 300 -and $code -lt 400) {
                $loc = $resp.Headers['Location']
                if (-not $loc) { throw "Redirect without Location header" }
                $next = [Uri]::new($loc, [UriKind]::RelativeOrAbsolute)
                if (-not $next.IsAbsoluteUri) {
                    $next = [Uri]::new($current, $next)
                }
                $current = $next
                continue
            }
            if ($code -lt 200 -or $code -ge 300) {
                throw "HTTP status $code"
            }

            $ct = [string]$resp.ContentType
            $ext = Get-ExtensionFromContentType -ContentType $ct
            if (-not $ext) {
                throw "disallowed Content-Type: $ct (allowed: $($script:AllowedImageMimeToExt.Keys -join ', '))"
            }

            $len = [long]$resp.ContentLength
            if ($len -gt 0 -and $len -gt [long]$FetchOptions.MaxBytes) {
                throw "Content-Length $len exceeds maxBytes $($FetchOptions.MaxBytes)"
            }

            $stream = $resp.GetResponseStream()
            $ms = New-Object System.IO.MemoryStream
            try {
                $buf = New-Object byte[] 8192
                $read = 0L
                while (($n = $stream.Read($buf, 0, $buf.Length)) -gt 0) {
                    $read += $n
                    if ($read -gt [long]$FetchOptions.MaxBytes) {
                        throw "response exceeds maxBytes $($FetchOptions.MaxBytes)"
                    }
                    $ms.Write($buf, 0, $n)
                }
            } finally {
                $stream.Close()
            }

            $name = [guid]::NewGuid().ToString('n') + $ext
            $path = Join-Path $DestDir $name
            [System.IO.File]::WriteAllBytes($path, $ms.ToArray())
            return (Resolve-Path -LiteralPath $path).Path
        } finally {
            if ($resp) { $resp.Close() }
        }
    }
    throw "too many redirects (max $($FetchOptions.MaxRedirects))"
}
