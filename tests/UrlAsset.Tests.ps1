# SSRF / URL policy tests (no network). Run: Invoke-Pester .\tests\UrlAsset.Tests.ps1

$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
. (Join-Path $RepoRoot 'wps-driver\Wps.UrlAsset.ps1')

Describe 'Test-UriSafeForAssetFetch' {
    It 'blocks non-http(s) schemes' {
        $u = [Uri]'file:///C:/secret.png'
        $r = Test-UriSafeForAssetFetch -Uri $u -AllowedHosts @() -AllowInsecureHttp $false
        $r.Ok | Should Be $false
    }

    It 'blocks credentials in URL' {
        $u = [Uri]'https://user:pass@example.com/x.png'
        $r = Test-UriSafeForAssetFetch -Uri $u -AllowedHosts @() -AllowInsecureHttp $false
        $r.Ok | Should Be $false
    }

    It 'blocks http without env flag' {
        $u = [Uri]'http://example.com/x.png'
        $r = Test-UriSafeForAssetFetch -Uri $u -AllowedHosts @() -AllowInsecureHttp $false
        $r.Ok | Should Be $false
    }

    It 'enforces allowedHosts when set' {
        $u = [Uri]'https://evil.com/x.png'
        $r = Test-UriSafeForAssetFetch -Uri $u -AllowedHosts @('good.com') -AllowInsecureHttp $false
        $r.Ok | Should Be $false
    }

    It 'allows host in allow-list' {
        $u = [Uri]'https://good.com/x.png'
        $r = Test-UriSafeForAssetFetch -Uri $u -AllowedHosts @('good.com') -AllowInsecureHttp $false
        $r.Ok | Should Be $true
    }
}

Describe 'Test-IpEndpointRestricted' {
    It 'flags loopback IPv4' {
        $a = [System.Net.IPAddress]::Parse('127.0.0.1')
        Test-IpEndpointRestricted -Addr $a | Should Be $true
    }

    It 'flags RFC1918' {
        $a = [System.Net.IPAddress]::Parse('10.0.0.1')
        Test-IpEndpointRestricted -Addr $a | Should Be $true
    }

    It 'allows public unicast' {
        $a = [System.Net.IPAddress]::Parse('8.8.8.8')
        Test-IpEndpointRestricted -Addr $a | Should Be $false
    }
}
