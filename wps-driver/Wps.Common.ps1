# Shared helpers for OpenClaw WPS-PPT tools (dot-source from tools/*.ps1)

function Get-WpsProjectRoot {
    return (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
}

function Ensure-Dir {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Get-WpsSessionFile {
    $root = Get-WpsProjectRoot
    return Join-Path $root 'logs\wps-session.json'
}

function Read-WpsSession {
    $f = Get-WpsSessionFile
    if (-not (Test-Path -LiteralPath $f)) { return $null }
    try {
        Get-Content -LiteralPath $f -Raw -Encoding UTF8 | ConvertFrom-Json
    } catch { $null }
}

function Write-WpsSession {
    param([hashtable]$Data)
    $root = Get-WpsProjectRoot
    Ensure-Dir (Join-Path $root 'logs')
    $f = Get-WpsSessionFile
    $Data | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $f -Encoding UTF8
}

function Clear-WpsSession {
    $f = Get-WpsSessionFile
    if (Test-Path -LiteralPath $f) { Remove-Item -LiteralPath $f -Force }
}

function Read-OcArgs {
    param(
        [string]$ArgsJson,
        [string]$Stdin
    )
    $raw = $null
    if ($ArgsJson) { $raw = $ArgsJson }
    elseif ($env:OC_ARGS_JSON) { $raw = $env:OC_ARGS_JSON }
    elseif ($Stdin) { $raw = $Stdin }
    elseif ($input) { $raw = ($input | Out-String).Trim() }
    if (-not $raw) { return @{} }
    try {
        $obj = $raw | ConvertFrom-Json
        $ht = @{}
        $obj.PSObject.Properties | ForEach-Object { $ht[$_.Name] = $_.Value }
        return $ht
    } catch {
        throw "Invalid JSON args: $($_.Exception.Message)"
    }
}

function Get-OcErrorCode {
    param([string]$Message)
    if (-not $Message) { return 'RUNTIME_ERROR' }
    $m = $Message
    if ($m -match 'spec file not found|Invalid JSON args|Need specPath|slides\[|\.layout:|\.bullets:|\.points:|unknown ''|\.layout: unknown') {
        return 'VALIDATION_FAILED'
    }
    if ($m -match 'Could not create WPS|COM|ProgID|Invalid class string') { return 'COM_UNAVAILABLE' }
    if ($m -match 'session|wps-session|presentationPath') { return 'NO_SESSION' }
    if ($m -match 'denied|access|Unauthorized|permission') { return 'SAVE_DENIED' }
    if ($m -match 'URL blocked|SSRF policy|disallowed Content-Type|exceeds maxBytes|too many redirects') {
        return 'ASSET_FETCH_DENIED'
    }
    return 'RUNTIME_ERROR'
}

function Write-OcResult {
    param(
        [bool]$Ok,
        $Data = $null,
        [string]$Error = $null,
        [string]$Trace = $null,
        [string]$Code = $null
    )
    $out = [ordered]@{ ok = $Ok }
    if ($Ok) {
        $out.data = $Data
    } else {
        $out.error = $Error
        if ($null -ne $Data) { $out.data = $Data }
        if ($Trace) { $out.trace = $Trace }
        if ($Code) { $out.code = $Code }
    }
    if ($env:OC_PRETTY_JSON -eq '1') {
        $out | ConvertTo-Json -Depth 12
    } else {
        $out | ConvertTo-Json -Depth 12 -Compress
    }
}

function Map-LayoutNameToInt {
    param([string]$Name)
    switch ($Name.ToLowerInvariant()) {
        'title' { return 1 }              # ppLayoutTitle
        'title-content' { return 2 }     # ppLayoutText
        'content' { return 2 }           # alias: same as title-content
        'chart' { return 2 }             # advisory chart meta → bullets; layout still text
        'section' { return 3 }           # ppLayoutSectionHeader
        'two-content' { return 4 }        # ppLayoutTwoColumnText
        'blank' { return 12 }            # ppLayoutBlank
        'timeline' { return 12 }         # narrative: custom shapes on blank
        'comparison' { return 12 }
        'thesis-chain' { return 12 }
        'argument' { return 12 }
        'thesis-vertical' { return 12 }
        'swot' { return 12 }
        default { return 2 }
    }
}
