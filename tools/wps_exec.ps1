# action: exec — dot-source a .ps1 under project root only (no arbitrary Invoke-Expression)
param([string]$ArgsJson)
$ErrorActionPreference = 'Stop'
try {
    . "$PSScriptRoot\_bootstrap.ps1"
    $ProjectRoot = Get-WpsProjectRoot
    $rootFull = [System.IO.Path]::GetFullPath($ProjectRoot)
    $a = Read-OcArgs -ArgsJson $ArgsJson
    $rel = [string]$a.scriptPath
    if (-not $rel) { throw 'Missing required arg: scriptPath (relative to project root, must be .ps1)' }
    $rel = $rel.TrimStart('/', '\').Replace('/', [System.IO.Path]::DirectorySeparatorChar)
    $candidate = [System.IO.Path]::GetFullPath((Join-Path $ProjectRoot $rel))
    if (-not $candidate.ToLowerInvariant().StartsWith($rootFull.ToLowerInvariant() + '\') -and
        -not $candidate.Equals($rootFull, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "scriptPath escapes project root: $rel"
    }
    if (-not $candidate.ToLowerInvariant().EndsWith('.ps1')) {
        throw 'scriptPath must end with .ps1'
    }
    if (-not (Test-Path -LiteralPath $candidate)) {
        throw "Script not found: $candidate"
    }
    . $candidate
    Write-OcResult $true @{ executed = $candidate }
} catch {
    $code = Get-OcErrorCode -Message $_.Exception.Message
    Write-OcResult $false $null $_.Exception.Message $_.ScriptStackTrace $code
}
