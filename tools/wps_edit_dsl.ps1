# action: edit-dsl — apply PPT Edit DSL v1 to a PPT DSL v2 file.
# Phase 1 edits the DSL source, then optionally invokes full run-dsl rendering.
param([string]$ArgsJson)
$ErrorActionPreference = 'Stop'
try {
    $toolsDir = $PSScriptRoot
    $ProjectRoot = (Resolve-Path (Join-Path $toolsDir '..')).Path
    . (Join-Path $ProjectRoot 'wps-driver\Wps.Common.ps1')
    . (Join-Path $ProjectRoot 'wps-driver\Wps.Dsl.ps1')
    . (Join-Path $ProjectRoot 'wps-driver\Wps.Edit.ps1')

    $a = Read-OcArgs -ArgsJson $ArgsJson
    $edit = $null
    if ($a.editPath) {
        $ep = [string]$a.editPath
        if (-not [System.IO.Path]::IsPathRooted($ep)) { $ep = Join-Path $ProjectRoot $ep }
        if (-not (Test-Path -LiteralPath $ep)) { throw "edit file not found: $ep" }
        $edit = Get-Content -LiteralPath $ep -Raw -Encoding UTF8 | ConvertFrom-Json
    } elseif ($a.edit) {
        if ($a.edit -is [string]) { $edit = $a.edit | ConvertFrom-Json }
        else { $edit = $a.edit }
    } else {
        throw 'Need editPath or edit'
    }

    $dslPath = $null
    if ($a.dslPath) { $dslPath = [string]$a.dslPath }
    elseif ($edit.target -and $edit.target.dslPath) { $dslPath = [string]$edit.target.dslPath }
    else { throw 'Need dslPath or edit.target.dslPath' }
    if (-not [System.IO.Path]::IsPathRooted($dslPath)) { $dslPath = Join-Path $ProjectRoot $dslPath }
    if (-not (Test-Path -LiteralPath $dslPath)) { throw "dsl file not found: $dslPath" }

    $outPath = $dslPath
    if ($a.outPath) { $outPath = [string]$a.outPath }
    elseif ($edit.options -and $edit.options.outPath) { $outPath = [string]$edit.options.outPath }
    if (-not [System.IO.Path]::IsPathRooted($outPath)) { $outPath = Join-Path $ProjectRoot $outPath }

    $validateAfter = $true
    if ($edit.options -and $null -ne $edit.options.validateAfterEdit) {
        try { $validateAfter = [bool]$edit.options.validateAfterEdit } catch { $validateAfter = $true }
    }
    if ($null -ne $a.validateAfterEdit) {
        try { $validateAfter = [bool]$a.validateAfterEdit } catch { }
    }

    $renderAfter = $false
    if ($edit.options -and $null -ne $edit.options.renderAfterEdit) {
        try { $renderAfter = [bool]$edit.options.renderAfterEdit } catch { $renderAfter = $false }
    }
    if ($null -ne $a.renderAfterEdit) {
        try { $renderAfter = [bool]$a.renderAfterEdit } catch { }
    }

    $dsl = Get-Content -LiteralPath $dslPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $result = Invoke-PptDslEdit -Dsl $dsl -Edit $edit -ValidateAfterEdit:$validateAfter

    $outDir = Split-Path -Parent $outPath
    if ($outDir) { Ensure-Dir $outDir }
    $result.dsl | ConvertTo-Json -Depth 100 | Set-Content -LiteralPath $outPath -Encoding UTF8

    $renderResult = $null
    if ($renderAfter) {
        $args = @{ dslPath = $outPath } | ConvertTo-Json -Compress
        $renderJson = & (Join-Path $toolsDir 'wps_run_dsl.ps1') -ArgsJson $args
        try { $renderResult = $renderJson | ConvertFrom-Json } catch { $renderResult = $renderJson }
    }

    Write-OcResult $true @{
        dslPath      = (Resolve-Path -LiteralPath $outPath).Path
        applied      = $result.applied
        rendered     = [bool]$renderAfter
        renderResult = $renderResult
    }
} catch {
    $code = Get-OcErrorCode -Message $_.Exception.Message
    Write-OcResult $false $null $_.Exception.Message $_.ScriptStackTrace $code
}
