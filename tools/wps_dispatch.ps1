# Single entry for OpenClaw Gateway: -Action <name> -ArgsJson <json>
param(
    [Parameter(Mandatory = $true)]
    [string]$Action,
    [string]$ArgsJson
)
$ErrorActionPreference = 'Stop'
$map = [ordered]@{
    'new'            = 'wps_new.ps1'
    'open'           = 'wps_open.ps1'
    'add-slide'      = 'wps_add_slide.ps1'
    'set-text'       = 'wps_set_text.ps1'
    'set-notes'      = 'wps_set_notes.ps1'
    'insert-image'   = 'wps_insert_image.ps1'
    'save'           = 'wps_save.ps1'
    'close'          = 'wps_close.ps1'
    'run-spec'       = 'wps_run_spec.ps1'
    'compile-dsl'    = 'wps_compile_dsl.ps1'
    'run-dsl'        = 'wps_run_dsl.ps1'
    'edit-dsl'       = 'wps_edit_dsl.ps1'
    'render-edit'    = 'wps_render_edit.ps1'
    'validate-spec'  = 'wps_validate_spec.ps1'
    'validate-agent' = 'wps_validate_agent.ps1'
    'doctor'         = 'wps_doctor.ps1'
    'exec'           = 'wps_exec.ps1'
}
$key = $Action.Trim().ToLowerInvariant()
if (-not $map.Contains($key)) {
    . "$PSScriptRoot\_bootstrap.ps1"
    Write-OcResult $false $null "Unknown action: $Action. Valid: $($map.Keys -join ', ')" $null 'UNKNOWN_ACTION'
    exit 0
}
$child = Join-Path $PSScriptRoot $map[$key]
& $child -ArgsJson $ArgsJson
