# WPS Presentation COM connection (Kwpp.Application and fallbacks)

$script:WpsProgIds = @(
    'Kwpp.Application',
    'kwpp.Application',
    'KWPP.Application',
    'wpp.Application',
    'KWPS.Application'
)

function Get-WpsApplication {
    $lastErr = $null
    foreach ($progId in $script:WpsProgIds) {
        try {
            $app = $null
            try {
                $app = [Runtime.InteropServices.Marshal]::GetActiveObject($progId)
            } catch {
                $app = New-Object -ComObject $progId
            }
            if ($null -ne $app) {
                $app.Visible = $true
                return $app
            }
        } catch {
            $lastErr = $_.Exception.Message
            continue
        }
    }
    throw "Could not create WPS Presentation COM. Tried: $($script:WpsProgIds -join ', '). Last error: $lastErr"
}

function Release-ComObjectSafe {
    param($Obj)
    if ($null -eq $Obj) { return }
    try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Obj) } catch { }
}
