# Renderer-facing plan helpers. Phase 1 keeps WPS COM execution behind the
# existing legacy spec renderer, but callers now depend on RenderPlan.

function Convert-PptDslToRenderPlan {
    param([object]$Dsl)

    $legacy = Convert-PptDslToLegacySpec -Dsl $Dsl
    return [ordered]@{
        renderPlanVersion = '1.0'
        renderer          = 'wps-com'
        source            = [ordered]@{
            dslVersion = [string]$Dsl.dslVersion
            compiledAt = (Get-Date).ToUniversalTime().ToString('o')
        }
        legacySpec        = $legacy
    }
}

function Test-PptRenderPlanObject {
    param([object]$RenderPlan)

    $errors = [System.Collections.Generic.List[string]]::new()
    if ($null -eq $RenderPlan) {
        [void]$errors.Add('RenderPlan is null')
        return $errors
    }
    if ([string]$RenderPlan.renderPlanVersion -ne '1.0') {
        [void]$errors.Add('renderPlanVersion: expected 1.0')
    }
    if ([string]$RenderPlan.renderer -ne 'wps-com') {
        [void]$errors.Add('renderer: expected wps-com')
    }
    if ($null -eq $RenderPlan.legacySpec) {
        [void]$errors.Add('legacySpec: required in phase 1 render plans')
    }
    return $errors
}

function Invoke-WpsDeckFromRenderPlan {
    param(
        [object]$RenderPlan,
        [string]$ProjectRoot
    )

    $errors = Test-PptRenderPlanObject -RenderPlan $RenderPlan
    if ($errors.Count -gt 0) {
        throw "RenderPlan validation failed: $($errors -join '; ')"
    }
    return Invoke-WpsDeckFromSpec -Spec $RenderPlan.legacySpec -ProjectRoot $ProjectRoot
}
