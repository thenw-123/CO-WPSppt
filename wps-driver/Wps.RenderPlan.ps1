# Renderer-facing plan helpers. RenderPlan is the execution source. The normal
# path is DSL -> Layout Engine -> atomic operations; legacySpec remains only as
# an explicit compatibility format.

function Convert-PptDslToRenderPlan {
    param([object]$Dsl)

    return Convert-PptDslToAtomicRenderPlan -Dsl $Dsl
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
    if ([string]$RenderPlan.renderer -notin @('wps-com-atomic', 'wps-com')) {
        [void]$errors.Add('renderer: expected wps-com-atomic')
    }
    if ([string]$RenderPlan.renderer -eq 'wps-com-atomic') {
        if ($null -eq $RenderPlan.slides) {
            [void]$errors.Add('slides: required for atomic render plans')
        }
    } elseif ($null -eq $RenderPlan.legacySpec) {
        [void]$errors.Add('legacySpec: required for legacy render plans')
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
    if ([string]$RenderPlan.renderer -eq 'wps-com-atomic') {
        return Invoke-WpsDeckFromAtomicRenderPlan -RenderPlan $RenderPlan -ProjectRoot $ProjectRoot
    }
    return Invoke-WpsDeckFromSpec -Spec $RenderPlan.legacySpec -ProjectRoot $ProjectRoot
}
