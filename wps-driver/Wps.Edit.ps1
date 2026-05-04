# Edit helpers for PPT DSL v2. Phase 1 edits the DSL source file and can
# optionally trigger the existing full run-dsl renderer.

function Test-PptEditRequestObject {
    param([object]$Edit)

    $errors = [System.Collections.Generic.List[string]]::new()
    if ($null -eq $Edit) {
        [void]$errors.Add('Edit request is null')
        return $errors
    }
    if ([string]$Edit.editVersion -ne '1.0') {
        [void]$errors.Add('editVersion: expected 1.0')
    }
    if ($null -eq $Edit.target) {
        [void]$errors.Add('target: required')
    } elseif (-not $Edit.target.dslPath) {
        [void]$errors.Add('target.dslPath: required')
    }
    if ($null -eq $Edit.operations) {
        [void]$errors.Add('operations: required')
        return $errors
    }
    if ($Edit.operations -is [string] -or $Edit.operations -isnot [System.Collections.IEnumerable]) {
        [void]$errors.Add('operations: must be a JSON array')
        return $errors
    }

    $idx = 0
    foreach ($op in @($Edit.operations)) {
        $idx++
        $prefix = "operations[$idx]"
        $opName = if ($op.op) { ([string]$op.op).ToLowerInvariant() } else { '' }
        if ($opName -notin @('replace', 'update')) {
            [void]$errors.Add("${prefix}.op: use replace or update")
        }
        if ($null -eq $op.target -or -not $op.target.slideId) {
            [void]$errors.Add("${prefix}.target.slideId: required")
        }
        if ($opName -eq 'replace' -and $null -eq $op.value) {
            [void]$errors.Add("${prefix}.value: required for replace")
        }
        if ($opName -eq 'update' -and $null -eq $op.patch) {
            [void]$errors.Add("${prefix}.patch: required for update")
        }
        if ($opName -eq 'update' -and ($null -eq $op.target -or -not $op.target.elementId)) {
            [void]$errors.Add("${prefix}.target.elementId: required for update")
        }
    }
    return $errors
}

function Get-PptDslSlideIndexById {
    param([object]$Dsl, [string]$SlideId)
    $slides = @($Dsl.slides)
    for ($i = 0; $i -lt $slides.Count; $i++) {
        if ([string]$slides[$i].id -eq $SlideId) { return $i }
    }
    return -1
}

function Get-PptDslElementIndexById {
    param([object]$Slide, [string]$ElementId)
    $elements = @($Slide.elements)
    for ($i = 0; $i -lt $elements.Count; $i++) {
        if ([string]$elements[$i].id -eq $ElementId) { return $i }
    }
    return -1
}

function Set-PptObjectProperty {
    param(
        [object]$Object,
        [string]$Name,
        $Value
    )
    $prop = $Object.PSObject.Properties[$Name]
    if ($null -eq $prop) {
        $Object | Add-Member -NotePropertyName $Name -NotePropertyValue $Value
    } else {
        $Object.$Name = $Value
    }
}

function Update-PptObjectShallow {
    param(
        [object]$Object,
        [object]$Patch
    )
    foreach ($p in $Patch.PSObject.Properties) {
        Set-PptObjectProperty -Object $Object -Name $p.Name -Value $p.Value
    }
}

function Invoke-PptEditOperation {
    param(
        [object]$Dsl,
        [object]$Operation
    )

    $opName = ([string]$Operation.op).ToLowerInvariant()
    $slideId = [string]$Operation.target.slideId
    $slideIndex = Get-PptDslSlideIndexById -Dsl $Dsl -SlideId $slideId
    if ($slideIndex -lt 0) { throw "Slide not found: $slideId" }

    $slides = @($Dsl.slides)
    $slide = $slides[$slideIndex]

    if ($opName -eq 'replace' -and -not $Operation.target.elementId) {
        $newSlide = $Operation.value
        if (-not $newSlide.id) {
            Set-PptObjectProperty -Object $newSlide -Name 'id' -Value $slideId
        } elseif ([string]$newSlide.id -ne $slideId) {
            throw "Replacement slide id must match target slideId: $slideId"
        }
        $slides[$slideIndex] = $newSlide
        $Dsl.slides = @($slides)
        return @{
            op      = 'replace'
            slideId = $slideId
            scope   = 'slide'
        }
    }

    $elementId = [string]$Operation.target.elementId
    if (-not $elementId) { throw "target.elementId is required for $opName element operation" }
    $elementIndex = Get-PptDslElementIndexById -Slide $slide -ElementId $elementId
    if ($elementIndex -lt 0) { throw "Element not found: $slideId/$elementId" }

    $elements = @($slide.elements)
    if ($opName -eq 'replace') {
        $newElement = $Operation.value
        if (-not $newElement.id) {
            Set-PptObjectProperty -Object $newElement -Name 'id' -Value $elementId
        } elseif ([string]$newElement.id -ne $elementId) {
            throw "Replacement element id must match target elementId: $elementId"
        }
        $elements[$elementIndex] = $newElement
        $slide.elements = @($elements)
        return @{
            op        = 'replace'
            slideId   = $slideId
            elementId = $elementId
            scope     = 'element'
        }
    }

    Update-PptObjectShallow -Object $elements[$elementIndex] -Patch $Operation.patch
    $slide.elements = @($elements)
    return @{
        op        = 'update'
        slideId   = $slideId
        elementId = $elementId
        scope     = 'element'
    }
}

function Invoke-PptDslEdit {
    param(
        [object]$Dsl,
        [object]$Edit,
        [switch]$ValidateAfterEdit
    )

    $editErrors = Test-PptEditRequestObject -Edit $Edit
    if ($editErrors.Count -gt 0) {
        throw "Edit validation failed: $($editErrors -join '; ')"
    }

    $applied = @()
    foreach ($op in @($Edit.operations)) {
        $applied += Invoke-PptEditOperation -Dsl $Dsl -Operation $op
    }

    if ($ValidateAfterEdit) {
        $dslErrors = Test-PptDslObject -Dsl $Dsl
        if ($dslErrors.Count -gt 0) {
            throw "Edited DSL validation failed: $($dslErrors -join '; ')"
        }
    }

    return @{
        dsl     = $Dsl
        applied = @($applied)
    }
}
