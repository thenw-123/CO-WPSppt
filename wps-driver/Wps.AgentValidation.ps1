# Validation helpers for the two-stage AI Agent pipeline:
# Planner Agent -> outline, Writer Agent -> PPT DSL v2.

$script:PptPlannerSlideTypes = [string[]]@(
    'cover', 'agenda', 'section', 'content', 'argument', 'comparison', 'timeline', 'data', 'swot', 'custom'
)
$script:PptPlannerLayouts = [string[]]@(
    'auto', 'title', 'title-content', 'section', 'two-column', 'image-right',
    'comparison-matrix', 'timeline-horizontal', 'claim-with-evidence', 'swot-grid', 'data-insight'
)
$script:PptPlannerElementTypes = [string[]]@('text', 'bullets', 'image', 'chart', 'table')
$script:PptPlannerElementRoles = [string[]]@(
    'title', 'subtitle', 'body', 'claim', 'evidence', 'insight', 'caption', 'quote', 'footer', 'decorative'
)

function Invoke-PptJsonSchemaValidation {
    param(
        [object]$Value,
        [string]$SchemaPath
    )

    $errors = [System.Collections.Generic.List[string]]::new()
    if (-not (Test-Path -LiteralPath $SchemaPath)) {
        [void]$errors.Add("schema file not found: $SchemaPath")
        return $errors
    }

    $cmd = Get-Command Test-Json -ErrorAction SilentlyContinue
    if ($null -eq $cmd -or -not $cmd.Parameters.ContainsKey('Schema')) {
        [void]$errors.Add('JSON Schema validation unavailable in this PowerShell; using structural fallback')
        return $errors
    }

    try {
        $json = $Value | ConvertTo-Json -Depth 100
        $schema = Get-Content -LiteralPath $SchemaPath -Raw -Encoding UTF8
        $ok = Test-Json -Json $json -Schema $schema -ErrorAction Stop
        if (-not $ok) { [void]$errors.Add('JSON Schema validation failed') }
    } catch {
        [void]$errors.Add("JSON Schema validation failed: $($_.Exception.Message)")
    }
    return $errors
}

function Test-PptPlannerOutlineObject {
    param([object]$Outline)

    $errors = [System.Collections.Generic.List[string]]::new()
    if ($null -eq $Outline) {
        [void]$errors.Add('Planner outline is null')
        return $errors
    }
    if ([string]$Outline.outlineVersion -ne '1.0') {
        [void]$errors.Add('outlineVersion: expected 1.0')
    }
    if ($null -eq $Outline.meta) {
        [void]$errors.Add('meta: required')
    } else {
        foreach ($key in @('title', 'audience', 'purpose', 'scenario', 'language', 'targetSlideCount')) {
            if ($null -eq $Outline.meta.$key -or [string]$Outline.meta.$key -eq '') {
                [void]$errors.Add("meta.${key}: required")
            }
        }
        if ($Outline.meta.scenario -and ([string]$Outline.meta.scenario) -notin @('defense', 'report', 'project-showcase')) {
            [void]$errors.Add('meta.scenario: use defense, report, or project-showcase')
        }
        if ($null -ne $Outline.meta.targetSlideCount) {
            try {
                $n = [int]$Outline.meta.targetSlideCount
                if ($n -lt 1) { [void]$errors.Add('meta.targetSlideCount: must be >= 1') }
            } catch {
                [void]$errors.Add('meta.targetSlideCount: must be an integer')
            }
        }
    }
    if ($null -eq $Outline.themeIntent) {
        [void]$errors.Add('themeIntent: required')
    } else {
        foreach ($key in @('tone', 'visualStyle')) {
            if ($null -eq $Outline.themeIntent.$key -or [string]$Outline.themeIntent.$key -eq '') {
                [void]$errors.Add("themeIntent.${key}: required")
            }
        }
    }
    if ($null -eq $Outline.slides) {
        [void]$errors.Add('slides: required')
        return $errors
    }
    if ($Outline.slides -is [string] -or $Outline.slides -isnot [System.Collections.IEnumerable]) {
        [void]$errors.Add('slides: must be a JSON array')
        return $errors
    }

    $seen = @{}
    $idx = 0
    foreach ($slide in @($Outline.slides)) {
        $idx++
        $prefix = "slides[$idx]"
        if (-not $slide.slideId) {
            [void]$errors.Add("${prefix}.slideId: required")
        } else {
            $sid = [string]$slide.slideId
            if ($sid -notmatch '^[a-z][a-z0-9-]*$') {
                [void]$errors.Add("${prefix}.slideId: use kebab-case")
            }
            if ($seen.ContainsKey($sid)) {
                [void]$errors.Add("${prefix}.slideId: duplicate '$sid'")
            }
            $seen[$sid] = $true
        }
        $type = if ($slide.type) { ([string]$slide.type).ToLowerInvariant() } else { '' }
        if (-not $type) {
            [void]$errors.Add("${prefix}.type: required")
        } elseif ($script:PptPlannerSlideTypes -notcontains $type) {
            [void]$errors.Add("${prefix}.type: unknown '$type'")
        }
        if (-not $slide.layout -or -not $slide.layout.name) {
            [void]$errors.Add("${prefix}.layout.name: required")
        } else {
            $layout = ([string]$slide.layout.name).ToLowerInvariant()
            if ($script:PptPlannerLayouts -notcontains $layout) {
                [void]$errors.Add("${prefix}.layout.name: unknown '$layout'")
            }
        }
        if (-not $slide.message) {
            [void]$errors.Add("${prefix}.message: required")
        }
        if ($null -eq $slide.contentIntent) {
            [void]$errors.Add("${prefix}.contentIntent: required")
        } else {
            foreach ($key in @('goal', 'mustInclude', 'avoid')) {
                if ($null -eq $slide.contentIntent.$key) {
                    [void]$errors.Add("${prefix}.contentIntent.${key}: required")
                }
            }
        }
        if ($null -eq $slide.elementPlan) {
            [void]$errors.Add("${prefix}.elementPlan: required")
            continue
        }
        if ($slide.elementPlan -is [string] -or $slide.elementPlan -isnot [System.Collections.IEnumerable]) {
            [void]$errors.Add("${prefix}.elementPlan: must be a JSON array")
            continue
        }
        $ei = 0
        foreach ($element in @($slide.elementPlan)) {
            $ei++
            $etype = if ($element.type) { ([string]$element.type).ToLowerInvariant() } else { '' }
            $role = if ($element.role) { ([string]$element.role).ToLowerInvariant() } else { '' }
            if ($script:PptPlannerElementTypes -notcontains $etype) {
                [void]$errors.Add("${prefix}.elementPlan[$ei].type: unknown '$etype'")
            }
            if ($script:PptPlannerElementRoles -notcontains $role) {
                [void]$errors.Add("${prefix}.elementPlan[$ei].role: unknown '$role'")
            }
        }
    }
    return $errors
}

function Test-PptWriterDslObject {
    param(
        [object]$Dsl,
        [string]$SchemaPath
    )

    $errors = [System.Collections.Generic.List[string]]::new()
    if ($SchemaPath) {
        $schemaErrors = Invoke-PptJsonSchemaValidation -Value $Dsl -SchemaPath $SchemaPath
        foreach ($err in $schemaErrors) {
            if ([string]$err -notmatch '^JSON Schema validation unavailable') {
                [void]$errors.Add($err)
            }
        }
    }
    $structural = Test-PptDslObject -Dsl $Dsl
    foreach ($err in $structural) { [void]$errors.Add([string]$err) }
    if ($null -eq $Dsl.meta) { [void]$errors.Add('meta: required for Writer PPT DSL v2') }
    $idx = 0
    foreach ($slide in @($Dsl.slides)) {
        $idx++
        if ($null -eq $slide.elements) {
            [void]$errors.Add("slides[$idx].elements: required for Writer PPT DSL v2")
        }
    }
    return $errors
}

function Test-PptAgentStageOutput {
    param(
        [ValidateSet('planner', 'writer')]
        [string]$Stage,
        [object]$Value,
        [string]$ProjectRoot
    )

    if ($Stage -eq 'planner') {
        $schemaPath = Join-Path $ProjectRoot 'src\intent\planner.schema.json'
        $schemaErrors = Invoke-PptJsonSchemaValidation -Value $Value -SchemaPath $schemaPath
        $errors = Test-PptPlannerOutlineObject -Outline $Value
        foreach ($err in $schemaErrors) {
            if ([string]$err -notmatch '^JSON Schema validation unavailable') {
                [void]$errors.Add([string]$err)
            }
        }
        return $errors
    }

    $dslSchemaPath = Join-Path $ProjectRoot 'specs\dsl\ppt-dsl.schema.json'
    return Test-PptWriterDslObject -Dsl $Value -SchemaPath $dslSchemaPath
}
