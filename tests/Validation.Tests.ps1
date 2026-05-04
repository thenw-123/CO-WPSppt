# Pester 3/5 compatible (no top-level BeforeAll). Run: Invoke-Pester .\tests\Validation.Tests.ps1
# Pester 5+ optional: Install-Module Pester -MinimumVersion 5.0.0 -Force

$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
. (Join-Path $RepoRoot 'wps-driver\Wps.Common.ps1')
. (Join-Path $RepoRoot 'wps-driver\Wps.ValidateSpec.ps1')
. (Join-Path $RepoRoot 'wps-driver\Wps.Dsl.ps1')
. (Join-Path $RepoRoot 'wps-driver\Wps.RenderPlan.ps1')

Describe 'Map-LayoutNameToInt' {
    It 'maps known layouts' {
        Map-LayoutNameToInt -Name 'title' | Should Be 1
        Map-LayoutNameToInt -Name 'title-content' | Should Be 2
        Map-LayoutNameToInt -Name 'content' | Should Be 2
        Map-LayoutNameToInt -Name 'chart' | Should Be 2
        Map-LayoutNameToInt -Name 'section' | Should Be 3
        Map-LayoutNameToInt -Name 'blank' | Should Be 12
        Map-LayoutNameToInt -Name 'timeline' | Should Be 12
        Map-LayoutNameToInt -Name 'argument' | Should Be 12
        Map-LayoutNameToInt -Name 'thesis-vertical' | Should Be 12
        Map-LayoutNameToInt -Name 'swot' | Should Be 12
    }
}

Describe 'Test-PptDeckSpecObject' {
    It 'accepts minimal valid spec' {
        $s = '{"slides":[{"layout":"title","title":"x"}]}' | ConvertFrom-Json
        $e = Test-PptDeckSpecObject -Spec $s
        $e.Count | Should Be 0
    }

    It 'rejects unknown layout' {
        $s = '{"slides":[{"layout":"nope","title":"x"}]}' | ConvertFrom-Json
        $e = Test-PptDeckSpecObject -Spec $s
        $e.Count | Should BeGreaterThan 0
    }

    It 'validates chart_data series length' {
        $s = @'
{"slides":[{"layout":"chart","title":"t","chart_type":"bar","chart_data":{"labels":["a","b"],"series":[{"values":[1]}]}}]}
'@ | ConvertFrom-Json
        $e = Test-PptDeckSpecObject -Spec $s
        $e.Count | Should BeGreaterThan 0
    }

    It 'accepts consistent chart_data' {
        $s = @'
{"slides":[{"layout":"chart","title":"t","chart_type":"bar","chart_data":{"labels":["a","b"],"series":[{"values":[1,2]}]}}]}
'@ | ConvertFrom-Json
        $e = Test-PptDeckSpecObject -Spec $s
        $e.Count | Should Be 0
    }
}

Describe 'Get-OcErrorCode' {
    It 'classifies COM message' {
        Get-OcErrorCode -Message 'Could not create WPS' | Should Be 'COM_UNAVAILABLE'
    }

    It 'classifies asset fetch policy' {
        Get-OcErrorCode -Message 'URL blocked (SSRF policy): test' | Should Be 'ASSET_FETCH_DENIED'
    }
}

Describe 'Test-PptDeckSpecObject extended' {
    It 'rejects image and imageUrl together' {
        $s = @'
{"slides":[{"layout":"title-content","title":"x","image":"a.png","imageUrl":"https://a.com/b.png"}]}
'@ | ConvertFrom-Json
        $e = Test-PptDeckSpecObject -Spec $s
        $e.Count | Should BeGreaterThan 0
    }

    It 'rejects chart_engine com without chart_data' {
        $s = @'
{"slides":[{"layout":"chart","title":"t","chart_engine":"com","chart_type":"bar"}]}
'@ | ConvertFrom-Json
        $e = Test-PptDeckSpecObject -Spec $s
        $e.Count | Should BeGreaterThan 0
    }

    It 'accepts narrative timeline spec' {
        $s = @'
{"slides":[{"layout":"timeline","title":"演进","timeline":[{"mark":"A","text":"起点"},{"mark":"B","text":"转折"}]}]}
'@ | ConvertFrom-Json
        $e = Test-PptDeckSpecObject -Spec $s
        $e.Count | Should Be 0
    }

    It 'rejects narrative layout with chart_data' {
        $s = @'
{"slides":[{"layout":"argument","title":"t","argument":{"thesis":"x","therefore":"y"},"chart_data":{"labels":["a"],"series":[{"values":[1]}]}}]}
'@ | ConvertFrom-Json
        $e = Test-PptDeckSpecObject -Spec $s
        $e.Count | Should BeGreaterThan 0
    }

    It 'accepts thesis-vertical with three steps' {
        $s = @'
{"slides":[{"layout":"thesis-vertical","title":"T","thesisVertical":{"claim":"C","steps":[{"label":"1","text":"a"},{"label":"2","text":"b"},{"label":"3","text":"c"}]}}]}
'@ | ConvertFrom-Json
        $e = Test-PptDeckSpecObject -Spec $s
        $e.Count | Should Be 0
    }

    It 'rejects thesis-vertical when steps count is not 3' {
        $s = @'
{"slides":[{"layout":"thesis-vertical","thesisVertical":{"claim":"C","steps":[{"text":"a"},{"text":"b"}]}}]}
'@ | ConvertFrom-Json
        $e = Test-PptDeckSpecObject -Spec $s
        $e.Count | Should BeGreaterThan 0
    }

    It 'accepts swot with four quadrants' {
        $s = @'
{"slides":[{"layout":"swot","title":"S","swot":{"strengths":"a","weaknesses":"b","opportunities":"c","threats":"d"}}]}
'@ | ConvertFrom-Json
        $e = Test-PptDeckSpecObject -Spec $s
        $e.Count | Should Be 0
    }
}

Describe 'PPT DSL compiler' {
    It 'accepts minimal DSL' {
        $dsl = @'
{"dslVersion":"2.0","slides":[{"kind":"cover","title":"T"}]}
'@ | ConvertFrom-Json
        $e = Test-PptDslObject -Dsl $dsl
        $e.Count | Should Be 0
    }

    It 'rejects unknown DSL slide kind' {
        $dsl = @'
{"dslVersion":"2.0","slides":[{"kind":"unknown","title":"T"}]}
'@ | ConvertFrom-Json
        $e = Test-PptDslObject -Dsl $dsl
        $e.Count | Should BeGreaterThan 0
    }

    It 'compiles argument DSL to legacy narrative spec' {
        $dsl = @'
{"dslVersion":"2.0","slides":[{"kind":"argument","message":"M","content":{"claim":"C","points":["a","b"],"conclusion":"Z"}}]}
'@ | ConvertFrom-Json
        $spec = Convert-PptDslToLegacySpec -Dsl $dsl
        $spec.slides[0].layout | Should Be 'argument'
        $spec.slides[0].argument.thesis | Should Be 'C'
        $spec.slides[0].argument.because.Count | Should Be 2
        $legacyErrors = Test-PptDeckSpecObject -Spec $spec
        $legacyErrors.Count | Should Be 0
    }

    It 'wraps compiled DSL in a render plan' {
        $dsl = @'
{"dslVersion":"2.0","slides":[{"kind":"content","title":"T","content":{"points":["a"]}}]}
'@ | ConvertFrom-Json
        $plan = Convert-PptDslToRenderPlan -Dsl $dsl
        $plan.renderPlanVersion | Should Be '1.0'
        $plan.renderer | Should Be 'wps-com'
        $plan.legacySpec.slides[0].layout | Should Be 'title-content'
    }

    It 'compiles simple PPT DSL v2 example' {
        $dsl = Get-Content -LiteralPath (Join-Path $RepoRoot 'specs\dsl\example-simple.json') -Raw -Encoding UTF8 | ConvertFrom-Json
        $plan = Convert-PptDslToRenderPlan -Dsl $dsl
        $plan.legacySpec.title | Should Be '季度项目进展汇报'
        $plan.legacySpec.slides.Count | Should Be 2
        $plan.legacySpec.slides[1].bullets.Count | Should Be 3
    }

    It 'compiles complex PPT DSL v2 chart and table elements' {
        $dsl = Get-Content -LiteralPath (Join-Path $RepoRoot 'specs\dsl\example-complex.json') -Raw -Encoding UTF8 | ConvertFrom-Json
        $plan = Convert-PptDslToRenderPlan -Dsl $dsl
        $chartSlide = @($plan.legacySpec.slides | Where-Object { $_.layout -eq 'chart' } | Select-Object -First 1)[0]
        $chartSlide.chart_type | Should Be 'bar'
        $chartSlide.chart_data.labels.Count | Should Be 3
    }
}
