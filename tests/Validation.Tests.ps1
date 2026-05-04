# Pester 3/5 compatible (no top-level BeforeAll). Run: Invoke-Pester .\tests\Validation.Tests.ps1
# Pester 5+ optional: Install-Module Pester -MinimumVersion 5.0.0 -Force

$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
. (Join-Path $RepoRoot 'wps-driver\Wps.Common.ps1')
. (Join-Path $RepoRoot 'wps-driver\Wps.ValidateSpec.ps1')
. (Join-Path $RepoRoot 'wps-driver\Wps.Dsl.ps1')
. (Join-Path $RepoRoot 'wps-driver\Wps.LayoutEngine.ps1')
. (Join-Path $RepoRoot 'wps-driver\Wps.AtomicRenderer.ps1')
. (Join-Path $RepoRoot 'wps-driver\Wps.RenderPlan.ps1')
. (Join-Path $RepoRoot 'wps-driver\Wps.AgentValidation.ps1')
. (Join-Path $RepoRoot 'wps-driver\Wps.Edit.ps1')
. (Join-Path $RepoRoot 'wps-driver\Wps.SlideMap.ps1')
. (Join-Path $RepoRoot 'wps-driver\Wps.RenderEdit.ps1')

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
{"dslVersion":"2.0","meta":{"title":"T"},"slides":[{"id":"s1","type":"content","elements":[{"id":"title","type":"text","role":"title","text":"T"}]}]}
'@ | ConvertFrom-Json
        $plan = Convert-PptDslToRenderPlan -Dsl $dsl
        $plan.renderPlanVersion | Should Be '1.0'
        $plan.renderer | Should Be 'wps-com-atomic'
        $plan.slides[0].operations[0].op | Should Be 'addText'
    }

    It 'compiles simple PPT DSL v2 example' {
        $dsl = Get-Content -LiteralPath (Join-Path $RepoRoot 'specs\dsl\example-simple.json') -Raw -Encoding UTF8 | ConvertFrom-Json
        $plan = Convert-PptDslToRenderPlan -Dsl $dsl
        $plan.meta.title | Should Be '季度项目进展汇报'
        $plan.slides.Count | Should Be 2
        $plan.slides[1].operations[1].items.Count | Should Be 3
    }

    It 'compiles complex PPT DSL v2 chart and table elements' {
        $dsl = Get-Content -LiteralPath (Join-Path $RepoRoot 'specs\dsl\example-complex.json') -Raw -Encoding UTF8 | ConvertFrom-Json
        $legacy = Convert-PptDslToLegacySpec -Dsl $dsl
        $chartSlide = @($legacy.slides | Where-Object { $_.layout -eq 'chart' } | Select-Object -First 1)[0]
        $chartSlide.chart_type | Should Be 'bar'
        $chartSlide.chart_data.labels.Count | Should Be 3
    }

    It 'lays out title-content with concrete coordinates' {
        $dsl = Get-Content -LiteralPath (Join-Path $RepoRoot 'specs\dsl\example-simple.json') -Raw -Encoding UTF8 | ConvertFrom-Json
        $plan = Convert-PptDslToRenderPlan -Dsl $dsl
        $titleOp = $plan.slides[1].operations[0]
        $bodyOp = $plan.slides[1].operations[1]
        $titleOp.box.x | Should Be 40
        $titleOp.box.y | Should Be 40
        $titleOp.box.w | Should Be 880
        $titleOp.box.h | Should Be 80
        $bodyOp.box.x | Should Be 40
        $bodyOp.box.y | Should Be 140
        $bodyOp.box.h | Should Be 360
        $bodyOp.op | Should Be 'addBullets'
    }

    It 'lays out image-right as left text and right image' {
        $dsl = @'
{"dslVersion":"2.0","meta":{"title":"T"},"slides":[{"id":"s1","type":"content","layout":{"name":"image-right"},"elements":[{"id":"title","type":"text","role":"title","text":"T"},{"id":"points","type":"bullets","role":"body","items":["a"]},{"id":"img","type":"image","role":"decorative","source":{"kind":"local","path":"assets/a.png"}}]}]}
'@ | ConvertFrom-Json
        $plan = Convert-PptDslToRenderPlan -Dsl $dsl
        $textOp = @($plan.slides[0].operations | Where-Object { $_.op -eq 'addBullets' })[0]
        $imageOp = @($plan.slides[0].operations | Where-Object { $_.op -eq 'addImage' })[0]
        $textOp.box.x | Should Be 40
        $textOp.box.w | Should Be 430
        $imageOp.box.x | Should Be 490
        $imageOp.box.w | Should Be 430
    }

    It 'compiles content slide with slide.title and bullets content.items to positioned atomic ops' {
        $dsl = @'
{"dslVersion":"2.0","meta":{"title":"Deck"},"slides":[{"id":"s1","type":"content","title":"测试","elements":[{"id":"b1","type":"bullets","content":{"items":["A","B"]}}]}]}
'@ | ConvertFrom-Json
        $e = Test-PptDslObject -Dsl $dsl
        $e.Count | Should Be 0
        $plan = Convert-PptDslToRenderPlan -Dsl $dsl
        $ops = @($plan.slides[0].operations)
        $ops.Count | Should Be 2
        $ops[0].op | Should Be 'addText'
        $ops[0].text | Should Be '测试'
        $ops[0].box.x | Should Be 40
        $ops[0].box.y | Should Be 40
        $ops[1].op | Should Be 'addBullets'
        $ops[1].items.Count | Should Be 2
        $ops[1].box.x | Should Be 40
        $ops[1].box.y | Should Be 140
    }

    It 'lays out three-column content into three equal columns' {
        $dsl = @'
{"dslVersion":"2.0","meta":{"title":"Deck"},"slides":[{"id":"s1","type":"content","title":"三列","layout":{"name":"three-column"},"elements":[{"id":"c1","type":"text","text":"A"},{"id":"c2","type":"text","text":"B"},{"id":"c3","type":"text","text":"C"}]}]}
'@ | ConvertFrom-Json
        $plan = Convert-PptDslToRenderPlan -Dsl $dsl
        $ops = @($plan.slides[0].operations | Where-Object { $_.role -ne 'title' })
        $ops.Count | Should Be 3
        $ops[0].box.x | Should Be 40
        $ops[0].box.w | Should Be 280
        $ops[1].box.x | Should Be 340
        $ops[2].box.x | Should Be 640
    }

    It 'lays out grid cards as 2 columns and multiple rows' {
        $dsl = @'
{"dslVersion":"2.0","meta":{"title":"Deck"},"slides":[{"id":"s1","type":"content","title":"卡片","layout":{"name":"grid"},"elements":[{"id":"g1","type":"text","text":"A"},{"id":"g2","type":"text","text":"B"},{"id":"g3","type":"text","text":"C"},{"id":"g4","type":"text","text":"D"}]}]}
'@ | ConvertFrom-Json
        $plan = Convert-PptDslToRenderPlan -Dsl $dsl
        $ops = @($plan.slides[0].operations | Where-Object { $_.role -ne 'title' })
        $ops.Count | Should Be 4
        $ops[0].box.x | Should Be 40
        $ops[1].box.x | Should Be 490
        $ops[2].box.y | Should Be 330
        $ops[0].box.w | Should Be 430
    }

    It 'lays out grid with layout.columns and layout.gap' {
        $dsl = @'
{"dslVersion":"2.0","meta":{"title":"Deck"},"slides":[{"id":"s1","type":"content","title":"网格","layout":{"name":"grid","columns":4,"gap":24},"elements":[{"id":"a","type":"text","text":"1"},{"id":"b","type":"text","text":"2"},{"id":"c","type":"text","text":"3"},{"id":"d","type":"text","text":"4"}]}]}
'@ | ConvertFrom-Json
        $plan = Convert-PptDslToRenderPlan -Dsl $dsl
        $bodyOps = @($plan.slides[0].operations | Where-Object { $_.role -ne 'title' })
        $bodyOps.Count | Should Be 4
        $expectedW = ((880 - 3 * 24) / 4)
        $bodyOps[0].box.x | Should Be 40
        [math]::Round([double]$bodyOps[0].box.w, 4) | Should Be ([math]::Round($expectedW, 4))
    }

    It 'emits addRoundedRect and inset text when layout.card uses surface tokens' {
        $dsl = @'
{"dslVersion":"2.0","meta":{"title":"Deck"},"slides":[{"id":"s1","type":"content","title":"卡片主题","layout":{"name":"grid","columns":2,"gap":16,"card":{"surface":"muted","padding":12}},"elements":[{"id":"g1","type":"text","text":"Alpha"},{"id":"g2","type":"text","text":"Beta"}]}]}
'@ | ConvertFrom-Json
        $plan = Convert-PptDslToRenderPlan -Dsl $dsl
        $all = @($plan.slides[0].operations)
        $chrome = @($all | Where-Object { $_.op -eq 'addRoundedRect' })
        $chrome.Count | Should Be 2
        $chrome[0].fill | Should Be '#F1F5F9'
        $t1 = @($all | Where-Object { $_.elementId -eq 'g1' -and $_.op -eq 'addText' })[0]
        $t1.box.x | Should Be 52
        $t1.box.y | Should Be 152
    }

    It 'auto-layout selects title-content for short content slide without layout' {
        $dsl = @'
{"dslVersion":"2.0","meta":{"title":"D"},"slides":[{"id":"s1","type":"content","title":"智能布局选择","elements":[{"id":"b","type":"bullets","content":{"items":["自动选择布局","支持图片","灵活扩展"]}}]}]}
'@ | ConvertFrom-Json
        $plan = Convert-PptDslToRenderPlan -Dsl $dsl
        $plan.slides[0].layout.key | Should Be 'title-content'
        $plan.slides[0].layout.autoLayout | Should Be $true
        $ops = @($plan.slides[0].operations)
        $ops[0].op | Should Be 'addText'
        $ops[1].op | Should Be 'addBullets'
    }

    It 'auto-layout selects image-right when slide has an image and text is not too long' {
        $dsl = @'
{"dslVersion":"2.0","meta":{"title":"D"},"slides":[{"id":"s1","type":"content","title":"短","elements":[{"id":"t","type":"text","text":"左列说明"},{"id":"i","type":"image","source":{"kind":"local","path":"assets/x.png"}}]}]}
'@ | ConvertFrom-Json
        $plan = Convert-PptDslToRenderPlan -Dsl $dsl
        $plan.slides[0].layout.key | Should Be 'left-right'
        $plan.slides[0].layout.autoLayout | Should Be $true
    }

    It 'auto-layout selects comparison for exactly two body elements' {
        $dsl = @'
{"dslVersion":"2.0","meta":{"title":"D"},"slides":[{"id":"s1","type":"content","title":"对比","elements":[{"id":"a","type":"text","text":"左"},{"id":"b","type":"text","text":"右"}]}]}
'@ | ConvertFrom-Json
        $plan = Convert-PptDslToRenderPlan -Dsl $dsl
        $plan.slides[0].layout.key | Should Be 'comparison'
        $plan.slides[0].layout.autoLayout | Should Be $true
        $bodyOps = @($plan.slides[0].operations | Where-Object { $_.role -ne 'title' })
        $bodyOps[0].box.x | Should Be 40
        $bodyOps[1].box.x | Should Be 490
    }

    It 'auto-layout selects grid when many elements and text length over 100' {
        $chunk = '1234567890123456789012345'
        $dsl = @'
{"dslVersion":"2.0","meta":{"title":"D"},"slides":[{"id":"s1","type":"content","title":"T","elements":[{"id":"e1","type":"text","text":"x"},{"id":"e2","type":"text","text":"x"},{"id":"e3","type":"text","text":"x"},{"id":"e4","type":"text","text":"x"}]}]}
'@ | ConvertFrom-Json
        foreach ($j in 0..3) { $dsl.slides[0].elements[$j].text = $chunk }
        $plan = Convert-PptDslToRenderPlan -Dsl $dsl
        $plan.slides[0].layout.key | Should Be 'grid'
        $plan.slides[0].layout.autoLayout | Should Be $true
        $plan.slides[0].layout.name | Should Be 'grid'
    }

    It 'explicit layout.name overrides auto selection' {
        $dsl = @'
{"dslVersion":"2.0","meta":{"title":"D"},"slides":[{"id":"s1","type":"content","title":"T","layout":{"name":"three-column"},"elements":[{"id":"a","type":"text","text":"A"},{"id":"b","type":"text","text":"B"},{"id":"c","type":"text","text":"C"}]}]}
'@ | ConvertFrom-Json
        $plan = Convert-PptDslToRenderPlan -Dsl $dsl
        $plan.slides[0].layout.key | Should Be 'three-column'
        $plan.slides[0].layout.autoLayout | Should Be $null
    }
}

Describe 'Atomic renderer helpers' {
    It 'shrinks font size when text is too long for box' {
        $box = @{ x = 40; y = 40; w = 200; h = 60 }
        $size = Get-WpsAtomicFittedFontSize -Text ('这是一段很长很长的文本。' * 8) -Box $box -Role 'body' -Style $null
        $size | Should BeLessThan 20
        $size | Should BeGreaterThan 11
    }

    It 'computes contain image rect while keeping ratio' {
        $box = @{ x = 40; y = 140; w = 430; h = 360 }
        $fit = Get-WpsAtomicContainImageBox -Box $box -ImageWidth 1920 -ImageHeight 1080
        $fit.w | Should Be 430
        $fit.h | Should Be 241.875
        $fit.y | Should Be 199.0625
    }
}

Describe 'Two-stage Agent validation' {
    It 'accepts Planner outline example' {
        $outline = Get-Content -LiteralPath (Join-Path $RepoRoot 'src\intent\example-outline.json') -Raw -Encoding UTF8 | ConvertFrom-Json
        $errors = Test-PptAgentStageOutput -Stage planner -Value $outline -ProjectRoot $RepoRoot
        $errors.Count | Should Be 0
    }

    It 'rejects Planner outline without stable slide id' {
        $outline = @'
{"outlineVersion":"1.0","meta":{"title":"T","audience":"A","purpose":"P","scenario":"report","language":"zh-CN","targetSlideCount":1},"themeIntent":{"tone":"business","visualStyle":"clean"},"slides":[{"slideId":"Bad Id","type":"cover","layout":{"name":"title"},"message":"M","contentIntent":{"goal":"G","mustInclude":[],"avoid":[]},"elementPlan":[{"type":"text","role":"title"}]}]}
'@ | ConvertFrom-Json
        $errors = Test-PptAgentStageOutput -Stage planner -Value $outline -ProjectRoot $RepoRoot
        $errors.Count | Should BeGreaterThan 0
    }

    It 'accepts Writer PPT DSL v2 example' {
        $dsl = Get-Content -LiteralPath (Join-Path $RepoRoot 'specs\dsl\example-simple.json') -Raw -Encoding UTF8 | ConvertFrom-Json
        $errors = Test-PptAgentStageOutput -Stage writer -Value $dsl -ProjectRoot $RepoRoot
        $errors.Count | Should Be 0
    }

    It 'rejects Writer output missing elements' {
        $dsl = @'
{"dslVersion":"2.0","meta":{"title":"T"},"slides":[{"type":"cover"}]}
'@ | ConvertFrom-Json
        $errors = Test-PptAgentStageOutput -Stage writer -Value $dsl -ProjectRoot $RepoRoot
        $errors.Count | Should BeGreaterThan 0
    }
}

Describe 'PPT DSL edit operations' {
    It 'updates one element without changing other slides' {
        $dsl = @'
{"dslVersion":"2.0","meta":{"title":"T"},"slides":[{"id":"s1","type":"content","elements":[{"id":"title","type":"text","role":"title","text":"Old"},{"id":"body","type":"bullets","role":"body","items":["a"]}]},{"id":"s2","type":"content","elements":[{"id":"title","type":"text","role":"title","text":"Keep"}]}]}
'@ | ConvertFrom-Json
        $edit = @'
{"editVersion":"1.0","target":{"dslPath":"unused.json"},"operations":[{"op":"update","target":{"slideId":"s1","elementId":"title"},"patch":{"text":"New"}}]}
'@ | ConvertFrom-Json
        $result = Invoke-PptDslEdit -Dsl $dsl -Edit $edit -ValidateAfterEdit
        $result.dsl.slides[0].elements[0].text | Should Be 'New'
        $result.dsl.slides[1].elements[0].text | Should Be 'Keep'
    }

    It 'replaces an element by stable id' {
        $dsl = @'
{"dslVersion":"2.0","meta":{"title":"T"},"slides":[{"id":"s1","type":"content","elements":[{"id":"body","type":"bullets","role":"body","items":["a"]}]}]}
'@ | ConvertFrom-Json
        $edit = @'
{"editVersion":"1.0","target":{"dslPath":"unused.json"},"operations":[{"op":"replace","target":{"slideId":"s1","elementId":"body"},"value":{"id":"body","type":"bullets","role":"body","items":["b","c"]}}]}
'@ | ConvertFrom-Json
        $result = Invoke-PptDslEdit -Dsl $dsl -Edit $edit -ValidateAfterEdit
        $result.dsl.slides[0].elements[0].items.Count | Should Be 2
        $result.dsl.slides[0].elements[0].items[0] | Should Be 'b'
    }

    It 'replaces a slide by stable id' {
        $dsl = @'
{"dslVersion":"2.0","meta":{"title":"T"},"slides":[{"id":"s1","type":"content","elements":[{"id":"title","type":"text","role":"title","text":"Old"}]},{"id":"s2","type":"content","elements":[{"id":"title","type":"text","role":"title","text":"Keep"}]}]}
'@ | ConvertFrom-Json
        $edit = @'
{"editVersion":"1.0","target":{"dslPath":"unused.json"},"operations":[{"op":"replace","target":{"slideId":"s1"},"value":{"id":"s1","type":"argument","elements":[{"id":"claim","type":"text","role":"claim","text":"New claim"}]}}]}
'@ | ConvertFrom-Json
        $result = Invoke-PptDslEdit -Dsl $dsl -Edit $edit -ValidateAfterEdit
        $result.dsl.slides[0].type | Should Be 'argument'
        $result.dsl.slides[0].elements[0].id | Should Be 'claim'
        $result.dsl.slides[1].id | Should Be 's2'
    }
}

Describe 'PPT slide sidecar map' {
    It 'creates stable slide id to index map' {
        $dsl = @'
{"dslVersion":"2.0","meta":{"title":"T"},"slides":[{"id":"cover","type":"cover","elements":[{"id":"title","type":"text","role":"title","text":"T"}]},{"id":"summary","type":"content","elements":[{"id":"body","type":"bullets","role":"body","items":["a"]}]}]}
'@ | ConvertFrom-Json
        $map = New-PptSlideMap -Dsl $dsl -PresentationPath 'output/test.pptx' -DslPath 'specs/test.json'
        $map.slides.Count | Should Be 2
        $map.slides[0].slideId | Should Be 'cover'
        $map.slides[1].index | Should Be 2
    }

    It 'extracts unique affected slide ids from edit request' {
        $edit = @'
{"editVersion":"1.0","target":{"dslPath":"unused.json"},"operations":[{"op":"update","target":{"slideId":"s1","elementId":"a"},"patch":{"text":"x"}},{"op":"replace","target":{"slideId":"s1","elementId":"b"},"value":{"id":"b","type":"text","text":"y"}},{"op":"replace","target":{"slideId":"s2"},"value":{"id":"s2","type":"content","elements":[{"id":"t","type":"text","text":"z"}]}}]}
'@ | ConvertFrom-Json
        $ids = Get-PptAffectedSlideIdsFromEdit -Edit $edit
        $ids.Count | Should Be 2
        $ids[0] | Should Be 's1'
        $ids[1] | Should Be 's2'
    }
}
