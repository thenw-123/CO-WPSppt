# PPT DSL v2 工程规范

PPT DSL v2 的目标是让 AI 输出语义结构，而不是输出 `add_slide`、`set_text`、坐标、COM placeholder 等操作指令。Renderer 根据语义、布局偏好和主题 token 决定具体 WPS COM 调用。

## 顶层结构

必填字段：

- `dslVersion`: 固定为 `2.0`。
- `meta`: 演示稿元信息，`meta.title` 必填。
- `slides`: 幻灯片数组，至少 1 页。

可选字段：

- `theme`: 全局视觉 token，包括 `preset`、`template`、`fonts`、`colors`、`motion`。
- `defaults`: 全局默认行为，包括 `layout`、`imagePlacement`、`chartEngine`、`assetFetch`。

Schema 文件：`specs/dsl/ppt-dsl.schema.json`。

## Slide 结构

每页使用语义字段：

- `type`: 必填。枚举：`cover`、`agenda`、`section`、`content`、`argument`、`comparison`、`timeline`、`data`、`swot`、`custom`。
- `elements`: 必填。按阅读和叙事顺序排列，不代表绘制顺序。
- `layout`: 可选。默认 `auto`，只表达版式偏好。
- `message`: 可选。一句话结论，供布局和讲稿使用。
- `style`: 可选。页级视觉 token。
- `notes`: 可选。演讲者备注。
- `motion`: 可选。页级动画和转场偏好。

默认行为：

- `layout.name` 省略或为 `auto` 时，Renderer 根据 `type` 和 `elements` 自动选版式。
- `title` 省略时，Renderer 优先使用 `role: title` 的 text element。
- `message` 可作为标题兜底，但不强制渲染。
- 图片、图表、表格是否可视化失败时，由 renderer fallback 策略决定。

## Element 系统

所有 element 都有：

- `type`: 必填。枚举：`text`、`bullets`、`image`、`chart`、`table`。
- `role`: 可选。语义角色，如 `title`、`body`、`claim`、`evidence`、`insight`、`caption`、`decorative`。
- `style`: 可选。元素级视觉 token。
- `layoutHint`: 可选。表达 `placement`、`emphasis`、`span` 等偏好。

### text

必填：`text`。

用于标题、副标题、论点、洞察、引用、页脚等。由 `role` 决定默认字号和位置。

### bullets

必填：`items`。

`items` 可以是字符串数组，也可以是对象数组：

```json
{ "text": "集成测试环境仍是主要风险", "level": 1, "evidence": "影响回归节奏" }
```

默认行为：`level` 默认为 1，Renderer 根据页级 `constraints.maxBullets` 决定截断、缩小字号或拆页。

### image

必填：`source`。

`source.kind` 支持：

- `local`: 本地路径，使用 `path`。
- `remote`: 远程图片，使用 `url`，受 `defaults.assetFetch` 约束。
- `commons`: Wikimedia Commons，使用 `title`。
- `generated`: 生成式图片意图，使用 `prompt`。
- `template`: 模板内置素材，使用 `assetId`。

默认行为：`layoutHint.placement` 缺省为 `auto`，Renderer 按版式选择右侧、底部、背景或全幅。

### chart

必填：`chart.kind`、`chart.data.labels`、`chart.data.series`。

支持 `bar`、`line`、`donut`、`pie`、`area`、`scatter`。当前 WPS 兼容渲染优先支持 `bar`、`line`、`donut`，其他图表可由后续 renderer 扩展或降级。

默认行为：`chart.render.engine` 为 `auto` 时，Renderer 可选择 COM 图表或 PNG 图表；失败时按 `fallback` 降级到文字或 PNG。

### table

必填：`columns`、`rows`。

用于对比矩阵、参数表、决策表。`highlight.rows` 和 `highlight.columns` 使用 0-based 索引，表达语义强调，不是绘图命令。

## 示例

- 简单示例：`specs/dsl/example-simple.json`
- 复杂示例：`specs/dsl/example-complex.json`

编译验证：

```powershell
.\tools\wps_dispatch.ps1 -Action compile-dsl -ArgsJson '{"dslPath":"specs/dsl/example-simple.json"}'
.\tools\wps_dispatch.ps1 -Action compile-dsl -ArgsJson '{"dslPath":"specs/dsl/example-complex.json"}'
```

## 两阶段 Agent 校验

Planner Agent 输出 outline 后先校验：

```powershell
.\tools\wps_dispatch.ps1 -Action validate-agent -ArgsJson '{"stage":"planner","outlinePath":"src/intent/example-outline.json"}'
```

Writer Agent 输出 PPT DSL v2 后再校验：

```powershell
.\tools\wps_dispatch.ps1 -Action validate-agent -ArgsJson '{"stage":"writer","dslPath":"specs/dsl/example-simple.json"}'
```

Planner Schema 位于 `src/intent/planner.schema.json`，Writer 使用 `specs/dsl/ppt-dsl.schema.json`。PowerShell 版本支持 `Test-Json -Schema` 时会优先执行 JSON Schema 校验；否则使用同等约束的结构化 fallback 校验，保证本机 WPS 自动化环境仍可运行。

## 第一阶段编辑能力

编辑请求使用 `src/edit/ppt-edit.schema.json`，当前支持：

- `replace` slide：按 `slideId` 替换整页。
- `replace` element：按 `slideId + elementId` 替换元素。
- `update` element：按 `slideId + elementId` 浅合并字段。

示例请求见 `src/edit/example-edit.json`。只修改 DSL、不启动 WPS：

```powershell
.\tools\wps_dispatch.ps1 -Action edit-dsl -ArgsJson '{"editPath":"src/edit/example-edit.json","outPath":"output/edited-dsl.json","renderAfterEdit":false}'
```

编辑后触发现有全量 `run-dsl` 渲染：

```powershell
.\tools\wps_dispatch.ps1 -Action edit-dsl -ArgsJson '{"editPath":"src/edit/example-edit.json","renderAfterEdit":true}'
```

## 第二阶段局部渲染

`run-dsl` 全量生成 PPTX 后会自动写入 sidecar slide map：

```text
output/demo.pptx
output/demo.pptx.map.json
```

map 记录 `slideId -> index`，例如：

```json
{
  "mapVersion": "1.0",
  "presentationPath": "output/demo.pptx",
  "dslPath": "specs/dsl/example-simple.json",
  "slides": [
    { "slideId": "cover", "index": 1, "title": "季度项目进展汇报" },
    { "slideId": "summary", "index": 2, "title": "本季度结论" }
  ]
}
```

局部更新使用 `render-edit`。它会先把 edit 应用到 DSL 文件，再打开已有 PPTX，按 sidecar map 定位受影响页，只重建这些页并保存：

```powershell
.\tools\wps_dispatch.ps1 -Action render-edit -ArgsJson '{"editPath":"src/edit/example-edit.json","presentationPath":"output/dsl-simple-demo.pptx"}'
```

如果 map 不在默认位置，可显式传入：

```powershell
.\tools\wps_dispatch.ps1 -Action render-edit -ArgsJson '{"editPath":"src/edit/example-edit.json","presentationPath":"output/dsl-simple-demo.pptx","mapPath":"output/dsl-simple-demo.pptx.map.json"}'
```

当前局部重建实现会用现有 renderer 临时生成单页 PPTX，再复制替换到目标 PPTX 的原 index，因此复用现有 WPS COM 版式、图片、图表和动画逻辑，同时避免重建其他页面。
