# AI 生成 PPT 三层架构

本项目现在支持两条入口：

- Legacy：`run-spec` 继续消费 `specs/ppt-spec.schema.json`，直接进入 WPS COM 渲染。
- DSL：`run-dsl` 消费 `specs/dsl/ppt-dsl.schema.json`，先编译为 RenderPlan，再进入 WPS COM 渲染。

## 分层

```mermaid
flowchart LR
  userInput["User Input"] --> intentLayer["Intent Layer"]
  intentLayer --> pptDsl["PPT DSL"]
  pptDsl --> compiler["DSL Compiler"]
  compiler --> renderPlan["RenderPlan"]
  renderPlan --> wpsRenderer["WPS Renderer"]
  wpsRenderer --> pptx["PPTX"]
```

## Intent Layer

Intent Layer 面向 agent 和产品能力，输出的是演示规划：标题、受众、章节、每页要传达的 message、证据和素材意图。它不应该输出 WPS COM、PowerShell action、坐标、placeholder index 或 shapeIndex。

当前仓库先落地了 DSL 之后的可执行部分，`src/intent/` 作为后续 agent prompt contract 和规划器代码的边界。

## DSL Layer

DSL 是 agent 和渲染系统之间的稳定契约：

- Schema：`specs/dsl/ppt-dsl.schema.json`
- 示例：`specs/dsl/example-dsl.json`
- 类型：`src/dsl/types.ts`
- 编译器：`wps-driver/Wps.Dsl.ps1`

DSL 使用语义化 slide kind，例如 `cover`、`content`、`argument`、`comparison`、`timeline`、`dataChart`、`swot`。agent 只需要选择语义和内容，不需要知道 WPS layout int。

## Renderer Layer

Renderer Layer 负责执行。第一阶段的 RenderPlan 位于 `specs/dsl/render-plan.schema.json`，由 `wps-driver/Wps.RenderPlan.ps1` 生成和执行。

当前主链路已经收敛为：

```text
PptDsl -> Layout Engine -> atomic RenderPlan -> WPS atomic renderer -> WPS COM
```

Layout Engine 位于 `wps-driver/Wps.LayoutEngine.ps1`，是纯函数：输入无坐标的 DSL slide，输出带 `box: {x,y,w,h}` 的原子操作。当前支持 `title-content` 和 `image-right` 两个最小布局。

Renderer 位于 `wps-driver/Wps.AtomicRenderer.ps1`，只执行 `addText`、`addBullets`、`addImage` 等 RenderPlan operation，不再判断 slide type、layout 或业务语义。Legacy spec 仍可通过 `compile-dsl` 的 `format=legacySpec` 显式输出，用于兼容旧路径。

## Actions

```powershell
.\tools\wps_dispatch.ps1 -Action compile-dsl -ArgsJson '{"dslPath":"specs/dsl/example-dsl.json"}'
.\tools\wps_dispatch.ps1 -Action compile-dsl -ArgsJson '{"dslPath":"specs/dsl/example-dsl.json","format":"legacySpec"}'
.\tools\wps_dispatch.ps1 -Action run-dsl -ArgsJson '{"dslPath":"specs/dsl/example-dsl.json"}'
.\tools\wps_dispatch.ps1 -Action compile-dsl -ArgsJson '{"dslPath":"specs/dsl/layout-engine-minimal.json"}'
```

`run-spec` 保持兼容，legacy schema 快照在 `specs/legacy/ppt-spec.schema.json`。

## 扩展新页型

1. 在 `specs/dsl/ppt-dsl.schema.json` 和 `src/dsl/types.ts` 中加入新的 `kind`。
2. 在 `wps-driver/Wps.Dsl.ps1` 中把新 `kind` 编译为 RenderPlan/legacy spec。
3. 如果 legacy spec 无法表达，再扩展 `Wps.RenderPlan.ps1` 和 WPS renderer。
4. 为 DSL 编译和 legacy spec 校验补测试。
