---
name: drive-wps-ppt
description: 通过 OpenClaw wps-ppt 工具用 COM 直驱 WPS 做演示稿；含项目答辩类内容的蒸馏叙事规则与 JSON spec 输出约定。
---

# Drive WPS PPT（OpenClaw）

## 何时使用

用户要在 **WPS 演示**里自动化做幻灯片、改字、插图、保存——且希望 **WPS 窗口实时变化**。

当用户要做 **项目答辩 / 汇报 PPT**、且需要「故事化、低密度、带备注与图表建议」的成稿时：先生成符合本文 **「答辩内容生成器」** 规则的 JSON，再 **`run-spec`**（或两阶段：先大纲再逐章）。

## 答辩内容生成器（蒸馏规则）

你是 **项目答辩 PPT 内容生成器**。目标不是堆砌信息，而是讲一个有说服力的故事。

### 1. 故事框架（卡片式叙事）

- 整体顺序：**痛点/背景 → 解决方案概览 → 技术实现细节 → 数据/成果验证 → 未来展望**。
- **封面**：项目名 + **一句动词开头、≤10 字**的价值主张（放在 `subtitle`）。
- **内容页**：一页只讲一个核心观点；章节/目录标题用 **动宾短语**（例：「破解分拣困局」「用算法代替人脑」）。

### 2. 内容密度

- **标题**：≤10 个汉字。
- **要点**：每页 ≤5 条，每条 ≤15 个汉字（对应 `bullets` 或 `points`）。
- 禁止大段正文；长句拆成要点。
- **数据页**：必须包含 **图表类型建议** + **一句核心洞察**（见下文字段）；必要时在 `notes` 里写口播如何强调数字。

### 3. 文案润色

- 技术术语写成 **业务价值**；每条要点能回答「所以呢？」。
- 推荐句式：**通过 [手段]，实现 [可衡量结果]，支撑 [更大目标]**。
- 避免「我们觉得」等主观句，用数据或事实。
- **不要**在最终 JSON 里输出本规则全文或思维链；只输出合法 spec JSON（或用户明确只要大纲时的简化 JSON）。

### 4. 图表与可视化（字段约定）

| 场景 | chart_type | 说明 |
|------|------------|------|
| 对比 | `bar` | 柱状图 |
| 趋势 | `line` | 折线图 |
| 占比 | `donut` | 环形图 |
| 流程/逻辑 | `flowchart` | 流程图或步骤卡片 |

数据页须包含 **`chart_type`、`data_summary`、`insight`**（可与 `layout`: `chart` 同用）。若同时提供 `bullets`/`points`，以列表为准，图表字段可作补充或省略（优先保证页面上有洞察）。

**PNG 真图（优先）**：在当页增加 **`chart_data`**：`{ "labels": [...], "series": [ { "name"?: "...", "values": [数字, ...] } ] }`，且 `chart_type` 为 **`bar` | `line` | `donut`**，`run-spec` 会用 Python+matplotlib 生成 PNG 并插入（需本机 `pip install -r requirements-charts.txt`）。`chart_render: false` 可关闭。无 Python 时仍保留文字要点。

**二期图表**：**`chart_engine`: `"com"`** 时优先 WPS 内嵌图（失败则视 **`chart_fallback_png`** 尝试 PNG）。**远程图**：**`imageUrl`**（HTTPS，建议根级 **`assetFetch.allowedHosts`** 白名单）；勿与 **`image`** 同页共用。安全约束见仓库 README「安全」节。

### 5. 演讲备注（notes）

**每一页**都要有 `notes`，建议包含：本页建议时长、与上一页的过渡语、可能被追问的问题及简短答法。语气口语化，适合口述。

### 6. 输出 JSON（WPS `run-spec` 专用）

必须符合 [specs/ppt-spec.schema.json](../../specs/ppt-spec.schema.json)。与通用 LLM 示例的映射关系：

- `slides[].layout`：使用 `title` | `title-content` | `content` | `chart` | `section` | `two-content` | `blank`，以及叙事版式 **`timeline` | `comparison` | `thesis-chain` | `argument` | `thesis-vertical` | `swot`**（空白版式 + 形状排版，减轻「每页都是标题+要点+配图」的单调感）。其中 **`content` / `chart` 与 `title-content` 同款版式**；`chart` 页可在无要点时用 `chart_type` + `data_summary` + `insight` 自动生成要点行。
- **叙事版式字段**（与 `twoColumns`、`chart_data` **互斥**，校验会报错）：
  - `timeline`：`timeline: [{ "mark"|"year", "text"|"caption" }, …]`（最多渲染 6 点）。
  - `comparison`：`comparison: { "leftHeader", "rightHeader", "rows": [{ "left", "right" }] }`。
  - `thesis-chain`：`chain: [{ "label", "text" }, …]`（至少 2 项，最多 4 卡）。
  - `argument`：`argument: { "thesis", "because": ["…"], "therefore" }`（一页「总—因—收」）。
  - `thesis-vertical`：`thesisVertical: { "claim"|"thesis", "steps": [ 恰好 3 项 `{ "label"?, "text" }` ] }`。
  - `swot`：`swot: { "strengths", "weaknesses", "opportunities", "threats" }`（每项 string 或 string[]；至少一栏非空）；可选 `headers` 覆盖象限标题。
- 完整示例见 [specs/narrative-layouts-demo.json](../../specs/narrative-layouts-demo.json)。
- 列表字段：优先 **`bullets`**；也可用 **`points`**（与 bullets 等价，二者不要同时塞两套不同内容）。
- `theme.font`：可单独指定（如 `微软雅黑`），会同时作用于标题与正文；需要区分时再写 `titleFont` / `bodyFont`。
- **`theme.colors`**：仅作 **设计参考**（Agent/人工套主题）；当前 COM 链路 **不会**自动把 hex 刷进母版，需要配色时请用 WPS 主题或 `theme.themePath`（`.thmx`）。

示例结构（字段名以 schema 为准）：

```json
{
  "savePath": "output/defense-deck.pptx",
  "autoAgenda": true,
  "agendaTitle": "议程",
  "theme": {
    "name": "答辩主题",
    "font": "微软雅黑",
    "colors": { "primary": "#1E3A8A", "secondary": "#0F172A", "accent": "#3B82F6" }
  },
  "slides": [
    {
      "layout": "title",
      "title": "智能仓储调度系统",
      "subtitle": "算法驱动分拣提效",
      "notes": "开场约30秒：点题+为什么听众要关心。过渡：先看行业痛点。"
    },
    {
      "layout": "title-content",
      "title": "为何必须升级调度？",
      "bullets": ["人工分拣峰值顶不住", "错分导致逆向物流成本", "扩产能不能线性堆人"],
      "notes": "约60秒，只抛问题不细讲方案。准备回答：数据来源。"
    },
    {
      "layout": "chart",
      "title": "效率提升对比",
      "chart_type": "bar",
      "data_summary": "人工分拣 vs 算法调度（同仓同 SKU）",
      "insight": "调度算法将分拣效率提升约40%",
      "notes": "强调40%与测试口径；追问：环境边界与样本量。"
    }
  ]
}
```

## 两阶段生成（边确认边做）

1. **大纲阶段**：用户或 Agent 声明只做大纲时，输出的 `slides` 仅保留 **封面** + **各章 `section` 或短 `title-content`（只有章节标题）**，**不写**详细要点与图表字段；`notes` 可写「待展开」。
2. **成稿阶段**：用户确认章节后，按章补全 `bullets`/`points`、`chart_*`、`notes`，再 **`run-spec`** 或 **`open` + 逐页 `set-text` / `set-notes`**。

**参数约定**：生成大纲 JSON 时可在 `args_json` 里加 **`"stage":"outline"`**（供 Agent 自解释）；工具链不解析该字段，仅用于提示模型「只出章节、不出细项」。

## 实机预览后的视觉 QC（反「AI 味」）

用户看过 WPS 里的幻灯片后，Agent 应用下面清单做 **一轮修订**（改 spec 或 `set-text` / `set-notes`），再保存：

- 每页是否 **只有一个核心观点**；标题是否 **动词化、≤10 字**。
- 要点是否 **≤5 条、每条 ≤15 字**；有无 **空洞形容词**、无数据支撑的断言。
- 数据页：**图表类型与数据关系是否一致**（对比→柱、趋势→线、占比→环）；`insight` 是否 **一句人话**、口径可在 `notes` 里备答。
- 过渡：章节之间 **口播衔接** 是否写在 `notes` 里。

## 自测清单（能力验收）

| 输入意图 | 通过标准 |
|----------|----------|
| 智能仓储答辩：分拣效率低 + AGV 调度算法 | 含痛点→方案→技术→**带 chart 字段的数据页**→展望；每页要点 ≤5、每条 ≤15 字；**每页有 notes** |
| 「给数据页加个图」 | 对应页出现 **`chart_type`**（对比用 `bar` 等）及 **`insight`** |
| 「首页标题更吸引人」 | **标题动词化、≤10 字**（如「指挥百台 AGV」类） |

## 标准流程（工具调用）

1. **澄清**：主题、页数、模板/配色、`savePath`（默认 `output/deck-时间戳.pptx`）。
2. **选型**：
   - 一稿成型：**`run-spec`**，JSON 符合 [specs/ppt-spec.schema.json](specs/ppt-spec.schema.json)（可放 `specs/`）。
   - 边聊边改：**`new`** 或 **`open`**，再 **`add-slide` / `set-text` / `set-notes` / `insert-image`**，最后 **`save`**。
3. **调用 OpenClaw**：`openclaw_invoke`，`tool = wps-ppt`，`action` 见 [manifest.json](../../manifest.json)，参数进 **`args_json`**。单脚本路由时用 [`tools/wps_dispatch.ps1`](../../tools/wps_dispatch.ps1)：`-Action` + `-ArgsJson`。
4. **读返回**：解析 stdout 一行 JSON；`ok` 为 false 时把 `error` 与 **`code`**（若有）告诉用户，并按 [README.md](../../README.md) 中的 **`code` 说明表** 建议下一步（如先 `validate-spec`、或 `doctor`）。
5. **交付**：给出 `data.presentationPath`；说明文件在 `output/`。

**诊断**：自动化异常时优先 **`doctor`**（`args_json`: `{}` 或 `{"comProbe":true}`）；大 spec 在 **`run-spec` 前** 可 **`validate-spec`**（同 `specPath` / `spec`）。

### args_json 示例

**run-spec（文件）**

```json
{"specPath":"specs/example-spec.json"}
```

**run-spec（内联 spec 片段）**

```json
{"spec":{"savePath":"output/x.pptx","slides":[{"layout":"title","title":"示例"}]}}
```

**run-spec（跳过校验，仅排障时）**

```json
{"specPath":"specs/example-spec.json","skipValidation":true}
```

**validate-spec**

```json
{"specPath":"specs/example-spec.json"}
```

**doctor**

```json
{}
```

**doctor（尝试连接 WPS）**

```json
{"comProbe":true}
```

**set-notes**

```json
{"slide":2,"notes":"这里讲 30 秒 COM 可见性。"}
```

**new**

```json
{"outPath":"output/my-deck.pptx"}
```

**set-text**

```json
{"slide":2,"title":"第二节","bullets":["要点一","要点二"]}
```

## 故障时

- 检查已装 **WPS 演示**，且本机已注册 `Kwpp.Application` 等 ProgID。
- 无会话时先 **`new`** 或 **`open`**。
- 32/64 位或权限问题：见 [README.md](../../README.md) 排查节。

## 进阶：规则热更新

可将「密度 / 图表映射 / 章节结构」抽成仓库内独立配置文件，在 Agent 侧读取后拼进 system prompt；本仓库以 **本 SKILL** 为默认真源，改 SKILL 即升级默认行为。
