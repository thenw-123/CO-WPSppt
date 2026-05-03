# Specs（JSON 规格与 Schema）

面向 **`run-spec` / `validate-spec`** 的 JSON 与机器可读 Schema。

| 文件 | 用途 |
|------|------|
| [ppt-spec.schema.json](ppt-spec.schema.json) | Deck 规格 JSON Schema（与校验脚本对齐）。 |
| [example-spec.json](example-spec.json) | 最小可读示例；扩展内置「Run example」默认指向此文件。 |
| [narrative-layouts-demo.json](narrative-layouts-demo.json) | 叙事版式：`timeline` / `comparison` / `thesis-chain` / `argument` / `thesis-vertical` / `swot`。 |
| [example-chart-png.json](example-chart-png.json) | `chart_data` + matplotlib PNG。 |
| [example-chart-com.json](example-chart-com.json) | `chart_engine`: `com` 内嵌图示例。 |
| [example-imageurl.json](example-imageurl.json) | 安全 `imageUrl` + `assetFetch` 示例。 |
| [templates/](templates/) | 大纲 / 答辩骨架模板，见 [templates/README.md](templates/README.md)。 |

完整主题演示稿（艺术哲学等）见仓库 **[examples/decks/](../examples/decks/)**。
