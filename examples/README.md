# Examples（示例与完整演示稿）

本目录存放**体积较大或用于端到端试跑**的素材，与 `specs/` 中面向工具链的**小型契约示例**区分。

| 路径 | 说明 |
|------|------|
| [decks/](decks/) | 完整 deck 规格 JSON（从仓库根目录执行 `run-spec` 时 `specPath` 写相对路径）。 |
| [assets/](assets/) | 与示例稿配套的本地 PNG 等（由脚本生成，非运行时必需）。 |
| [scripts/](scripts/) | 生成示例资源的辅助脚本（matplotlib 等）。 |

## 《艺术哲学》完整示例

- 规格：[decks/art-philosophy-beauty.json](decks/art-philosophy-beauty.json)（含 `imageCommons`、叙事版式、主题等）。
- 可选配图：运行 `python examples/scripts/gen_art_philosophy_assets.py` 可在 `examples/assets/art-philosophy/` 下再生成示意图（当前 deck 以维基共享资源为主，脚本为可选增强）。

```powershell
# 在仓库根目录
.\tools\wps_dispatch.ps1 -Action validate-spec -ArgsJson (ConvertTo-Json @{ specPath = 'examples/decks/art-philosophy-beauty.json' })
.\tools\wps_dispatch.ps1 -Action run-spec -ArgsJson (ConvertTo-Json @{ specPath = 'examples/decks/art-philosophy-beauty.json' })
```

**说明**：CO-WPSppt 扩展的 `bundled/` **默认不包含** `examples/`，以控制 VSIX 体积；完整示例请在 **clone 后的仓库根** 使用上述命令。
