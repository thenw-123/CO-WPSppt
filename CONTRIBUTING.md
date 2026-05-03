# Contributing to CO-WPSppt

感谢你愿意改进本项目。以下为维护者期望的贡献方式（中英对照要点）。

## 行为准则

请保持专业、尊重与建设性讨论；议题与 PR 请围绕**可复现的问题**或**清晰的改进目标**。

## 如何提交变更

1. **Fork** 仓库并基于 `main` 创建分支（建议命名：`fix/…`、`feat/…`、`docs/…`）。
2. 尽量**小步提交**，一条 PR 解决一类问题，避免无关格式化大杂烩。
3. 变更 **PowerShell 校验逻辑**时，请同步更新或补充 **`tests/Validation.Tests.ps1`**（及相关的 `UrlAsset` 等测试），并在本机执行：
   ```powershell
   Invoke-Pester .\tests\Validation.Tests.ps1
   Invoke-Pester .\tests\UrlAsset.Tests.ps1
   ```
4. 变更 **JSON Schema 或 `Wps.ValidateSpec.ps1`** 时，请同步 **`specs/ppt-spec.schema.json`**，避免文档与实现漂移。
5. 修改 **CO-WPSppt 扩展**时，在 `extension/` 下执行 `npm run compile`；发版前 `npm run vscode:package` 应能通过。

## 目录约定

| 区域 | 用途 |
|------|------|
| `tools/`、`wps-driver/` | 网关与扩展共用的运行时脚本（保持向后兼容的 `action` 与 stdout JSON 契约）。 |
| `specs/` | 小型契约示例与 Schema；**大型完整 deck** 放在 `examples/decks/`。 |
| `examples/` | 试跑素材、生成脚本、配套静态资源。 |
| `releases/` | 已发布的 `.vsix` 与安装说明（可选，与 GitHub Releases 二选一或并存）。 |

## 提交信息

建议使用简短英文或中文动宾短语，例如：`fix: validate thesis-vertical steps`、`docs: update specs README`。

## License

向本仓库提交代码即表示你同意在 **MIT License**（见根目录 [LICENSE](LICENSE)）下授权你的贡献。
