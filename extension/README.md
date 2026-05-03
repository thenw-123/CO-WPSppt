# CO-WPSppt（Cursor / VS Code 扩展）

**显示名称：CO-WPSppt** — 在编辑器内调用打包好的 PowerShell + **WPS 演示 COM** 工具链（`run-spec`、`validate-spec`、`doctor` 等）。

- **扩展 ID**：`co-wpsppt.co-wpsppt`（安装后在「已安装扩展」里可见）
- **配置项命名空间**：仍为 `wpsPpt`（与旧版兼容，无需改设置键名）

## 环境要求

- **Windows**，已安装 **WPS 演示**
- **PowerShell 5.1+**（或把设置里的 `wpsPpt.powershellExecutable` 改成 `pwsh.exe`）
- 可选：本机 **Python 3** + 对打包目录执行 `pip install -r requirements-charts.txt` 以生成 PNG 图表

## 本地打包 `.vsix`

```bash
cd extension
npm install
npm run vscode:package
```

生成的文件位于本目录，文件名形如：

`co-wpsppt-<version>.vsix`

（`<version>` 与 `package.json` 中 `version` 一致。）

## 安装

- **Cursor / VS Code**：命令面板 → **Extensions: Install from VSIX…** → 选择上述 `.vsix`
- 或 CLI：`code --install-extension co-wpsppt-1.6.0.vsix`（版本号请按实际文件名替换）

## 命令（命令面板里搜「CO-WPSppt」）

| 命令 | 说明 |
|------|------|
| **CO-WPSppt: Doctor** | 环境检查（COM、Python、路径等） |
| **CO-WPSppt: 校验 spec JSON…** | 选择 spec 文件做结构校验 |
| **CO-WPSppt: Run spec（生成演示稿）…** | 选择 spec 生成 `.pptx`（输出见 Output 通道中的 `presentationPath`） |
| **CO-WPSppt: 运行示例 spec** | 快速冒烟测试 |
| **打开 manifest / Skill / specs** | 文档与示例辅助 |

输出默认在扩展自带的 **`bundled/output/`**；也可在 spec 里写**工作区绝对路径** 的 `savePath`。

## 发布到 GitHub Releases

见仓库根目录 **[docs/CO-WPSppt-GitHub发布.md](../docs/CO-WPSppt-GitHub发布.md)**：推送 `v1.x.x` 标签后，GitHub Actions 会自动构建并上传 `.vsix` 到 Release。

发布前请把 **`package.json` 里的 `repository` / `bugs` / `homepage`** 改成你的 GitHub 地址。

## 安全说明

spec 中的远程 **`imageUrl`** 经 SSRF 加固；生产环境建议配置 **`assetFetch.allowedHosts`**。详见仓库根目录 `README.md`。
