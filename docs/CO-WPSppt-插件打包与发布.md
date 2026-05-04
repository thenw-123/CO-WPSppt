# CO-WPSppt 插件：打包、安装与发布（完整说明）

本文说明如何把本仓库打成 **VS Code / Cursor 可用的 `.vsix` 完整插件包**，以及如何走 **GitHub Release（仓库存档）** 自动构建。

---

## 一、插件里装了什么

扩展根目录在 `extension/`，发布前会执行 `scripts/copy-bundled.mjs`，把运行时资源复制到 **`extension/bundled/`**，再打 VSIX。当前会打包：

| 目录/文件 | 说明 |
|-----------|------|
| `tools/` | `wps_dispatch.ps1` 及各 action 脚本（含 `compile-dsl`、`run-dsl` 等） |
| `wps-driver/` | COM 驱动、**Layout Engine**、Atomic Renderer、DSL 校验等 |
| `specs/` | 示例与 JSON Schema（含 DSL / RenderPlan） |
| `themes/` | 主题相关资源（若存在） |
| `tests/` | Pester 等（便于随包排查） |
| `skills/` | Skill 文档（若存在） |
| `manifest.json` | OpenClaw 网关工具清单（`wps-ppt`） |
| `docs/` | 项目文档（含本说明） |
| `requirements-charts.txt` | 图表可选依赖说明 |

扩展本体逻辑在 `extension/src/extension.ts`，编译输出为 `extension/out/`。

---

## 二、本地环境要求

- **Node.js** 18+（推荐 20，与 CI 一致）
- **npm**
- **全局或本地的 `@vscode/vsce`**（已在 `extension/package.json` 的 `devDependencies` 中，用 `npx vsce` 即可）

**仅打插件包、不跑 WPS**：不需要安装 WPS；若要在本机试生成 PPT，需 Windows + 已安装 WPS 演示。

---

## 三、本地完整打包（生成 `.vsix`）

在仓库根目录执行：

```powershell
cd "d:\第一个系统\cursor-openclaw-wps-ppt\extension"
npm ci
npm run vscode:package
```

成功后在 **`extension/`** 下得到类似：

`co-wpsppt-1.x.x.vsix`

说明：

- `vscode:prepublish` 会先 **`tsc` 编译**扩展，再执行 **`node ../scripts/copy-bundled.mjs`** 刷新 `bundled/`。
- `vscode:package` 使用 **`vsce package --no-dependencies`**，避免把未声明的依赖打进包（与当前工程约定一致）。

---

## 四、安装到 Cursor / VS Code

1. 打开命令面板：**Extensions: Install from VSIX…**
2. 选择上面生成的 **`co-wpsppt-*.vsix`**
3. 重载窗口后，可用命令例如：
   - **CO-WPSppt: Doctor（环境检查）**
   - **CO-WPSppt: Run spec（生成演示稿）…**
   - **CO-WPSppt: 打开打包内的 manifest.json**

扩展通过设置里的 **`wpsPpt.powershellExecutable`** 调用打包内的 `tools/wps_dispatch.ps1`（路径相对于扩展安装目录下的 `bundled`）。

---

## 五、GitHub 仓库存档与自动 Release（打 Tag）

工作流文件：`.github/workflows/release-vsix.yml`

- **触发条件**：推送符合 **`v*`** 的 **tag**（例如 `v1.7.0`）。
- **运行环境**：`windows-latest`。
- **步骤**：在 **`extension/`** 下 `npm ci` → `npm run vscode:package` → 把 **`extension/*.vsix`** 作为附件发布到 **GitHub Release**。

### 推荐发布流程（示例）

```powershell
# 1. 已更新 extension/package.json 与根目录 manifest.json 的 version，并提交
git add -A
git commit -m "chore: release 1.7.0"
git tag v1.7.0
git push origin main
git push origin v1.7.0
```

推送 tag 后，在仓库 **Actions / Releases** 中查看构建产物。

---

## 六、版本号对齐（避免混淆）

发布前请对齐两处版本（至少主版本号一致）：

1. **`extension/package.json`** → `"version"`
2. **仓库根 `manifest.json`** → `"version"`（OpenClaw 工具元数据）

VSIX 文件名中的版本来自 **`extension/package.json`**。

---

## 七、OpenClaw 网关与 manifest

网关通过 **`manifest.json`** 发现 `wps-ppt` 工具，入口脚本为 **`tools/wps_dispatch.ps1`**。

若插件安装在 Cursor 扩展目录，网关侧需指向 **含 `bundled` 的仓库克隆** 或自行同步的目录；具体以你当前 OpenClaw 配置为准（参见仓库 `docs/` 中 MCP 与网关说明）。

---

## 八、常见问题

**Q：`npm run vscode:package` 报错找不到 vsce**  
A：在 `extension` 目录执行 `npx @vscode/vsce package`，或先 `npm ci` 再 `npm run vscode:package`。

**Q：打出来的 VSIX 里没有最新 `wps-driver`**  
A：确认执行过 `vscode:prepublish`（已包含 `copy-bundled`）；不要手动删掉 `extension/bundled` 后又跳过复制步骤。

**Q：只想验证 PowerShell 管线，不打 VSIX**  
A：在仓库根运行 `Invoke-Pester .\tests\Validation.Tests.ps1`，或使用 `tools/wps_compile_dsl.ps1` / `tools/wps_run_dsl.ps1`。

---

## 九、与「布局引擎 / DSL」的关系

当前 **`wps-driver/Wps.LayoutEngine.ps1`** 负责：DSL → 坐标化 **RenderPlan**；**`Wps.AtomicRenderer.ps1`** 仅执行原子操作。上述脚本均随 **`bundled/wps-driver`** 进入 VSIX，无需在扩展里单独配置。
