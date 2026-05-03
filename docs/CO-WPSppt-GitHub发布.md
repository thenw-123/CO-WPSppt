# CO-WPSppt — 在 GitHub 上发布扩展（.vsix）

扩展在商店里的**显示名称**为 **CO-WPSppt**；技术标识为 `co-wpsppt.co-wpsppt`（`publisher` + `name`）。本地 `vsce package` 打出来的文件名一般为：

`co-wpsppt-<version>.vsix`

（版本号见 `extension/package.json` 的 `version`。）

## 一、首次推到 GitHub

1. 若尚未建库：可用 **GitHub MCP** `create_repository`（本仓库已对应 **https://github.com/thenw-123/CO-WPSppt**），或网页新建空仓库（不要勾选自动 README，避免首次推送冲突）。
2. 在项目根目录初始化/绑定远程并推送：

   ```powershell
   cd "d:\第一个系统\cursor-openclaw-wps-ppt"
   git init
   git checkout -b main
   git remote add origin https://github.com/thenw-123/CO-WPSppt.git
   git add -A
   git commit -m "chore: initial CO-WPSppt"
   git push -u origin main
   ```

3. **`extension/package.json`** 里的 `repository` / `bugs` / `homepage` 应与上述远程一致（当前已指向 `thenw-123/CO-WPSppt`）。

## 二、本地打包（不上传 Release 时）

```powershell
Set-Location -LiteralPath "d:\第一个系统\cursor-openclaw-wps-ppt\extension"
npm install
npm run vscode:package
```

在 `extension` 目录下得到 `.vsix`，用 Cursor：**Extensions: Install from VSIX…** 安装。

## 三、用 GitHub Releases 自动发布（推荐）

本仓库已配置 **打 `v*` 标签即构建并上传 VSIX** 的工作流（见 `.github/workflows/release-vsix.yml`）。

1. 确认 `extension/package.json` 里 **`version`** 与本次发布一致（例如 `1.6.0`）。
2. 提交并推送代码到 `main`（或你的默认分支）。
3. 创建并推送标签（示例）：

   ```powershell
   cd "d:\第一个系统\cursor-openclaw-wps-ppt"
   git add -A
   git commit -m "release: CO-WPSppt v1.6.0"
   git push
   git tag v1.6.0
   git push origin v1.6.0
   ```

4. 打开 GitHub 仓库 → **Actions**，等待 **Release VSIX (CO-WPSppt)** 跑完。
5. 打开 **Releases**，应看到对应版本及附件中的 `.vsix`。用户下载后 **Install from VSIX** 即可。

**权限**：工作流使用仓库默认 `GITHUB_TOKEN` 创建 Release；若你关闭了 Actions 写权限，需在仓库 **Settings → Actions → General** 中允许工作流读写内容。

## 四、CI 仅验证构建（不上传 Release）

推送/PR 修改 `extension/`、`tools/`、`wps-driver/` 等路径时，**Extension build** 工作流会构建 VSIX 并上传为 **Artifact**（供下载检查，不自动发 Release）。

## 五、与 VS Marketplace 的区别

本文档描述的是 **GitHub Releases 分发 .vsix**。若还要上架 **Visual Studio Marketplace**，需单独注册 Publisher、改 `publisher` 字段并执行 `npx vsce publish`（见 [官方文档](https://code.visualstudio.com/api/working-with-extensions/publishing-extension)）。
