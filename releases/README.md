# CO-WPSppt 安装包（方式 B · 仓库内分发）

**当前扩展版本：1.6.0**

## 下载 VSIX

1. 在本目录中打开文件 **`co-wpsppt-1.6.0.vsix`**。
2. 点击 **Download**（或 **Raw** 另存为）。

直链（`main` 分支）：

```
https://github.com/thenw-123/CO-WPSppt/raw/main/releases/co-wpsppt-1.6.0.vsix
```

## 安装

在 **Cursor** 或 **VS Code** 中：**命令面板** → **Extensions: Install from VSIX…** → 选择下载的 `.vsix`。

## 说明

- 因当前 GitHub MCP **不提供**「创建 Release / 上传 Release 附件」接口，方式 B 采用 **将 `.vsix` 放在 `releases/` 目录** 并随仓库推送，效果与在 Release 里挂附件相同（用户同样通过浏览器下载）。
- 若你更习惯 **GitHub Releases 页面**：可在网页上 **Draft a new release**，再把本地的 `co-wpsppt-1.6.0.vsix` 拖入 **Assets**。
