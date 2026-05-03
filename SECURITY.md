# Security Policy

## 支持的版本

我们以 **`main` 分支最新提交** 为活跃维护线；历史 tag 仅作快照，不保证回溯修补。

## 报告漏洞

若你认为发现**可利用的安全问题**（尤其是远程代码执行、凭据窃取、SSRF 绕过 `imageUrl` 策略等），请：

1. **不要**在公开 Issue 中披露利用细节；或
2. 在 GitHub 使用 **Private vulnerability report**（若仓库已开启），或发邮件至 **then_132423@qq.com** 说明影响范围与复现思路。

我们会在合理时间内确认并修复；公开披露前会协调修复版本与说明。

## 已知设计取舍

- **`imageUrl`**：实现包含 SSRF 防护（私网/回环拒绝、重定向重检、体积与 MIME 限制等）。生产环境请配置 **`assetFetch.allowedHosts`** 白名单。
- **`exec` action**：仅应用于**可信**仓库内脚本路径；勿对不可信输入开放任意路径执行。
- **COM 自动化**：依赖本机 WPS 与用户权限；本工具不提升进程权限。

## 依赖与供应链

- 扩展侧依赖见 `extension/package-lock.json`；Python 图表依赖见 `requirements-charts.txt`。升级依赖时请注意 CI 与 `doctor` 输出。
