# 用 Cursor MCP 调用 wps-ppt（OpenClaw Gateway）

在 Cursor 里配置 **OpenClaw Gateway** 的 MCP 后，对话里可以让 Agent 走 **`openclaw_invoke`**，而不用手写本机 PowerShell（与 [CO-WPSppt 扩展](CO-WPSppt-GitHub发布.md) / 本地 `wps_dispatch.ps1` 并行）。

## 1. MCP 工具（Gateway 提供）

| 工具 | 作用 |
|------|------|
| **`openclaw_discover`** | 列出网关上当前可用的工具；**应先调用**，确认是否有 `wps-ppt`。 |
| **`openclaw_skill`** | 拉取某个工具的完整说明（`tool` 填 `wps-ppt`）。 |
| **`openclaw_invoke`** | 按工具名调用；`wps-ppt` 时传入 `action` 与 **`args_json`**（字符串形式的 JSON）。 |

`openclaw_invoke` 参数要点：

- **`tool`**（必填）：`wps-ppt`
- **`action`**：与根目录 [manifest.json](../manifest.json) 里 `actions[].name` 一致，例如 `run-spec`、`validate-spec`、`doctor`
- **`args_json`**：一行 JSON 字符串，例如校验 spec：  
  `{"specPath":"specs/example-spec.json"}`  
  路径一般相对**网关执行时的工程根**；若网关以本仓库为根，则与仓库内 `specs/` 一致。

## 2. 在 Cursor 里启用 Gateway MCP

1. 确认 **OpenClaw Gateway** 已安装并在本机/内网可访问（具体端口与启动方式以你的 Gateway 文档为准）。
2. 打开 **Cursor → Settings → MCP**，添加 Gateway 对应的 MCP Server（名称常见为 `openclaw-gateway` 或你自定义的 `user-openclaw-gateway`）。
3. 保存后可在 **MCP 资源/工具** 里看到 `openclaw_discover`、`openclaw_invoke` 等。

若使用项目级配置，部分环境支持仓库下的 **`.cursor/mcp.json`**（以你当前 Cursor 版本文档为准）。

## 3. 在对话里怎么用（给 Agent 的口令示例）

- 「先 **`openclaw_discover`**，确认有没有 **`wps-ppt`**。」
- 「用 **`openclaw_skill`** 查 **`wps-ppt`** 的 action 列表。」
- 「**`openclaw_invoke`**：`tool` = `wps-ppt`，`action` = `validate-spec`，`args_json` = `{\"specPath\":\"specs/art-philosophy-beauty.json\"}`。」
- 「同上，`action` 改成 `run-spec` 生成 pptx。」

注意：`args_json` 在部分客户端里需转义引号；若 Gateway 走 Shell 模板，需与 [manifest.json](../manifest.json) 里 `preferredDispatch` 的 JSON 约定一致。

## 4. 常见问题

### `openclaw_discover` 显示 Available tools (0)

说明 **Gateway 未连上插件列表**，或当前进程没有注册任何工具。请检查：

- Gateway 进程是否启动；
- 是否在 Gateway 中 **注册/加载** 了本仓库的 **wps-ppt**（manifest 路径、`tool` 名称是否为 `wps-ppt`）；
- Cursor 里 MCP Server URL / 命令是否与 Gateway 一致。

### `openclaw_skill` / `openclaw_invoke` 报 Not found: wps-ppt

网关上 **没有名为 `wps-ppt` 的工具**。需要在 OpenClaw Gateway 配置里把本项目的 [manifest.json](../manifest.json) 加为工具源，并重启或热加载。

### 与「本地脚本」的关系

| 方式 | 适用场景 |
|------|----------|
| **MCP + openclaw_invoke** | 已部署 Gateway，且已注册 `wps-ppt`。 |
| **`tools/wps_dispatch.ps1`** | 本机直接跑，不经过 Gateway。 |
| **CO-WPSppt 扩展** | 在 Cursor 里点命令跑打包后的 `bundled/tools/wps_dispatch.ps1`。 |

三者选其一即可跑同一套 `run-spec` / `validate-spec` 逻辑。
