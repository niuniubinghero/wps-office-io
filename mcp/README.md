# wps-office-io —— MCP Server

把「借助 WPS/Office COM 读写 Office 文档」的能力暴露成 MCP 工具，供任何支持 MCP
的客户端调用（WorkBuddy / Cursor / Claude Code / Trae / VS Code Copilot 等）。
适用于读取、修改被天锐绿盾（E-SafeNet）透明加密的文档。

## 架构

```
MCP 客户端 ──stdio/JSON-RPC──> mcp/server.py  ──subprocess──> scripts/wpsdoc.py ──COM──> WPS
                                (协议层，无需 pywin32)        (真正的读写逻辑)
```

`server.py` **只做协议层**，实际工作全部交给 `scripts/wpsdoc.py`，因此：
复用已验证的逻辑（加密检测、ProgID 回退、Find 位置参数修正），且 server 自身
不必安装 pywin32。

## 提供的工具

| 工具 | 作用 | 写操作 |
|---|---|---|
| `wps_probe` | 加密检测 + 页数/表格数/工作表 + 分页预览 | 否 |
| `wps_read` | 读 Word 全文或指定页 | 否 |
| `wps_tables` | 导出 Word 表格（md/csv） | 否 |
| `wps_sheet_read` | 读 Excel 工作表 | 否 |
| `wps_replace` | 替换 Word 文字（支持预演） | 是 |
| `wps_cell` | 写 Word 表格单元格（支持预演） | 是 |
| `wps_sheet_write` | 写 Excel 单元格（支持预演） | 是 |

写操作默认生成 `<原文件>.bak` 备份；建议一律先以 `dry_run=true` 预演。

## 启用步骤

配置已写入 `~/.workbuddy/mcp.json`（与既有 `esafe-fs` 并存）：

```json
"wps-office-io": {
  "command": "C:\\Users\\<USERNAME>\\.workbuddy\\binaries\\python\\envs\\wps-mcp\\Scripts\\python.exe",
  "args": ["C:\\Users\\<USERNAME>\\.workbuddy\\skills\\wps-office-io\\mcp\\server.py"],
  "env": { "WPSDOC_TIMEOUT": "180" }
}
```

**写完配置不会自动生效**：需在连接管理页面右上角的「自定义连接器」入口里，
对 `wps-office-io` 点一次「信任」才会启用。

## 依赖

独立 venv：`C:\Users\<USERNAME>\.workbuddy\binaries\python\envs\wps-mcp`

```
pip install "mcp<2"
```

> **必须锁 `<2`**：mcp 2.x 把 `FastMCP` 改名为 `MCPServer` 并调整了 API，
> 本 server 用的是 v1 写法（`from mcp.server.fastmcp import FastMCP`）。
> 本机实际安装版本 1.30.0。

`wpsdoc.py` 一侧需要 pywin32，它会自动寻找带 pywin32 的解释器（本机为 miniconda）。

## 环境变量

| 变量 | 说明 |
|---|---|
| `WPSDOC_PYTHON` | 指定执行 `wpsdoc.py` 的解释器；省略则用 server 自身解释器，由 wpsdoc 自动切换到带 pywin32 的那个 |
| `WPSDOC_TIMEOUT` | 单次调用超时秒数，默认 180 |

## 排错要点（都是实测踩过的）

1. **子进程必须隔离 stdin**：`subprocess.run(..., stdin=subprocess.DEVNULL)`。
   若让 WPS/COM 子进程继承 MCP 的 stdio 管道，实测会挂起直到超时（单独执行却只要 6 秒）。
2. **server 绝不向 stdout 打印**：stdio 传输用 stdout 承载 JSON-RPC，污染即协议崩溃。
   诊断信息一律 `print(..., file=sys.stderr)`。
3. **mcp 版本**：装到 2.x 会报 `No module named 'mcp.server.fastmcp'`，需 `pip install "mcp<2"`。
4. 调用慢是 WPS 冷启动所致（通常 5–15 秒），不是故障；超时阈值给到 180 秒较稳妥。

## 自检

```bash
python .workbuddy/mcp_selftest.py
```

会完成：启动 server → `initialize` → `tools/list` → 实地调用 `wps_probe`。
