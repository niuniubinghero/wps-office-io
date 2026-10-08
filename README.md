# wps-office-io

通过 WPS / MS Office 的 COM 自动化读写 Windows 上的 Office 文档，专门解决天锐绿盾
（E-SafeNet / EsafeNet Cobra DocGuard）等**透明加密（DLP）环境**下 `.doc/.docx/.xls/.xlsx/.ppt/.pptx`
无法读写的问题。

同一份读写内核 `scripts/wpsdoc.py`，提供两个入口：

| 入口 | 适用场景 | 特点 | 文档 |
|---|---|---|---|
| Agent Skill | 支持 Agent Skills 的客户端（如 WorkBuddy） | 零配置，加载即用，无需常驻进程 | [SKILL.md](SKILL.md) |
| MCP Server | 任何 MCP 客户端（WorkBuddy / Cursor / Claude Code / Trae 等） | 常驻服务，工具在所有会话可见 | [mcp/README.md](mcp/README.md) |

两个入口调用的是同一份逻辑，**不存在双份维护**：MCP server 只做协议层，实际工作全部
转交给 `scripts/wpsdoc.py`。

## 为什么必须有它

绿盾是**透明加密**：白名单进程（WPS / Word）读取时自动解密，写回时自动再加密。由此有三条铁律：

1. **读**必须经 COM。直接读磁盘字节只能拿到密文（`python-docx` / `openpyxl` / `olefile` 一律报错或返回乱码）。
2. **写**必须经 COM。保存时驱动自动再加密，**不会破坏文件的加密状态**。
3. **禁止**绕过 WPS 直接改字节——会把密文封装写坏，文件直接报废。

顺带一个实用结论：常见的绿盾策略往往只对 Word 类加密，`.xls`/`.xlsx` 可能仍是明文，
处理前建议先用 `probe` 确认，能省掉一次 WPS 冷启动。

## 仓库结构

```
wps-office-io/
├── README.md              # 本文档
├── requirements.txt       # mcp / pywin32
├── SKILL.md               # Agent Skill 入口定义
├── scripts/
│   └── wpsdoc.py          # 读写内核（CLI）：probe/read/tables/replace/cell/sheet-read/sheet-write
├── mcp/
│   ├── server.py          # MCP Server（仅协议层，转发到 wpsdoc.py）
│   └── README.md          # MCP 配置、排错、自检
└── references/
    └── com-api.md         # COM 常量、ProgID 回退、分页与文本结构符原理
```

## 安装

### 方式一：作为 Agent Skill

把整个目录放到 skills 目录下（`~/.workbuddy/skills/wps-office-io/`）即可，无需任何配置。

### 方式二：作为 MCP Server

在客户端的 MCP 配置文件中加入（路径按本机实际替换）：

```json
{
  "mcpServers": {
    "wps-office-io": {
      "command": "<PYTHON_WITH_MCP>",
      "args": ["<THIS_REPO>/mcp/server.py"],
      "env": { "WPSDOC_TIMEOUT": "180" }
    }
  }
}
```

依赖安装见下方「依赖」一节。写完配置通常需要在客户端的自定义连接器入口点一次「信任」才会启用。

## 能力一览

| 能力 | 说明 | 写操作 |
|---|---|---|
| `probe` | 加密检测 + 页数/段落/表格/工作表清单 + 分页预览 | 否 |
| `read` | 读 Word 全文或**指定页**（普通库做不到分页） | 否 |
| `tables` | 导出 Word 表格为 markdown / csv | 否 |
| `sheet-read` | 读 Excel 工作表 | 否 |
| `replace` | 替换 Word 文字（支持预演） | 是 |
| `cell` | 写 Word 表格单元格（适合补签日期、签署栏） | 是 |
| `sheet-write` | 写 Excel 单元格 | 是 |

## 操作规范

- 改动前先 `probe` + `read`，确认目标内容真实存在、页码正确。
- 写操作**必须先预演**（CLI 的 `--dry-run` / MCP 的 `dry_run=true`），确认命中数与预期一致再执行。
- 写操作默认生成 `<原文件>.bak` 备份。
- 改完后用 `read` 或 `tables` **读回校验**，不要仅凭返回值下结论。

## 依赖

- Windows + WPS（或 MS Office）
- `wpsdoc.py` 需要 `pywin32`（会自动寻找带 pywin32 的解释器，也可设 `WPSDOC_PYTHON` 指定）
- MCP Server 需要 `mcp>=1.0,<2.0`（**必须锁 2.x 以下**，2.x 已移除 `FastMCP` 写法）

```bash
pip install -r requirements.txt
```

## 已知坑

1. **COM 代码不要用 Bash heredoc 内联执行**——会被沙箱以 SIGTERM 终止且无任何输出。写成脚本文件再跑。
2. 单次调用通常 5–15 秒，是 WPS 冷启动所致，不是故障。
3. WPS 会**静默忽略** `Find.Execute` 的关键字参数（不报错但只查不换），必须用位置参数。

更多排错详见 [references/com-api.md](references/com-api.md) 与 [mcp/README.md](mcp/README.md)。
