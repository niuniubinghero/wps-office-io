---
name: wps-office-io
description: 通过 WPS/MS Office 的 COM 自动化读写 Windows 上的 Office 文档（doc/docx/xls/xlsx/ppt）。当文档被天锐绿盾（E-SafeNet / E-Safe / 绿盾）透明加密、用 python-docx 或 openpyxl 读到乱码或直接报错、或需要经白名单进程安全读写 Office 文件时使用。支持加密检测、按页码读取、Word 表格导出、文字替换、Excel 单元格读写。
agent_created: true
---

# WPS Office 文档读写

## 用途

在装有天锐绿盾等透明加密（DLP）系统的 Windows 机器上，安全地读取和修改 Office 文档。
核心手段是调用 `scripts/wpsdoc.py`，由它驱动 WPS 的 COM 接口完成读写。

## 何时使用

满足任一条件即应使用本 skill：

- 需要读取或修改 `.doc / .docx / .xls / .xlsx / .ppt / .pptx`
- 文件打开是乱码、或 `python-docx` / `openpyxl` / `olefile` 报错、或读到的文本毫无意义
- 需要读取**指定页码**的内容（普通库做不到分页）
- 需要修改加密文档并保持其加密状态不被破坏

## 核心原理（务必理解）

绿盾是**透明加密**：白名单进程（WPS / Word）读取时自动解密，写入时自动再加密。

由此得出三条铁律：

1. **读**：必须经 WPS COM 取文本。直接读磁盘字节只能拿到密文。
2. **写**：必须经 WPS COM 写回。保存时绿盾会自动再加密，**不会破坏加密状态**。
3. **禁止**绕过 WPS 用 `python-docx` / `openpyxl` / `olefile` 直接改字节——会把密文封装覆盖破坏，文件直接报废。

## 快速开始

```bash
# 1. 先探测：是否加密 + 结构 + 分页预览
python scripts/wpsdoc.py probe "C:/path/to/文件.doc"

# 2. 读取（先看分页预览定位目标页，再按页读，省 token）
python scripts/wpsdoc.py read  "C:/path/to/文件.doc" --page 2

# 3. 表格导出为 markdown
python scripts/wpsdoc.py tables "C:/path/to/文件.doc" --format md

# 4. 修改前一律先预演
python scripts/wpsdoc.py replace "C:/path/to/文件.doc" --old "旧文字" --new "新文字" --dry-run
```

## 命令一览

| 命令 | 作用 | 关键参数 |
|---|---|---|
| `probe` | 加密检测 + 页数/段落/表格/工作表 + 分页预览 | `--preview N` |
| `read` | 读 Word 全文或指定页 | `--page N`、`--out 文件` |
| `tables` | 导出 Word 表格 | `--format md\|csv`、`--out` |
| `replace` | 替换 Word 文字 | `--old`、`--new`、`--first`、`--dry-run` |
| `cell` | 写 Word 表格单元格 | `--table T --row R --col C --value V` |
| `sheet-read` | 读 Excel 工作表 | `--sheet 名`、`--max-rows N`、`--format` |
| `sheet-write` | 写 Excel 单元格 | `--sheet 名 --cell A1 --value V` |

`--out` 省略时结果打印到标准输出。

## 操作规范

- **改动前先 `probe` + `read`**，确认目标内容真实存在、页码正确。
- **写操作必须先 `--dry-run` 预演**，确认替换处数与预期一致，再去掉该参数执行。
- 写操作默认生成 `<原文件>.bak` 备份；确认无误后再决定是否删除备份。
- 批量替换时先小范围试一次，确认无误再放开。
- 修改完成后，用 `read` 或 `tables` **读回校验**，不要凭脚本返回值就下结论。
- 文档若正被人工在 WPS 中打开，脚本用 `DispatchEx` 开的是独立实例，一般不会冲突，但仍建议提醒对方先关闭。

## 环境与执行注意事项

- 依赖 `pywin32`。脚本会自动寻找带 pywin32 的解释器重跑；也可设环境变量 `WPSDOC_PYTHON` 指定。
- 若已知某解释器带 pywin32（如 miniconda），直接指定它可省去一次自动探测。
- **调用方式：直接执行脚本文件**。不要把 COM 代码用 Bash heredoc 内联执行——那样会被沙箱以 SIGTERM 终止且无任何输出。若前台执行仍被终止，改用后台方式运行并读取输出。
- 脚本保证 COM 实例 `Visible=False`，且在退出路径中 `Quit()`，不会残留 WPS 进程。

## MCP 版（可选）

同一能力另提供 MCP server（供 Cursor / Claude Code / Trae 等外部 MCP 客户端调用）：
`mcp/server.py`，启用方式与排错见 `mcp/README.md`。
两者共用同一份 `scripts/wpsdoc.py`，因此行为一致。

在 WorkBuddy 会话内直接处理文件用本 skill 的脚本；需要给外部客户端调用时用 MCP 版。

### 与 esafe-fs MCP 的选型规则（本机实证）

本机同时装了 `esafe-fs`（纯文本用途）和 `wps-office-io`（Office 用途）两个 MCP，分工如下：

| 目标文件 | 用哪个 |
|---|---|
| `.doc/.docx/.xls/.xlsx/.ppt/.pptx`（**无论是否加密**） | **一律用 `wps-office-io`** |
| `.c/.h/.md/.txt/.log/.csv/.json/.py` 等纯文本 | 用 `esafe-fs`（`read_file` / `read_file_lines` / `scan_directory`） |
| 只需要找文件、看大小行数 | 两者皆可，优先 `esafe-fs` 的 `list_directory` / `scan_directory` |

**为什么 Office 不能交给 esafe-fs**（2026-10-08 实测，同一份加密 .doc 对照）：

- `esafe-fs` 的 `get_file_info` 返回 `esafe_encrypted: false`——**假阴性**。根因是其源码 `ESAFE_MAGIC = b"E-SafeNetLOCK"` 为精确子串匹配，而真实文件头是 `E-SafeNet\x00\x00\x00LOCK`（中间 3 个空字节），永远匹配不上。
- 于是它跳过 decrypt 桥接直接 `open()` 读，`read_file_lines` 吐出的首行就是 `...E-SafeNet...LOCK...` 等明文标记夹乱码的**密文**。
- 其降级桥接 `cmd /c type` 在本机同样不可用：返回 rc=1、0 字节、stderr 为 GBK 的「命令语法不正确。」（已排除中文路径因素，纯 ASCII 路径同样失败）。
- 对照 `wps_probe`：正确报「加密：是 —— E-SafeNet」+ 页数/段落/表格数 + 分页预览；`wps_read --page 2` 返回干净明文。

一句话：**看到 Office 扩展名就无脑选 `wps-office-io`，别在 esafe-fs 上试错。** 反过来，纯文本也别用 WPS COM 去读——慢且没必要。

### MCP 标准作业流程（SOP）

1. **`wps_probe`** —— 确认加密状态、页数/表格数，用分页预览定位目标内容在第几页。
2. **`wps_read --page N`** —— 按页精读，避免整篇返回浪费上下文。
3. 要表格用 **`wps_tables`**；要改字先 **`wps_replace --dry-run`** 看命中处数，确认后才真跑，跑完 **`wps_read`** 读回校验。
4. Excel 用 **`wps_sheet_read` / `wps_sheet_write`**；Word 表格单元改写用 **`wps_cell`**。

## 详细参考

COM 常量、ProgID 回退顺序、分页与文本结构符原理、常见报错处理见 `references/com-api.md`。
