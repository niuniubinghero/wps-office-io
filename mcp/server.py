#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
wps-office-io —— MCP Server（stdio 传输）

把「借助 WPS/Office COM 读写 Office 文档」的能力暴露为 MCP 工具，
用于读取/修改被天锐绿盾(E-SafeNet)透明加密的文档。

设计要点
--------
* 本进程**只做协议层**：参数校验 + 调度子进程。真正的 COM 操作全部交给
  `scripts/wpsdoc.py`，从而复用已验证的逻辑（加密检测、ProgID 回退、
  Find 位置参数修正等），本进程因此无需安装 pywin32。
* **绝不向 stdout 打印任何东西**：MCP stdio 用 stdout 承载 JSON-RPC，
  一旦被污染协议即崩。调试信息一律走 stderr。
* 写操作默认由 wpsdoc.py 生成 .bak 备份，并建议先以 dry_run=true 预演。

环境变量
--------
  WPSDOC_PYTHON  指定执行 wpsdoc.py 的解释器（会自动切换到带 pywin32 的那个）
  WPSDOC_TIMEOUT 单次调用超时秒数，默认 180
"""

import os
import subprocess
import sys

from mcp.server.fastmcp import FastMCP

HERE = os.path.dirname(os.path.abspath(__file__))
SKILL_DIR = os.path.dirname(HERE)
WPSDOC = os.path.join(SKILL_DIR, "scripts", "wpsdoc.py")

PY = os.environ.get("WPSDOC_PYTHON") or sys.executable
TIMEOUT = int(os.environ.get("WPSDOC_TIMEOUT", "180"))

mcp = FastMCP("wps-office-io")


def _run(args):
    """调用 wpsdoc.py 并返回其输出文本。"""
    if not os.path.exists(WPSDOC):
        return "[错误] 找不到 wpsdoc.py: %s" % WPSDOC
    cmd = [PY, WPSDOC] + [str(a) for a in args]
    # 注意：必须把子进程的 stdin 与 MCP 的 stdio 管道隔离。
    # 若让 WPS/COM 子进程继承本进程的 stdin 管道，实测会挂起直到超时。
    print("[wps-office-io] exec: %s" % " ".join(cmd), file=sys.stderr, flush=True)
    try:
        p = subprocess.run(
            cmd, capture_output=True, text=True, stdin=subprocess.DEVNULL,
            encoding="utf-8", errors="replace", timeout=TIMEOUT)
    except subprocess.TimeoutExpired:
        return "[错误] 调用超时（%s 秒）" % TIMEOUT
    except Exception as e:
        return "[错误] 调用失败 %s: %s" % (type(e).__name__, e)
    out = (p.stdout or "").strip()
    err = (p.stderr or "").strip()
    if p.returncode != 0:
        return "[失败 rc=%d]\n%s\n%s" % (p.returncode, out, err)
    return out or err or "(无输出)"


# ------------------------------------------------------------------ 只读工具


@mcp.tool()
def wps_probe(file: str, preview: int = 50) -> str:
    """探测 Office 文档：是否被加密、页数、段落数、表格数、工作表清单，以及分页预览。

    读写任何 Office 文档之前都应先调用它，用来确认加密状态，并通过「分页预览」
    定位目标内容在第几页，避免整篇读取浪费上下文。

    Args:
        file: 文档绝对路径（.doc/.docx/.xls/.xlsx 等）
        preview: 分页预览截取的字符数，默认 50
    """
    return _run(["probe", file, "--preview", str(preview)])


@mcp.tool()
def wps_read(file: str, page: int = None, out: str = None) -> str:
    """读取 Word 文档全文或指定页的正文。

    先用 wps_probe 确认页码，再用 page 精读，可显著减少返回内容。

    Args:
        file: Word 文档路径
        page: 页码（从 1 起）；省略则读取全文
        out: 可选，把结果写入该文件的路径；省略则直接返回文本
    """
    args = ["read", file]
    if page is not None:
        args += ["--page", str(page)]
    if out:
        args += ["--out", out]
    return _run(args)


@mcp.tool()
def wps_tables(file: str, format: str = "md", out: str = None) -> str:
    """把 Word 文档中的所有表格导出为 markdown 或 csv。

    表单类文档（如 ECN/评审记录）用它能一次拿到规整的表格数据。

    Args:
        file: Word 文档路径
        format: md 或 csv，默认 md
        out: 可选输出文件路径
    """
    args = ["tables", file, "--format", format]
    if out:
        args += ["--out", out]
    return _run(args)


@mcp.tool()
def wps_sheet_read(file: str, sheet: str = None, format: str = "md",
                   max_rows: int = None, out: str = None) -> str:
    """读取 Excel 工作表内容。

    Args:
        file: Excel 文件路径
        sheet: 工作表名；省略则取第一个
        format: md 或 csv，默认 md
        max_rows: 最多返回行数
        out: 可选输出文件路径
    """
    args = ["sheet-read", file, "--format", format]
    if sheet:
        args += ["--sheet", sheet]
    if max_rows is not None:
        args += ["--max-rows", str(max_rows)]
    if out:
        args += ["--out", out]
    return _run(args)


# ------------------------------------------------------------------ 写入工具


@mcp.tool()
def wps_replace(file: str, old: str, new: str, first: bool = False,
                dry_run: bool = False, no_backup: bool = False) -> str:
    """替换 Word 文档中的文字。

    **务必先以 dry_run=true 预演**，确认替换处数符合预期后再实际执行。
    实际执行时默认生成 <原文件>.bak 备份。

    Args:
        file: Word 文档路径
        old: 待替换原文
        new: 替换为的新文字
        first: true 则只替换第一处，默认替换全部
        dry_run: true 为预演，不保存
        no_backup: true 则不生成 .bak 备份
    """
    args = ["replace", file, "--old", old, "--new", new]
    if first:
        args.append("--first")
    if dry_run:
        args.append("--dry-run")
    if no_backup:
        args.append("--no-backup")
    return _run(args)


@mcp.tool()
def wps_cell(file: str, table: int, row: int, col: int, value: str,
             dry_run: bool = False, no_backup: bool = False) -> str:
    """写入 Word 文档中指定表格的单元格（索引从 1 开始）。

    适合填写表单里空缺的日期、签署等单元格。建议先 dry_run 预演。

    Args:
        file: Word 文档路径
        table: 表格序号（从 1 起，可用 wps_probe 查看表格数）
        row: 行号（从 1 起）
        col: 列号（从 1 起）
        value: 要写入的内容
        dry_run: true 为预演，不保存
        no_backup: true 则不生成 .bak 备份
    """
    args = ["cell", file, "--table", str(table), "--row", str(row),
            "--col", str(col), "--value", value]
    if dry_run:
        args.append("--dry-run")
    if no_backup:
        args.append("--no-backup")
    return _run(args)


@mcp.tool()
def wps_sheet_write(file: str, cell: str, value: str, sheet: str = None,
                    dry_run: bool = False, no_backup: bool = False) -> str:
    """写入 Excel 指定单元格（如 A1）。建议先 dry_run 预演。

    Args:
        file: Excel 文件路径
        cell: 单元格地址，如 A1
        value: 要写入的值
        sheet: 工作表名；省略则取第一个
        dry_run: true 为预演，不保存
        no_backup: true 则不生成 .bak 备份
    """
    args = ["sheet-write", file, "--cell", cell, "--value", value]
    if sheet:
        args += ["--sheet", sheet]
    if dry_run:
        args.append("--dry-run")
    if no_backup:
        args.append("--no-backup")
    return _run(args)


if __name__ == "__main__":
    mcp.run()
