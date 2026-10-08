#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
wpsdoc.py —— 借助 WPS / MS Office 的 COM 自动化读写 Office 文档。

为什么需要它
------------
在装有天锐绿盾(E-SafeNet / E-Safe)透明加密的 Windows 机器上，Office 文档在磁盘上
是密文（头部带 "E-SafeNet" + "LOCK" 标记，无 OLE2 魔数）。python-docx / openpyxl /
olefile 等库直接读字节只能拿到密文。WPS / Word 属于绿盾白名单进程，打开时会被
自动解密，因此通过 COM 驱动 WPS 是唯一既能读到明文、又能安全写回的途径
（写回时绿盾会自动再加密，不会破坏加密状态）。

子命令
------
  probe       检测加密状态 + 文档结构（页数 / 表格数 / 工作表 / 分页预览）
  read        读取 Word 全文或指定页码
  tables      导出 Word 中的表格（markdown / csv）
  replace     替换 Word 文字（支持 --dry-run 预演、自动备份）
  cell        写入 Word 表格单元格（自动备份）
  sheet-read  读取 Excel 工作表
  sheet-write 写入 Excel 单元格（自动备份）

安全约定
--------
  * 所有写操作默认先生成 <原文件>.bak 备份，--no-backup 可关闭
  * COM 实例一律 Visible=False / DisplayAlerts=0，并在 finally 中 Quit，不留残留进程
  * 读操作一律 ReadOnly=True，绝不落盘

依赖
----
  需要 pywin32。若当前解释器没有，脚本会自动寻找带 pywin32 的 Python 重新执行；
  也可用环境变量 WPSDOC_PYTHON 显式指定。
"""

import argparse
import io
import os
import shutil
import subprocess
import sys

# ---------------------------------------------------------------- 基础工具


def _force_utf8():
    try:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")
    except Exception:
        pass


_force_utf8()


def _log(msg):
    print(msg, flush=True)


def _die(msg, code=2):
    print("[ERROR] %s" % msg, file=sys.stderr, flush=True)
    sys.exit(code)


def _emit(text, out_path=None):
    if out_path:
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(text)
        _log("[ok] 已写入 -> %s" % out_path)
    else:
        _log(text)


# ------------------------------------------------- pywin32 解释器自动切换


def _python_candidates():
    env = os.environ.get("WPSDOC_PYTHON")
    cands = [env] if env else []
    home = os.path.expanduser("~")
    for base in (home, r"C:\ProgramData"):
        for name in ("miniconda3", "anaconda3", "Miniconda3", "Anaconda3"):
            cands.append(os.path.join(base, name, "python.exe"))
    return [c for c in cands if c]


def _has_win32com(pyexe):
    try:
        r = subprocess.run(
            [pyexe, "-c", "import win32com.client"],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=60)
        return r.returncode == 0
    except Exception:
        return False


def ensure_win32com():
    """若当前解释器缺少 pywin32，自动换用带 pywin32 的解释器重跑本脚本。"""
    try:
        import win32com.client  # noqa: F401
        return
    except ImportError:
        pass
    if os.environ.get("WPSDOC_REEXEC") == "1":
        _die("仍未找到带 pywin32 的解释器。请 pip install pywin32，"
             "或设置环境变量 WPSDOC_PYTHON 指向可用的 python.exe")
    for py in _python_candidates():
        if not os.path.exists(py):
            continue
        if os.path.abspath(py) == os.path.abspath(sys.executable):
            continue
        if _has_win32com(py):
            _log("[info] 当前解释器无 pywin32，改用 %s 重新执行" % py)
            env = dict(os.environ, WPSDOC_REEXEC="1")
            sys.exit(subprocess.call([py] + sys.argv, env=env))
    _die("未找到带 pywin32 的 Python。请 pip install pywin32，或设置 WPSDOC_PYTHON。")


# ------------------------------------------------------------ COM 应用


PROGIDS = {
    # 优先 WPS 官方 ProgID（绿盾白名单），回退到 MS Office
    "word":  ["KWps.Application", "WPS.Application", "Word.Application"],
    "excel": ["KET.Application", "ET.Application", "Excel.Application"],
    "ppt":   ["KWPP.Application", "WPP.Application", "PowerPoint.Application"],
}

EXT_KIND = {
    ".doc": "word", ".docx": "word", ".wps": "word", ".dot": "word", ".dotx": "word",
    ".xls": "excel", ".xlsx": "excel", ".et": "excel", ".xlsm": "excel", ".xlsb": "excel",
    ".ppt": "ppt", ".pptx": "ppt", ".dps": "ppt",
}


def kind_of(path):
    ext = os.path.splitext(path)[1].lower()
    if ext not in EXT_KIND:
        _die("不支持的扩展名: %s（支持: %s）" % (ext, " ".join(sorted(EXT_KIND))))
    return EXT_KIND[ext]


def new_app(kind):
    import win32com.client as win32
    last = None
    for prog in PROGIDS[kind]:
        try:
            app = win32.DispatchEx(prog)   # DispatchEx = 新开独立实例，避免复用已有窗口
            app.Visible = False
            try:
                app.DisplayAlerts = 0
            except Exception:
                pass
            return app, prog
        except Exception as e:
            last = e
    _die("无法创建 %s 的 COM 实例（已尝试 %s）。最后错误: %s" % (kind, PROGIDS[kind], last))


def close_all(app, *docs):
    for d in docs:
        try:
            if d is not None:
                d.Close(0)
        except Exception:
            pass
    try:
        app.Quit()
    except Exception:
        pass


# ---------------------------------------------------------- 加密检测


OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


def detect_encryption(path):
    """读前 64 字节判断封装类型。OLE2 魔数存在=正常文档；否则看是否有绿盾标记。"""
    with open(path, "rb") as f:
        head = f.read(64)
    is_ole = head[:8] == OLE_MAGIC
    esafe = (b"E-SafeNet" in head) or (b"LOCK" in head)
    info = {"head_hex": head[:32].hex()}
    if esafe and not is_ole:
        info.update(encrypted=True, engine="E-SafeNet(天锐绿盾)")
    elif is_ole:
        info.update(encrypted=False, engine=None)
    else:
        info.update(encrypted=None, engine="未知封装（可能是其他加密或已损坏）")
    return info


# ------------------------------------------------------------ 文本处理


def clean_text(t):
    """把 Word 取回的原始文本整理成可读形式：\x07=单元格分隔，\r=单元格/段落结束。"""
    t = t.replace("\x07", " | ")
    t = t.replace("\r", "\n")
    t = t.replace("\x0b", "\n").replace("\x0c", "\n")
    while "\n\n\n" in t:
        t = t.replace("\n\n\n", "\n\n")
    return t.strip()


def cell_text(raw):
    """表格单元格文本：去掉结尾的 \r\x07 结构符。"""
    return raw.replace("\r", "").replace("\x07", "").strip()


# ------------------------------------------------------------ Word 助手


def word_page_count(doc):
    try:
        doc.Repaginate()
    except Exception:
        pass
    try:
        return int(doc.ComputeStatistics(2))   # 2 = wdStatisticPages
    except Exception:
        return None


def word_page_text(doc, n):
    """取第 n 页文本：第 n 页起点 -> 第 n+1 页起点（末页取到文末）。"""
    start = doc.GoTo(1, 1, n).Start
    try:
        end = doc.GoTo(1, 1, n + 1).Start
    except Exception:
        end = doc.Content.End
    return doc.Range(start, end).Text


def word_page_map(doc, n_pages, preview=50):
    """生成每页起始预览，便于定位内容在哪一页。"""
    lines = []
    for i in range(1, (n_pages or 0) + 1):
        try:
            t = clean_text(word_page_text(doc, i)).replace("\n", " ")
            lines.append("  p%-3d %s" % (i, t[:preview]))
        except Exception as e:
            lines.append("  p%-3d <读取失败: %s>" % (i, e))
    return "\n".join(lines)


def backup_file(path, enabled=True):
    if not enabled:
        return None
    bak = path + ".bak"
    shutil.copy2(path, bak)
    return bak


# ------------------------------------------------------------ 子命令


def cmd_probe(args):
    path = os.path.abspath(args.file)
    if not os.path.exists(path):
        _die("文件不存在: %s" % path)
    kind = kind_of(path)
    enc = detect_encryption(path)
    _log("文件      : %s" % path)
    _log("大小      : %d 字节" % os.path.getsize(path))
    _log("类型      : %s" % kind)
    _log("加密      : %s" % ("是 —— %s" % enc["engine"] if enc["encrypted"]
                             else ("否" if enc["encrypted"] is False else enc["engine"])))
    _log("文件头    : %s" % enc["head_hex"])

    ensure_win32com()
    app, prog = new_app(kind)
    _log("COM ProgID: %s" % prog)
    doc = None
    try:
        if kind == "word":
            doc = app.Documents.Open(path, ReadOnly=True)
            n_pages = word_page_count(doc)
            _log("页数      : %s" % n_pages)
            _log("段落数    : %s" % doc.Paragraphs.Count)
            _log("表格数    : %s" % doc.Tables.Count)
            if n_pages:
                _log("分页预览  :")
                _log(word_page_map(doc, n_pages, args.preview))
        elif kind == "excel":
            doc = app.Workbooks.Open(path, ReadOnly=True)
            _log("工作表    : %s" % doc.Worksheets.Count)
            for i in range(1, doc.Worksheets.Count + 1):
                ws = doc.Worksheets(i)
                try:
                    addr = ws.UsedRange.Address(False, False)
                except Exception:
                    addr = "-"
                _log("  [%d] %s  已用区域=%s" % (i, ws.Name, addr))
        else:
            _log("PPT 仅支持有限探测，未做结构展开。")
    finally:
        close_all(app, doc)


def cmd_read(args):
    path = os.path.abspath(args.file)
    kind = kind_of(path)
    if kind != "word":
        _die("read 仅支持 Word 文档；Excel 请用 sheet-read")
    ensure_win32com()
    app, prog = new_app(kind)
    doc = None
    try:
        doc = app.Documents.Open(path, ReadOnly=True)
        n_pages = word_page_count(doc)
        if args.page:
            if n_pages and args.page > n_pages:
                _die("指定页 %d 超出范围（共 %d 页）" % (args.page, n_pages))
            chunks = ["### 共 %s 页，以下为第 %d 页\n" % (n_pages, args.page),
                      clean_text(word_page_text(doc, args.page))]
        else:
            parts = ["### 共 %s 页，以下为全文\n" % n_pages]
            for i in range(1, (n_pages or 1) + 1):
                parts.append("\n----- 第 %d 页 -----\n" % i)
                parts.append(clean_text(word_page_text(doc, i)))
            chunks = parts
        _emit("\n".join(chunks), args.out)
    finally:
        close_all(app, doc)


def cmd_tables(args):
    path = os.path.abspath(args.file)
    kind = kind_of(path)
    if kind != "word":
        _die("tables 仅支持 Word 文档")
    ensure_win32com()
    app, prog = new_app(kind)
    doc = None
    try:
        doc = app.Documents.Open(path, ReadOnly=True)
        count = doc.Tables.Count
        out = ["### 表格总数: %d" % count]
        for t in range(1, count + 1):
            tbl = doc.Tables(t)
            rows, cols = tbl.Rows.Count, tbl.Columns.Count
            out.append("\n----- 表 %d (%d行 x %d列) -----" % (t, rows, cols))
            grid = []
            for r in range(1, rows + 1):
                grid.append([cell_text(tbl.Cell(r, c).Range.Text) for c in range(1, cols + 1)])
            if args.format == "csv":
                for row in grid:
                    out.append(",".join('"%s"' % v.replace('"', '""') for v in row))
            else:
                out.append("| " + " | ".join(grid[0]) + " |")
                out.append("|" + "|".join(["---"] * len(grid[0])) + "|")
                for row in grid[1:]:
                    out.append("| " + " | ".join(row) + " |")
        _emit("\n".join(out), args.out)
    finally:
        close_all(app, doc)


WD_FIND_STOP = 0      # 搜到文末即停
WD_FIND_CONTINUE = 1  # 搜到文末从头继续
WD_REPLACE_ONE = 1
WD_REPLACE_ALL = 2


def _configure_find(rng, text, replace_text=None):
    """配置 Find 对象。

    务必使用 Word 自带的 Find 做替换：它内部自行计算插入点。
    不要改用「Content.Text 下标 + Range(pos, ...)」——文档含表格时，
    单元格结束符在 Text 中占 2 个字符(\\r\\x07)、在 Range 位置中只占 1 个，
    下标会随表格数量累积漂移，导致替换落到错误位置、静默写坏文档。
    """
    f = rng.Find
    f.ClearFormatting()
    try:
        f.Replacement.ClearFormatting()
    except Exception:
        pass
    f.Text = text
    if replace_text is not None:
        f.Replacement.Text = replace_text
    f.Forward = True
    f.Wrap = WD_FIND_STOP
    f.Format = False
    f.MatchCase = False
    f.MatchWholeWord = False
    f.MatchWildcards = False
    f.MatchSoundsLike = False
    f.MatchAllWordForms = False
    return f


def count_occurrences(doc, text):
    if not text:
        return 0
    f = _configure_find(doc.Content, text)
    n = 0
    while n < 10000 and f.Execute():
        n += 1
    return n


def do_replace(doc, old, new, replace_all=True):
    # Wrap 一律用「不回绕」：若用回绕(wdFindContinue)，当 new 包含 old 时
    # （如 old='A', new='A(B)'），ReplaceAll 绕回文首会再次命中刚插入的文字，
    # 把替换内容重复叠加，实测出现 'A(B)(B)'。不回绕即只在单趟前向扫描中替换。
    wrap = WD_FIND_STOP
    mode = WD_REPLACE_ALL if replace_all else WD_REPLACE_ONE
    f = _configure_find(doc.Content, old, new)
    f.Wrap = wrap
    # 必须用位置参数：WPS 会静默忽略 Execute 的关键字参数（不报错，但只查找不替换）。
    # 参数顺序: FindText, MatchCase, MatchWholeWord, MatchWildcards, MatchSoundsLike,
    #           MatchAllWordForms, Forward, Wrap, Format, ReplaceWith, Replace
    return bool(f.Execute(old, False, False, False, False, False, True, wrap, False, new, mode))


def cmd_replace(args):
    path = os.path.abspath(args.file)
    kind = kind_of(path)
    if kind != "word":
        _die("replace 仅支持 Word 文档")
    if args.old == args.new:
        _die("--old 与 --new 相同，无需替换")
    if args.old in args.new:
        _log("[提示] --new 中包含 --old，已启用不回绕模式，避免替换内容被重复叠加")
    ensure_win32com()
    app, prog = new_app(kind)
    doc = None
    try:
        doc = app.Documents.Open(path, ReadOnly=bool(args.dry_run))
        total = count_occurrences(doc, args.old)
        if args.dry_run:
            _log("[预演] 将替换 %d 处：'%s' -> '%s'（未保存）" % (total, args.old, args.new))
        elif total == 0:
            _log("[跳过] 未找到 '%s'，文件未改动" % args.old)
        else:
            # 不信任 Execute 的返回值：替换后比对全文，确认真的变了再保存。
            before = doc.Content.Text
            do_replace(doc, args.old, args.new, not args.first)
            if doc.Content.Text == before:
                _log("[失败] 替换未生效（文档内容无变化），已放弃保存")
            else:
                bak = backup_file(path, not args.no_backup)
                doc.Save()
                _log("[完成] 替换 %d 处：'%s' -> '%s'" % (total, args.old, args.new))
                if bak:
                    _log("        备份: %s" % bak)
    finally:
        close_all(app, doc)


def cmd_cell(args):
    path = os.path.abspath(args.file)
    kind = kind_of(path)
    if kind != "word":
        _die("cell 仅支持 Word 文档")
    ensure_win32com()
    app, prog = new_app(kind)
    doc = None
    try:
        doc = app.Documents.Open(path, ReadOnly=False)
        cnt = doc.Tables.Count
        if args.table < 1 or args.table > cnt:
            _die("表序号 %d 越界（共 %d 个表）" % (args.table, cnt))
        tbl = doc.Tables(args.table)
        if args.row < 1 or args.row > tbl.Rows.Count:
            _die("行号越界（共 %d 行）" % tbl.Rows.Count)
        if args.col < 1 or args.col > tbl.Columns.Count:
            _die("列号越界（共 %d 列）" % tbl.Columns.Count)
        old = cell_text(tbl.Cell(args.row, args.col).Range.Text)
        if args.dry_run:
            _log("[预演] 表%d (%d,%d): '%s' -> '%s'（未保存）"
                 % (args.table, args.row, args.col, old, args.value))
        else:
            bak = backup_file(path, not args.no_backup)
            tbl.Cell(args.row, args.col).Range.Text = args.value
            doc.Save()
            _log("[完成] 表%d (%d,%d): '%s' -> '%s'" % (args.table, args.row, args.col, old, args.value))
            if bak:
                _log("        备份: %s" % bak)
    finally:
        close_all(app, doc)


def _sheet_of(wb, name):
    if name:
        try:
            return wb.Worksheets(name)
        except Exception:
            _die("找不到工作表: %s" % name)
    return wb.Worksheets(1)


def _to_grid(val):
    if val is None:
        return []
    if not isinstance(val, tuple):
        return [[val]]
    if len(val) > 0 and not isinstance(val[0], tuple):
        return [list(val)]
    return [list(row) for row in val]


def _fmt(grid, fmt):
    out = []
    if not grid:
        return "(空表)"
    if fmt == "csv":
        for row in grid:
            out.append(",".join('"%s"' % ("" if v is None else str(v)).replace('"', '""') for v in row))
    else:
        head = ["" if v is None else str(v) for v in grid[0]]
        out.append("| " + " | ".join(head) + " |")
        out.append("|" + "|".join(["---"] * len(head)) + "|")
        for row in grid[1:]:
            out.append("| " + " | ".join("" if v is None else str(v) for v in row) + " |")
    return "\n".join(out)


def cmd_sheet_read(args):
    path = os.path.abspath(args.file)
    kind = kind_of(path)
    if kind != "excel":
        _die("sheet-read 仅支持 Excel 文件")
    ensure_win32com()
    app, prog = new_app(kind)
    wb = None
    try:
        wb = app.Workbooks.Open(path, ReadOnly=True)
        ws = _sheet_of(wb, args.sheet)
        grid = _to_grid(ws.UsedRange.Value)
        if args.max_rows:
            grid = grid[:args.max_rows]
        header = "### 工作表: %s（%d 行）\n" % (ws.Name, len(grid))
        _emit(header + _fmt(grid, args.format), args.out)
    finally:
        close_all(app, wb)


def cmd_sheet_write(args):
    path = os.path.abspath(args.file)
    kind = kind_of(path)
    if kind != "excel":
        _die("sheet-write 仅支持 Excel 文件")
    ensure_win32com()
    app, prog = new_app(kind)
    wb = None
    try:
        wb = app.Workbooks.Open(path, ReadOnly=bool(args.dry_run))
        ws = _sheet_of(wb, args.sheet)
        rng = ws.Range(args.cell)
        old = rng.Value
        if args.dry_run:
            _log("[预演] %s!%s: %r -> %r（未保存）" % (ws.Name, args.cell, old, args.value))
        else:
            bak = backup_file(path, not args.no_backup)
            rng.Value = args.value
            wb.Save()
            _log("[完成] %s!%s: %r -> %r" % (ws.Name, args.cell, old, args.value))
            if bak:
                _log("        备份: %s" % bak)
    finally:
        close_all(app, wb)


# ------------------------------------------------------------ CLI


def build_parser():
    p = argparse.ArgumentParser(
        prog="wpsdoc.py",
        description="借助 WPS/Office COM 读写 Office 文档（支持天锐绿盾透明加密文件）")
    sub = p.add_subparsers(dest="cmd", required=True)

    def add_backup(sp):
        sp.add_argument("--no-backup", action="store_true", help="写前不生成 .bak 备份（默认会备份）")

    sp = sub.add_parser("probe", help="检测加密状态与文档结构")
    sp.add_argument("file")
    sp.add_argument("--preview", type=int, default=50, help="分页预览字符数")
    sp.set_defaults(func=cmd_probe)

    sp = sub.add_parser("read", help="读取 Word 全文或指定页")
    sp.add_argument("file")
    sp.add_argument("--page", type=int, help="只读取指定页码")
    sp.add_argument("--out", help="输出到文件（默认打印到标准输出）")
    sp.set_defaults(func=cmd_read)

    sp = sub.add_parser("tables", help="导出 Word 表格")
    sp.add_argument("file")
    sp.add_argument("--format", choices=["md", "csv"], default="md")
    sp.add_argument("--out")
    sp.set_defaults(func=cmd_tables)

    sp = sub.add_parser("replace", help="替换 Word 文字")
    sp.add_argument("file")
    sp.add_argument("--old", required=True)
    sp.add_argument("--new", required=True)
    sp.add_argument("--first", action="store_true", help="只替换第一处（默认替换全部）")
    sp.add_argument("--dry-run", action="store_true", help="预演，不保存")
    add_backup(sp)
    sp.set_defaults(func=cmd_replace)

    sp = sub.add_parser("cell", help="写入 Word 表格单元格")
    sp.add_argument("file")
    sp.add_argument("--table", type=int, required=True, help="表格序号，从 1 开始")
    sp.add_argument("--row", type=int, required=True)
    sp.add_argument("--col", type=int, required=True)
    sp.add_argument("--value", required=True)
    sp.add_argument("--dry-run", action="store_true")
    add_backup(sp)
    sp.set_defaults(func=cmd_cell)

    sp = sub.add_parser("sheet-read", help="读取 Excel 工作表")
    sp.add_argument("file")
    sp.add_argument("--sheet", help="工作表名，省略则取第一个")
    sp.add_argument("--format", choices=["md", "csv"], default="md")
    sp.add_argument("--max-rows", type=int)
    sp.add_argument("--out")
    sp.set_defaults(func=cmd_sheet_read)

    sp = sub.add_parser("sheet-write", help="写入 Excel 单元格")
    sp.add_argument("file")
    sp.add_argument("--sheet")
    sp.add_argument("--cell", required=True, help="单元格地址，如 A1")
    sp.add_argument("--value", required=True)
    sp.add_argument("--dry-run", action="store_true")
    add_backup(sp)
    sp.set_defaults(func=cmd_sheet_write)

    return p


def main():
    args = build_parser().parse_args()
    if not os.path.exists(os.path.abspath(args.file)):
        _die("文件不存在: %s" % args.file)
    args.func(args)


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        raise
    except Exception as e:
        print("[ERROR] %s: %s" % (type(e).__name__, e), file=sys.stderr, flush=True)
        sys.exit(1)
