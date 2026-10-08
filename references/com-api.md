# WPS / Office COM 自动化参考

## ProgID 与回退顺序

脚本按下列顺序逐个尝试，第一个成功的即为所用引擎（优先 WPS，因为它是绿盾白名单进程）。

| 文档类型 | ProgID 尝试顺序 |
|---|---|
| Word（.doc/.docx/.wps） | `KWps.Application` → `WPS.Application` → `Word.Application` |
| Excel（.xls/.xlsx/.et） | `KET.Application` → `ET.Application` → `Excel.Application` |
| PPT（.ppt/.pptx/.dps） | `KWPP.Application` → `WPP.Application` → `PowerPoint.Application` |

实测可用的组合：`KWps.Application`（WPS 文字）、`KET.Application`（WPS 表格）。

用 `DispatchEx` 而非 `Dispatch`——前者开独立实例，不会劫持用户已打开的 WPS 窗口。

## 标准开闭模板

```python
import win32com.client as win32

app = win32.DispatchEx("KWps.Application")
app.Visible = False
try:
    app.DisplayAlerts = 0        # 抑制另存为/兼容性等弹窗
except Exception:
    pass

doc = None
try:
    doc = app.Documents.Open(path, ReadOnly=True)   # 读：True；写：False
    # ... 操作 ...
finally:
    if doc is not None:
        doc.Close(0)     # 0 = wdDoNotSaveChanges；已显式 Save() 过就用 0
    app.Quit()
```

务必在 `finally` 中 Quit，否则残留 WPS 进程会锁住文件，后续 `Rename-Item` / 删除都会「拒绝访问」。

## 常用常量

| 常量 | 值 | 用途 |
|---|---|---|
| `wdStatisticPages` | 2 | `doc.ComputeStatistics(2)` 取总页数 |
| `wdGoToPage` | 1 | `doc.GoTo(What, Which, Count)` 的 What 参数 |
| `wdGoToAbsolute` | 1 | 同上 Which 参数 |
| `wdDoNotSaveChanges` | 0 | `doc.Close(0)` |
| `wdSaveChanges` | -1 | `doc.Close(-1)` |
| `wdFormatDocument` | 0 | `SaveAs(path, 0)` 存为 .doc |

## 分页读取原理

Word 的 COM 没有直接的「取第 N 页文本」接口，标准做法是用页起点做区间：

```python
start = doc.GoTo(1, 1, n).Start          # 第 n 页起点
try:
    end = doc.GoTo(1, 1, n + 1).Start    # 第 n+1 页起点
except Exception:
    end = doc.Content.End                # 末页：取到文末
text = doc.Range(start, end).Text
```

调用前先 `doc.Repaginate()` 触发重新分页，页数才准确。

## 文本结构符

COM 取回的文本含 Word 内部符号，展示前必须清洗：

| 符号 | 含义 | 建议替换 |
|---|---|---|
| `\x07` | 表格单元格结束符 | 单行展示替换为 `" \| "`；还原表格时作为**列分隔符** |
| `\r` (CR) | 单元格/行结束符 | 替换为 `\n`；还原表格时作为**行分隔符** |
| `\x0b` `\x0c` | 手动换行/分页符 | 替换为 `\n` |

单元格文本清洗：`raw.replace("\r", "").replace("\x07", "").strip()`。

## 写入要点

- **替换文字（两个已踩过的坑，务必遵守）**：
  1. **必须用 Word 的 `Find` 对象，不能用「`Content.Text` 下标 + `Range(pos,...)`」**。
     含表格的文档里，单元格结束符在 `Text` 中占 2 个字符（`\r\x07`）、在 Range 位置中只占 1 个，
     下标会随表格数量累积漂移，替换会落到**错误位置**、静默写坏文档。
  2. **`Find.Execute` 必须用位置参数，不能用关键字参数**。
     WPS 会静默忽略 `Execute(Replace=2)` 这类关键字参数——不报错、Execute 返回 True，
     但只做查找不做替换。实测有效写法：

     ```python
     f = doc.Content.Find
     f.ClearFormatting(); f.Replacement.ClearFormatting()
     f.Text = old; f.Replacement.Text = new
     f.Forward = True; f.Wrap = 1; f.Format = False
     f.MatchCase = f.MatchWholeWord = f.MatchWildcards = False
     f.MatchSoundsLike = f.MatchAllWordForms = False
     f.Execute(old, False, False, False, False, False, True, 1, False, new, 2)
     # 位置: FindText, MatchCase, MatchWholeWord, MatchWildcards, MatchSoundsLike,
     #       MatchAllWordForms, Forward, Wrap, Format, ReplaceWith, Replace
     # Replace: 2=全部替换  1=替换一处
     ```
  3. **不要信任 `Execute` 的返回值判断成败**。替换后比对 `doc.Content.Text` 是否真的变化，
     没变就放弃保存并报错。（`new` 包含 `old` 时，用"是否仍含 old"判断会误判，只能比对全文。）
- **表格单元格**：`doc.Tables(t).Cell(r, c).Range.Text = value`，索引从 1 开始。
- **Excel 单元格**：`ws.Range("A1").Value = value`；读取用 `ws.UsedRange.Value`（返回元组套元组；单格返回标量；空表返回 `None`）。

## 加密状态判断

读前 64 字节即可判断，无需打开文档：

- 正常 Office 文档是 OLE2 复合文档，魔数必为 `D0 CF 11 E0 A1 B1 1A E1`
- 绿盾加密后魔数消失，头部出现 ASCII 标记 `E-SafeNet` 与 `LOCK`
- 可辅助验证：全文搜不到 `WordDocument` / `Root Entry` 等 OLE2 流签名；正文熵接近 8

注意：同一目录下不同文件加密状态可能不同（例如 .doc 被加密而 .xls 未加密），
必须逐个 `probe`，不能按目录推断。

## 常见报错与处理

| 现象 | 原因 | 处理 |
|---|---|---|
| `No module named 'win32com'` | 当前解释器缺 pywin32 | 脚本会自动切换解释器；或设 `WPSDOC_PYTHON` |
| 被 SIGTERM 终止、无任何输出 | COM 代码用 Bash heredoc 内联执行 | 改为执行 `.py` 脚本文件；仍失败则后台运行 |
| `无法创建 COM 实例` | WPS 未安装或 ProgID 不匹配 | 确认 WPS 已装；检查 ProgID 回退列表 |
| 文件被占用 / 保存失败 | 文档正被人工打开，或上次运行残留进程 | 关闭 WPS；确认脚本已 Quit |
| 页数不准 | 未重新分页 | 调用 `doc.Repaginate()` |
| `UsedRange.Address` 失败 | 工作表受保护或为空 | 已做兜底显示 `-`，不影响 `sheet-read` 取数 |

## 环境依赖

`pywin32` 是唯一依赖。若托管 Python 未安装，优先使用系统里已带 pywin32 的解释器
（例如 miniconda）。不要为了这个 skill 污染全局环境。
