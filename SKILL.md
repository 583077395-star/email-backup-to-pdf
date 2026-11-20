---
name: email-backup-to-pdf
description: 把邮箱导出的 .eml 备份（含 messages_package 分卷 zip）批量转成 PDF，文件名按「邮件时间+主题」命名，附件集中归档并生成索引。当用户提到邮箱备份、邮件备份解压、.eml 转 PDF、邮件归档整理、按时间主题命名邮件、Outlook/企业邮箱导出包，或要求把一堆邮件整理成可检索的 PDF 时使用。
---

# 邮箱备份转 PDF

把一堆 `.eml`（通常在 `messages_package x_y.zip` 分卷里）转成「一封信一个 PDF」，附件单独归档，并产出 `索引.csv`。

## 产出结构

```
<out>/
├── PDF/YYYYMMDD_HHMM_主题.pdf
├── 附件/<与 PDF 同名的目录>/<原始附件文件>
└── 索引.csv   序号,邮件时间,主题,发件人,收件人,PDF文件,附件数,来源包,原件名
```

## 前置条件

- Windows / macOS / Linux 均可，需装 Microsoft Edge、Google Chrome 或 Chromium 之一（脚本自动探测，也可 `--edge` 指定）
- Python 3 + `pymupdf`（`python -m pip install pymupdf`，仅体检脚本需要）

## 用法

```bash
# 全量转换（首次）
python scripts/eml_to_pdf.py --zip-dir "E:\...\邮箱备份" --out "E:\...\邮箱备份_PDF"

# 增量：补转新增分卷/上次失败的邮件（重复跑同一条命令即可，按邮件粒度续跑）
python scripts/eml_to_pdf.py --zip-dir "E:\...\邮箱备份" --out "E:\...\邮箱备份_PDF"

# 先小批量验证版面（后续必须再跑一次全量，否则剩余邮件不会补上）
python scripts/eml_to_pdf.py --zip-dir "<dir>" --out "<tmp_out>" --limit 3

# 常用参数
--date-format "%Y%m%d_%H%M"           # 文件名时间格式，默认即此
--att-mode collection|per-mail|none   # 附件：集中目录 / 与 PDF 同目录 / 不导出
--fresh                               # 忽略已有索引，全部重转
--allow-remote                        # 允许加载邮件里的远程图片（默认剥离，见坑 #7）
```

体检：

```bash
python scripts/qa_pdfs.py --out "<out>"
```

## 必须遵守的坑（这些坑会直接毁掉产出）

1. **绝对不要用 Word COM 导出 PDF。** 装了终端加密/DLP 的企业电脑上，Word `SaveAs2`/`ExportAsFixedFormat` 生成的 ".pdf" 实际是加密密文（文件头不是 `%PDF-`，只有白名单程序能打开，PyMuPDF 报 `no objects found`）。统一走浏览器 headless 打印：`msedge --headless=new --print-to-pdf`，产物是标准可搜索文本 PDF。
2. **正文要先清洗再套模板。** 大量邮件（尤其 Outlook 转发/回复、转载的 Word 邮件）正文本身就是一整份带 `<html><head>`、`@page`、`WordSection1 {page:...}` 的完整文档。直接拼进模板会强制新起一页，表现为「第一页只有抬头、正文全在第二页」。`clean_body()` 已处理，改脚本时不要删。
3. **分卷必须先下载完整。** 同步盘（百度网盘等）下载中会给文件加独占锁且体积持续增长，此时解压得到残缺目录或 `BadZipFile`。脚本遇到打不开的分卷直接报错退出，不要绕过它"尽力解析"；等体积稳定、能列出条目后再转。
4. **zip 内中文文件名是 cp936。** Python 默认按 cp437 解码是乱码，用 `name.encode('cp437').decode('cp936')` 还原（脚本已内置）；起名一律以邮件头 `Subject` 为准，不依赖文件名。
5. **命名与去重**：`时间_主题`，`\ / : * ? " < |` 替换为 `_`，超长截断；重名加 `_2`、`_3`。判重必须按「(来源包, 原件名)」，**不能用 (邮件时间, 主题)**——同一天的同名主题（日报、系统通知、审批提醒）是不同邮件，按主题判重会把后一封静默丢掉。
6. **索引只追加不重写**，序号连续；只有 `--fresh` 才清空重建。失败邮件写进 `失败清单.csv`，修好重跑同一条命令即可补齐，已成功的不重做。
7. **默认剥离邮件里的远程图片**。headless 打印等于真把邮件打开一次，企业邮件中的跟踪像素会回连发件方服务器，把「已读」时间和 IP 泄露出去；`strip_remote()` 只保留 `cid:` / `data:` 内联图，确需保留外链时才加 `--allow-remote`。

## 质量抽检（交付前必做）

1. 跑 `qa_pdfs.py`：确认数量、打不开的 PDF、异常空首页、附件缺失。
2. 抽 2–3 个 PDF 渲染首页看图：

```python
import pymupdf
d = pymupdf.open(path)
d[0].get_pixmap(matrix=pymupdf.Matrix(1.5, 1.5)).save("p1.png")
```

看三件事：正文是否从第一页开始、内嵌图片是否还在、签名档与引用历史是否完整。首页文字很少但邮件本身就是短通知（验证码、审批提醒）属正常，不要误判为丢内容。

## 汇报口径

交付时给出：PDF 数量与总页数、附件目录/文件数与体积、时间跨度、分卷覆盖情况、失败/跳过数量；存在 `失败清单.csv` 时必须读出来逐条说明原因，不要只说「失败 N 封」。同时提示缺失分卷或需人工确认的短邮件。
