# email-backup-to-pdf

把邮箱导出的 `.eml` 备份批量转成 PDF：一封信一个文件，命名「邮件时间+主题」，附件集中归档，并生成可检索的 `索引.csv`。

An agent skill that turns a mailbox export (`.eml` files, often split into `messages_package x_y.zip` volumes) into one searchable PDF per email, named `YYYYMMDD_HHMM_<subject>.pdf`, with attachments archived and an index CSV produced.

## 它能做什么

- 读取 `messages_package x_y.zip` 分卷，自动还原 zip 内 cp936 中文名
- 每封邮件 → 一个 A4 PDF：抬头（发件人/收件人/抄送/时间）+ 正文 + 内嵌图片 + 附件清单
- 附件按「与 PDF 同名」的子目录集中归档，可选与 PDF 同目录或不导出
- 生成 `索引.csv`：序号、邮件时间、主题、发件人、收件人、PDF 文件、附件数、来源包、原件名
- **增量**：重复执行同一命令即可补转新增分卷与上次失败的邮件。判重按「(来源包, 原件名)」，中途失败、`--limit` 只跑一部分都能续上，已成功的不重做
- **失败可见**：渲染失败/异常的邮件逐条写入 `失败清单.csv`（来源包、原件名、原因），修好后重跑即补齐
- **隐私保护**：默认剥离邮件中的远程图片与外部 CSS 资源，避免 headless 打印时回连发件方、泄露「已读」状态（确需保留时加 `--allow-remote`）
- 自带体检脚本，交付前查空首页、打不开的 PDF、附件缺失

## 安装

任选一种，装完重开对话即可被识别：

```bash
# 1) 直接拷贝文件夹到技能目录
#    Claude 生态 : ~/.claude/skills/
#    通用        : <workspace>/.user_skills/

# 2) 用 skills CLI 安装（需要 node/npx）
npx skills add <owner>/<repo>@email-backup-to-pdf

# 3) 没有 node：直接下载 Release 里的 zip
#    https://github.com/583077395-star/email-backup-to-pdf/releases
#    解压后把 email-backup-to-pdf-<version> 文件夹放进上面任一个技能目录
```

## 依赖

| 依赖 | 用途 | 说明 |
|---|---|---|
| Python 3.8+ | 运行脚本 | 转换本身只用标准库 |
| Edge / Chrome / Chromium | HTML→PDF | 自动探测 Windows/macOS/Linux 常见路径，也可 `--edge` 指定 |
| pymupdf | 仅体检脚本 | `pip install pymupdf` |

## 用法

```bash
# 全量
python scripts/eml_to_pdf.py --zip-dir "./邮箱备份" --out "./邮箱备份_PDF"

# 增量（补转新增分卷/上次失败的邮件，直接重跑）
python scripts/eml_to_pdf.py --zip-dir "./邮箱备份" --out "./邮箱备份_PDF"

# 先小批量验证版面（之后记得再跑一次全量，否则剩余邮件不会补上）
python scripts/eml_to_pdf.py --zip-dir "./邮箱备份" --out "/tmp/test" --limit 3

# 可选参数
--date-format "%Y%m%d_%H%M"           # 文件名时间格式
--att-mode collection|per-mail|none   # 附件策略
--fresh                               # 忽略已有索引，全部重转
--allow-remote                        # 保留邮件中的远程图片（默认剥离）
```

体检：

```bash
python scripts/qa_pdfs.py --out "./邮箱备份_PDF"
```

输出结构：

```
<out>/
├── PDF/YYYYMMDD_HHMM_主题.pdf
├── 附件/<与 PDF 同名的目录>/<原始附件>
├── 索引.csv
└── 失败清单.csv      # 仅在出现失败时生成：时间, 来源包, 原件名, 原因
```

## 增量与重跑语义

判重键是 **（来源包, 原件名）**，不是（邮件时间, 主题）。这样：

- 同一天收到两封同名主题的邮件（日报、系统通知）都会各自成一份 PDF，不会互相覆盖或被当成重复丢弃
- 中途失败、中途 Ctrl-C、`--limit` 只跑了几封，重跑同一条命令都会把缺的补齐
- 文件名冲突仍用 `_2`、`_3` 兜底，索引只追加不重写（序号连续）

## 为什么不用 Word 转 PDF

在企业装了终端加密/DLP 的电脑上，Word COM 导出的 `.pdf` 实际是加密密文（文件头不是 `%PDF-`，除白名单程序外都打不开）。本项目统一用浏览器 headless 打印，产物是标准、可搜索、跨机器可读的 PDF。

另外，很多邮件（尤其 Outlook 转发/回复）正文本身是一整份带 `@page`、`WordSection` 的 HTML 文档，直接套模板会强制另起一页（首页只剩抬头）。脚本里 `clean_body()` 负责清洗，改代码时不要删。

## 许可

本项目采用 [MIT License](LICENSE) 开源。
