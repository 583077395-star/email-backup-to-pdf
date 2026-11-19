#!/usr/bin/env python3
"""邮箱备份 .eml 批量转 PDF + 附件归档。

用法：
  python eml_to_pdf.py --zip-dir <备份目录> --out <输出目录> [--limit N] [--date-format FMT]
                       [--att-mode collection|per-mail|none] [--edge PATH] [--fresh]

行为：
  * 自动识别 zip 分卷；遇到未下载完整/损坏的分卷直接报错退出（不生成残缺文件）。
  * 已存在 索引.csv 时默认增量：跳过已转换过的分卷，文件名去重，索引用追加模式。
  * 输出：<out>/PDF/*.pdf  <out>/附件/<PDF同名>/*  <out>/索引.csv
"""
import argparse
import base64
import csv
import email
import html as htmlmod
import os
import re
import shutil
import subprocess
import sys
import zipfile
from email import policy
from email.utils import parsedate_to_datetime

ILLEGAL = re.compile(r'[\\/:*?"<>|\r\n\t]+')
EDGE_CANDIDATES = [
    # Windows
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    # macOS
    "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    "/Applications/Chromium.app/Contents/MacOS/Chromium",
    # Linux
    "/usr/bin/microsoft-edge",
    "/usr/bin/google-chrome",
    "/usr/bin/chromium",
    "/usr/bin/chromium-browser",
]

STYLE = """
@page { size: A4; margin: 16mm 14mm; }
body { font-family: "Microsoft YaHei","SimSun",sans-serif; font-size: 10.5pt; line-height: 1.55; margin:0; color:#1a1a1a; }
h1.s { font-size: 14pt; margin: 0 0 10px 0; line-height:1.35; }
table.m { border-collapse: collapse; width:100%; margin-bottom: 12px; }
table.m td { border-bottom: 1px solid #d8d8d8; padding: 4px 6px; vertical-align: top; }
table.m td.k { width: 56px; color:#666; white-space: nowrap; }
hr { border:none; border-top:1px solid #dcdcdc; margin: 10px 0 14px 0; }
img { max-width: 100%; height:auto; }
.atts { margin-top: 18px; border-top:1px solid #dcdcdc; padding-top:8px; font-size:9.5pt; color:#333; }
.atts ul { margin:4px 0 0 18px; padding:0; }
.foot { margin-top: 16px; color:#999; font-size:8.5pt; }
pre { white-space: pre-wrap; font-family: inherit; }
"""


def sanitize(text, maxlen=80):
    text = ILLEGAL.sub("_", (text or "").strip())
    text = re.sub(r"\s+", " ", text).strip(" ._")
    return text[:maxlen].rstrip(" ._") or "无主题"


def clean_body(content):
    """去掉正文里内嵌的整份 HTML 文档、@page 与 Word 分节定义（否则会强制另起一页）。"""
    content = re.sub(r"(?is)<!DOCTYPE[^>]*>", "", content)
    content = re.sub(r"(?is)<head[^>]*>.*?</head>", "", content)
    content = re.sub(r"(?is)<html[^>]*>|</html>|<body[^>]*>|</body>", "", content)
    content = re.sub(r"(?is)@page\s*[^{]*\{[^}]*\}", "", content)
    content = re.sub(r"(?i)\bpage\s*:\s*[A-Za-z0-9_]+\s*;", "", content)
    content = re.sub(r"(?i)mso-[a-z-]+\s*:\s*[^;\"']+;?", "", content)
    content = re.sub(r"(?is)<script.*?</script>", "", content)
    content = re.sub(r"(?is)<link[^>]*>|<meta[^>]*>", "", content)
    return content


def decode_name(raw):
    """zip 内中文名按 cp936 还原。"""
    try:
        return raw.encode("cp437").decode("cp936")
    except Exception:
        return raw


def build_html(msg, orig_name, att_list, date_fmt):
    subject = str(msg.get("Subject") or "")
    sender = str(msg.get("From") or "")
    to = str(msg.get("To") or "")
    cc = str(msg.get("Cc") or "")
    try:
        dt = parsedate_to_datetime(msg.get("Date")).astimezone()
        date_str, ts = dt.strftime("%Y-%m-%d %H:%M"), dt.strftime(date_fmt)
    except Exception:
        date_str, ts = str(msg.get("Date") or ""), "00000000_0000"

    body = msg.get_body(preferencelist=("html", "plain"))
    if body is None:
        content = "<p>（无正文）</p>"
    elif body.get_content_type() == "text/html":
        content = body.get_content()
    else:
        content = "<pre>" + htmlmod.escape(body.get_content()) + "</pre>"

    for part in msg.walk():
        cid = part.get("Content-ID")
        if not cid:
            continue
        data = part.get_payload(decode=True)
        if not data:
            continue
        uri = "data:%s;base64,%s" % (part.get_content_type(), base64.b64encode(data).decode("ascii"))
        key = "cid:" + cid.strip().strip("<>")
        content = content.replace(key, uri).replace(key.upper(), uri)

    content = re.sub(r"\son\w+\s*=\s*\"[^\"]*\"", "", content, flags=re.I)
    content = clean_body(content)

    att_html = ""
    if att_list:
        items = "".join("<li>%s</li>" % htmlmod.escape(a) for a in att_list)
        att_html = "<div class='atts'><b>附件（%d）</b><ul>%s</ul></div>" % (len(att_list), items)

    return ts, date_str, subject, """<!DOCTYPE html>
<html><head><meta charset="utf-8"><title>%s</title><style>%s</style></head><body>
<h1 class="s">%s</h1>
<table class="m">
<tr><td class="k">发件人</td><td>%s</td></tr>
<tr><td class="k">收件人</td><td>%s</td></tr>
%s
<tr><td class="k">时间</td><td>%s</td></tr>
</table>
<hr>
%s
%s
<div class="foot">原件 %s</div>
</body></html>""" % (
        htmlmod.escape(subject), STYLE, htmlmod.escape(subject), htmlmod.escape(sender),
        htmlmod.escape(to),
        ("<tr><td class='k'>抄送</td><td>%s</td></tr>" % htmlmod.escape(cc)) if cc else "",
        htmlmod.escape(date_str), content, att_html, htmlmod.escape(orig_name))


def find_browser(explicit=None):
    candidates = [explicit] if explicit else []
    candidates += EDGE_CANDIDATES
    # PATH 兜底：各平台常见命令名
    for cmd in ("msedge", "google-chrome", "chromium", "chromium-browser", "chrome"):
        p = shutil.which(cmd)
        if p:
            candidates.append(p)
    for p in candidates:
        if p and os.path.exists(p):
            return p
    sys.exit("未找到 Edge/Chrome/Chromium，请用 --edge 指定浏览器可执行文件路径")


def print_pdf(browser, html_path, pdf_path, profile):
    from pathlib import Path
    url = Path(os.path.abspath(html_path)).as_uri()
    cmd = [browser, "--headless=new", "--disable-gpu", "--no-pdf-header-footer",
           "--user-data-dir=" + profile, "--print-to-pdf=" + pdf_path, url]
    subprocess.run(cmd, capture_output=True, timeout=300)
    if not os.path.exists(pdf_path):
        return False
    with open(pdf_path, "rb") as f:
        good = f.read(5) == b"%PDF-"
    if not good:
        os.remove(pdf_path)
    return good


def open_zip(path):
    try:
        z = zipfile.ZipFile(path)
        z.infolist()
        return z
    except Exception as e:
        sys.exit("分卷不可用（可能仍在同步/下载或已损坏）：%s\n原因：%s\n请等下载完成或改用完整副本后重试。" % (path, e))


def load_index(path):
    rows, stems, keys, pkgs = [], set(), set(), set()
    if os.path.exists(path):
        with open(path, encoding="utf-8-sig") as f:
            for r in csv.DictReader(f):
                rows.append(r)
                if r.get("来源包"):
                    pkgs.add(r["来源包"])
                if r.get("PDF文件"):
                    stems.add(r["PDF文件"][:-4])
                keys.add((r.get("邮件时间", ""), r.get("主题", "")))
    return rows, stems, keys, pkgs


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--zip-dir", required=True, help="存放 messages_package x_y.zip 的目录")
    ap.add_argument("--out", required=True, help="输出目录")
    ap.add_argument("--limit", type=int, default=0, help="只处理前 N 封（调试用）")
    ap.add_argument("--date-format", default="%Y%m%d_%H%M", help="文件名时间格式")
    ap.add_argument("--att-mode", default="collection", choices=["collection", "per-mail", "none"],
                    help="collection=附件/<PDF同名>/文件；per-mail=与 PDF 同目录；none=不导出附件")
    ap.add_argument("--edge", default=None)
    ap.add_argument("--fresh", action="store_true", help="忽略已有索引，全部重转")
    ap.add_argument("--keep-work", action="store_true", help="保留临时 html/浏览器 profile")
    args = ap.parse_args()

    pdf_dir = os.path.join(args.out, "PDF")
    att_dir = os.path.join(args.out, "附件")
    work = os.path.join(args.out, ".work")
    index_path = os.path.join(args.out, "索引.csv")
    for d in (pdf_dir, work):
        os.makedirs(d, exist_ok=True)
    if args.att_mode == "collection":
        os.makedirs(att_dir, exist_ok=True)
    if args.fresh and os.path.exists(index_path):
        os.remove(index_path)

    browser = find_browser(args.edge)
    rows, used, seen, done_pkgs = load_index(index_path)
    if args.fresh:
        rows, used, seen, done_pkgs = [], set(), set(), set()
    header_written = os.path.exists(index_path)
    seq = len(rows)
    added = skipped = failed = 0

    zips = sorted(f for f in os.listdir(args.zip_dir) if f.lower().endswith(".zip"))
    if not zips:
        sys.exit("--zip-dir 下没有 .zip 分卷")

    with open(index_path, "a", encoding="utf-8-sig", newline="") as idx:
        w = csv.writer(idx)
        if not header_written:
            w.writerow(["序号", "邮件时间", "主题", "发件人", "收件人", "PDF文件", "附件数", "来源包", "原件名"])
        for zp in zips:
            if zp in done_pkgs:
                print("跳过已处理分卷：%s" % zp, flush=True)
                continue
            zpath = os.path.join(args.zip_dir, zp)
            z = open_zip(zpath)
            with z:
                for info in z.infolist():
                    if args.limit and added >= args.limit:
                        break
                    orig = decode_name(info.filename)
                    raw = z.read(info)
                    try:
                        msg = email.message_from_bytes(raw, policy=policy.default)
                        att = [p.get_filename() for p in msg.walk()
                               if p.get_filename() and not (p.get("Content-ID") and p.get_content_maintype() == "image")]
                        ts, date_str, subject, doc = build_html(msg, orig, att, args.date_format)

                        if (date_str, subject) in seen:
                            skipped += 1
                            print("SKIP 重复：%s %s" % (date_str, subject), flush=True)
                            continue

                        base = "%s_%s" % (ts, sanitize(subject))
                        stem, n = base, 2
                        while stem in used:
                            stem = "%s_%d" % (base, n)
                            n += 1
                        used.add(stem)
                        seen.add((date_str, subject))

                        if args.att_mode != "none":
                            for p in msg.walk():
                                fn = p.get_filename()
                                if not fn or (p.get("Content-ID") and p.get_content_maintype() == "image"):
                                    continue
                                payload = p.get_payload(decode=True)
                                if not payload:
                                    continue
                                target_dir = os.path.join(att_dir, stem) if args.att_mode == "collection" else pdf_dir
                                os.makedirs(target_dir, exist_ok=True)
                                safe = sanitize(fn, 100)
                                fp = os.path.join(target_dir, safe)
                                b, e = os.path.splitext(safe)
                                k = 2
                                while os.path.exists(fp):
                                    fp = os.path.join(target_dir, "%s_%d%s" % (b, k, e))
                                    k += 1
                                with open(fp, "wb") as fh:
                                    fh.write(payload)

                        hp = os.path.join(work, "t.html")
                        with open(hp, "w", encoding="utf-8") as fh:
                            fh.write(doc)
                        tmp = os.path.join(work, "_t.pdf")
                        if os.path.exists(tmp):
                            os.remove(tmp)
                        out_pdf = os.path.join(pdf_dir, stem + ".pdf")
                        if print_pdf(browser, hp, tmp, os.path.join(work, "profile")):
                            os.replace(tmp, out_pdf)
                            seq += 1
                            added += 1
                            w.writerow([seq, date_str, subject, str(msg.get("From") or ""),
                                        str(msg.get("To") or ""), stem + ".pdf", len(att), zp, orig])
                            idx.flush()
                            print("[%d] OK %s att=%d" % (seq, stem, len(att)), flush=True)
                        else:
                            failed += 1
                            print("FAIL 渲染失败：%s" % stem, flush=True)
                    except Exception as e:
                        failed += 1
                        print("ERR %s: %s" % (orig, e), flush=True)

    if not args.keep_work:
        import shutil
        shutil.rmtree(work, ignore_errors=True)

    print("\n完成：新增 %d，跳过重复 %d，失败 %d；输出目录 %s" % (added, skipped, failed, args.out))
    if failed:
        sys.exit(1)


if __name__ == "__main__":
    main()
