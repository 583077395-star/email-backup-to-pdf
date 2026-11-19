#!/usr/bin/env python3
"""转换结果体检：页数、空首页、过短 PDF、附件目录对不上号。

用法：python qa_pdfs.py --out <输出目录> [--sparse-threshold 400]
退出码 1 表示发现问题，便于批量流程断言。
"""
import argparse
import csv
import glob
import os
import sys


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", required=True)
    ap.add_argument("--sparse-threshold", type=int, default=400)
    args = ap.parse_args()

    try:
        import pymupdf
    except ImportError:
        sys.exit("需要 PyMuPDF：python -m pip install pymupdf")

    pdf_dir = os.path.join(args.out, "PDF")
    index = os.path.join(args.out, "索引.csv")
    pdfs = sorted(glob.glob(os.path.join(pdf_dir, "*.pdf")))
    if not pdfs:
        sys.exit("没有找到 PDF，检查 --out 是否正确")

    sparse, tiny, bad_end = [], [], []
    total_pages = 0
    for p in pdfs:
        try:
            d = pymupdf.open(p)
        except Exception as e:
            bad_end.append((os.path.basename(p), str(e)))
            continue
        total_pages += d.page_count
        t1 = len(d[0].get_text())
        if d.page_count >= 2 and t1 < args.sparse_threshold and len(d[1].get_text()) > t1:
            sparse.append(os.path.basename(p))
        if t1 < 120:
            tiny.append((os.path.basename(p), t1))
        d.close()

    rows = []
    if os.path.exists(index):
        with open(index, encoding="utf-8-sig") as f:
            rows = list(csv.DictReader(f))

    att_root = os.path.join(args.out, "附件")
    missing_att = []
    if os.path.isdir(att_root) and rows:
        for r in rows:
            if r.get("附件数") and int(r["附件数"]) > 0:
                stem = r["PDF文件"][:-4]
                if not os.path.isdir(os.path.join(att_root, stem)):
                    missing_att.append(stem)

    print("PDF 数量：%d，总页数：%d" % (len(pdfs), total_pages))
    print("索引行数：%d" % len(rows))
    print("打不开的 PDF：%d %s" % (len(bad_end), bad_end[:5]))
    print("首页异常空白（疑似强制分页）：%d %s" % (len(sparse), sparse[:5]))
    print("首页文字过少（多为短通知，需人工确认）：%d %s" % (len(tiny), tiny[:5]))
    print("附件缺失：%d %s" % (len(missing_att), missing_att[:5]))
    if bad_end or sparse or missing_att:
        sys.exit(1)


if __name__ == "__main__":
    main()
