# -*- coding: utf-8 -*-
"""PDF のページを PNG にして、目で確認できるようにする。

    python render_pages.py 計画書.pdf --out <作業フォルダ> [--pages 1 2] [--find 研究の背景 文献リスト] [--dpi 110]

--find を付けると、その語を含むページを探して書き出す。何も指定しなければ全ページ。
書き出したら Read ツールで画像を開いて見る（表紙・本文1ページ・同意文書・文献・様式は必ず見る）。
"""
import argparse
import io
import sys
from pathlib import Path

import fitz  # PyMuPDF

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf")
    ap.add_argument("--out", required=True)
    ap.add_argument("--pages", type=int, nargs="*", default=[], help="1始まり")
    ap.add_argument("--find", nargs="*", default=[])
    ap.add_argument("--dpi", type=int, default=110)
    a = ap.parse_args()
    out = Path(a.out)
    out.mkdir(parents=True, exist_ok=True)
    doc = fitz.open(a.pdf)
    print(f"{Path(a.pdf).name}: {doc.page_count} ページ")
    want = set(p - 1 for p in a.pages)
    for i in range(doc.page_count):
        text = doc[i].get_text()
        for w in a.find:
            if w in text:
                print(f"  「{w}」 → {i + 1} ページ")
                want.add(i)
    if not a.pages and not a.find:
        want = set(range(doc.page_count))
    stem = Path(a.pdf).stem[:20]
    for i in sorted(want):
        p = out / f"{stem}_p{i + 1:02d}.png"
        doc[i].get_pixmap(dpi=a.dpi).save(str(p))
        print("saved", p)


if __name__ == "__main__":
    main()
