# -*- coding: utf-8 -*-
"""倫理審査の書類づくりに使う配布物の最新版を、配布元のページから毎回取ってくる。

  python get_docs.py --out <研究のフォルダ>/様式                 # 全部
  python get_docs.py --out <研究のフォルダ>/様式 --only templates shishin
  python get_docs.py --list                                     # 取らずに、今の版だけ見る

取ってくるもの（配布元のページを毎回読み、リンクの名前で探す。ファイルの場所は改訂のたびに
変わるので決め打ちしない）:
  templates  北大病院の雛形 HT1〜HT4（研究計画書・同意説明文書・情報公開用文書・適切な同意の説明文書）
             令和8年版の指針（2026-12-01 施行）で「適切な同意」がなくなる。北大が HT4 の配布をやめたら「見つからない」と出るのは正常
  forms      北大病院の委員会様式・病院様式（生命・医学系研究）
  rules      北大病院 生命・医学系研究倫理審査委員会内規
  shishin    厚労省の倫理指針（本文）とガイダンス。最新版が施行前なら、いま効いている版も取る。PDF は .txt も作る

書き出すもの（--out）:
  ファイル（例: HT1_研究計画書ひな形_202412.docx、倫理指針_本文_20260827.pdf）
  取得元.md / sources.json  名前・版・URL・取った日・前回からの変化（新しい版が出ていたら「更新」）
古いファイルは消さない（新しい版は版の入った別の名前で置く）。

配布物は配布元の機関のもの。このスキルには入れない（再配布しない）。
厚労省のものは「出典：厚生労働省ホームページ（URL）」。
"""
import argparse
import hashlib
import io
import json
import re
import sys
import time
import unicodedata
import urllib.parse
import urllib.request
from datetime import date, datetime
from html import unescape
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
UA = {"User-Agent": "irb-docs/1.0 (+https://github.com/lfb15345-del/irb-docs)"}

HOKUDAI = "https://helios.huhp.hokudai.ac.jp/dcra/"
PAGES = {
    "templates": HOKUDAI + "seimei/研究実施計画書・同意説明文書・情報公開用文書/",
    "forms_irb": HOKUDAI + "seimei/生命・医学系研究倫理審査委員会様式一覧/",
    "forms_hosp": HOKUDAI + "rinsyou/生命科学・医学系研究の実施手続き/生命・医学系指針様式一覧/",
    "rules": HOKUDAI + "seimei/生命医学系研究倫理審査委員会/",
    "mhlw": "https://www.mhlw.go.jp/stf/seisakunitsuite/bunya/hokabunya/kenkyujigyou/i-kenkyu/index.html",
}
# (グループ, ページ, リンクの名前, 保存名)。保存名が None ならリンクの名前をそのまま使う（様式）
ITEMS = [
    ("templates", "templates", r"^研究計画書\s*雛形", "HT1_研究計画書ひな形"),
    ("templates", "templates", r"^同意説明文書\s*雛形", "HT2_同意説明文書ひな形"),
    ("templates", "templates", r"^情報公開用文書\s*雛形", "HT3_情報公開用文書ひな形"),
    ("templates", "templates", r"^適切な同意の説明文書\s*雛形", "HT4_適切な同意説明文書ひな形"),
    ("forms", "forms_irb", r"^(委員会様式|病院様式|参考様式)", None),
    ("forms", "forms_hosp", r"^病院様式", None),
    ("rules", "rules", r"倫理審査委員会内規", "北大病院_倫理審査委員会内規"),
    ("shishin", "mhlw", r"^人を対象とする生命科学・医学系研究に関する倫理指針（本文）", "倫理指針_本文"),
    ("shishin", "mhlw", r"^人を対象とする生命科学・医学系研究に関する倫理指針ガイダンス", "倫理指針_ガイダンス"),
]
DOC_RX = re.compile(r"\.(docx?|xlsx?|pdf)$", re.I)
ERA = {"令和": 2018, "平成": 1988}


def quote_url(u):
    return urllib.parse.quote(u, safe=":/%?=&#")


def fetch(url, tries=3):
    for k in range(tries):
        try:
            with urllib.request.urlopen(urllib.request.Request(quote_url(url), headers=UA), timeout=60) as r:
                return r.read()
        except Exception:
            if k == tries - 1:
                raise
            time.sleep(2 * (k + 1))


def links(page_url):
    html = fetch(page_url).decode("utf-8", "replace")
    for m in re.finditer(r'<a\s[^>]*href="([^"]+)"[^>]*>(.*?)</a>', html, re.S | re.I):
        href = urllib.parse.unquote(urllib.parse.urljoin(page_url, unescape(m.group(1))))
        text = re.sub(r"\s+", " ", unescape(re.sub(r"<[^>]+>", "", m.group(2)))).strip()
        if DOC_RX.search(href.split("?")[0]):
            yield text, href


def wareki(text):
    """'令和8年8月27日' → date。無ければ None。全角の数字・空白も読む。"""
    text = re.sub(r"\s+", "", unicodedata.normalize("NFKC", text))
    m = re.search(r"(令和|平成)(\d+|元)年(\d+)月(\d+)日", text)
    if not m:
        return None
    y = 1 if m.group(2) == "元" else int(m.group(2))
    return date(ERA[m.group(1)] + y, int(m.group(3)), int(m.group(4)))


def version_of(text, href):
    """版: リンクの名前の『2024年12月版』、和暦の日付、無ければ置き場所の年月（uploads/2023/03 → 202303）。"""
    text = unicodedata.normalize("NFKC", text)
    m = re.search(r"(\d{4})年(\d{1,2})月版", text)
    if m:
        return f"{m.group(1)}{int(m.group(2)):02d}"
    d = wareki(text)
    if d:
        return d.strftime("%Y%m%d")
    m = re.search(r"/uploads/(\d{4})/(\d{2})/", href)
    return f"{m.group(1)}{m.group(2)}" if m else ""


def safe(name):
    return re.sub(r'[\\/:*?"<>|\s]+', "_", name).strip("_")


def pdf_text(data):
    try:
        import fitz  # pymupdf
    except ImportError:
        return None
    doc = fitz.open(stream=data, filetype="pdf")
    return "\n".join(p.get_text() for p in doc)


def collect(groups):
    """取るものの一覧（まだダウンロードしない）。見つからないものは missing に。"""
    cache, found, missing, seen = {}, [], [], {}
    for grp, page, rx, stem in ITEMS:
        if grp not in groups:
            continue
        if page not in cache:
            try:
                cache[page] = list(links(PAGES[page]))
            except Exception as e:
                missing.append(f"{grp}: ページを読めない（{PAGES[page]}）: {e}")
                cache[page] = []
                continue
        hits = [(t, h) for t, h in cache[page] if re.search(rx, t)]
        if not hits:
            missing.append(f"{grp}: 「{rx.lstrip('^')}」が見つからない（ページの作りが変わったかもしれない: {PAGES[page]}）")
            continue
        if grp == "shishin":                      # 版がいくつも並ぶ。新しい順に
            hits = sorted(hits, key=lambda th: wareki(th[0]) or date.min, reverse=True)
            found.append(dict(group=grp, name=stem, text=hits[0][0], url=hits[0][1], page=PAGES[page],
                              older=[dict(text=t, url=h) for t, h in hits[1:]]))
            continue
        for t, h in hits:
            name = stem or re.sub(r"\s+", "", t.split("「")[0])
            if name in seen:                      # 同じ名前が別のページにもある（病院様式02 など）
                if seen[name]["url"] != h:
                    seen[name].setdefault("also", []).append(h)
                continue
            item = dict(group=grp, name=name, text=t, url=h, page=PAGES[page])
            seen[name] = item
            found.append(item)
            if stem:
                break
    return found, missing


def main():
    ap = argparse.ArgumentParser(description="倫理審査の配布物（北大病院の雛形・様式・内規、厚労省の倫理指針）の最新版を取る")
    ap.add_argument("--out", help="保存先（研究のフォルダの 様式/ など）")
    ap.add_argument("--only", nargs="*", default=["templates", "forms", "rules", "shishin"],
                    choices=["templates", "forms", "rules", "shishin"])
    ap.add_argument("--list", action="store_true", help="取らずに、今の版を一覧にする")
    a = ap.parse_args()
    if not a.list and not a.out:
        ap.error("--out か --list を付ける")

    found, missing = collect(set(a.only))
    if a.list:
        for it in found:
            print(f"{it['group']:<9} {it['name']:<22} {version_of(it['text'], it['url']):<9} {it['text']}")
        for m in missing:
            print("見つからない:", m)
        return 1 if missing else 0

    out = Path(a.out)
    out.mkdir(parents=True, exist_ok=True)
    prev = {}                                     # 名前 → 前回の行（指針は版が2つ並ぶことがある）
    if (out / "sources.json").exists():
        for r in json.loads((out / "sources.json").read_text(encoding="utf-8"))["items"]:
            prev.setdefault(r["name"], []).append(r)
    today = date.today()
    rows, notes = [], []

    def save(it, name, text, url, role=""):
        ver = version_of(text, url)
        ext = Path(urllib.parse.urlparse(url).path).suffix.lower()
        fname = f"{safe(name)}_{ver}{ext}" if ver else f"{safe(name)}{ext}"
        data = fetch(url)
        (out / fname).write_bytes(data)
        sha = hashlib.sha256(data).hexdigest()
        row = dict(group=it["group"], name=name, version=ver, file=fname, link_text=text, url=url, page=it["page"],
                   sha256=sha, fetched=datetime.now().isoformat(timespec="seconds"), role=role,
                   also=it.get("also", []))
        olds = prev.get(name, [])
        if not olds:
            row["change"] = "新規"
        elif sha in {o["sha256"] for o in olds}:
            row["change"] = "同じ"
        else:
            row["change"] = "更新（前回 " + "・".join(o.get("version") or "版なし" for o in olds) + "）"
        if ext == ".pdf":
            txt = pdf_text(data)
            if txt is not None:
                (out / (Path(fname).stem + ".txt")).write_text(txt, encoding="utf-8")
                row["txt"] = Path(fname).stem + ".txt"
                flat = re.sub(r"\s+", "", unicodedata.normalize("NFKC", txt))
                m = re.search(r"この指針は、((令和|平成)(\d+|元)年\d+月\d+日)から施行", flat)
                if m and wareki(m.group(1)):
                    row["effective"] = wareki(m.group(1)).isoformat()
        rows.append(row)
        print(f"{row['change']:<14} {fname}")
        time.sleep(0.5)
        return row

    for it in found:
        try:
            r = save(it, it["name"], it["text"], it["url"])
            if it["group"] == "shishin" and r.get("effective") and date.fromisoformat(r["effective"]) > today:
                r["role"] = f"次の版（{r['effective']} から施行）"
                notes.append(f"{it['name']}: 最新版は {r['effective']} から施行。それまでは前の版が効いている（両方を置いた）。")
                if it["older"]:
                    o = it["older"][0]
                    save(it, it["name"], o["text"], o["url"], role="いま効いている版")
        except Exception as e:
            missing.append(f"{it['name']}: 取れない（{it['url']}）: {e}")

    done = {r["group"] for r in rows}
    keep = [r for rs in prev.values() for r in rs if r["group"] not in done]
    (out / "sources.json").write_text(json.dumps({"fetched": today.isoformat(), "items": rows + keep}, ensure_ascii=False,
                                                 indent=1), encoding="utf-8")
    lines = [f"# 取得元（{today.isoformat()} に取得）", "",
             "配布元のページから最新版を取ったもの。配布物は配布元の機関のもの（再配布しない）。",
             "厚労省のもの: 出典：厚生労働省ホームページ（下の URL）。", ""]
    lines += [f"- {n}" for n in notes]
    lines += ["", "| 変化 | ファイル | 版・名前 | 役割 | 取得元 |", "|---|---|---|---|---|"]
    lines += [f"| {r['change']} | {r['file']} | {r['link_text']} | {r['role']} | {r['url']} |" for r in rows]
    alts = [r for r in rows if r["also"]]
    if alts:
        lines += ["", "同じ名前の別ファイルが別のページにもある（中身を見比べて、新しい方を使う）:"]
        lines += [f"- {r['name']}: {', '.join(r['also'])}" for r in alts]
    if missing:
        lines += ["", "## 取れなかったもの", ""] + [f"- {m}" for m in missing]
    (out / "取得元.md").write_text("\n".join(lines) + "\n", encoding="utf-8")
    upd = [r for r in rows if r["change"].startswith("更新")]
    print(f"\n{len(rows)} 件 → {out}（新規 {sum(r['change'] == '新規' for r in rows)}・更新 {len(upd)}）")
    for n in notes:
        print("注意:", n)
    for m in missing:
        print("取れなかった:", m)
    return 1 if missing else 0


if __name__ == "__main__":
    sys.exit(main())
