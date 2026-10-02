# -*- coding: utf-8 -*-
"""引用文献を探し、抄録で主張を確かめ、RISをダウンロードして文献リストを作る。

  1) 探す      python fetch_refs.py search "physician work style reform Japan" [--max 20]
  2) 確かめる  python fetch_refs.py abstract 39134926 20222990
  3) 作る      python fetch_refs.py build refs.json --out <この研究の文献フォルダ> --title "<計画書名>" [--clean]

連絡先は環境変数 IRB_DOCS_MAILTO で渡す（NCBI・Crossref の推奨。無くても動くが回数制限が厳しい）。
--clean は出力先の古い NN_*.ris を消す。別の研究の文献フォルダを出力先にしない。

refs.json（本文での登場順に並べる。番号はこの順で振る）:
[
  {"key": "Wada2025", "pmid": "39134926", "cited_at": "1章 背景（上限規制と客観的な覚醒度）"},
  {"key": "MaasHox2005", "doi": "10.1027/1614-2241.1.3.86", "abbr": "Methodology", "cited_at": "13章 症例数"}
]
- PubMed 収載なら pmid を書く（NCBI Literature Citation Exporter から RIS を取る）。
- 非収載なら doi を書く（Crossref から RIS を取る）。Crossref の RIS は雑誌略名を持たないので abbr を添える。
- ウェブページ等で RIS が取れない文献は "manual_ris" に RIS 本文を書く（文献リスト.md に手作りと明記される）。

build が書き出すもの（--out）:
  NN_key.ris / references_all.ris / references_meta.json（整形済み表記つき）/ 文献リスト.md
生成スクリプトは references_meta.json の "reference" をそのまま文献リストに使う。

書式（北大病院の承認済み計画書の例に合わせた）:
  著者（6名を超えたら先頭6名, et al）. 表題. 雑誌略名 年;巻:頁. doi:….
  表題の大文字小文字は取得元のまま。頁は PubMed 本体の表記を全桁に展開（1043-65 → 1043–1065）。
"""
import argparse
import io
import json
import os
import re
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
# NCBI と Crossref は連絡先の付いたアクセスを推奨している。連絡先は環境変数 IRB_DOCS_MAILTO で渡す
_MAILTO = os.environ.get("IRB_DOCS_MAILTO", "").strip()
UA = {"User-Agent": "irb-docs-refs/1.1" + (f" (mailto:{_MAILTO})" if _MAILTO else "")}
EUTILS = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/"
CTXP = "https://api.ncbi.nlm.nih.gov/lit/ctxp/v1/pubmed/?format=ris&id="


def get(url, tries=4):
    """回数制限（429）や一時的な失敗は待って再試行する。NCBI は連絡先なしだと毎秒3回まで。"""
    for k in range(tries):
        try:
            with urllib.request.urlopen(urllib.request.Request(url, headers=UA), timeout=40) as r:
                return r.read().decode("utf-8", "replace")
        except urllib.error.HTTPError as e:
            if e.code not in (429, 500, 502, 503, 504) or k == tries - 1:
                raise
        except urllib.error.URLError:
            if k == tries - 1:
                raise
        time.sleep(1.5 * (k + 1))


def norm_newlines(txt):
    """PubMed は CRLF で返す。Windows でそのまま書くと CR CR LF になるので、いったん LF にそろえる。"""
    return txt.replace("\r\n", "\n").replace("\r", "\n")


def write_ris(path, txt):
    """RIS は CRLF で書く（EndNote 等の取り込みで一般的な形）。"""
    Path(path).write_text(norm_newlines(txt).strip() + "\n", encoding="utf-8", newline="\r\n")


def esummary(pmids):
    if not pmids:
        return {}
    js = json.loads(get(EUTILS + "esummary.fcgi?db=pubmed&retmode=json&id=" + ",".join(pmids)))
    return js["result"]


def doi_of(d):
    return next((a["value"] for a in d.get("articleids", []) if a["idtype"] == "doi"), None)


# ------------------------------------------------------------------ search / abstract
def cmd_search(a):
    q = urllib.parse.quote(a.query)
    ids = json.loads(get(f"{EUTILS}esearch.fcgi?db=pubmed&retmode=json&sort=relevance&retmax={a.max}&term={q}"))[
        "esearchresult"]["idlist"]
    print(f"{len(ids)} 件: {a.query}")
    time.sleep(0.4)
    s = esummary(ids)
    for i in ids:
        d = s[i]
        kinds = [t for t in d.get("pubtype", []) if t not in ("Journal Article", "English Abstract")]
        kind = "/".join(kinds)[:30] or "Article"
        print(f"  PMID {i} | {d.get('pubdate', '')[:4]} | {d.get('source', '')[:26]:<26} | {kind:<18} | "
              f"{d.get('title', '')[:110]} | doi:{doi_of(d)}")
    if len(ids) < 3:
        print("  （少ないときは語を減らすか、[tiab]・[pt]・[dp] で絞り方を変える。citations.md の検索例を参照）")


def cmd_abstract(a):
    x = get(EUTILS + "efetch.fcgi?db=pubmed&rettype=abstract&retmode=text&id=" + ",".join(a.pmids))
    x = re.sub(r"Author information:.*?\n\n", "", x, flags=re.S)
    print(x)


# ------------------------------------------------------------------ build
def parse_ris(txt):
    rec = {}
    for line in txt.splitlines():
        m = re.match(r"^([A-Z][A-Z0-9])  - ?(.*)$", line)
        if not m:
            continue
        k, v = m.group(1), m.group(2).strip()
        if k == "ER":
            break
        rec.setdefault(k, []).append(v)
    return rec


def initials(given):
    out = ""
    for p in re.split(r"[\s.]+", given.strip()):
        for q in p.split("-"):
            if q:
                out += q[0].upper()
    return out


SUFFIX_RX = re.compile(r"^(Jr|Sr|II|III|IV|V|VI|\d+(st|nd|rd|th))\.?$", re.I)


def fmt_author(a):
    """'Shaffer, F' / 'Maas, Cora J. M.' / 'Kivimäki, Mika' → 'Shaffer F'
    'Bogardus, Sidney T, Jr' → 'Bogardus ST Jr'、'Brown, Charles H, 4th' → 'Brown CH 4th'（接尾辞を頭文字に混ぜない）"""
    if "," not in a:
        return a.strip()
    parts = [x.strip() for x in a.split(",")]
    last, rest = parts[0], [x for x in parts[1:] if x]
    suffix = [x.rstrip(".") for x in rest if SUFFIX_RX.match(x)]
    given = " ".join(x for x in rest if not SUFFIX_RX.match(x))
    ini = given.replace(" ", "") if re.fullmatch(r"[A-Z](\s?[A-Z])*", given) else initials(given)
    out = f"{last} {ini}".strip()
    if suffix:
        out += " " + " ".join(suffix)
    return out


SIMPLE_PAGES_RX = re.compile(r"[A-Za-z]?\d+(-[A-Za-z]?\d+)?")


def complex_pages(pg):
    """単純な「頁」「頁-頁」でない（330.e1-6、869-75.e1-2、318-30, 330.e1-6 など）。"""
    pg = (pg or "").strip()
    return bool(pg) and not SIMPLE_PAGES_RX.fullmatch(pg)


def nlm_pages(pg):
    pg = (pg or "").strip()
    if not pg:
        return ""
    if complex_pages(pg):
        return pg.replace("-", "–")      # 複合頁は PubMed の表記のまま（全桁に展開しない）
    m = re.fullmatch(r"([A-Za-z]?)(\d+)-([A-Za-z]?)(\d+)", pg)
    if m:
        pre, sp, _, ep = m.groups()
        if len(ep) < len(sp):
            ep = sp[:len(sp) - len(ep)] + ep
        return f"{pre}{sp}–{pre}{ep}"
    return pg


def format_ref(rec, doi, pages=None, abbr=None, authors=None):
    """authors: PubMed の esummary の著者名（'Westendorp RG'、'Bogardus ST Jr'）。あればそれを使う。
    RIS の名（'Westendorp, Rudi G J'）から頭文字を作ると、PubMed の表記と食い違うことがある。"""
    title = (rec.get("T1") or rec.get("TI") or [""])[0].rstrip(".")
    authors = list(authors or []) or [fmt_author(x) for x in rec.get("AU", []) + rec.get("A1", [])]
    if not authors and ". Task Force of" in title:        # PubMed は団体著者を表題の後ろに置く
        title, corp = title.split(". Task Force of", 1)
        authors = ["Task Force of" + corp]
    if not authors and rec.get("A2"):
        authors = rec["A2"]
    au = ", ".join(authors[:6]) + (", et al" if len(authors) > 6 else "")
    j = abbr or (rec.get("J2") or [""])[0] or (rec.get("JA") or [""])[0] or (rec.get("T2") or rec.get("JF") or [""])[0]
    year = re.search(r"\d{4}", " ".join(rec.get("Y1", []) + rec.get("PY", []) + rec.get("DA", [])))
    year = year.group(0) if year else ""
    vol = (rec.get("VL") or [""])[0]
    if pages is None:
        sp = (rec.get("SP") or [""])[0]
        ep = (rec.get("EP") or [""])[0]
        if "-" in sp and not ep:
            sp, ep = sp.split("-", 1)
        pages = f"{sp}–{ep}" if sp and ep and sp != ep else sp
    s = f"{au}. {title}. {j} {year}".strip()
    if vol:
        s += f";{vol}"
    if pages:
        s += f":{pages}"
    s += "."
    if doi:
        s += f" doi:{doi}."
    return s


def cmd_build(a):
    items = json.loads(Path(a.refs).read_text(encoding="utf-8"))
    out = Path(a.out)
    out.mkdir(parents=True, exist_ok=True)
    summ = esummary([it["pmid"] for it in items if it.get("pmid")])
    if a.clean:
        for old in out.glob("[0-9][0-9]_*.ris"):
            old.unlink()
    meta, combined = [], []
    for no, it in enumerate(items, 1):
        key = f"{no:02d}_{it['key']}"
        fixes, pages, doi, names = [], None, it.get("doi"), None
        if it.get("manual_ris"):
            ris, src = norm_newlines(it["manual_ris"]), "manual"
            fixes.append("RIS が取得できないため手作り")
        elif it.get("pmid"):
            ris, src = norm_newlines(get(CTXP + it["pmid"])), "pubmed"
            d = summ[it["pmid"]]
            doi = doi or doi_of(d)
            pages = nlm_pages(d.get("pages")) or None
            names = [x["name"] for x in d.get("authors", []) if x.get("name")]
            if complex_pages(d.get("pages")):
                # Citation Exporter は複合頁の終頁を壊す（330.e1-6 → 330.e3306、869-75.e1-2 → 75.e752）ので PubMed 本体の表記に置換
                ris = re.sub(r"\nSP  - [^\n]*", lambda _: "\nSP  - " + d["pages"], ris, count=1)
                ris = re.sub(r"\nEP  - [^\n]*", "", ris, count=1)
                fixes.append(f"頁欄を PubMed 本体の表記（{d['pages']}）に置換")
        else:
            q = urllib.parse.quote(doi, safe="")
            ris = norm_newlines(get(f"https://api.crossref.org/works/{q}/transform/application/x-research-info-systems"))
            src = "crossref"
            if not re.search(r"\nSP  - ", ris):
                art = json.loads(get(f"https://api.crossref.org/works/{q}"))["message"].get("article-number")
                if art:
                    ris = re.sub(r"\nER  -", lambda _: f"\nSP  - {art}\nER  -", ris, count=1)
                    pages = art
                    fixes.append(f"頁欄に論文番号（{art}）を追加")
        rec = parse_ris(ris)
        if not rec.get("TY"):
            raise SystemExit(f"RIS が取れない: {key}")
        if doi and "DO" not in rec:
            ris = re.sub(r"\nER  -", lambda _: f"\nDO  - {doi}\nER  -", ris, count=1)
            fixes.append("DOI を追加")
        write_ris(out / f"{key}.ris", ris)
        combined.append(ris.strip())
        ref = format_ref(rec, doi, pages, it.get("abbr"), names)
        meta.append({"no": no, "file": f"{key}.ris", "pmid": it.get("pmid"), "doi": doi, "source": src,
                     "cited_at": it.get("cited_at", ""), "ris_fixes": fixes, "reference": ref})
        print(f"{no:>2}. {ref}")
        time.sleep(0.4)
    write_ris(out / "references_all.ris", "\n\n".join(combined))
    (out / "references_meta.json").write_text(json.dumps(meta, ensure_ascii=False, indent=1), encoding="utf-8")
    n_pm = sum(m["source"] == "pubmed" for m in meta)
    n_cr = sum(m["source"] == "crossref" for m in meta)
    n_mn = sum(m["source"] == "manual" for m in meta)
    got = []
    if n_pm:
        got.append(f"PubMed 収載の{n_pm}件は NCBI の Literature Citation Exporter から")
    if n_cr:
        got.append(f"PubMed 非収載の{n_cr}件は Crossref から")
    how = "、".join(got) + " RIS をダウンロードした。" if got else ""
    if n_mn:
        how += f"RIS が取得できない{n_mn}件は手で書いた（下の手直しの一覧を参照）。"
    lines = [
        f"# 参考文献（{a.title}）", "",
        "番号は研究計画書の本文での登場順。各文献の RIS は同じフォルダにある（EndNote / Mendeley / Zotero に取り込める）。",
        how,
        f"`references_all.ris` は{len(meta)}件をまとめたもの。`references_meta.json` は番号・PMID・DOI・引用箇所・整形済みの表記。",
        "", "| # | ファイル | PMID | 取得元 | 引用箇所 |", "|---|---|---|---|---|",
    ]
    lines += [f"| {m['no']} | {m['file']} | {m['pmid'] or '—'} | {m['source']} | {m['cited_at']} |" for m in meta]
    fx = [m for m in meta if m["ris_fixes"]]
    if fx:
        lines += ["", "## ダウンロードした RIS に加えた手直し", ""]
        lines += [f"- {m['file']}: {'、'.join(m['ris_fixes'])}" for m in fx]
    lines += ["", "## 計画書の文献リスト", ""] + [f"{m['no']}. {m['reference']}" for m in meta]
    lines += ["", "## 取り直し方", "", "```bash",
              f'python "{Path(__file__).resolve()}" build "{Path(a.refs).resolve()}" --out "{out.resolve()}" '
              f'--title "{a.title}" --clean',
              "```", ""]
    (out / "文献リスト.md").write_text("\n".join(lines), encoding="utf-8")
    print(f"saved {len(meta)} → {out}")


def main():
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="cmd", required=True)
    s = sub.add_parser("search")
    s.add_argument("query")
    s.add_argument("--max", type=int, default=20)
    s.set_defaults(fn=cmd_search)
    b = sub.add_parser("abstract")
    b.add_argument("pmids", nargs="+")
    b.set_defaults(fn=cmd_abstract)
    c = sub.add_parser("build")
    c.add_argument("refs")
    c.add_argument("--out", required=True)
    c.add_argument("--title", default="研究計画書")
    c.add_argument("--clean", action="store_true", help="古い NN_*.ris を消してから書く")
    c.set_defaults(fn=cmd_build)
    a = ap.parse_args()
    a.fn(a)


if __name__ == "__main__":
    main()
