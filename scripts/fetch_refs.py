# -*- coding: utf-8 -*-
"""引用文献を探し、抄録で主張を確かめ、RISをダウンロードして文献リストを作る。

  1) 探す      python fetch_refs.py search "physician work style reform Japan" [--max 20]
  2) 確かめる  python fetch_refs.py abstract 39134926 20222990
  3) 作る      python fetch_refs.py build refs.json --out <この研究の文献フォルダ> --title "<計画書名>" [--clean]

連絡先は環境変数 IRB_DOCS_MAILTO で渡す（NCBI・Crossref の推奨。無くても動くが回数制限が厳しい）。
PubMed につながらないときは、PubMed から保存した .nbib（や雑誌のページの .ris）を --local で渡す（abstract と build）。
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


# ネットに出られない環境（Claude.ai のコード実行は既定で PubMed・Crossref に出られないことがある）。
# IRB_DOCS_OFFLINE=1 で最初から接続しない。接続できなかったら以後は接続を試さない。
_OFFLINE = os.environ.get("IRB_DOCS_OFFLINE", "").strip() not in ("", "0")
OFFLINE_HELP = """PubMed / Crossref に接続できない（ネットの制限がある環境）。文献を記憶で書かずに、書く人に次を頼む:
  PubMed 収載: PubMed で文献を開き「Cite」→「Download .nbib」。
    複数なら検索結果でチェックを付けて「Send to」→「Citation manager」→「Create file」（1つの .nbib にまとまる）。
  PubMed 非収載: 雑誌のページの「Download citation」などで RIS を保存。
  受け取ったファイルを --local に渡す（.nbib の AB 欄が抄録）:
    python fetch_refs.py abstract <PMID> … --local pubmed-xxx.nbib
    python fetch_refs.py build refs.json --out <文献フォルダ> --local pubmed-xxx.nbib other.ris
  検索語は Claude が提案し、PubMed での検索と選択は書く人にしてもらう。"""


class Offline(Exception):
    pass


def get(url, tries=4):
    """回数制限（429）や一時的な失敗は待って再試行する。NCBI は連絡先なしだと毎秒3回まで。"""
    global _OFFLINE
    if _OFFLINE:
        raise Offline(url)
    for k in range(tries):
        try:
            with urllib.request.urlopen(urllib.request.Request(url, headers=UA), timeout=40) as r:
                return r.read().decode("utf-8", "replace")
        except urllib.error.HTTPError as e:
            if e.code == 403:             # 出口の制限（プロキシ）で止められた
                _OFFLINE = True
                raise Offline(url) from e
            if e.code not in (429, 500, 502, 503, 504) or k == tries - 1:
                raise
        except urllib.error.URLError as e:
            if k == tries - 1:
                _OFFLINE = True
                raise Offline(url) from e
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
    try:
        ids = json.loads(get(f"{EUTILS}esearch.fcgi?db=pubmed&retmode=json&sort=relevance&retmax={a.max}&term={q}"))[
            "esearchresult"]["idlist"]
    except Offline:
        raise SystemExit(f"検索語: {a.query}\n" + OFFLINE_HELP)
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
    by_pmid, _ = load_local(a.local)
    rest = [p for p in a.pmids if p not in by_pmid]
    for p in a.pmids:
        if p in by_pmid:
            r = by_pmid[p]
            g = lambda k: (r.get(k) or [""])[0]
            print(f"PMID {p}（手元のファイル）\n{g('TI')}\n{', '.join(r.get('AU', []))}. {g('TA')} {g('DP')};{g('VI')}:{g('PG')}\n\n"
                  f"{g('AB') or '（抄録なし）'}\n")
    if not rest:
        return
    try:
        x = get(EUTILS + "efetch.fcgi?db=pubmed&rettype=abstract&retmode=text&id=" + ",".join(rest))
    except Offline:
        raise SystemExit(f"抄録を取れない PMID: {' '.join(rest)}\n{OFFLINE_HELP}")
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


# ------------------------------------------------------------------ 手元のファイル（ネットに出られないとき）
def parse_medline(txt):
    """PubMed の .nbib（MEDLINE 形式）。レコードは空行で区切られ、行は 'TAG - 値'、続きの行は空白6つで始まる。"""
    recs, cur, last = [], {}, None
    for line in norm_newlines(txt).split("\n"):
        if not line.strip():
            if cur:
                recs.append(cur)
            cur, last = {}, None
            continue
        m = re.match(r"^([A-Z]{2,4})\s*- (.*)$", line)
        if m:
            last = m.group(1)
            cur.setdefault(last, []).append(m.group(2).strip())
        elif last and line.startswith("      "):
            cur[last][-1] += " " + line.strip()
    if cur:
        recs.append(cur)
    return [r for r in recs if r.get("PMID")]


def medline_doi(r):
    for v in r.get("LID", []) + r.get("AID", []):
        if v.endswith("[doi]"):
            return v[:-5].strip()
    return None


def medline_pages(r):
    """頁。電子版だけの雑誌は PG が無く、論文番号が種類なしの LID 行にある（'LID - 2705'）。"""
    pg = (r.get("PG") or [""])[0]
    return pg or next((v for v in r.get("LID", []) if not v.endswith("]")), "")


def medline_to_ris(r):
    g = lambda k: (r.get(k) or [""])[0]
    out = ["TY  - JOUR"]
    out += [f"AU  - {x}" for x in (r.get("FAU") or r.get("AU") or [])]
    out += [f"AU  - {x}" for x in r.get("CN", [])]
    out.append(f"TI  - {g('TI')}")
    if g("JT"):
        out.append(f"T2  - {g('JT')}")
    if g("TA"):
        out.append(f"J2  - {g('TA')}")
    y = re.search(r"\d{4}", g("DP"))
    if y:
        out.append(f"PY  - {y.group(0)}")
    if g("VI"):
        out.append(f"VL  - {g('VI')}")
    if g("IP"):
        out.append(f"IS  - {g('IP')}")
    pg = medline_pages(r)
    if complex_pages(pg):
        out.append(f"SP  - {pg}")
    elif pg:
        sp, _, ep = nlm_pages(pg).partition("–")
        out.append(f"SP  - {sp}")
        if ep:
            out.append(f"EP  - {ep}")
    if g("AB"):
        out.append(f"AB  - {g('AB')}")
    if medline_doi(r):
        out.append(f"DO  - {medline_doi(r)}")
    out.append(f"AN  - {g('PMID')}")
    out.append("ER  - ")
    return "\n".join(out)


def load_local(paths):
    """--local のファイル（.nbib／MEDLINE 形式の .txt、.ris）を PMID と DOI で引けるようにする。"""
    by_pmid, by_doi = {}, {}
    for p in paths or []:
        txt = norm_newlines(Path(p).read_text(encoding="utf-8-sig", errors="replace"))
        if re.search(r"^PMID- ", txt, re.M):
            for r in parse_medline(txt):
                by_pmid[r["PMID"][0]] = r
                if medline_doi(r):
                    by_doi[medline_doi(r).lower()] = ("medline", r)
        for block in re.findall(r"^TY  - .*?^ER  -[^\n]*", txt, re.M | re.S):
            rec = parse_ris(block)
            d = (rec.get("DO") or [""])[0].strip()
            if d:
                by_doi.setdefault(d.lower(), ("ris", block))
    return by_pmid, by_doi


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


def fetch_item(it, summ, by_pmid, by_doi):
    """1件の RIS と、整形に要るもの（取得元・手直し・頁・DOI・著者名）。ネットに出られなければ Offline を投げる。"""
    fixes, pages, doi, names = [], None, it.get("doi"), None
    pmid = it.get("pmid")
    if it.get("manual_ris"):
        return norm_newlines(it["manual_ris"]), "manual", ["RIS が取得できないため手作り"], None, doi, None
    local = by_pmid.get(pmid) if pmid else None
    if local is None and doi and not pmid:
        kind, v = by_doi.get(doi.lower(), (None, None))
        if kind == "ris":
            return norm_newlines(v), "ris-file", [], None, doi, None
        local = v if kind == "medline" else None
    if local is not None:                  # PubMed から保存した .nbib（MEDLINE 形式）
        pg = medline_pages(local)
        return (medline_to_ris(local), "pubmed-nbib", [], nlm_pages(pg) or None,
                doi or medline_doi(local), local.get("AU", []) + local.get("CN", []))
    if pmid:
        ris = norm_newlines(get(CTXP + pmid))
        d = summ.get(pmid)
        if not d or d.get("error"):
            raise SystemExit(f"PMID {pmid} が PubMed に無い（番号を確かめる）")
        doi = doi or doi_of(d)
        pages = nlm_pages(d.get("pages")) or None
        names = [x["name"] for x in d.get("authors", []) if x.get("name")]
        if complex_pages(d.get("pages")):
            # Citation Exporter は複合頁の終頁を壊す（330.e1-6 → 330.e3306、869-75.e1-2 → 75.e752）ので PubMed 本体の表記に置換
            ris = re.sub(r"\nSP  - [^\n]*", lambda _: "\nSP  - " + d["pages"], ris, count=1)
            ris = re.sub(r"\nEP  - [^\n]*", "", ris, count=1)
            fixes.append(f"頁欄を PubMed 本体の表記（{d['pages']}）に置換")
        return ris, "pubmed", fixes, pages, doi, names
    q = urllib.parse.quote(doi, safe="")
    ris = norm_newlines(get(f"https://api.crossref.org/works/{q}/transform/application/x-research-info-systems"))
    if not re.search(r"\nSP  - ", ris):
        art = json.loads(get(f"https://api.crossref.org/works/{q}"))["message"].get("article-number")
        if art:
            ris = re.sub(r"\nER  -", lambda _: f"\nSP  - {art}\nER  -", ris, count=1)
            pages = art
            fixes.append(f"頁欄に論文番号（{art}）を追加")
    return ris, "crossref", fixes, pages, doi, None


def cmd_build(a):
    items = json.loads(Path(a.refs).read_text(encoding="utf-8"))
    out = Path(a.out)
    by_pmid, by_doi = load_local(a.local)
    need = [it["pmid"] for it in items if it.get("pmid") and not it.get("manual_ris") and it["pmid"] not in by_pmid]
    summ = {}
    try:
        summ = esummary(need)
    except Offline:
        pass
    # 先に全部そろえる。1件でも取れなければ何も書き出さない（--clean で古いファイルを消してしまわない）
    got, missing = [], []
    for it in items:
        try:
            got.append(fetch_item(it, summ, by_pmid, by_doi))
            if got[-1][1] in ("pubmed", "crossref"):
                time.sleep(0.4)
        except Offline:
            missing.append(it)
    if missing:
        lst = "\n".join(f"  {it['key']}: " + (f"PMID {it['pmid']}" if it.get("pmid") else f"DOI {it.get('doi')}")
                        for it in missing)
        raise SystemExit(f"次の{len(missing)}件を取れなかった（何も書き出していない）:\n{lst}\n{OFFLINE_HELP}")
    out.mkdir(parents=True, exist_ok=True)
    if a.clean:
        for old in out.glob("[0-9][0-9]_*.ris"):
            old.unlink()
    meta, combined = [], []
    for no, (it, (ris, src, fixes, pages, doi, names)) in enumerate(zip(items, got), 1):
        key = f"{no:02d}_{it['key']}"
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
    write_ris(out / "references_all.ris", "\n\n".join(combined))
    (out / "references_meta.json").write_text(json.dumps(meta, ensure_ascii=False, indent=1), encoding="utf-8")
    cnt = lambda s: sum(m["source"] == s for m in meta)
    parts = []
    if cnt("pubmed"):
        parts.append(f"PubMed 収載の{cnt('pubmed')}件は NCBI の Literature Citation Exporter から RIS をダウンロードした")
    if cnt("crossref"):
        parts.append(f"PubMed 非収載の{cnt('crossref')}件は Crossref から RIS をダウンロードした")
    if cnt("pubmed-nbib"):
        parts.append(f"{cnt('pubmed-nbib')}件は PubMed から保存した .nbib（MEDLINE 形式）を RIS に変換した")
    if cnt("ris-file"):
        parts.append(f"{cnt('ris-file')}件は雑誌のページ等から保存した RIS を使った")
    how = "。".join(parts) + "。" if parts else ""
    if cnt("manual"):
        how += f"RIS が取得できない{cnt('manual')}件は手で書いた（下の手直しの一覧を参照）。"
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
    loc = "".join(f' "{Path(p).resolve()}"' for p in (a.local or []))
    lines += ["", "## 取り直し方", "", "```bash",
              f'python "{Path(__file__).resolve()}" build "{Path(a.refs).resolve()}" --out "{out.resolve()}" '
              f'--title "{a.title}" --clean' + (f" --local{loc}" if loc else ""),
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
    b.add_argument("--local", nargs="*", default=[], help="PubMed から保存した .nbib など（ネットに出られないとき）")
    b.set_defaults(fn=cmd_abstract)
    c = sub.add_parser("build")
    c.add_argument("refs")
    c.add_argument("--out", required=True)
    c.add_argument("--title", default="研究計画書")
    c.add_argument("--clean", action="store_true", help="古い NN_*.ris を消してから書く")
    c.add_argument("--local", nargs="*", default=[], help="手元の .nbib（PubMed から保存）や .ris。ここにある文献はネットで取らない")
    c.set_defaults(fn=cmd_build)
    a = ap.parse_args()
    a.fn(a)


if __name__ == "__main__":
    main()
