# -*- coding: utf-8 -*-
"""倫理審査のWordひな形に転記するための python-docx ヘルパー。

使い方（生成スクリプトの先頭で）:
    import sys; from pathlib import Path
    sys.path.insert(0, str(Path.home() / ".claude" / "skills" / "irb-docs" / "scripts"))
    from docx_tools import *

基本の流れ:
    doc = Document(ひな形)
    keep = protect(doc, [(r"^ここに研究課題名を記載$", 題名, 14), ...])   # 表紙などを先に書き換えて守る
    cut_before(doc, r"^表紙の最初の行$")   # 冒頭の「使用上の注意」ページを消す
    purge_guidance(doc, keep, mode=...)   # 赤字（注意書き）・青字（例文）・＊行を消す。見出しは残す
    purge_examples(doc, keep)             # 黒字に混ざった例文（（例、←、○○…）と例示の表を消す
    purge_tables(doc)                     # 表の中の色と「例）」の行
    ins = Inserter(doc, style_anchor=r"^本研究に携わる全ての関係者は", size=11)
    ins.after(r"^1．研究の背景$", ["本文…{{1}}。", ...])   # {{n}} は上付きの引用番号になる
    ins.report()                          # 見つからなかった見出しを表示（黙って飛ばさない）
    finalize(doc)                         # 先頭の白紙・ヘッダー/フッターの例示を片付ける
    safe_save(doc, 出力先)                # 人が Word で直した版は上書きしない

mode の選び方（ひな形の作りで決める。dump_template.py --runs で確かめる）:
    "paragraph"  注意書き・例文が段落ごと色分けされているひな形（北大 HT1/HT2）。
                 黒字の段落に混ざった色字は黒に戻すだけ（あとで replace_words で書き換える前提）。
    "inline"     1行に黒字のラベルと色字の値が混ざるひな形（北大 HT3 など）。
                 赤字は消し、青字は選択肢（「／…」や choice_words）なら消し、それ以外は【　　】にする。

位置は必ず見出しの文字（正規表現）で探す。段落の番号（索引）は削除のたびにずれるので使わない。
"""
import copy
import re
import unicodedata

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor
from docx.table import Table
from docx.text.paragraph import Paragraph

# 北大（2024年12月版・2025年4月版）: 赤=記載上の注意、青=例文、黒=定型文。
# 機関によっては緑・紫も手引き（注意書き・例文）の色に使うので、既定値に含めてある。
RED = {"FF0000", "C00000", "CC0000"}
BLUE = {"0000FF", "0432FF", "0042AA", "0070C0"}
GREEN_ETC = {"00B050", "70AD47", "7030A0"}
GUIDE_COLORS = RED | BLUE | GREEN_ETC

SLOT = "【　　】"
PLACEHOLDER = "（本文を記載）"
HEADING_RX = re.compile(r"^(\d+[．.]\s*\S|（\d+）|\(\d+\)|【)")        # 旧版との互換用。判定は heading_level を使う
EXAMPLE_PAT = ("（例", "(例", "←", "○○", "〇〇", "◯◯", "○歳", "○例", "〇例", "○年", "○か月",
               "△△", "××", "20XX", "XX年", "＊", "※", "◇", "■")
EXAMPLE_TABLE_KEYS = ("ヶ月間", "許容範囲", "Day", "調査開始前", "観察開始前")
EXAMPLE_ROW_HEADS = ("例）", "例)", "（例", "(例")
# 「＊」で始まっても消してはいけない定型文（オプトアウトの拒否の機会など）
KEEP_RX = re.compile(r"ご了解いただけない場合|利用を拒否|利用の停止を求め|ご連絡ください。?$")
CITE_RX = re.compile(r"\{\{([^}]*)\}\}")
PARTICLE_RX = re.compile(r"[ぁ-ん、・]{1,2}")      # 青字と青字の間の「に、」などは青字の一部として扱う


# ------------------------------------------------------------------ 基本
def blocks(doc):
    """本文直下の段落と表を文書順に返す。"""
    out = []
    for ch in doc.element.body.iterchildren():
        t = ch.tag.split("}")[-1]
        if t == "p":
            out.append(Paragraph(ch, doc._body))
        elif t == "tbl":
            out.append(Table(ch, doc._body))
    return out


def all_paragraphs(doc):
    """本文と表の中の段落をすべて返す。"""
    for b in blocks(doc):
        if isinstance(b, Paragraph):
            yield b
        else:
            for row in b.rows:
                for c in distinct_cells(row):
                    yield from c.paragraphs


def find(bl, pattern, start=0):
    rx = re.compile(pattern)
    for i in range(start, len(bl)):
        b = bl[i]
        if isinstance(b, Paragraph) and rx.search(b.text.strip()):
            return i
    raise KeyError(pattern)


def find_all(bl, pattern):
    rx = re.compile(pattern)
    return [i for i, b in enumerate(bl) if isinstance(b, Paragraph) and rx.search(b.text.strip())]


def rcolor(r):
    rp = r._r.rPr
    if rp is None:
        return None
    c = rp.find(qn("w:color"))
    return c.get(qn("w:val")).upper() if c is not None and c.get(qn("w:val")) else None


def style_color(p):
    """段落スタイルに付いた文字色（ランに色が無いときの見た目の色）。"""
    st = p.style
    while st is not None:
        try:
            c = st.font.color
            if c is not None and c.type is not None and c.rgb is not None:
                return str(c.rgb).upper()
        except (AttributeError, ValueError):
            pass
        st = st.base_style
    return None


def run_class(r, p=None):
    """'red' / 'blue'（青・緑・紫を含む手引きの色）/ 'black'。ランに色が無ければ段落スタイルの色を見る。"""
    c = rcolor(r)
    if c is None and p is not None:
        c = style_color(p)
    if c in RED:
        return "red"
    if c in BLUE | GREEN_ETC:
        return "blue"
    return "black"


def clear_marks(r):
    rp = r._r.rPr
    if rp is None:
        return
    for tag in ("w:color", "w:highlight"):
        el = rp.find(qn(tag))
        if el is not None:
            rp.remove(el)


def clear_mark_color(p):
    """段落記号の色を落とす。残すと、その行の末尾で改行したとき青字で入力される。"""
    ppr = p._p.pPr
    rpr = None if ppr is None else ppr.find(qn("w:rPr"))
    if rpr is None:
        return
    col = rpr.find(qn("w:color"))
    if col is not None and (col.get(qn("w:val")) or "").upper() in GUIDE_COLORS:
        rpr.remove(col)
    hl = rpr.find(qn("w:highlight"))
    if hl is not None:
        rpr.remove(hl)


def remove(obj):
    el = obj._element if hasattr(obj, "_element") else obj
    if el.getparent() is not None:
        el.getparent().remove(el)


def set_font(r, font):
    r.font.name = font
    rpr = r._r.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rpr.insert(0, rf)
    for a in ("w:eastAsia", "w:ascii", "w:hAnsi"):
        rf.set(qn(a), font)


def set_text(p, text, size=None, font=None):
    """段落の文字を置き換える（段落書式は残し、ランは作り直す）。"""
    for ch in list(p._p):
        if ch.tag != qn("w:pPr"):
            p._p.remove(ch)
    _add_runs(p, text, size, font)
    return p


def _add_runs(p, text, size=None, font=None):
    """{{n}} を上付きの引用番号にしながらランを足す。"""
    pos = 0
    for m in list(CITE_RX.finditer(text)) + [None]:
        seg = text[pos:(m.start() if m else len(text))]
        if seg:
            r = p.add_run(seg)
            if size:
                r.font.size = Pt(size)
            if font:
                set_font(r, font)
        if m:
            r = p.add_run(m.group(1))
            r.font.superscript = True
            if size:
                r.font.size = Pt(size)
            if font:
                set_font(r, font)
            pos = m.end()


def insert_after(anchor_p, lines, ref_p=None, size=None, font=None, parent=None):
    """anchor_p の直後に段落を足す。ref_p の段落書式（字下げ等）を写す。"""
    last = anchor_p._p
    ppr = copy.deepcopy(ref_p._p.pPr) if (ref_p is not None and ref_p._p.pPr is not None) else None
    made = []
    for line in lines:
        np_ = OxmlElement("w:p")
        if ppr is not None:
            np_.append(copy.deepcopy(ppr))
        last.addnext(np_)
        p = Paragraph(np_, parent if parent is not None else anchor_p._parent)
        _add_runs(p, line, size, font)
        clear_mark_color(p)
        last = np_
        made.append(p)
    return made


def has_page_break(p):
    return any(br.get(qn("w:type")) == "page" for br in p._p.iter(qn("w:br")))


# ------------------------------------------------------------------ 見出し
_WIDE = str.maketrans("０１２３４５６７８９．（）", "0123456789.()")


def _narrow(t):
    """全角の数字・ピリオド・括弧だけを半角にする（丸数字①は変えない。NFKC だと①が1になる）。"""
    return t.translate(_WIDE)


def heading_level(t):
    """見出しの深さ。0 は見出しでない。
    1「1．」「1.」 / 2「4.1」 / 2.5「（1）」 / 3「6.4.1」「【…】」「○…」 / 2「[…]」（同じ行に値があれば 0）
    深さは「親の直後に子が来ているか」の判定だけに使う（本文のない見出しの検出）。"""
    t = t.strip()
    if not t or len(t) >= 60 or "\t" in t or t.endswith("。"):
        return 0
    n = _narrow(t)
    m = re.match(r"^(\d+(?:\.\d+)+)\.?\s*[^\d\s.]", n)          # 4.1 / 6.4.1
    if m:
        return float(m.group(1).count(".") + 1)
    if re.match(r"^\d+\.\s*[^\d\s.]", n):                        # 4．研究の方法 / 4. 研究の方法
        return 1.0
    if re.match(r"^\(\d+\)", n):
        return 2.5
    if t.startswith("【") and t.endswith("】"):
        return 3.0
    if t.startswith("["):
        close = t.find("]")
        if close > 0:
            return 0.0 if t[close + 1:].strip() else 2.0
    if re.match(r"^[○〇◯](?![○〇◯])\S", t) and len(t) < 40:
        return 3.0
    return 0.0


def is_heading(t):
    return heading_level(t) > 0 or (t.strip().startswith("[") and "]" in t and len(t.strip()) < 60)


def sanitize_heading(p):
    """見出しに付いた注意書き（＊以降、（←…）、（…場合に記載）など）を落とし、色を黒にする。"""
    t = p.text
    t2 = re.split(r"＊|\*", t)[0]
    t2 = re.sub(r"（←[^）]*）|（[^）]*場合に記載）|（[^）]*場合、[^）]*）|（[^）]*がある場合）", "", t2).rstrip()
    if t2 != t and t2.strip():
        set_text(p, t2)
    for r in p.runs:
        clear_marks(r)
    clear_mark_color(p)


# ------------------------------------------------------------------ 色の片付け（行の中）
def segments(p):
    """[['red'|'blue'|'black', [ラン…]], …]。同じ色の連続はまとめ、青字に挟まれた助詞・読点は青字に含める。"""
    segs = []
    for r in p.runs:
        if not r.text:
            continue
        c = run_class(r, p)
        if segs and segs[-1][0] == c:
            segs[-1][1].append(r)
        else:
            segs.append([c, [r]])
    merged, i = [], 0
    while i < len(segs):
        c, rs = segs[i]
        if (c == "black" and merged and merged[-1][0] == "blue" and i + 1 < len(segs)
                and segs[i + 1][0] == "blue" and PARTICLE_RX.fullmatch("".join(r.text for r in rs))):
            merged[-1][1].extend(rs + segs[i + 1][1])
            i += 2
            continue
        merged.append([c, list(rs)])
        i += 1
    return merged


def apply_inline(p, choice_words=(), slot=SLOT, recolor_only=False):
    """行の中の色字を片付ける。赤字は消す。青字は、選択肢（「／…」で始まる、または choice_words）なら消し、
    それ以外は slot（【　　】）に置き換える。recolor_only=True なら青字は黒に戻すだけ。変更の記録を返す。"""
    log = []
    for c, rs in segments(p):
        text = "".join(r.text for r in rs)
        if c == "black":
            continue
        if c == "red":
            for r in rs:
                remove(r)
            log.append(("赤字を削除", text))
        elif recolor_only or not text.strip():
            for r in rs:
                clear_marks(r)
        elif text.startswith("／") or text in choice_words:
            for r in rs:
                remove(r)
            log.append(("青字の選択肢を削除", text))
        else:
            rs[0].text = slot
            clear_marks(rs[0])
            for r in rs[1:]:
                remove(r)
            log.append(("青字を空欄に", text))
    clear_mark_color(p)
    if log:
        for r in reversed(p.runs):              # 末尾の空白を整える
            if r.text.strip():
                r.text = r.text.rstrip()
                break
    return log


# ------------------------------------------------------------------ 削除
def protect(doc, items):
    """[(正規表現, 新しい文字, 文字サイズ)] を書き換え、その段落を以後の削除から守る。"""
    keep = set()
    bl = blocks(doc)
    for pat, txt, size in items:
        p = bl[find(bl, pat)]
        set_text(p, txt, size)
        for r in p.runs:
            clear_marks(r)
        clear_mark_color(p)
        keep.add(id(p._p))
    return keep


def purge_guidance(doc, keep_ids=(), mode="paragraph", choice_words=(), keep_rx=KEEP_RX, log=None):
    """赤字だけ・青字だけの段落と「＊」で始まる段落を削除する。見出しは消さずに色と注意書きだけ落とす
    （消すと「該当なし」を書く場所がなくなる）。keep_rx に合う段落（拒否の機会の定型文など）は消さない。
    mode は冒頭の説明を参照。log にリストを渡すと、消した段落と書き換えた段落を記録する。"""
    n = 0
    for b in blocks(doc):
        if not isinstance(b, Paragraph) or id(b._p) in keep_ids:
            continue
        t = b.text.strip()
        if not t:
            clear_mark_color(b)
            continue
        kinds = {c for c, rs in segments(b) if "".join(r.text for r in rs).strip()}
        if is_heading(t):
            if mode == "inline" and "black" in kinds:
                apply_inline(b, choice_words)          # 黒字の見出しに付いた色字（注意書き・例）だけ片付ける
            sanitize_heading(b)
            continue
        if keep_rx is not None and keep_rx.search(t):
            apply_inline(b, choice_words, recolor_only=(mode == "paragraph"))
            for r in b.runs:
                clear_marks(r)
            for r in b.runs:                         # 先頭の「＊」は注意の印なので外す（文は残す）
                if r.text:
                    if r.text.startswith(("＊", "*")):
                        r.text = r.text[1:]
                    break
            if log is not None:
                log.append(("残した定型文", t))
            continue
        if t.startswith(("＊", "*")) or "black" not in kinds:
            if log is not None:
                log.append(("削除", "赤字" if kinds == {"red"} else ("青字" if "black" not in kinds else "＊行"), t))
            remove(b)
            n += 1
            continue
        before = b.text
        apply_inline(b, choice_words, recolor_only=(mode == "paragraph"))
        if log is not None and b.text != before:
            log.append(("書き換え", before.strip(), b.text.strip()))
    return n


def purge_examples(doc, keep_ids=(), patterns=EXAMPLE_PAT, table_keys=EXAMPLE_TABLE_KEYS, log=None):
    """色では拾えない例文の行（黒字の「（例」「○歳」など）と、例示の表を削除する。
    mode="inline" で【　　】にした後なら、○○ は残っていない。"""
    n = 0
    for b in blocks(doc):
        if isinstance(b, Table):
            txt = " ".join(c.text for row in b.rows for c in row.cells)
            if table_keys and any(k in txt for k in table_keys):
                remove(b)
                n += 1
            continue
        if id(b._p) in keep_ids:
            continue
        t = b.text.strip()
        if is_heading(t):
            sanitize_heading(b)
            continue
        if KEEP_RX.search(t):
            continue
        if t and any(k in t for k in patterns):
            if log is not None:
                log.append(("削除", "黒字の例文", t))
            remove(b)
            n += 1
    return n


def purge_tables(doc, mode="recolor", choice_words=(), row_heads=EXAMPLE_ROW_HEADS, log=None):
    """表の中の片付け。「例）」で始まる行を消し、色字は mode="recolor" なら黒に戻す（表の見出し行が
    青字のひな形があるため）、mode="inline" なら apply_inline と同じ扱いにする。"""
    n = 0
    for b in blocks(doc):
        if not isinstance(b, Table):
            continue
        for row in list(b.rows):
            first = distinct_cells(row)[0].text.strip()
            if first.startswith(row_heads):
                if log is not None:
                    log.append(("表の例示行を削除", " | ".join(c.text.strip() for c in distinct_cells(row))))
                row._tr.getparent().remove(row._tr)
                n += 1
        for row in b.rows:
            for c in distinct_cells(row):
                for p in c.paragraphs:
                    for r in list(p.runs):
                        if run_class(r, p) == "red":
                            remove(r)
                    apply_inline(p, choice_words, recolor_only=(mode == "recolor"))
    return n


def delete_between(doc, start_pat, end_pat, keep=()):
    """start の段落と end の段落の間を削除する（keep の文字列で始まる段落は残す）。
    挿入した本文の後ろに残ったひな形の空見出し・空ラベルを片付けるのに使う。"""
    bl = blocks(doc)
    i = find(bl, start_pat)
    j = find(bl, end_pat, i + 1)
    n = 0
    for b in bl[i + 1:j]:
        t = b.text.strip() if isinstance(b, Paragraph) else ""
        if any(t.startswith(k) for k in keep):
            continue
        remove(b)
        n += 1
    return n


def cut_before(doc, pattern):
    """pattern の段落より前（ひな形の使用上の注意ページ等）を削除する。セクション区切りは残す。"""
    bl = blocks(doc)
    i = find(bl, pattern)
    n = 0
    for b in bl[:i]:
        el = b._element if hasattr(b, "_element") else b._p
        if el.findall(".//" + qn("w:sectPr")):
            continue
        remove(b)
        n += 1
    return n


def cut_from(doc, pattern):
    """pattern の段落以降（巻末の用語集や、使わない方の版のページ）を削除し、末尾の空行・改ページも消す。"""
    bl = blocks(doc)
    for b in bl[find(bl, pattern):]:
        remove(b)
    drop_trailing_blanks(doc)


def drop_trailing_blanks(doc):
    """文書末の空段落（改ページだけの段落を含む）を消す。残すと最後に白紙のページが出る。"""
    while True:
        bl = blocks(doc)
        if not bl or isinstance(bl[-1], Table) or bl[-1].text.strip():
            return
        if bl[-1]._p.findall(".//" + qn("w:sectPr")) or bl[-1]._p.findall(".//" + qn("w:txbxContent")):
            return
        remove(bl[-1])


def replace_words(doc, mapping):
    """ランの境界をまたぐ語も置換する（段落の文字を組み立て直すので、段落内の書式は先頭ランに揃う）。"""
    for p in all_paragraphs(doc):
        t = p.text
        if not t:
            continue
        t2 = t
        for a, c in mapping.items():
            t2 = t2.replace(a, c)
        if t2 != t and p.runs:
            p.runs[0].text = t2
            for r in p.runs[1:]:
                r.text = ""


# ------------------------------------------------------------------ 挿入
class Inserter:
    """見出しの文字で位置を探して本文を足す。見つからない見出しは記録して最後に報告する。
    size=None なら文字サイズは段落の書式に任せる（ひな形ごとに本文の大きさが違う。HT1 は11pt、HT3 は12pt）。"""

    def __init__(self, doc, style_anchor=None, size=None, font=None):
        self.doc, self.size, self.font = doc, size, font
        self.missing = []
        self.ref = None
        if style_anchor:
            bl = blocks(doc)
            self.ref = bl[find(bl, style_anchor)]

    def after(self, pat, lines, all_matches=False, ref=None):
        """all_matches=True なら一致した見出しすべての後ろに足す（①単独版と②多機関版のページが並ぶひな形など）。"""
        bl = blocks(self.doc)
        idx = find_all(bl, pat)
        if not idx:
            self.missing.append(("after", pat))
            return []
        made = []
        for i in (idx if all_matches else idx[:1]):
            made += insert_after(bl[i], lines, ref or self.ref, self.size, self.font, parent=self.doc._body)
        return made

    def replace(self, pat, text, size=None):
        bl = blocks(self.doc)
        try:
            i = find(bl, pat)
        except KeyError:
            self.missing.append(("replace", pat))
            return None
        return set_text(bl[i], text, size or self.size)

    def report(self):
        if self.missing:
            print("!! 見つからなかった見出し:")
            for m in self.missing:
                print("   ", m)
        return self.missing


def placeholders_for_empty_headings(doc, text=PLACEHOLDER, size=None):
    """本文のない見出し（次がより浅いか同じ深さの見出し、表でない、ページの終わり）の直後に text を置く。
    下書きを作るとき用。置いた見出しの一覧を返す。提出前に audit_docs.py が（本文を記載）の残りを拾う。"""
    done = []
    bl = blocks(doc)
    for i, b in enumerate(bl):
        if not isinstance(b, Paragraph):
            continue
        t = b.text.strip()
        lv = heading_level(t)
        if not lv:
            continue
        nxt = None
        for nb in bl[i + 1:]:
            if isinstance(nb, Table):
                nxt = "table"
                break
            if nb.text.strip():
                nxt = heading_level(nb.text)
                break
            if has_page_break(nb):
                break
        if nxt == "table" or (nxt is not None and (nxt == 0 or nxt > lv)):
            continue
        insert_after(b, [text], None, size, parent=doc._body)
        done.append(t)
    return done


def left_align(paragraphs):
    """英文の文献リストなど。両端揃えのままだと語間が不自然に広がる。"""
    for p in paragraphs:
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT


# ------------------------------------------------------------------ 表・チェック欄
def distinct_cells(row):
    """結合セルは row.cells に同じセルが繰り返し出るので、重複を除いて返す。"""
    seen, out = set(), []
    for c in row.cells:
        if id(c._tc) in seen:
            continue
        seen.add(id(c._tc))
        out.append(c)
    return out


def value_cell(row):
    """ラベル｜値 の行で、値を書くセル（右端の別セル）を返す。"""
    return distinct_cells(row)[-1]


def put(cell, text, size=10.5, font=None):
    cell.text = ""
    r = cell.paragraphs[0].add_run(text)
    r.font.size = Pt(size)
    if font:
        set_font(r, font)


def tick(cell, label, mark="☑"):
    """「□はい」の□を☑にする。□と語が別のランに分かれていても動く。"""
    target = "□" + label
    for p in cell.paragraphs:
        if target not in p.text:
            continue
        pos = p.text.index(target)
        off = 0
        for r in p.runs:
            nxt = off + len(r.text)
            if off <= pos < nxt:
                i = pos - off
                r.text = r.text[:i] + mark + r.text[i + 1:]
                return True
            off = nxt
    return False


# ------------------------------------------------------------------ ヘッダー・フッター・表紙の枠
def chrome_parts(doc):
    """定義のあるヘッダー/フッターだけを返す。定義の無い部品（前のセクションと同じ扱い）に触ると、
    python-docx が空の部品を新しく作ってしまうので触らない。"""
    out = []
    for si, sec in enumerate(doc.sections):
        for name in ("header", "footer", "first_page_header", "first_page_footer",
                     "even_page_header", "even_page_footer"):
            part = getattr(sec, name)
            if part.is_linked_to_previous:
                continue
            out.append((f"s{si}.{name}", part))
    return out


def clean_chrome(doc, mode="inline", choice_words=()):
    """ヘッダー/フッターの例示を片付ける。赤字は消し、青字は【　　】（mode="recolor" なら黒に戻すだけ）、
    ○○ や 20XX のような黒字の例示も【　　】にする。変更の数を返す。"""
    n = 0
    for _, part in chrome_parts(doc):
        for p in part.paragraphs:
            before = p.text
            apply_inline(p, choice_words, recolor_only=(mode == "recolor"))
            for r in p.runs:
                if re.search(r"[〇○◯]+|X{2,}", r.text):
                    r.text = re.sub(r"[〇○◯]+|X{2,}", SLOT, r.text)
                if run_class(r, p) != "black":
                    clear_marks(r)
            n += p.text != before
    return n


def blacken_chrome(doc):
    """ヘッダー/フッターの色字を黒に戻す（文字は変えない）。旧版との互換用。"""
    n = 0
    for _, part in chrome_parts(doc):
        for par in part.paragraphs:
            for r in par.runs:
                if r.font.color is not None and r.font.color.type is not None:
                    r.font.color.rgb = RGBColor(0, 0, 0)
                    n += 1
    return n


def drop_empty_title_page(doc):
    """「先頭ページのみ別指定」で先頭ページのヘッダーが空なら、その指定を外す。
    冒頭の注意ページを消すと、本文の1ページ目に空のヘッダーが出て作成日・版数が消えるため。
    表紙を残すひな形（HT1 など）では呼ばない。"""
    n = 0
    for sec in doc.sections:
        if not sec.different_first_page_header_footer:
            continue
        first = sec.first_page_header
        if first.is_linked_to_previous or not any(p.text.strip() for p in first.paragraphs):
            sec.different_first_page_header_footer = False
            n += 1
    return n


def fix_textboxes(doc, repl):
    """テキストボックス（表紙の期間の枠など）の段落を置換/削除する。
    repl: {元の文字列の一部: 新しい文字列 または None(削除)}"""
    for tb in doc.element.body.findall(".//" + qn("w:txbxContent")):
        for p in list(tb.findall(qn("w:p"))):
            txt = "".join(n.text or "" for n in p.findall(".//" + qn("w:t")))
            for key, new in repl.items():
                if key not in txt:
                    continue
                if new is None:
                    p.getparent().remove(p)
                else:
                    ts = p.findall(".//" + qn("w:t"))
                    if ts:
                        ts[0].text = new
                        for t2 in ts[1:]:
                            t2.text = ""
                    for r in p.findall(".//" + qn("w:r")):
                        rpr = r.find(qn("w:rPr"))
                        if rpr is not None:
                            for tag in ("w:color", "w:highlight"):
                                el = rpr.find(qn(tag))
                                if el is not None:
                                    rpr.remove(el)
                break


def set_footer_text(doc, key, new):
    for _, part in chrome_parts(doc):
        for p in part.paragraphs:
            if key in p.text and p.runs:
                p.runs[0].text = new
                for r in p.runs[1:]:
                    r.text = ""


def drop_leading_blanks(doc):
    """冒頭の空段落と、最初の本文段落に残った改ページを消す（白紙の1ページ目対策）。"""
    body = doc.element.body
    for ch in list(body.iterchildren()):
        tag = ch.tag.split("}")[-1]
        if tag == "bookmarkEnd":
            continue
        if tag != "p":
            break
        p = Paragraph(ch, doc._body)
        if p.text.strip():
            for br in ch.findall(".//" + qn("w:br")):
                if br.get(qn("w:type")) == "page":
                    br.getparent().remove(br)
            break
        if ch.findall(".//" + qn("w:txbxContent")) or ch.findall(".//" + qn("w:sectPr")):
            break                                   # 表紙の枠やセクション区切りを抱えた段落は残す
        body.remove(ch)


def unify_list_fonts(doc, start_pat, end_pat, pt=11, item_rx=r"^\d+\."):
    """start〜end の間の番号付き行の文字サイズと字体を揃える（同意文書の説明項目一覧など）。"""
    on, rows = False, []
    for b in blocks(doc):
        if not isinstance(b, Paragraph):
            continue
        t = b.text.strip()
        if re.match(start_pat, t):
            on = True
            continue
        if on and re.match(end_pat, t):
            break
        if on and re.match(item_rx, t):
            rows.append(b)
    donor = None
    for b in rows:
        for r in b.runs:
            rpr = r._element.rPr
            f = None if rpr is None else rpr.find(qn("w:rFonts"))
            if f is not None:
                donor = f
                break
        if donor is not None:
            break
    n = 0
    for b in rows:
        for r in b.runs:
            r.font.size = Pt(pt)
            n += 1
            if donor is None:
                continue
            rpr = r._element.get_or_add_rPr()
            f = rpr.find(qn("w:rFonts"))
            if f is None:
                rpr.insert(0, copy.deepcopy(donor))
            else:
                for k, v in donor.attrib.items():
                    f.set(k, v)
    return n


def finalize(doc, chrome_mode="recolor", cover_removed=False):
    """仕上げ。先頭の白紙を消し、ヘッダー/フッターを片付ける。
    cover_removed=True（冒頭の注意ページや表紙を消した）なら、空の「先頭ページのみ別指定」も外す。"""
    drop_leading_blanks(doc)
    n = clean_chrome(doc, mode=chrome_mode)
    if cover_removed:
        drop_empty_title_page(doc)
    return n


# ------------------------------------------------------------------ PI が Word で整えた空行の配置を写す
def _is_plain_blank(el):
    """文字も改ページもセクション区切りも図も無い、ただの空段落か。"""
    if el.tag != qn("w:p"):
        return False
    if "".join(t.text or "" for t in el.iter(qn("w:t"))).strip():
        return False
    for tag in ("w:sectPr", "w:br", "w:drawing", "w:pict", "w:txbxContent", "w:object"):
        if el.find(".//" + qn(tag)) is not None:
            return False
    return True


def _layout_key(el):
    if el.tag == qn("w:tbl"):
        tr = el.find(".//" + qn("w:tr"))
        return "表:" + ("".join(t.text or "" for t in tr.iter(qn("w:t")))[:40] if tr is not None else "")
    t = "".join(x.text or "" for x in el.iter(qn("w:t"))).strip()
    if t:
        return t
    return "特殊:区切り" if el.find(".//" + qn("w:sectPr")) is not None else "特殊:その他"


def _layout_profile(body):
    items, run = [], []
    for el in list(body.iterchildren()):
        if el.tag not in (qn("w:p"), qn("w:tbl")):
            continue
        if _is_plain_blank(el):
            run.append(el)
            continue
        items.append((el, _layout_key(el), run))
        run = []
    return items


def apply_blank_layout(doc, ref_path, skip=()):
    """ref_path（PI が Word で空行を足し引きして整えた版）と同じ文の段落の前に、同じ数の空行を置く。
    表紙の位置合わせやブロックの間の空行を、作り直した版に写すのに使う。skip で始まる段落は触らない
    （改ページで新しいページから始める段落など）。変えた箇所の数を返す。"""
    from docx import Document
    table = {}
    for _, k, run in _layout_profile(Document(str(ref_path)).element.body):
        table.setdefault(k, []).append(run)
    n = 0
    for el, k, run in _layout_profile(doc.element.body):
        if any(k.startswith(x) for x in skip) or not table.get(k):
            continue
        want = table[k].pop(0)
        diff = len(want) - len(run)
        if diff < 0:
            for b in run[:-diff]:
                b.getparent().remove(b)
            n += 1
        elif diff > 0:
            for b in want[len(want) - diff:]:
                el.addprevious(copy.deepcopy(b))
            n += 1
    return n


def new_page_before(doc, pattern):
    """pattern の段落を新しいページから始める（空行を重ねてページを送るより崩れにくい）。前の空行は消す。"""
    n = 0
    for b in blocks(doc):
        if isinstance(b, Paragraph) and re.search(pattern, b.text.strip()):
            b.paragraph_format.page_break_before = True
            prev = b._p.getprevious()
            while prev is not None and _is_plain_blank(prev):
                nxt = prev.getprevious()
                prev.getparent().remove(prev)
                prev = nxt
            n += 1
    return n


GENERATOR = "irb-docs generator"


def safe_save(doc, out, force=False):
    """生成した書類を保存する。人が Word で手直しした版や、開かれている書類は上書きしない。

    生成時に作成者欄（lastModifiedBy）へ GENERATOR を入れておき、次に保存するとき既存ファイルの
    作成者欄が別人（= Word で誰かが保存した）なら止める。手直しを生成スクリプトに反映してから
    force=True で作り直す。2026-09-17、PI が Word で直した4件を危うく上書きしかけたので入れた。"""
    from datetime import datetime
    from pathlib import Path
    from docx import Document

    out = Path(out)
    locks = [out.with_name("~$" + out.name), out.with_name("~$" + out.name[2:])]
    if any(p.exists() for p in locks):
        raise SystemExit(f"{out.name} は Word で開かれている（~$ のロックファイルがある）。閉じてもらってから保存する。")
    if out.exists() and not force:
        cp = Document(str(out)).core_properties
        if (cp.last_modified_by or "") != GENERATOR:
            raise SystemExit(
                f"{out.name} は「{cp.last_modified_by}」が {cp.modified} に保存した版で、手直しが入っている可能性がある。"
                "上書きすると消えるので止めた。手直しを生成スクリプトに反映してから force=True（--force）で作り直す。")
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.core_properties.last_modified_by = GENERATOR
    doc.core_properties.modified = datetime.now()
    doc.save(str(out))
    return out


__all__ = [_n for _n in list(globals()) if not _n.startswith("_")] + ["_add_runs"]
