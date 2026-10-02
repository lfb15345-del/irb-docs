# -*- coding: utf-8 -*-
"""倫理審査書類（.docx）の提出前チェック。書類は読むだけで、書き換えない。

    python audit_docs.py 計画書.docx --icf 説明文書.docx --form 様式.docx ...
        [--template 計画書=HT1.docx 説明文書=HT2.docx] [--refs-heading "参考資料・文献リスト"]
        [--ban 事業化 <対象者の実名> ...] [--warn 語 ...] [--allow 語 ...]

「要修正」が1件でもあれば終了コード1。「要確認」は人が読んで判断する（誤検出がありうる）。
これで分かるのは形の不備まで。中身（実態との一致、引用の中身、書類間の説明の食い違い）は SKILL.md の
「点検の手順」で人が見る。

見るもの
  A 色が残っている（本文・表・ヘッダー/フッター）            … ひな形の注意書き/例文の消し残し
  B ひな形の痕跡（（例、←、○○、○歳、20XX、＊、（本文を記載）…）… 黒字に混ざった例文、下書きの残り
  C 見出しだけで本文がない                                     … 青字の定型文を色で消した副作用
  D 同じ見出しが二重にある／同じ趣旨の文が二度ある              … 挿入した本文と、ひな形の文の両方が残った
  E 未記入の【　】【氏名】【A／B】                              … 提出前に埋める欄
  F 引用番号（上付き・登場順・未引用の文献・「17)」表記）         … 北大雛形は「文献番号を右肩に」
  G □のままのチェック欄                                         … ラン境界で置換が空振りしていないか
  H 説明文書が指針の説明事項（第8の6。令和5年版は第8の5）に触れているか（--icf） … キーワードによる目安
  I 書類間で数字・言い方（期間・人数・保管年数・安静の分数・端末）が揃っているか … 一覧を目で見る
  J ひな形の定型文が見当たらない（--template）                   … 消しすぎの目安
"""
import argparse
import difflib
import io
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

from docx import Document  # noqa: E402
from docx.table import Table  # noqa: E402
from docx.text.paragraph import Paragraph  # noqa: E402

from docx_tools import (GUIDE_COLORS, blocks, chrome_parts, distinct_cells, heading_level,  # noqa: E402
                        rcolor, run_class)

TRACE = ("（例", "(例", "←", "○○", "〇〇", "◯◯", "○歳", "○例", "〇例", "○年", "△△", "××",
         "20XX", "XX年", "＊", "ここに研究", "記載すること", "削除してください", "記入してください",
         "本文を記載")
WARN = ("してください", "※", "患者さん", "検体", "カルテ", "診療情報", "治療に影響", "医療を提供")
WARN_NOT_FOR_ICF = ("してください",)         # 説明文書では対象者への指示として普通に使う
AUTHOR_NOTE = ("雛形", "ご利用ください", "ご使用ください", "避けてください", "記載", "削除", "選択して", "適宜")
PLACEHOLDER_RX = re.compile(r"【([\s　]*|[^】]*／[^】]*|氏名|職名|所属|電話番号?|連絡先|日付)】")
PLAIN_CITE_RX = re.compile(r"(?<=[^\d（(．.\s０-９])(\d{1,2}(?:[,，\-–]\d{1,2})*)\)(?=[、。がはをにとも])")
CONSENT_MARK = re.compile(r"^同\s*意\s*(文\s*)?書$")
SKIP_REST_RX = re.compile(r"☑\s*指名しない|以下、?回答不要")

# 指針の説明事項（第8の6。令和5年版は第8の5）の目安。説明文書のどこかに語があれば「触れている」とみなす
ICF_ITEMS = [
    ("① 研究の名称・機関の長の許可", ["許可"]),
    ("② 機関名・研究責任者（多機関なら代表者・共同機関）", ["研究代表者", "研究責任者"]),
    ("③ 目的・意義", ["目的"]),
    ("④ 方法・期間（情報の利用目的）", ["方法", "期間"]),
    ("⑤ 選定された理由", ["お願いしているのは", "お願いする理由", "選ばれ", "選定", "対象となる方"]),
    ("⑥ 負担・リスク・利益", ["負担", "不利益"]),
    ("⑦ いつでも撤回できる", ["撤回"]),
    ("⑧ 同意しない・撤回しても不利益なし", ["不利益を受け", "不利益な扱い", "不利益になる", "不利益はありません"]),
    ("⑨ 情報公開の方法", ["公表", "学会", "論文"]),
    ("⑩ 計画書等の入手・閲覧", ["研究計画", "閲覧", "お知らせ"]),
    ("⑪ 個人情報等の取扱い", ["個人情報", "研究用の番号", "研究用ID"]),
    ("⑫ 保管・廃棄の方法", ["保管", "廃棄"]),
    ("⑬ 資金源・利益相反", ["研究資金", "利益相反"]),
    ("⑭ 研究で得られた結果の取扱い", ["結果"]),
    ("⑮ 相談等への対応", ["相談"]),
    ("⑰ 経済的負担・謝礼", ["費用", "謝礼", "謝金"]),
    ("⑭ 偶発的所見の扱い（ある研究のみ）", ["偶発", "所見"]),
    ("㉑ 将来の研究・他機関への提供", ["将来", "他の研究", "二次利用"]),
]


def table_paragraphs(tbl):
    for row in tbl.rows:
        for c in distinct_cells(row):
            yield c, c.paragraphs


class Report:
    def __init__(self, name):
        self.name, self.fix, self.check = name, defaultdict(list), defaultdict(list)

    def must(self, key, msg):
        self.fix[key].append(msg)

    def maybe(self, key, msg):
        self.check[key].append(msg)

    def show(self):
        print(f"\n==================== {self.name}")
        if not self.fix and not self.check:
            print("  問題なし")
        for title, store in (("要修正", self.fix), ("要確認", self.check)):
            for k, v in store.items():
                print(f"  [{title}] {k}: {len(v)}件")
                for m in v[:12]:
                    print("      -", m)
                if len(v) > 12:
                    print(f"      …ほか{len(v) - 12}件")
        return sum(len(v) for v in self.fix.values())


def check_text(where, t, args, rep, is_icf, form):
    for m in PLACEHOLDER_RX.finditer(t):
        rep.maybe("E 未記入の欄", f"{where} {m.group(0)}  …{t[:50]}")
    if form:
        return
    for k in TRACE + tuple(args.ban):
        if k in t:
            rep.must("B ひな形の痕跡・禁止語", f"{where} 「{k}」 {t[:60]}")
            break
    for k in WARN + tuple(args.warn):
        if is_icf and k in WARN_NOT_FOR_ICF:
            continue
        if k in t and not any(ok in t for ok in args.allow):
            rep.maybe("B' 研究に合っているか（患者・治療・検体・カルテ等）", f"{where} 「{k}」 {t[:60]}")
            break


def audit_one(path, args, rep, form=False, is_icf=False):
    """form=True（委員会様式・病院様式）は A 色 / E 未記入 / G チェック欄だけを見る。
    様式の脚注（「…記載すること」「（１）本院の教員…」「*1」）は正規の文言なので痕跡扱いしない。"""
    doc = Document(path)
    bl = blocks(doc)
    paras = [(i, b) for i, b in enumerate(bl) if isinstance(b, Paragraph)]

    # A 色 / B 痕跡 / E 未記入（本文・表）
    texts = []
    for i, b in enumerate(bl):
        ps = [b] if isinstance(b, Paragraph) else [p for _, pp in table_paragraphs(b) for p in pp]
        for p in ps:
            hit = [r.text for r in p.runs if r.text.strip() and run_class(r, p) != "black"]
            if hit:
                rep.must("A 色が残っている", f"#{i} {''.join(hit)[:40]}")
            t = p.text.strip()
            if t:
                texts.append(t)
                check_text(f"#{i}", t, args, rep, is_icf, form)
    # ヘッダー/フッター（定義のある部品だけ。触ると空の部品ができるため）
    for name, part in chrome_parts(doc):
        for p in part.paragraphs:
            t = p.text.strip()
            if not t:
                continue
            if any(r.text.strip() and (rcolor(r) or "") in GUIDE_COLORS for r in p.runs):
                rep.must("A 色が残っている", f"{name}: {t[:40]}")
            check_text(name, t, args, rep, is_icf, form)

    if not form:
        audit_structure(bl, paras, args, rep)
        audit_repeats(paras, rep)
    audit_checkboxes(bl, rep)
    return doc, texts


def audit_structure(bl, paras, args, rep):
    # C 本文のない見出し / D 見出しの重複
    consent_started = False
    seen1, seen2 = Counter(), Counter()
    heads_seen = set()
    chapter = ""
    for i, p in paras:
        t = p.text.strip()
        if CONSENT_MARK.match(t):
            consent_started = True
        lv = heading_level(t)
        if not lv:
            continue
        norm = re.sub(r"[\s　]", "", t)
        repeat = norm in heads_seen
        heads_seen.add(norm)
        if lv == 1 and not repeat:
            chapter = t
        if not consent_started:
            if lv == 1:
                seen1[t] += 1
            elif lv < 3:
                seen2[(chapter, t)] += 1
        if repeat:
            continue                             # 再掲（重複はDで数える。同意文書の項目一覧もここで除外）
        nxt_level, j = None, i + 1
        while j < len(bl):
            nb = bl[j]
            if isinstance(nb, Table):
                nxt_level = -1                   # 表が続く＝中身あり
                break
            nt = nb.text.strip()
            if nt:
                nxt_level = heading_level(nt)
                break
            j += 1
        if nxt_level is None:
            rep.must("C 見出しだけで本文がない", f"#{i} {t}（文書末）")
        elif 0 < nxt_level <= lv:
            rep.must("C 見出しだけで本文がない", f"#{i} {t} → 次も見出し「{bl[j].text.strip()[:30]}」")
    for t, n in seen1.items():
        if n > 1:
            rep.must("D 見出しの重複", f"{t} ×{n}")
    for (ch, t), n in seen2.items():
        if n > 1:
            rep.must("D 見出しの重複", f"{ch[:12]} の {t} ×{n}")

    # F 引用
    order = []
    refs_start = None
    for i, p in paras:
        if args.refs_heading and args.refs_heading in p.text and heading_level(p.text):
            refs_start = i
            break
        for r in p.runs:
            if r.font.superscript and r.text.strip():
                for part in re.split(r"[,，]", r.text.strip()):
                    m = re.fullmatch(r"(\d+)\s*[-–]\s*(\d+)", part.strip())
                    nums = range(int(m.group(1)), int(m.group(2)) + 1) if m else (
                        [int(part)] if part.strip().isdigit() else [])
                    for n in nums:
                        if n not in order:
                            order.append(n)
        for m in PLAIN_CITE_RX.finditer(p.text):
            rep.must("F 引用番号が上付きでない", f"#{i} 「{m.group(0)}」 …{p.text[max(0, m.start() - 20):m.start()]}")
    refs = []
    if refs_start is not None:
        for i, p in paras:
            if i <= refs_start:
                continue
            t = p.text.strip()
            if heading_level(t) == 1:
                break
            if re.match(r"^\d+[\.．]\s*", t):
                refs.append(t)
    if order and order != list(range(1, len(order) + 1)):
        rep.must("F 引用番号が登場順でない", f"初出の順: {order}")
    if refs:
        cited = set(order)
        uncited = [n for n in range(1, len(refs) + 1) if n not in cited]
        if uncited:
            rep.must("F 本文で引用されていない文献", f"番号 {uncited}")
        if order and max(order) > len(refs):
            rep.must("F 文献リストにない番号", f"最大 {max(order)} / リスト {len(refs)} 件")
        for t in refs:
            if not re.search(r"\b(19|20)\d\d\b", t):
                rep.maybe("F 文献の書誌", f"年が見当たらない: {t[:60]}")
    elif order:
        rep.maybe("F 文献リスト", f"--refs-heading で見出しを指定すると照合する（引用 {len(order)} 件）")


def sentences(t):
    """「。」で終わる20字以上の文（住所や署名欄のような「。」の無い行は比べない）。"""
    return [s.strip() for s in re.split(r"(?<=。)", t) if s.strip().endswith("。") and len(s.strip()) >= 20]


def audit_repeats(paras, rep, ratio=0.9):
    """D' 同じ章の中で、行の頭が重なる段落（②測定項目：該当なし／②測定項目）や、ほぼ同じ文が二度ある。
    同意文書・撤回書（署名欄が繰り返す）と見出しの行は比べない。"""
    chapter, prev = "", ""
    by_chapter = defaultdict(list)
    for i, p in paras:
        t = p.text.strip()
        if not t:
            continue
        if CONSENT_MARK.match(t):
            break
        lv = heading_level(t)
        if lv == 1:
            chapter = t
        if (prev and len(t) >= 4 and len(prev) >= 4 and t != prev and not prev.endswith("。")
                and (prev.startswith(t) or t.startswith(prev))):
            rep.maybe("D' 重複の疑い（行の頭が同じ）", f"#{i} 「{prev[:30]}」の次に「{t[:30]}」")
        prev = t
        if lv or "\t" in t:
            continue
        for s in sentences(t):
            by_chapter[chapter].append((i, s))
    for ch, items in by_chapter.items():
        for a in range(len(items)):
            for b in range(a + 1, len(items)):
                (ia, sa), (ib, sb) = items[a], items[b]
                if ia == ib:
                    continue
                if difflib.SequenceMatcher(None, sa, sb).ratio() >= ratio:
                    rep.maybe("D' 同じ趣旨の文が二度ある", f"{ch[:14]}: #{ia}「{sa[:28]}…」と #{ib}")


def audit_checkboxes(bl, rep):
    # G チェック欄（「☑指名しない（以下、回答不要）」の後ろと、名前の空いた予備の行は数えない）
    for i, b in enumerate(bl):
        if not isinstance(b, Table):
            continue
        skip_rest = False
        for row in b.rows:
            cells = distinct_cells(row)
            row_text = " ".join(c.text for c in cells)
            if skip_rest:
                if re.match(r"^\s*[ⅠⅡⅢⅣ]", cells[0].text):
                    skip_rest = False
                else:
                    continue
            if SKIP_REST_RX.search(row_text):
                skip_rest = True
                continue
            if cells[0].text.strip() in ("該当なし", "なし"):
                continue                          # 「該当なし」の行
            if len(cells) > 2 and not cells[0].text.strip() and all(
                    not re.sub(r"[□☑■✓✔☒\s　]|医師|歯科医師|その他|（\s*）|該当|非該当|有効|期限切れ|未取得|/", "", c.text)
                    for c in cells[1:]):
                continue                          # 予備の空行
            for c in cells:
                ct = c.text
                if "□" in ct and not re.search(r"[☑■✓✔☒]", ct):
                    label = cells[0].text.strip().replace("\n", " ")[:30]
                    rep.maybe("G □のままのチェック欄", f"表#{i} {label} | {ct.strip().replace(chr(10), ' ')[:40]}")


def icf_items(texts, rep):
    body = "\n".join(texts)
    for name, kws in ICF_ITEMS:
        if not any(k in body for k in kws):
            (rep.maybe if "ある研究のみ" in name else rep.must)("H 説明事項が見当たらない", f"{name}（語: {'/'.join(kws)}）")


def template_lines(path, min_len=12):
    """ひな形の黒字の定型文。[(比べる文, 種類)] を返す。
    種類 'label' は「臨床研究課題名：「○○」」のように空欄を含む行で、ラベルの語があるかだけを見る。
    使用上の注意など作成者向けの文、赤字の注記（←…）は除く。"""
    doc = Document(path)
    out = []
    for b in blocks(doc):
        if not isinstance(b, Paragraph):
            continue
        t = b.text.strip()
        if not t or heading_level(t) or t.startswith(("＊", "*", "・")):
            continue
        black = "".join(r.text for r in b.runs if run_class(r, b) == "black").strip()
        black = re.sub(r"（←[^）]*）|←.*$|＊.*$", "", black).strip()
        if not black or any(k in black for k in AUTHOR_NOTE):
            continue
        if re.search(r"[○〇◯]{2,}", t):                # 「臨床研究課題名：「○○」」のように空欄（色字のことも）を含む行
            label = re.sub(r"[○〇◯]+|[「」（）()：:、\s　]", "", black)
            if len(label) >= 5:
                out.append((label[:8], "label"))
            continue
        if len(black) >= min_len and "（例" not in black:
            out.append((black, "text"))
    return out


def audit_template(doc_texts, tpl_path, rep, limit=30):
    """J ひな形の黒字の定型文・ラベルが、出来上がった書類に見当たらないもの（消しすぎの目安）。
    ひな形の冒頭（使用上の注意など）は、書類と最初に一致する行が出てくるまで比べない。"""
    joined = "\n".join(doc_texts)
    missing = []
    started = False
    for t, kind in template_lines(tpl_path):
        if not started:
            if kind == "text" and max((difflib.SequenceMatcher(None, t, x).ratio() for x in doc_texts), default=0) >= 0.9:
                started = True
            continue
        if kind == "label":
            if t not in joined:
                missing.append(f"ラベル「{t}」の行")
            continue
        best = max((difflib.SequenceMatcher(None, t, x).ratio() for x in doc_texts), default=0)
        if best < 0.6:
            missing.append(t)
    for t in missing[:limit]:
        rep.maybe("J ひな形の定型文が見当たらない（消しすぎの疑い）", t[:70])
    if len(missing) > limit:
        rep.maybe("J ひな形の定型文が見当たらない（消しすぎの疑い）", f"…ほか{len(missing) - limit}件")


def numbers(texts):
    out = defaultdict(Counter)
    pats = {
        "日付": r"20\d\d年\d{1,2}月\d{1,2}日",
        "人数": r"\d+\s*(?:例|名)(?!誉)",
        "年数": r"\d+\s*年(?:間|が経過)",
        "期間": r"実施許可日[～〜~][^、。）]*",
        "分数": r"\d+\s*[～〜~]\s*\d+\s*分",
        "端末": r"iPhone|Android|スマートフォン|携帯電話",
        "生年": r"生年月日|生年(?!月)",
    }
    for t in texts:
        for k, rx in pats.items():
            for m in re.findall(rx, t):
                out[k][m] += 1
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("docs", nargs="*")
    ap.add_argument("--icf", nargs="*", default=[], help="説明文書（説明事項の目安も見る）")
    ap.add_argument("--form", nargs="*", default=[], help="委員会様式・病院様式（色・未記入・チェック欄だけ見る）")
    ap.add_argument("--template", nargs="*", default=[], help="書類名の一部=ひな形.docx（定型文の消しすぎを見る）")
    ap.add_argument("--refs-heading", default="参考資料・文献リスト")
    ap.add_argument("--ban", nargs="*", default=[], help="書いてはいけない語（例: 事業化 実名の対象者 システムの商品名）")
    ap.add_argument("--warn", nargs="*", default=[])
    ap.add_argument("--allow", nargs="*", default=[], help="WARN を出さない文に含まれる語")
    args = ap.parse_args()
    tpl = {}
    for x in args.template:
        if "=" in x:
            k, v = x.split("=", 1)
            tpl[k] = v
    total, nums = 0, {}
    for path in list(args.docs) + list(args.icf) + list(args.form):
        rep = Report(Path(path).name)
        is_form, is_icf = path in args.form, path in args.icf
        _, texts = audit_one(path, args, rep, form=is_form, is_icf=is_icf)
        if is_icf:
            icf_items(texts, rep)
        for key, tp in tpl.items():
            if key in Path(path).name:
                audit_template(texts, tp, rep)
        total += rep.show()
        if not is_form:
            nums[Path(path).name] = numbers(texts)
    print("\n==================== I 書類間の数字・言い方（目で突き合わせる）")
    for kind in ("期間", "日付", "人数", "年数", "分数", "端末", "生年"):
        for name, d in nums.items():
            vals = ", ".join(f"{k}×{v}" for k, v in d[kind].most_common(8))
            if vals:
                print(f"  {kind:<3} {name[:28]:<28} {vals}")
    print(f"\n要修正 合計 {total} 件")
    print("※ 形の点検だけ。実態との一致・引用の中身・説明の食い違いは SKILL.md の「点検の手順」で見る。")
    sys.exit(1 if total else 0)


if __name__ == "__main__":
    main()
