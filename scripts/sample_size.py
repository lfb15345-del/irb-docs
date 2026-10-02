# -*- coding: utf-8 -*-
"""症例数の根拠を、よく使う設計について式で出し、計画書に貼れる文の下書きまで作る。

シンプルな解析の研究で「現実的な数を、冗長にならずに」示すための道具。反復測定・個人内の
集中的縦断デザインのように式で表しにくいものは power_sim.py（シミュレーション）を使う。

使い方（両側有意水準5%・検出力80%が既定。--alpha / --power で変える）:
  python sample_size.py two-proportions --p1 0.20 --p2 0.10 [--dropout 0.1] [--ratio 1]
  python sample_size.py two-means --delta 5 --sd 10 [--dropout 0.1] [--ratio 1]
  python sample_size.py paired-means --delta 3 --sd-diff 6
  python sample_size.py correlation --r 0.3
  python sample_size.py proportion-precision --p 0.2 --half-width 0.05
  python sample_size.py mean-precision --sd 10 --half-width 2
  python sample_size.py logistic-events --vars 5 --event-rate 0.15 [--epv 10]
  python sample_size.py survival --hr 0.7 --event-prob 0.3 [--alloc 0.5]
  どれにも --per-year N --years Y を付けると、期間中に集まる見込みの数と並べて出す（実施可能性）。

出力の文は下書き。【　　】は根拠（文献・自施設のデータ）を書き足す所。仮定は楽観的に置かない。
"""
import argparse
import io
import math
import sys

from scipy import stats

if hasattr(sys.stdout, "buffer"):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")


def z(p):
    return stats.norm.ppf(p)


def ceil(x):
    return int(math.ceil(x - 1e-9))


def with_dropout(n, rate):
    return ceil(n / (1 - rate)) if rate else n


def pct(x):
    return f"{x * 100:.0f}%" if abs(x * 100 - round(x * 100)) < 1e-9 else f"{x * 100:.1f}%"


def common(a):
    za, zb = z(1 - a.alpha / 2), z(a.power)
    return za, zb


def feasibility(a, total):
    if not (a.per_year and a.years):
        return ""
    have = a.per_year * a.years
    ok = "足りる" if have >= total else "足りない（期間を延ばす・多施設にする・仮定を見直す）"
    return (f"\n実施可能性: 年間 {a.per_year:g} 例 × {a.years:g} 年 = {have:g} 例 → 必要 {total} 例に{ok}。"
            f"\n  文例: 当院の年間の対象症例は約 {a.per_year:g} 例であり【根拠: 過去○年の実績】、"
            f"{a.years:g} 年間で約 {have:g} 例の登録が見込まれる。")


def head(a):
    return f"両側有意水準 {pct(a.alpha)}、検出力 {pct(a.power)}"


def two_proportions(a):
    za, zb = common(a)
    p1, p2, k = a.p1, a.p2, a.ratio            # k = 群2 / 群1
    d = abs(p1 - p2)
    pbar = (p1 + k * p2) / (1 + k)
    n1 = (za * math.sqrt((1 + 1 / k) * pbar * (1 - pbar)) + zb * math.sqrt(p1 * (1 - p1) + p2 * (1 - p2) / k)) ** 2 / d ** 2
    # Fleiss の連続性補正（保守的な方。計画書にはこちらを使うことが多い）
    n1cc = n1 / 4 * (1 + math.sqrt(1 + 2 * (k + 1) / (k * n1 * d))) ** 2
    n1, n1cc = ceil(n1), ceil(n1cc)
    n2cc = ceil(n1cc * k)
    tot = n1cc + n2cc
    t1, t2 = with_dropout(n1cc, a.dropout), with_dropout(n2cc, a.dropout)
    out = [f"2群の割合の比較（{head(a)}、群の比 1:{k:g}）",
           f"  補正なし: 群1 {n1} 例 / 補正あり(Fleiss): 群1 {n1cc} 例・群2 {n2cc} 例・合計 {tot} 例"]
    if a.dropout:
        out.append(f"  脱落 {pct(a.dropout)} を見込むと: 群1 {t1} 例・群2 {t2} 例・合計 {t1 + t2} 例")
    fin = t1 + t2 if a.dropout else tot
    out.append("文例: 主要評価項目である【　　】の割合を、対照群 "
               f"{pct(p1)}【根拠】、試験群 {pct(p2)} と仮定し、{head(a)} で2群の差を検出するのに必要な症例数は"
               f"各群 {n1cc} 例（連続性補正あり）である。"
               + (f"脱落を {pct(a.dropout)} 見込み、合計 {fin} 例とした。" if a.dropout else f"合計 {fin} 例とした。"))
    return "\n".join(out) + feasibility(a, fin)


def _t_iterate(base, df_of, za_f, zb_f):
    """z の式の n から始めて、t 分布で n を合わせる（自由度が小さいときの補正）。"""
    n = max(2.0, base(za_f, zb_f))
    for _ in range(50):
        df = max(1.0, df_of(n))
        ta, tb = stats.t.ppf(za_f, df), stats.t.ppf(zb_f, df)
        n_new = base(ta, tb)
        if abs(n_new - n) < 1e-6:
            break
        n = n_new
    return n


def two_means(a):
    k = a.ratio
    pa, pb = 1 - a.alpha / 2, a.power
    base = lambda qa, qb: (1 + 1 / k) * (qa + qb) ** 2 * a.sd ** 2 / a.delta ** 2
    n1 = ceil(_t_iterate(base, lambda n: n + n * k - 2, pa, pb))
    n2 = ceil(n1 * k)
    t1, t2 = with_dropout(n1, a.dropout), with_dropout(n2, a.dropout)
    fin = t1 + t2
    out = [f"2群の平均の比較（{head(a)}、群の比 1:{k:g}、t 分布で補正）",
           f"  群1 {n1} 例・群2 {n2} 例・合計 {n1 + n2} 例" + (f" → 脱落 {pct(a.dropout)} を見込み 合計 {fin} 例" if a.dropout else ""),
           f"  効果量 d = {a.delta / a.sd:.2f}",
           "文例: 主要評価項目である【　　】の群間差を "
           f"{a.delta:g}【単位】、標準偏差を {a.sd:g}【根拠】と仮定し、{head(a)} で必要な症例数は各群 {n1} 例である。"
           + (f"脱落を {pct(a.dropout)} 見込み、合計 {fin} 例とした。" if a.dropout else f"合計 {fin} 例とした。")]
    return "\n".join(out) + feasibility(a, fin)


def paired_means(a):
    pa, pb = 1 - a.alpha / 2, a.power
    base = lambda qa, qb: (qa + qb) ** 2 * a.sd_diff ** 2 / a.delta ** 2
    n = ceil(_t_iterate(base, lambda n: n - 1, pa, pb))
    fin = with_dropout(n, a.dropout)
    out = [f"対応のある平均の比較（前後比較など。{head(a)}、t 分布で補正）",
           f"  {n} 例" + (f" → 脱落 {pct(a.dropout)} を見込み {fin} 例" if a.dropout else ""),
           "文例: 【前後】の差の平均を "
           f"{a.delta:g}【単位】、差の標準偏差を {a.sd_diff:g}【根拠】と仮定し、{head(a)} で必要な症例数は {n} 例である。"
           + (f"脱落を {pct(a.dropout)} 見込み {fin} 例とした。" if a.dropout else "")]
    return "\n".join(out) + feasibility(a, fin)


def correlation(a):
    za, zb = common(a)
    c = 0.5 * math.log((1 + a.r) / (1 - a.r))
    n = ceil(((za + zb) / c) ** 2 + 3)
    fin = with_dropout(n, a.dropout)
    return ("\n".join([f"相関係数が0でないことの検出（{head(a)}、Fisher の z 変換）",
                       f"  {n} 例" + (f" → 欠測 {pct(a.dropout)} を見込み {fin} 例" if a.dropout else ""),
                       f"文例: 【変数A】と【変数B】の相関係数を {a.r:g}【根拠】と仮定し、{head(a)} で相関を検出するのに必要な症例数は {n} 例である。"])
            + feasibility(a, fin))


def proportion_precision(a):
    zc = z(1 - a.alpha / 2)
    n = ceil(zc ** 2 * a.p * (1 - a.p) / a.half_width ** 2)
    fin = with_dropout(n, a.dropout)
    conf = pct(1 - a.alpha)
    return ("\n".join([f"割合を推定する精度（{conf} 信頼区間 ±{pct(a.half_width)}）",
                       f"  {n} 例" + (f" → {fin} 例（欠測 {pct(a.dropout)} を見込み）" if a.dropout else ""),
                       f"文例: 本研究は探索的研究であり、検出力に基づく症例数設計は行わない。【　　】の割合を {pct(a.p)}【根拠】と見込むと、"
                       f"{n} 例で {conf} 信頼区間の幅は ±{pct(a.half_width)} となり、推定に十分な精度が得られる。"])
            + feasibility(a, fin))


def mean_precision(a):
    zc = z(1 - a.alpha / 2)
    n = ceil((zc * a.sd / a.half_width) ** 2)
    fin = with_dropout(n, a.dropout)
    conf = pct(1 - a.alpha)
    return ("\n".join([f"平均を推定する精度（{conf} 信頼区間 ±{a.half_width:g}）",
                       f"  {n} 例" + (f" → {fin} 例（欠測 {pct(a.dropout)} を見込み）" if a.dropout else ""),
                       f"文例: 本研究は探索的研究であり、検出力に基づく症例数設計は行わない。【　　】の標準偏差を {a.sd:g}【根拠】と見込むと、"
                       f"{n} 例で平均の {conf} 信頼区間の幅は ±{a.half_width:g}【単位】となる。"])
            + feasibility(a, fin))


def logistic_events(a):
    events = ceil(a.epv * a.vars)
    n = ceil(events / a.event_rate)
    fin = with_dropout(n, a.dropout)
    return ("\n".join([f"多変量ロジスティック回帰（説明変数 {a.vars} 個、1変数あたりのイベント数 {a.epv:g}・目安）",
                       f"  必要なイベント数 {events} → 発生割合 {pct(a.event_rate)} なら {n} 例" + (f"（欠測 {pct(a.dropout)} を見込み {fin} 例）" if a.dropout else ""),
                       "  注意: 1変数あたり10イベントは目安にすぎない。予測モデルを作る研究なら、より厳密な方法で根拠を出す。",
                       f"文例: 調整する説明変数を {a.vars} 個とし、1変数あたり {a.epv:g} 件のイベントを目安とすると {events} 件のイベントが必要である。"
                       f"【　　】の発生割合を {pct(a.event_rate)}【根拠】と見込み、{n} 例を目標とした。"])
            + feasibility(a, fin))


def survival(a):
    za, zb = common(a)
    pi = a.alloc
    d = ceil((za + zb) ** 2 / (pi * (1 - pi) * math.log(a.hr) ** 2))
    n = ceil(d / a.event_prob)
    fin = with_dropout(n, a.dropout)
    return ("\n".join([f"生存時間の2群比較（log-rank 検定、Schoenfeld の式、{head(a)}、割付 {pct(pi)}）",
                       f"  必要なイベント数 {d} → 観察期間中のイベント割合 {pct(a.event_prob)} なら合計 {n} 例" + (f"（脱落 {pct(a.dropout)} を見込み {fin} 例）" if a.dropout else ""),
                       f"文例: ハザード比を {a.hr:g}【根拠】と仮定し、{head(a)} で必要なイベント数は {d} 件である。"
                       f"観察期間中のイベント割合を {pct(a.event_prob)}【根拠】と見込み、合計 {n} 例を目標とした。"])
            + feasibility(a, fin))


def main():
    ap = argparse.ArgumentParser(description="よく使う設計の症例数（式）と計画書の文の下書き")
    sub = ap.add_subparsers(dest="cmd", required=True)

    def add(name, fn, *args):
        p = sub.add_parser(name)
        for spec in args:
            p.add_argument(*spec[0], **spec[1])
        p.add_argument("--alpha", type=float, default=0.05)
        p.add_argument("--power", type=float, default=0.80)
        p.add_argument("--dropout", type=float, default=0.0, help="脱落・欠測の見込み（0〜1）")
        p.add_argument("--per-year", type=float, default=0, help="期間中に集まる見込みの数（年あたり）")
        p.add_argument("--years", type=float, default=0)
        p.set_defaults(fn=fn)

    F = lambda *n, **k: (n, {"type": float, **k})
    add("two-proportions", two_proportions, F("--p1", required=True, help="対照群の割合"), F("--p2", required=True, help="比べる群の割合"), F("--ratio", default=1.0))
    add("two-means", two_means, F("--delta", required=True), F("--sd", required=True), F("--ratio", default=1.0))
    add("paired-means", paired_means, F("--delta", required=True), F("--sd-diff", required=True))
    add("correlation", correlation, F("--r", required=True))
    add("proportion-precision", proportion_precision, F("--p", required=True), F("--half-width", required=True))
    add("mean-precision", mean_precision, F("--sd", required=True), F("--half-width", required=True))
    add("logistic-events", logistic_events, (("--vars",), {"type": int, "required": True}), F("--event-rate", required=True), F("--epv", default=10.0))
    add("survival", survival, F("--hr", required=True), F("--event-prob", required=True), F("--alloc", default=0.5))
    a = ap.parse_args()
    print(a.fn(a))


if __name__ == "__main__":
    main()
