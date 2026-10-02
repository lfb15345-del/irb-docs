# -*- coding: utf-8 -*-
"""集中的縦断デザイン（1人を何十日も測る観察研究）の症例数シミュレーション。

個人ごとに「説明変数→目的変数」の傾きを最小二乗で推定し、DerSimonian-Laird のランダム効果メタ解析で
統合したとき、平均の傾きが0と異なると判定できる確率（検出力）と、個人ごとの傾きの95%区間の幅を出す。
予備データで得た値を「真値」として入れる。計画書の「目標症例数の設定根拠」にそのまま書ける形で出力する。

例（1人を毎日測る研究：当日の曝露（時間）→ 翌朝の測定値（対数）。値は予備データから入れる）
    python power_sim.py --n 10 20 30 --days 60 120 --mu -0.02 --tau 0.02 \
        --sigma 0.29 0.47 --x-sd 4.5 --sims 4000 --per 4 --baseline 17 --unit "ln/時間"
    → 10名 0.79、20名 0.97、30名 0.998（1人60日）

引数
  --mu      平均の傾き（検出したい大きさ）       --tau   個人間の傾きのSD
  --sigma   個人内の残差SD（複数なら人ごとに順に割り当て）
  --x-sd    説明変数の個人内SD                  --ar    残差の自己相関（AR(1)の係数、既定0）
  --test    z（既定）か t（自由度 k-1）          --per/--baseline  区間幅を実単位に直す（例: 4時間あたり、17 ms）
"""
import argparse
import io
import math
import sys

import numpy as np

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")


def simulate(n, days, a, rng):
    sims = a.sims
    slopes = rng.normal(a.mu, a.tau, size=(sims, n))
    sig = np.array([a.sigma[i % len(a.sigma)] for i in range(n)])
    x = rng.normal(0.0, a.x_sd, size=(sims, n, days))
    e = rng.normal(0.0, 1.0, size=(sims, n, days))
    if a.ar:
        for t in range(1, days):
            e[..., t] = a.ar * e[..., t - 1] + math.sqrt(1 - a.ar ** 2) * e[..., t]
    y = slopes[..., None] * x + sig[None, :, None] * e
    xc = x - x.mean(-1, keepdims=True)
    yc = y - y.mean(-1, keepdims=True)
    sxx = (xc ** 2).sum(-1)
    b = (xc * yc).sum(-1) / sxx
    resid = yc - b[..., None] * xc
    s2 = (resid ** 2).sum(-1) / (days - 2)
    v = s2 / sxx                                        # 個人ごとの傾きの分散
    w = 1 / v
    b_fe = (w * b).sum(-1) / w.sum(-1)
    q = (w * (b - b_fe[:, None]) ** 2).sum(-1)
    c = w.sum(-1) - (w ** 2).sum(-1) / w.sum(-1)
    tau2 = np.maximum(0, (q - (n - 1)) / c)
    ws = 1 / (v + tau2[:, None])
    mu_hat = (ws * b).sum(-1) / ws.sum(-1)
    se = np.sqrt(1 / ws.sum(-1))
    stat = np.abs(mu_hat / se)
    if a.test == "t":
        from scipy import stats
        crit = stats.t.ppf(1 - a.alpha / 2, n - 1)
    else:
        crit = 1.959963984540054 if a.alpha == 0.05 else float(np.abs(np.quantile(rng.standard_normal(10 ** 6), a.alpha / 2)))
    power = float((stat > crit).mean())
    se_i = np.sqrt(v)
    half = {"平均": float(1.959963984540054 * se_i.mean())}
    for s in sorted(set(a.sigma)):
        half[f"残差SD {s:g}"] = float(1.959963984540054 * se_i[:, sig == s].mean())
    return power, half


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--n", type=int, nargs="+", default=[10, 20, 30])
    ap.add_argument("--days", type=int, nargs="+", default=[60])
    ap.add_argument("--mu", type=float, required=True)
    ap.add_argument("--tau", type=float, required=True)
    ap.add_argument("--sigma", type=float, nargs="+", required=True)
    ap.add_argument("--x-sd", type=float, required=True)
    ap.add_argument("--ar", type=float, default=0.0)
    ap.add_argument("--sims", type=int, default=4000)
    ap.add_argument("--alpha", type=float, default=0.05)
    ap.add_argument("--test", choices=["z", "t"], default="z")
    ap.add_argument("--per", type=float, default=None, help="区間幅を『説明変数この量あたり』に直す")
    ap.add_argument("--baseline", type=float, default=None, help="対数の区間幅を実単位に直す基準値（例: 17 ms）")
    ap.add_argument("--unit", default="")
    ap.add_argument("--seed", type=int, default=20260916)
    a = ap.parse_args()
    rng = np.random.default_rng(a.seed)

    print(f"条件: 平均の傾き {a.mu}{a.unit}、個人間SD {a.tau}、残差SD {a.sigma}、説明変数SD {a.x_sd}、"
          f"AR(1) {a.ar}、{a.test}検定 両側{a.alpha}、各{a.sims}回")
    print("\n検出力（平均の傾き≠0）")
    print("  人数 \\ 1人あたり日数: " + "  ".join(f"{d:>6}" for d in a.days))
    halves = {}
    for n in a.n:
        row = []
        for d in a.days:
            p, h = simulate(n, d, a, rng)
            row.append(p)
            halves[d] = h
        print(f"  {n:>4}名             " + "  ".join(f"{p:>6.3f}" for p in row))
    print("\n個人ごとの傾きの95%区間（±）。残差SDが大きい人ほど広い")
    for d, hs in halves.items():
        for label, h in hs.items():
            s = f"  {d}日 {label}: ±{h:.3f}{a.unit}"
            if a.per:
                s += f" → 説明変数{a.per:g}あたり ±{h * a.per:.3f}"
                if a.baseline:
                    s += f"（基準値{a.baseline:g}なら約±{h * a.per * a.baseline:.1f}）"
            print(s)
    print("\n計画書に書くときは、真値にした予備データの値・シミュレーションの設計（人ごとに傾きを推定→"
          "ランダム効果メタ解析で統合）・有意水準・回数・結果をそのまま並べる。")


if __name__ == "__main__":
    main()
