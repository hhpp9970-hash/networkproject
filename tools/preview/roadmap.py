"""稟議用のロードマップ（v12）。VBA の DrawRoadmap と同じ手順で行・帯・マイルストーン・年度ごとの金額を計算し、
A4 横1枚の完成イメージを描く。

決めごと
・まとめごと（まとめが空欄の作業は工事ごと）に1行。目的があれば目的ごとに分け、目的の中は始まりの早い順（同じなら出てきた順）
・作業のある期間を帯で描く（14日までのすき間はつなぐ）。日数0の作業（マイルストーン）は◆と名前・日付
・目盛りは年度（年度の始まりの月。既定は4月）と月。最初の作業の年度の初めから、最後の作業の年度の終わりまで
・金額は作業ごと。年度をまたぐ作業は、年度ごとの工事する日の数で分ける（工事する日がない作業は始まりの年度）

使い方：python tools/preview/roadmap.py 出力先 [--big]（--big は 4期・約150件・4年の工程に金額を付けたもの）
"""
import calendar
import datetime as dt
import random
import sys

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Polygon, Rectangle

import network_logic3 as v3
import network_logic5 as v5
from network_logic import compute
from network_logic3 import collect_phases, phase_levels
from network_logic4 import collect_zones

NAVY, PURP_FILL, PHASE_FILL, GRID, MID, DARK, MILE = "#1F3864", "#D9E1F2", "#D9E1F2", "#D9D9D9", "#767676", "#595959", "#C00000"
PALETTE = ["#2F5597", "#38761D", "#7030A0", "#8C5A2B", "#1B7F86", "#B4407E", "#1F3864", "#7F7F00"]  # VBA の AutoColor
CH = 7.0  # Excel の列幅1文字分のおおよその幅（pt。プレビューの決めごと）


def md(d):
    d = dt.date.fromordinal(d)
    return f"{d.month}/{d.day}"


def span_text(s, e):
    """VBA の RoadSpan"""
    a, b = dt.date.fromordinal(s), dt.date.fromordinal(e)
    if e <= s:
        return f"{a.year}/{a.month}/{a.day}"
    if a.year == b.year:
        return f"{a.year}/{a.month}/{a.day}～{b.month}/{b.day}"
    return f"{a.year}/{a.month}/{a.day}～{b.year}/{b.month}/{b.day}"


def last_day(S, F, i):
    return F[i] - 1 if F[i] > S[i] else S[i]


def road_rows(acts, c, zn):
    """VBA の RoadRows：目的とまとめ（まとめが空欄なら工事）の組ごと。並びは目的の順 → 始まりの早い順 → 出てきた順"""
    S, F = c["S"], c["F"]
    rows, idx, act_rw = [], {}, []
    for i, a in enumerate(acts):
        g = a.get("group", "")
        key = (zn["pidx"][i], "G" if g else "K", g or a["kouji"])
        if key not in idx:
            idx[key] = len(rows)
            rows.append(dict(label=g or a["kouji"], purp=zn["pidx"][i], s=S[i], e=last_day(S, F, i), mile=False))
        r = rows[idx[key]]
        r["s"] = min(r["s"], S[i])
        r["e"] = max(r["e"], last_day(S, F, i))
        r["mile"] = r["mile"] or F[i] <= S[i]
        act_rw.append(idx[key])
    seq = sorted(range(len(rows)), key=lambda k: (rows[k]["purp"], rows[k]["s"], k))
    return rows, act_rw, seq


def fiscal_year(d, fy_m):
    d = dt.date.fromordinal(d)
    return d.year - (1 if d.month < fy_m else 0)


def road_amounts(acts, c, rows, act_rw, y0, n_fy, fy_m):
    """VBA の RoadAmounts：作業の金額を年度ごとの工事する日の数で分ける"""
    S, F = c["S"], c["F"]
    amt = [[0.0] * n_fy for _ in rows]
    for i, a in enumerate(acts):
        A = a.get("amount") or 0
        if not A:
            continue
        w = v3.work_days_in(S[i], F[i]) if F[i] > S[i] else 0
        for f in range(n_fy):
            fs = dt.date(y0 + f, fy_m, 1).toordinal()
            fe = dt.date(y0 + f + 1, fy_m, 1).toordinal()
            if w > 0:
                wf = v3.work_days_in(max(S[i], fs), min(F[i], fe))
                amt[act_rw[i]][f] += A * wf / w
            elif fs <= S[i] < fe:
                amt[act_rw[i]][f] += A
    tot = [sum(r[f] for r in amt) for f in range(n_fy)]
    return amt, tot + [sum(tot)]


def segments(acts, c, act_rw, rw):
    """VBA の RoadBars：作業のある期間（14日までのすき間はつなぐ）と、マイルストーン"""
    S, F = c["S"], c["F"]
    iv = sorted((S[i], F[i]) for i in range(len(acts)) if act_rw[i] == rw and F[i] > S[i])
    segs = []
    for s, f in iv:
        if segs and s - segs[-1][1] <= 14:
            segs[-1][1] = max(segs[-1][1], f)
        else:
            segs.append([s, f])
    miles = [i for i in range(len(acts)) if act_rw[i] == rw and F[i] <= S[i]]
    return segs, miles


def render(acts, c, ph, zn, path, title, unit="千円", fy_m=4):
    S, F = c["S"], c["F"]
    rows, act_rw, seq = road_rows(acts, c, zn)
    min_s, max_e = min(S), max(last_day(S, F, i) for i in range(len(acts)))
    y0, y_e = fiscal_year(min_s, fy_m), fiscal_year(max_e, fy_m)
    n_fy = y_e - y0 + 1
    n_mon = n_fy * 12
    use_amt = any(a.get("amount") for a in acts)
    if use_amt:
        amt, tot = road_amounts(acts, c, rows, act_rw, y0, n_fy, fy_m)

    # 列（pt）
    widths = []
    if zn["puse"]:
        widths.append(12 * CH)
    widths += [20 * CH, 18 * CH]
    c0 = len(widths)
    mon_w = max(1.6, 72 / n_mon) * CH
    widths += [mon_w] * n_mon
    cm = len(widths)
    if use_amt:
        widths += [11 * CH] * (n_fy + 1)
    xs = [10.0]
    for w in widths:
        xs.append(xs[-1] + w)

    # 行（pt）
    levels = phase_levels(ph) if ph["names"] else {}
    n_lv = max(levels.values()) if levels else 0
    hs = [24, 16, 8] + [16] * n_lv + [16, 16] + [30 if rows[k]["mile"] else 22 for k in seq] + \
         ([22] if use_amt else []) + [10, 16]
    ys = [0.0]
    for h in hs:
        ys.append(ys[-1] + h)
    r_h1 = 3 + n_lv
    r_h2 = r_h1 + 1
    r_first = r_h2 + 1
    r_last = r_first + len(rows) - 1
    r_tot = r_last + 1 if use_amt else None
    r_note = (r_tot or r_last) + 2
    W, H = xs[-1] + 10, ys[-1] + 4

    fig = plt.figure(figsize=(W / 72, H / 72), dpi=150)
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_xlim(0, W)
    ax.set_ylim(H, 0)
    ax.axis("off")

    def cell(r1, c1, r2, c2):
        return xs[c1], ys[r1], xs[c2 + 1] - xs[c1], ys[r2 + 1] - ys[r1]

    def head(r1, c1, r2, c2, txt):
        x, y, w, h = cell(r1, c1, r2, c2)
        ax.add_patch(Rectangle((x, y), w, h, facecolor=NAVY, edgecolor="white", lw=0.5, zorder=1))
        ax.text(x + w / 2, y + h / 2, txt, fontsize=8, color="white", weight="bold", ha="center", va="center", zorder=2)

    def road_x(d):
        d = dt.date.fromordinal(int(d))
        mi = (d.year - y0) * 12 + d.month - fy_m
        return xs[c0] + (mi + (d.day - 1) / calendar.monthrange(d.year, d.month)[1]) * mon_w

    ax.text(10, ys[0] + 12, f"{title}　ロードマップ", fontsize=14, weight="bold", va="center")
    ax.text(10, ys[1] + 8, f"第1版　作成日：2026/11/05　工期：{span_text(min_s, max_e)}", fontsize=9, va="center",
            color="#333333")

    # 見出し
    cp = 0 if zn["puse"] else None
    ci = 1 if zn["puse"] else 0
    if cp is not None:
        head(r_h1, cp, r_h2, cp, "目的")
    head(r_h1, ci, r_h2, ci, "項目")
    head(r_h1, ci + 1, r_h2, ci + 1, "期間")
    for f in range(n_fy):
        head(r_h1, c0 + f * 12, r_h1, c0 + f * 12 + 11, f"{y0 + f}年度" if fy_m != 1 else f"{y0 + f}年")
        for k in range(12):
            x, y, w, h = cell(r_h2, c0 + f * 12 + k, r_h2, c0 + f * 12 + k)
            ax.add_patch(Rectangle((x, y), w, h, facecolor="#F2F2F2", edgecolor="none", zorder=0.5))
            ax.text(x + w / 2, y + h / 2, str((fy_m - 1 + k) % 12 + 1), fontsize=7, ha="center", va="center")
    if use_amt:
        head(r_h1, cm, r_h1, cm + n_fy, f"金額（{unit}）" if unit else "金額")
        for f in range(n_fy):
            head(r_h2, cm + f, r_h2, cm + f, f"{y0 + f}年度" if fy_m != 1 else f"{y0 + f}年")
        head(r_h2, cm + n_fy, r_h2, cm + n_fy, "合計")

    # 罫線
    r_end = r_tot if use_amt else r_last
    for k in range(n_mon + 1):
        x = xs[c0 + k]
        if k % 12 == 0:
            ax.plot([x, x], [ys[r_h1], ys[r_end + 1]], color=MID, lw=0.8, zorder=1.2)
        else:
            ax.plot([x, x], [ys[r_h2], ys[r_last + 1]], color=GRID, lw=0.5, zorder=0.6)
    for r in range(r_first, r_end + 1):
        ax.plot([xs[0], xs[-1]], [ys[r + 1], ys[r + 1]], color=GRID, lw=0.5, zorder=0.6)
    for cc in range(1, c0 + 1):  # 左の欄の区切り（合計の行は結合しているので、目盛りの始まりだけ）
        ax.plot([xs[cc], xs[cc]], [ys[r_h1], ys[(r_end if cc == c0 else r_last) + 1]], color=MID, lw=0.6, zorder=1.2)
    if use_amt:
        ax.plot([xs[cm], xs[cm]], [ys[r_h1], ys[r_end + 1]], color=DARK, lw=1.5, zorder=1.3)
    ax.add_patch(Rectangle((xs[0], ys[r_h1]), xs[-1] - xs[0], ys[r_end + 1] - ys[r_h1], facecolor="none",
                           edgecolor=DARK, lw=0.8, zorder=1.4))
    ax.plot([xs[0], xs[-1]], [ys[r_h2 + 1], ys[r_h2 + 1]], color=DARK, lw=0.8, zorder=1.4)

    # 項目の行
    for q, rw in enumerate(seq):
        r = r_first + q
        R = rows[rw]
        x, y, w, h = cell(r, ci, r, ci)
        ax.text(x + 6, y + h / 2, R["label"], fontsize=8, va="center")
        x, y, w, h = cell(r, ci + 1, r, ci + 1)
        ax.text(x + w / 2, y + h / 2, span_text(R["s"], R["e"]), fontsize=7.5, ha="center", va="center")
        if use_amt:
            for f in range(n_fy + 1):
                v = amt[rw][f] if f < n_fy else sum(amt[rw])
                if v:
                    x, y, w, h = cell(r, cm + f, r, cm + f)
                    ax.text(x + w - 4, y + h / 2, f"{v:,.0f}", fontsize=7.5, ha="right", va="center",
                            weight="bold" if f == n_fy else "normal")
        yc = ys[r] + (19 if R["mile"] else 11)  # マイルストーンのある行は帯を少し下に（上に名前）
        segs, miles = segments(acts, c, act_rw, rw)
        for s, e in segs:
            ax.add_patch(Rectangle((road_x(s), yc - 5), max(road_x(e) - road_x(s), 2), 10,
                                   facecolor=PALETTE[(R["purp"] - 1) % 8], edgecolor="none", zorder=3))
        for i in miles:
            x = road_x(S[i])
            ax.add_patch(Polygon([(x, yc - 5), (x + 5, yc), (x, yc + 5), (x - 5, yc)], closed=True, facecolor=MILE,
                                 edgecolor="none", zorder=4))
            nm = (acts[i].get("disp") or acts[i]["name"]).replace("\n", " ")
            txt = f"{nm} {md(S[i])}"
            w = v5.tw(txt, 8) + 6
            xl = min(x - 4, xs[c0 + n_mon] - w)  # ◆の上に、少し左から。右端からはみ出さない（VBA と同じ見積もり）
            ax.text(xl, yc - 11, txt, fontsize=7.5, color=MILE, weight="bold", va="center", zorder=4)
    if zn["puse"]:  # 目的（続く行を結合して塗る）
        q0 = 0
        for q in range(len(seq)):
            if q == len(seq) - 1 or rows[seq[q + 1]]["purp"] != rows[seq[q]]["purp"]:
                x, y, w, h = cell(r_first + q0, 0, r_first + q, 0)
                ax.add_patch(Rectangle((x, y), w, h, facecolor=PURP_FILL, edgecolor="none", zorder=0.8))
                ax.text(x + w / 2, y + h / 2, zn["pnames"][rows[seq[q]]["purp"] - 1], fontsize=8, weight="bold",
                        color=NAVY, ha="center", va="center", zorder=2)
                ax.plot([x, x + w], [y + h, y + h], color=MID, lw=0.6, zorder=1.2)
                q0 = q + 1
    if use_amt:
        x, y, w, h = cell(r_tot, 0, r_tot, len(widths) - 1)
        ax.add_patch(Rectangle((x, y), w, h, facecolor="#F2F2F2", edgecolor="none", zorder=0.5))
        ax.plot([x, x + w], [y + 1, y + 1], color=DARK, lw=0.6, zorder=1.5)
        ax.plot([x, x + w], [y - 1, y - 1], color=DARK, lw=0.6, zorder=1.5)
        x, y, w, h = cell(r_tot, 0, r_tot, c0 - 1)
        ax.text(x + w / 2, y + h / 2, "合計", fontsize=8, weight="bold", ha="center", va="center")
        for f in range(n_fy + 1):
            x, y, w, h = cell(r_tot, cm + f, r_tot, cm + f)
            ax.text(x + w - 4, y + h / 2, f"{tot[f]:,.0f}", fontsize=7.5, weight="bold", ha="right", va="center")

    # 期の帯
    if n_lv:
        ax.text(14, ys[3] + 8, "期", fontsize=8, va="center")
        for p, nm in enumerate(ph["names"], start=1):
            x1, x2 = road_x(ph["minS"][p]), road_x(ph["maxF"][p])
            t = ys[2 + levels[p]] + 2
            ax.add_patch(Rectangle((x1, t), max(x2 - x1, 6), 12, facecolor=PHASE_FILL, edgecolor="none", zorder=1))
            ax.text(x1 + 3, t + 6, f"{nm}（{span_text(ph['minS'][p], ph['last'][p])}）", fontsize=7.5, weight="bold",
                    color=NAVY, va="center", zorder=2)
    ax.text(10, ys[r_note] + 8, "帯＝作業のある期間（14日までのすき間はつなぐ）　◆＝マイルストーン（日数0の作業）" +
            ("　金額は年度ごとの工事する日の数で分けています" if use_amt else ""), fontsize=8, color=DARK, va="center")
    fig.savefig(path, dpi=150, facecolor="white")
    plt.close(fig)
    return dict(rows=[(rows[k]["label"], span_text(rows[k]["s"], rows[k]["e"])) for k in seq],
                totals=(tot if use_amt else None), n_fy=n_fy, y0=y0)


def big_acts():
    """4期・約150件（2026～2029年度）。まとめあり。金額は作業ごとに適当に付ける"""
    import big_sample
    rows = big_sample.make_rows(use_groups=True)
    starts = {f"{p}期": dt.date(2025 + p, 4, 1).toordinal() for p in range(1, 5)}
    acts = big_sample.make_acts(rows, starts)
    rnd = random.Random(7)
    for a in acts:
        a["amount"] = 0 if a["dur"] == 0 else rnd.choice([500, 800, 1200, 2000, 3500, 6000])
    return acts


if __name__ == "__main__":
    outdir = sys.argv[1]
    big = "--big" in sys.argv[2:]
    if big:
        v3.set_calendar({})
        acts = big_acts()
    else:
        v3.set_calendar(v3.load_sample_calendar())
        acts = v5.make_acts()
    c = compute(acts, is_hol=v3.is_hol)
    ph = collect_phases(acts, c)
    zn = collect_zones(acts)
    res = render(acts, c, ph, zn, f"{outdir}/roadmap{'_big' if big else ''}.png", "○○設備更新工事（サンプル）")
    print(f"行 {len(res['rows'])}  年度 {res['y0']}～{res['y0'] + res['n_fy'] - 1}")
    for lbl, sp in res["rows"]:
        print(f"  {lbl:16s} {sp}")
    if res["totals"]:
        print("  年度ごとの金額:", [f"{v:,.0f}" for v in res["totals"]])
