"""VBA（NetworkSchedule.bas）と同じ手順で、計算・段の割り当て・ダミーの引き方を再現し、
サンプルの完成イメージを描くためのスクリプト。"""
import datetime as dt
import json
import sys

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Circle, Rectangle

# 日本語フォント：入っているものを使う（Linux は Noto、Windows は游ゴシック・メイリオ）
from matplotlib import font_manager as _fm
_HAVE = {f.name for f in _fm.fontManager.ttflist}
plt.rcParams["font.family"] = next((f for f in ["Noto Sans CJK JP", "Yu Gothic", "Meiryo", "MS Gothic"]
                                    if f in _HAVE), "sans-serif")

# ---------------------------------------------------------------- サンプル
BASE = dt.date(2027, 1, 12)  # 着工日
SAMPLE = [
    # 工事, 作業ID, 作業名, 会社, 日数, 先行作業
    ("準備・仮設工事", "A-010", "現地調査・墨出し", "○○建設", 3, ""),
    ("準備・仮設工事", "A-020", "仮設足場・養生", "○○建設", 4, "A-010"),
    ("準備・仮設工事", "A-030", "足場解体・片付け", "○○建設", 2, "F-010"),
    ("撤去工事", "B-010", "停電・安全措置", "○○電設", 1, "A-020"),
    ("撤去工事", "B-020", "既設機器撤去", "○○建設", 5, "B-010"),
    ("撤去工事", "B-030", "既設配線撤去", "○○電設", 3, "B-010"),
    ("基礎・架台工事", "C-010", "基礎補修", "○○建設", 6, "B-020"),
    ("基礎・架台工事", "C-020", "架台設置", "○○建設", 4, "C-010"),
    ("機器据付工事", "D-010", "新設機器搬入", "○○電機", 2, "C-020"),
    ("機器据付工事", "D-020", "機器据付・芯出し", "○○電機", 5, "D-010"),
    ("電気工事", "E-010", "配線ルート整備", "○○電設", 5, "B-030"),
    ("電気工事", "E-020", "新設配線", "○○電設", 6, "E-010,D-020"),
    ("電気工事", "E-030", "結線", "○○電設", 3, "E-020"),
    ("試験・調整", "F-010", "単体試験", "○○電機", 3, "E-030"),
    ("試験・調整", "F-020", "総合試運転", "○○電機", 3, "F-010"),
    ("試験・調整", "F-030", "竣工検査", "○○建設", 1, "F-020,A-030"),
    ("試験・調整", "F-040", "引渡し", "○○建設", 0, "F-030"),
]


def parse_preds(s):
    s = (s or "").replace("，", ",").replace("、", ",").replace("　", "").replace(" ", "")
    return [t for t in s.split(",") if t]


def make_sample_dates():
    """サンプル用：最早で開始・終了日を入れる（取引先が日付を入れて返してきた想定）"""
    ids = {r[1]: i for i, r in enumerate(SAMPLE)}
    fin = {}
    out = []
    remaining = list(range(len(SAMPLE)))
    while remaining:
        for i in list(remaining):
            preds = parse_preds(SAMPLE[i][5])
            if all(p in fin for p in preds):
                s = max([fin[p] for p in preds], default=BASE.toordinal())
                d = SAMPLE[i][4]
                f = s if d == 0 else s + d
                fin[SAMPLE[i][1]] = f
                remaining.remove(i)
    for k, i in ids.items():
        pass
    for r in SAMPLE:
        preds = parse_preds(r[5])
        s = max([fin[p] for p in preds], default=BASE.toordinal())
        d = r[4]
        start = dt.date.fromordinal(s)
        end = dt.date.fromordinal(s) if d == 0 else dt.date.fromordinal(s + d - 1)
        out.append(dict(kouji=r[0], id=r[1], name=r[2], co=r[3], dur=d, pred=r[5],
                        start=start, end=end, prog=0.0))
    return out


# ---------------------------------------------------------------- VBAと同じ手順
def compute(acts, base_day=None, is_hol=None):
    """is_hol を渡すと VBA と同じく、余裕を工事する日の数で数え、矢線の終点 FD を休工日の先まで伸ばす"""
    if is_hol is None:
        def is_hol(d):
            return False

    def wdays(s, f):
        return sum(1 for d in range(s, f) if not is_hol(d))

    def from_end(e, dur):  # VBA の StartFromEnd
        d, cnt = e + 1, 0
        while cnt < dur:
            d -= 1
            if not is_hol(d):
                cnt += 1
        return d

    def next_work(d):
        while is_hol(d):
            d += 1
        return d

    n = len(acts)
    msgs = [[] for _ in range(n)]
    idx = {}
    for i, a in enumerate(acts):
        if a["id"] in idx:
            msgs[i].append("作業IDが重複しています")
        else:
            idx[a["id"]] = i
    preds = [[] for _ in range(n)]
    succs = [[] for _ in range(n)]
    for i, a in enumerate(acts):
        for t in parse_preds(a["pred"]):
            if t in idx:
                j = idx[t]
                if j != i:
                    preds[i].append(j)
                    succs[j].append(i)
            else:
                msgs[i].append(f"先行作業「{t}」が見つかりません")
    indeg = [len(p) for p in preds]
    topo = [i for i in range(n) if indeg[i] == 0]
    h = 0
    while h < len(topo):
        i = topo[h]
        h += 1
        for j in succs[i]:
            indeg[j] -= 1
            if indeg[j] == 0:
                topo.append(j)
    if len(topo) < n:
        raise SystemExit("循環あり")
    if base_day is None:
        given = [a["start"].toordinal() for a in acts if a["start"]]
        base_day = min(given) if given else dt.date.today().toordinal()
    S = [0] * n
    F = [0] * n
    for i in topo:
        a = acts[i]
        d = a["dur"]
        if a["start"]:
            S[i] = a["start"].toordinal()
        elif a["end"] and d is not None:
            S[i] = a["end"].toordinal() + 1 - d
        else:
            S[i] = max([F[p] for p in preds[i]] + [base_day])
        if d == 0:
            F[i] = S[i]
        elif a["end"]:
            F[i] = a["end"].toordinal() + 1
        elif d is not None:
            F[i] = S[i] + d
        else:
            F[i] = S[i] + 1
        if F[i] < S[i]:
            F[i] = S[i]
    pe = max(F)
    lateF = [0] * n
    TF = [0] * n
    def late_start(j):  # 後の作業の最も遅い開始：最も遅い終わりから、工事する日の数だけさかのぼる
        w = wdays(S[j], F[j])
        return from_end(lateF[j] - 1, w) if w > 0 else lateF[j]

    for i in reversed(topo):
        if succs[i]:
            lf = min(late_start(j) for j in succs[i])
        else:
            lf = pe
        lateF[i] = lf
        TF[i] = wdays(F[i], lf) if lf >= F[i] else lf - F[i]
    crit = [t <= 0 for t in TF]
    FD = [next_work(F[i]) if F[i] > S[i] else F[i] for i in range(n)]  # 矢線の終点
    for i in range(n):
        for p in preds[i]:
            if S[i] < F[p]:
                msgs[i].append(f"先行作業「{acts[p]['id']}」の終了前に開始しています")
    return dict(n=n, preds=preds, succs=succs, S=S, F=F, FD=FD, TF=TF, crit=crit, msgs=msgs, topo=topo)


def assign_lanes(acts, c):
    n, S, F, crit, preds = c["n"], c["S"], c["F"], c["crit"], c["preds"]
    bands, band_of = [], []
    for a in acts:
        if a["kouji"] not in bands:
            bands.append(a["kouji"])
        band_of.append(bands.index(a["kouji"]))

    def before(a, b):  # VBAの ComesBefore と同じ：開始が早い→クリティカル優先→終了が早い
        if S[a] != S[b]:
            return S[a] < S[b]
        if crit[a] != crit[b]:
            return crit[a]
        return F[a] < F[b]

    lane_band, lane_lastF, lane_lastAct, occ = [], [], [], []
    act_lane = [0] * n
    band_first, band_last = [], []
    for b in range(len(bands)):
        members = []
        for i in range(n):
            if band_of[i] == b:
                members.append(i)
                k = len(members) - 1
                while k > 0 and before(members[k], members[k - 1]):
                    members[k - 1], members[k] = members[k], members[k - 1]
                    k -= 1
        first = len(lane_band)
        for i in members:
            placed = False
            for L in range(first, len(lane_band)):
                ok = lane_lastF[L] < S[i] or (lane_lastF[L] == S[i] and lane_lastAct[L] in preds[i])
                if ok:
                    act_lane[i] = L
                    lane_lastF[L] = F[i]
                    lane_lastAct[L] = i
                    occ[L].append((S[i], F[i]))
                    placed = True
                    break
            if not placed:
                lane_band.append(b)
                lane_lastF.append(F[i])
                lane_lastAct.append(i)
                occ.append([(S[i], F[i])])
                act_lane[i] = len(lane_band) - 1
        band_first.append(first)
        band_last.append(len(lane_band) - 1)
    return dict(bands=bands, band_first=band_first, band_last=band_last,
                act_lane=act_lane, occ=occ, nlane=len(lane_band))


def register_nodes(c, lanes):
    nodes = {}
    for i in range(c["n"]):
        for d in (c["S"][i], c["F"][i]):
            key = (lanes["act_lane"][i], d)
            nodes[key] = nodes.get(key, False) or c["crit"][i]
    return nodes


def lane_free(occ, L, a, b):
    return not any(v0 < b and a < v1 for v0, v1 in occ[L])


def route_links(acts, c, lanes, nodes):
    """戻り値：描画命令のリスト（kind, args）。座標は (日, 段) 単位。"""
    out = []
    S, F, crit, preds = c["S"], c["F"], c["crit"], c["preds"]
    occ = lanes["occ"]
    lane = lanes["act_lane"]
    for i in range(c["n"]):
        for p in preds[i]:
            La, Lb, fa, sb = lane[p], lane[i], F[p], S[i]
            if fa > sb:
                out.append(("diag", fa, La, sb, Lb, "red"))
            elif La == Lb:
                if fa < sb:
                    if lane_free(occ, La, fa, sb):
                        out.append(("float", fa, sb, La, True, True))
                        occ[La].append((fa, sb))
                    else:
                        out.append(("bridge", fa, sb, La))
            else:
                red = crit[p] and crit[i] and fa == sb
                if fa == sb:
                    out.append(("vert", fa, La, Lb, True, True, "red" if red else "gray"))
                elif lane_free(occ, La, fa, sb) and (La, sb) not in nodes:
                    out.append(("float", fa, sb, La, True, False))
                    occ[La].append((fa, sb))
                    out.append(("vert", sb, La, Lb, False, True, "gray"))
                elif lane_free(occ, Lb, fa, sb) and (Lb, fa) not in nodes:
                    out.append(("vert", fa, La, Lb, True, False, "gray"))
                    out.append(("float", fa, sb, Lb, False, True))
                    occ[Lb].append((fa, sb))
                else:
                    out.append(("diag", fa, La, sb, Lb, "gray"))
    return out


# ---------------------------------------------------------------- ラベル（VBAの NameLabel / SubLabel / TextWidth と同じ）
def text_w(txt, fs):
    w = 0.0
    for ch in txt:
        code = ord(ch)
        w += fs * 0.6 if (code < 0x2000 or 0xFF61 <= code <= 0xFF9F) else fs
    return w


def name_label(a):
    return a["id"] + (" " + a["name"] if a["name"] else "")


def sub_label(a, s, f):
    t = a["co"]
    if f > s:
        t += ("・" if t else "") + f"{f - s}日"
    if a["prog"] > 0:
        t += ("・" if t else "") + f"{a['prog']:.0%}"
    return t


def label_right_days(acts, c, lanes, view_start, day_w):
    S, F = c["S"], c["F"]
    nr = [-1e9] * lanes["nlane"]
    sr = [-1e9] * lanes["nlane"]
    mx = 0.0
    for i in sorted(range(len(acts)), key=lambda k: (S[k], F[k])):
        L = lanes["act_lane"][i]
        cx = ((S[i] - view_start) + (F[i] - view_start)) / 2
        w = (text_w(name_label(acts[i]), 8) + 6) / day_w
        lft = max(cx - w / 2, nr[L] + 2 / day_w)
        nr[L] = lft + w
        mx = max(mx, nr[L])
        sub = sub_label(acts[i], S[i], F[i])
        if sub:
            w = (text_w(sub, 7) + 6) / day_w
            lft = max(cx - w / 2, sr[L] + 2 / day_w)
            sr[L] = lft + w
            mx = max(mx, sr[L])
    return mx


# ---------------------------------------------------------------- 完成イメージ
def render(acts, c, lanes, nodes, links, path, title):
    DAY_W, LANE_H, ARROW_OFF, NODE_D = 24.0, 44.0, 24.0, 12.0
    R = NODE_D / 2
    LABEL_W = 110.0
    S, F = c["S"], c["F"]
    view_start = min(S) - 2
    view_end = max(F) + 2
    ncols = max(view_end - view_start + 1, 40)
    ncols = max(ncols, int(label_right_days(acts, c, lanes, view_start, DAY_W)) + 2)
    TOP_TITLE = 46
    ROW_M, ROW_D, ROW_W = 16, 14, 14
    hdr_top = TOP_TITLE
    lanes_top = hdr_top + ROW_M + ROW_D + ROW_W
    nlane = lanes["nlane"]
    W = LABEL_W + ncols * DAY_W + 20
    H = lanes_top + nlane * LANE_H + 80
    fig = plt.figure(figsize=(W / 72, H / 72), dpi=150)
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_xlim(0, W)
    ax.set_ylim(H, 0)
    ax.set_aspect("equal")
    ax.axis("off")
    X0 = 10.0

    def X(d):
        return X0 + LABEL_W + (d - view_start) * DAY_W

    def Y(L):
        return lanes_top + L * LANE_H + ARROW_OFF

    gray, red, dark = "#595959", "#C00000", "#404040"
    ax.text(X0, 18, f"{title}　ネットワーク工程表", fontsize=14, weight="bold", va="center")
    ax.text(X0, 34, "第1版　作成日：2026/11/05　目盛り：日単位", fontsize=9, va="center", color="#333333")
    chart_bottom = lanes_top + nlane * LANE_H
    week = "日月火水木金土"
    prev_month = None
    for k in range(ncols):
        d = dt.date.fromordinal(view_start + k)
        x = X(view_start + k)
        wd = (d.weekday() + 1) % 7  # 0=日
        if wd in (0, 6):
            ax.add_patch(Rectangle((x, hdr_top + ROW_M), DAY_W, chart_bottom - hdr_top - ROW_M,
                                   facecolor="#F2F2F2", edgecolor="none", zorder=0))
        if d.month != prev_month:
            ax.text(x + 2, hdr_top + ROW_M / 2, f"{d.year}/{d.month}", fontsize=8, va="center")
            ax.plot([x, x], [hdr_top, chart_bottom], color="#969696", lw=0.8, zorder=1)
            prev_month = d.month
        ax.text(x + DAY_W / 2, hdr_top + ROW_M + ROW_D / 2, str(d.day), fontsize=7, ha="center", va="center")
        ax.text(x + DAY_W / 2, hdr_top + ROW_M + ROW_D + ROW_W / 2, week[wd], fontsize=7, ha="center",
                va="center", color="#C00000" if wd == 0 else "#000000")
    ax.plot([X0, X(view_start + ncols)], [lanes_top, lanes_top], color="#7F7F7F", lw=0.8)
    ax.plot([X0 + LABEL_W, X0 + LABEL_W], [hdr_top, chart_bottom], color="#7F7F7F", lw=0.8)
    ax.text(X0 + 4, hdr_top + ROW_M + ROW_D / 2, "工事", fontsize=8, va="center")
    for b, name in enumerate(lanes["bands"]):
        t = lanes_top + lanes["band_first"][b] * LANE_H
        btm = lanes_top + (lanes["band_last"][b] + 1) * LANE_H
        ax.text(X0 + 6, (t + btm) / 2, name, fontsize=9, va="center")
        ax.plot([X0, X(view_start + ncols)], [btm, btm], color="#BFBFBF", lw=0.8, zorder=1)

    def arrow(x1, y1, x2, y2, color, lw, ls="-", head=True):
        if abs(x2 - x1) < 0.5 and abs(y2 - y1) < 0.5:
            return
        if head:
            ax.annotate("", xy=(x2, y2), xytext=(x1, y1), zorder=3,
                        arrowprops=dict(arrowstyle="-|>", color=color, lw=lw, linestyle=ls,
                                        shrinkA=0, shrinkB=0, mutation_scale=7))
        else:
            ax.plot([x1, x2], [y1, y2], color=color, lw=lw, linestyle=ls, zorder=3)

    def tw(txt, fs):  # VBAの TextWidth と同じ見積もり
        w = 0.0
        for ch in txt:
            code = ord(ch)
            w += fs * 0.6 if (code < 0x2000 or 0xFF61 <= code <= 0xFF9F) else fs
        return w

    name_right = [-1e9] * nlane
    sub_right = [-1e9] * nlane
    for i in sorted(range(len(acts)), key=lambda k: (S[k], F[k])):
        a = acts[i]
        L = lanes["act_lane"][i]
        y = Y(L)
        xs, xf = X(S[i]), X(F[i])
        cr = c["crit"][i]
        if xf - xs > NODE_D + 1:
            arrow(xs + R, y, xf - R, y, red if cr else dark, 2.0 if cr else 1.1)
        cx = (xs + xf) / 2
        name_text = name_label(a)
        w = tw(name_text, 8) + 6
        left = cx - w / 2
        if left < name_right[L] + 2:
            left = name_right[L] + 2
        name_right[L] = left + w
        ax.text(left + w / 2, y - R - 1, name_text, fontsize=8, ha="center", va="bottom",
                color=red if cr else "#000000", zorder=5)
        sub = sub_label(a, S[i], F[i])
        if sub:
            w = tw(sub, 7) + 6
            left = cx - w / 2
            if left < sub_right[L] + 2:
                left = sub_right[L] + 2
            sub_right[L] = left + w
            ax.text(left + w / 2, y + R + 0.5, sub, fontsize=7, ha="center", va="top", color=dark, zorder=5)

    for lk in links:
        kind = lk[0]
        if kind == "vert":
            _, d, La, Lb, from_node, to_node, col = lk
            y1, y2 = Y(La), Y(Lb)
            sg = 1 if y2 >= y1 else -1
            if from_node:
                y1 += sg * R
            if to_node:
                y2 -= sg * R
            arrow(X(d), y1, X(d), y2, red if col == "red" else gray, 1.3 if col == "red" else 0.9, "--")
        elif kind == "float":
            _, fa, sb, L, from_node, to_node = lk
            x1 = X(fa) + (R if from_node else 0)
            x2 = X(sb) - (R if to_node else 0)
            arrow(x1, Y(L), x2, Y(L), gray, 0.9, ":", head=False)
        elif kind == "bridge":
            _, fa, sb, L = lk
            y = Y(L) - R - 8
            arrow(X(fa), Y(L) - R, X(fa), y, gray, 0.9, "--", False)
            arrow(X(fa), y, X(sb), y, gray, 0.9, "--", False)
            arrow(X(sb), y, X(sb), Y(L) - R, gray, 0.9, "--")
        elif kind == "diag":
            _, fa, La, sb, Lb, col = lk
            x1, y1, x2, y2 = X(fa), Y(La), X(sb), Y(Lb)
            ln = ((x2 - x1) ** 2 + (y2 - y1) ** 2) ** 0.5
            ux, uy = (x2 - x1) / ln, (y2 - y1) / ln
            arrow(x1 + ux * R, y1 + uy * R, x2 - ux * R, y2 - uy * R,
                  red if col == "red" else gray, 0.9, "--")

    order = sorted(nodes.keys(), key=lambda k: (k[1], k[0]))
    for num, key in enumerate(order, start=1):
        L, d = key
        cr = nodes[key]
        ax.add_patch(Circle((X(d), Y(L)), R, facecolor="white", edgecolor=red if cr else dark,
                            lw=0.9, zorder=6))
        ax.text(X(d), Y(L), str(num), fontsize=6 if num < 100 else 5, ha="center", va="center", zorder=7)

    ly = chart_bottom + 30
    ax.text(X0 + 4, ly, "凡例", fontsize=8, va="center")
    x = X0 + LABEL_W
    arrow(x, ly, x + 36, ly, dark, 1.1)
    ax.text(x + 42, ly, "作業（矢線）", fontsize=8, va="center")
    x += 130
    ax.add_patch(Circle((x + R, ly), R, facecolor="white", edgecolor=dark, lw=0.9, zorder=6))
    ax.text(x + NODE_D + 6, ly, "結合点（番号付きの○）", fontsize=8, va="center")
    x += 150
    arrow(x, ly, x + 36, ly, gray, 0.9, "--")
    ax.text(x + 42, ly, "ダミー（他の工事とのつながり）", fontsize=8, va="center")
    x += 190
    arrow(x, ly, x + 36, ly, red, 2.0)
    ax.text(x + 42, ly, "クリティカルパス", fontsize=8, va="center")
    x += 130
    arrow(x, ly, x + 36, ly, gray, 0.9, ":", head=False)
    ax.text(x + 42, ly, "余裕日数", fontsize=8, va="center")
    ax.text(X0 + LABEL_W + 4, ly + 21, "表記：矢線の上＝作業ID・作業名、矢線の下＝会社・日数（進捗があれば％も）",
            fontsize=8, va="center", color=dark)
    fig.savefig(path, dpi=150, facecolor="white")
    return len(order)


if __name__ == "__main__":
    out_png = sys.argv[1]
    out_json = sys.argv[2]
    acts = make_sample_dates()
    c = compute(acts)
    lanes = assign_lanes(acts, c)
    nodes = register_nodes(c, lanes)
    links = route_links(acts, c, lanes, nodes)
    for i, a in enumerate(acts):
        print(f"{a['id']} {a['kouji']:<8} {a['name']:<10} S={dt.date.fromordinal(c['S'][i])} "
              f"E={dt.date.fromordinal(c['F'][i] - 1) if c['F'][i] > c['S'][i] else '-'} "
              f"TF={c['TF'][i]:>2} CP={'●' if c['crit'][i] else ' '} lane={lanes['act_lane'][i]} "
              f"{' / '.join(c['msgs'][i])}")
    print("lanes:", lanes["nlane"], "bands:", lanes["bands"])
    for lk in links:
        print("  link", lk)
    nn = render(acts, c, lanes, nodes, links, out_png, "○○設備更新工事（サンプル）")
    print("nodes:", nn)
    with open(out_json, "w", encoding="utf-8") as fp:
        json.dump([dict(a, start=a["start"].isoformat(), end=a["end"].isoformat()) for a in acts],
                  fp, ensure_ascii=False, indent=1)
