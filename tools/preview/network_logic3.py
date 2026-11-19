"""NetworkSchedule.bas（まとめ・会社の色・休工日 対応版）と同じ手順を再現し、完成イメージを描く。"""
import datetime as dt
import json
import sys

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Circle, Rectangle

from network_logic import compute, make_sample_dates, text_w

# 日本語フォント：入っているものを使う（Linux は Noto、Windows は游ゴシック・メイリオ）
from matplotlib import font_manager as _fm
_HAVE = {f.name for f in _fm.fontManager.ttflist}
plt.rcParams["font.family"] = next((f for f in ["Noto Sans CJK JP", "Yu Gothic", "Meiryo", "MS Gothic"]
                                    if f in _HAVE), "sans-serif")

KIND_ACT, KIND_FROM, KIND_TO, KIND_SUM = 0, 1, 2, 3
PHASE_OF = {
    "A-010": "1期", "A-020": "1期", "B-010": "1期", "B-020": "1期", "B-030": "1期",
    "C-010": "1期", "C-020": "1期", "E-010": "1期",
    "A-030": "2期", "D-010": "2期", "D-020": "2期", "E-020": "2期", "E-030": "2期",
    "F-010": "2期", "F-020": "2期", "F-030": "2期", "F-040": "2期",
}
GROUP_OF = {
    "A-010": "準備・仮設一式", "A-020": "準備・仮設一式",
    "B-010": "既設撤去一式", "B-020": "既設撤去一式", "B-030": "既設撤去一式",
    "C-010": "基礎・架台一式", "C-020": "基礎・架台一式",
    "D-010": "機器据付一式", "D-020": "機器据付一式",
    "E-010": "電気工事一式", "E-020": "電気工事一式", "E-030": "電気工事一式",
    "F-010": "試験・検査一式", "F-020": "試験・検査一式", "F-030": "試験・検査一式",
}
HOLIDAYS = [(dt.date(2026, 12, 29) + dt.timedelta(days=k), "年末年始") for k in range(6)] + \
           [(dt.date(2027, 1, 11), "成人の日（例）")]
CO_SETTING = {"○○建設": "#2F5597", "○○電設": "#38761D", "○○電機": "#7030A0"}
PALETTE = ["#2F5597", "#38761D", "#7030A0", "#8C5A2B", "#1B7F86", "#B4407E", "#1F3864", "#7F7F00"]
NEUTRAL, GHOST_C, RED, GRAY, BAND, HOL_FILL = "#404040", "#767676", "#C00000", "#595959", "#FF7C80", "#FCE4D6"

HOL = set()


def is_hol(d):
    return d in HOL


def next_work(d):
    while is_hol(d):
        d += 1
    return d


def add_work(s, dur):
    d, cnt = s, 0
    while cnt < dur:
        if not is_hol(d):
            cnt += 1
        d += 1
    return d


def start_from_end(e, dur):
    if dur <= 0:
        return e
    d, cnt = e + 1, 0
    while cnt < dur:
        d -= 1
        if not is_hol(d):
            cnt += 1
    return d


def last_day(c, i):
    return c["F"][i] - 1 if c["F"][i] > c["S"][i] else c["S"][i]


def check_dates(acts, c):
    msgs = {}
    for i, a in enumerate(acts):
        if a["start"] and a["end"] and c["F"][i] > c["S"][i]:
            hits = [d for d in range(c["S"][i], c["F"][i]) if is_hol(d)]
            m = []
            if hits:
                txt = "・".join(f"{dt.date.fromordinal(d).month}/{dt.date.fromordinal(d).day}" for d in hits[:3])
                m.append(f"休工日（{txt}{'ほか' if len(hits) > 3 else ''}）にかかっています")
            if a["dur"] and a["dur"] > 0 and c["F"][i] - c["S"][i] != a["dur"]:
                m.append(f"日数（{a['dur']}日）と開始〜終了（{c['F'][i] - c['S'][i]}日）が合いません")
            if m:
                msgs[a["id"]] = m
    return msgs


def collect_phases(acts, c):
    names, mins = [], {}
    for i, a in enumerate(acts):
        ph = a.get("phase", "")
        if ph:
            if ph not in mins:
                names.append(ph)
                mins[ph] = c["S"][i]
            else:
                mins[ph] = min(mins[ph], c["S"][i])
    order = sorted(names, key=lambda n: mins[n])
    idx = {n: k + 1 for k, n in enumerate(order)}
    act_phase = [idx.get(a.get("phase", ""), 0) for a in acts]
    minS = {p: None for p in idx.values()}
    maxF = {p: 0 for p in idx.values()}
    lastd = {p: 0 for p in idx.values()}
    for i, p in enumerate(act_phase):
        if p:
            minS[p] = c["S"][i] if minS[p] is None else min(minS[p], c["S"][i])
            maxF[p] = max(maxF[p], c["F"][i])
            lastd[p] = max(lastd[p], last_day(c, i))
    return dict(names=order, act_phase=act_phase, minS=minS, maxF=maxF, last=lastd)


def collect_groups(acts, c):
    names, gidx = [], []
    for a in acts:
        g = a.get("group", "")
        if g and g not in names:
            names.append(g)
        gidx.append(names.index(g) + 1 if g else 0)
    G = {}
    for g in range(1, len(names) + 1):
        mem = [i for i in range(len(acts)) if gidx[i] == g]
        cos = []
        for i in mem:
            if acts[i]["co"] and acts[i]["co"] not in cos:
                cos.append(acts[i]["co"])
        w = sum(c["F"][i] - c["S"][i] for i in mem)
        p = sum(acts[i]["prog"] * (c["F"][i] - c["S"][i]) for i in mem)
        G[g] = dict(name=names[g - 1], kouji=acts[mem[0]]["kouji"],
                    S=min(c["S"][i] for i in mem), F=max(c["F"][i] for i in mem),
                    crit=any(c["crit"][i] for i in mem),
                    co_label=(cos[0] if len(cos) == 1 else (f"{cos[0]}ほか{len(cos) - 1}社" if cos else "")),
                    co_name=(cos[0] if len(cos) == 1 else ""),
                    prog=(p / w if w else 0))
    return dict(idx=gidx, G=G)


def company_colors(acts):
    colors = dict(CO_SETTING)
    order, used = [], 0
    for a in acts:
        co = a["co"]
        if co:
            if co not in colors:
                colors[co] = PALETTE[used % len(PALETTE)]
                used += 1
            if co not in order:
                order.append(co)
    return colors, order


def build_sheet(acts, c, ph, gr, mode, day_w=24.0):
    """mode：0＝全体工程、-1＝詳細工程、1以上＝その期"""
    S, F, crit, preds, succs = c["S"], c["F"], c["crit"], c["preds"], c["succs"]
    n = len(acts)
    overall = (mode == 0)
    use_sum = overall and len(gr["G"]) > 0
    items, act_item, g_item = [], {}, {}
    for i in range(n):
        if mode <= 0 or ph["act_phase"][i] == mode:
            g = gr["idx"][i]
            if use_sum and g:
                if g not in g_item:
                    g_item[g] = len(items)
                    items.append(dict(act=i, kind=KIND_SUM, group=g, s=gr["G"][g]["S"], f=gr["G"][g]["F"], lane=None))
                act_item[i] = g_item[g]
            else:
                act_item[i] = len(items)
                items.append(dict(act=i, kind=KIND_ACT, group=0, s=S[i], f=F[i], lane=None))
    minS = min(it["s"] for it in items)
    maxF = max(it["f"] for it in items)
    vs, ve = minS - 2, maxF + 2

    def clamp(d):
        return min(max(d, vs + 1), ve)

    ghost = {}
    if mode > 0:
        for k in range(len(items)):
            a = items[k]["act"]
            for v in preds[a]:
                if v not in act_item and (KIND_FROM, v) not in ghost:
                    ghost[(KIND_FROM, v)] = len(items)
                    items.append(dict(act=v, kind=KIND_FROM, group=0, s=clamp(F[v]), f=clamp(F[v]), lane=None))
            for v in succs[a]:
                if v not in act_item and (KIND_TO, v) not in ghost:
                    ghost[(KIND_TO, v)] = len(items)
                    items.append(dict(act=v, kind=KIND_TO, group=0, s=clamp(S[v]), f=clamp(S[v]), lane=None))

    def icrit(it):
        if it["kind"] == KIND_SUM:
            return gr["G"][it["group"]]["crit"]
        if it["kind"] == KIND_ACT:
            return crit[it["act"]]
        return False

    # つながり（作業の並び順で、同じ項目どうしは余裕の最も少ない1本に）
    links, pair = [], {}

    def add_link(fk, tk, fa, sb, cr):
        key = (fk, tk)
        if key in pair:
            j = pair[key]
            L0 = links[j]
            if (sb - fa) < (L0["b"] - L0["a"]) or ((sb - fa) == (L0["b"] - L0["a"]) and fa > L0["a"]):
                L0.update(a=fa, b=sb, crit=cr)
        else:
            pair[key] = len(links)
            links.append(dict(f=fk, t=tk, a=fa, b=sb, crit=cr))

    for a in range(n):
        if a in act_item:
            tk = act_item[a]
            for v in preds[a]:
                if v in act_item:
                    fk, fa = act_item[v], F[v]
                else:
                    fk = ghost[(KIND_FROM, v)]
                    fa = items[fk]["f"]
                if fk != tk:
                    add_link(fk, tk, fa, S[a], crit[v] and crit[a] and F[v] == S[a])
            for v in succs[a]:
                if v not in act_item:
                    k = ghost[(KIND_TO, v)]
                    add_link(act_item[a], k, F[a], items[k]["s"], crit[a] and crit[v] and F[a] == S[v])

    def shares_node(lk_, k):
        return any(L0["f"] == lk_ and L0["t"] == k and L0["a"] == items[lk_]["f"] and L0["b"] == items[k]["s"]
                   for L0 in links)

    # 段
    order_all = sorted(range(len(items)), key=lambda k: (items[k]["s"], items[k]["f"]))

    def ghost_w(it):
        return max(text_w(name_label(acts, ph, gr, it), 8) + 6, text_w(sub_label(acts, c, gr, it), 7) + 6)

    def stack(kind, first):
        right = []
        for k in order_all:
            it = items[k]
            if it["kind"] != kind:
                continue
            lft = (it["s"] - vs) * day_w - 6
            for L, r in enumerate(right):
                if r + 2 <= lft:
                    break
            else:
                right.append(None)
                L = len(right) - 1
            it["lane"] = first + L
            right[L] = lft + 12 + 2 + ghost_w(it)
        return len(right)

    def ikouji(it):
        return gr["G"][it["group"]]["kouji"] if it["kind"] == KIND_SUM else acts[it["act"]]["kouji"]

    bands, band_of = [], {}
    for k, it in enumerate(items):
        if it["kind"] in (KIND_ACT, KIND_SUM):
            kk = ikouji(it)
            if kk not in bands:
                bands.append(kk)
            band_of[k] = bands.index(kk)

    occ, lastF, lastK = [], [], []
    n_from = stack(KIND_FROM, 0)
    for _ in range(n_from):
        occ.append([]); lastF.append(None); lastK.append(None)

    def before(x, y):
        ix, iy = items[x], items[y]
        if ix["s"] != iy["s"]:
            return ix["s"] < iy["s"]
        if icrit(ix) != icrit(iy):
            return icrit(ix)
        return ix["f"] < iy["f"]

    band_first, band_last = [], []
    for b in range(len(bands)):
        members = []
        for k, it in enumerate(items):
            if it["kind"] in (KIND_ACT, KIND_SUM) and band_of[k] == b:
                members.append(k)
                m = len(members) - 1
                while m > 0 and before(members[m], members[m - 1]):
                    members[m - 1], members[m] = members[m], members[m - 1]
                    m -= 1
        first = len(occ)
        for k in members:
            it = items[k]
            placed = False
            for L in range(first, len(occ)):
                if lastF[L] < it["s"] or (lastF[L] == it["s"] and shares_node(lastK[L], k)):
                    it["lane"] = L
                    lastF[L] = it["f"]; lastK[L] = k
                    occ[L].append((it["s"], it["f"]))
                    placed = True
                    break
            if not placed:
                occ.append([(it["s"], it["f"])]); lastF.append(it["f"]); lastK.append(k)
                it["lane"] = len(occ) - 1
        band_first.append(first)
        band_last.append(len(occ) - 1)
    first_to = len(occ)
    n_to = stack(KIND_TO, first_to)
    for _ in range(n_to):
        occ.append([]); lastF.append(None); lastK.append(None)

    # 結合点（項目の始点・終点と、つながりの付く点）
    nodes = {}

    def node(L, d, cr, g):
        if (L, d) in nodes:
            nodes[(L, d)]["crit"] = nodes[(L, d)]["crit"] or cr
        else:
            nodes[(L, d)] = dict(crit=cr, ghost=g)

    def sum_end_crit(it, at_start):
        mem = [i for i in range(n) if gr["idx"][i] == it["group"]]
        if at_start:
            return any(crit[i] and S[i] == it["s"] for i in mem)
        return any(crit[i] and F[i] == it["f"] for i in mem)

    for it in items:
        g = it["kind"] in (KIND_FROM, KIND_TO)
        if it["kind"] == KIND_SUM:
            node(it["lane"], it["s"], sum_end_crit(it, True), False)
            node(it["lane"], it["f"], sum_end_crit(it, False), False)
            continue
        cr = icrit(it) and not g
        node(it["lane"], it["s"], cr, g)
        node(it["lane"], it["f"], cr, g)
        if g:
            occ[it["lane"]].append((it["s"], it["f"]))
    for L0 in links:
        node(items[L0["f"]]["lane"], L0["a"], L0["crit"], False)
        node(items[L0["t"]]["lane"], L0["b"], L0["crit"], False)

    # 線の引き方（VBAの LinkItems と同じ判定）
    def free(L, a, b):
        return not any(v0 < b and a < v1 for v0, v1 in occ[L])

    def vclear(d, La, Lb):
        lo, hi = min(La, Lb), max(La, Lb)
        return not any((L, d) in nodes for L in range(lo + 1, hi))

    prims = []
    for L0 in links:
        La, Lb, fa, sb = items[L0["f"]]["lane"], items[L0["t"]]["lane"], L0["a"], L0["b"]
        if fa > sb:
            prims.append(("diag", fa, La, sb, Lb, "red"))
        elif La == Lb:
            if fa < sb:
                if free(La, fa, sb):
                    prims.append(("float", fa, sb, La, True, True)); occ[La].append((fa, sb))
                else:
                    prims.append(("bridge", fa, sb, La))
        else:
            col = "red" if (L0["crit"] and fa == sb) else "gray"
            if fa == sb:
                prims.append(("vert", fa, La, Lb, True, True, col))
            elif free(La, fa, sb) and (La, sb) not in nodes and vclear(sb, La, Lb):
                prims.append(("float", fa, sb, La, True, False)); occ[La].append((fa, sb))
                prims.append(("vert", sb, La, Lb, False, True, "gray"))
            elif free(Lb, fa, sb) and (Lb, fa) not in nodes and vclear(fa, La, Lb):
                prims.append(("vert", fa, La, Lb, True, False, "gray"))
                prims.append(("float", fa, sb, Lb, False, True)); occ[Lb].append((fa, sb))
            else:
                prims.append(("diag", fa, La, sb, Lb, "gray"))

    def seg_crit(it, e1, e2):
        if it["kind"] == KIND_ACT:
            return crit[it["act"]]
        return any(crit[i] and S[i] < e2 and F[i] > e1 for i in range(n) if gr["idx"][i] == it["group"])

    return dict(seg_crit=seg_crit, mode=mode, overall=overall, items=items, bands=bands, band_first=band_first, band_last=band_last,
                nlane=len(occ), from_rng=(0, n_from - 1) if n_from else None,
                to_rng=(first_to, first_to + n_to - 1) if n_to else None,
                nodes=nodes, prims=prims, links=links, view_start=vs, view_end=ve, icrit=icrit)


def phase_label(ph, i):
    p = ph["act_phase"][i]
    return ph["names"][p - 1] if p else "期未入力"


def name_label(acts, ph, gr, it):
    if it["kind"] == KIND_SUM:
        return gr["G"][it["group"]]["name"]
    a = acts[it["act"]]
    t = a["id"] + (" " + a["name"] if a["name"] else "")
    if it["kind"] != KIND_ACT:
        t += "（" + phase_label(ph, it["act"]) + "）"
    return t


def sub_label(acts, c, gr, it):
    i = it["act"]
    if it["kind"] == KIND_FROM:
        d = dt.date.fromordinal(last_day(c, i))
        return f"{d.month}/{d.day}終了"
    if it["kind"] == KIND_TO:
        d = dt.date.fromordinal(c["S"][i])
        return f"{d.month}/{d.day}開始"
    if it["kind"] == KIND_SUM:
        g = gr["G"][it["group"]]
        t = g["co_label"]
        if g["F"] > g["S"]:
            t += ("・" if t else "") + f"{g['F'] - g['S']}日"
        if g["prog"] > 0:
            t += ("・" if t else "") + f"{g['prog']:.0%}"
        return t
    a = acts[i]
    t = a["co"]
    if c["F"][i] > c["S"][i]:
        t += ("・" if t else "") + f"{c['F'][i] - c['S'][i]}日"
    if a["prog"] > 0:
        t += ("・" if t else "") + f"{a['prog']:.0%}"
    return t


def item_color(acts, gr, colors, it):
    if it["kind"] == KIND_SUM:
        nm = gr["G"][it["group"]]["co_name"]
        return colors.get(nm, NEUTRAL) if nm else NEUTRAL
    if it["kind"] == KIND_ACT:
        co = acts[it["act"]]["co"]
        return colors.get(co, NEUTRAL) if co else NEUTRAL
    return GHOST_C


def phase_levels(ph):
    lv = {}
    for p in range(1, len(ph["names"]) + 1):
        L = 1
        while any(lv[q] == L and ph["minS"][q] < ph["maxF"][p] and ph["minS"][p] < ph["maxF"][q]
                  for q in range(1, p)):
            L += 1
        lv[p] = L
    return lv


def render(acts, c, ph, gr, sh, colors, co_order, path, title, caption):
    DAY_W, LANE_H, ARROW_OFF, NODE_D = 24.0, 44.0, 24.0, 12.0
    R = NODE_D / 2
    LABEL_W, X0 = 110.0, 10.0
    items = sh["items"]
    vs = sh["view_start"]
    icrit = sh["icrit"]
    order = sorted(range(len(items)), key=lambda k: (items[k]["s"], items[k]["f"]))

    def label_base(it, cx_days, w_days):
        return cx_days - w_days / 2 if it["kind"] in (KIND_ACT, KIND_SUM) else cx_days + (R + 2) / DAY_W

    ncols = max(sh["view_end"] - vs + 1, 40)
    nr, sr, mx = {}, {}, 0.0
    for k in order:
        it = items[k]
        L = it["lane"]
        cx = ((it["s"] - vs) + (it["f"] - vs)) / 2
        w = (text_w(name_label(acts, ph, gr, it), 8) + 6) / DAY_W
        lft = max(label_base(it, cx, w), nr.get(L, -1e9) + 2 / DAY_W)
        nr[L] = lft + w
        mx = max(mx, nr[L])
        sub = sub_label(acts, c, gr, it)
        if sub:
            w = (text_w(sub, 7) + 6) / DAY_W
            lft = max(label_base(it, cx, w), sr.get(L, -1e9) + 2 / DAY_W)
            sr[L] = lft + w
            mx = max(mx, sr[L])
    ncols = max(ncols, int(mx) + 2)

    levels = phase_levels(ph) if (sh["overall"] and ph["names"]) else {}
    nph_rows = max(levels.values()) if levels else 0
    y = 0.0
    row_title, y = y, y + 24
    row_sub, y = y, y + 16
    y += 8
    row_phase = y
    y += 16 * nph_rows
    row_month, y = y, y + 16
    row_day, y = y, y + 14
    row_week, y = y, y + 14
    lanes_top = y
    nlane = sh["nlane"]
    chart_bottom = lanes_top + nlane * LANE_H
    legend_top = chart_bottom + 15
    W = X0 + LABEL_W + ncols * DAY_W + 20
    right = X0 + LABEL_W + ncols * DAY_W

    # 凡例の会社の行（右端で折り返す）
    co_rows, row_items, xx = [], [], X0 + LABEL_W + 4 + 30
    chips = [(nm, colors[nm]) for nm in co_order]
    if sh["overall"] and any(it["kind"] == KIND_SUM and not gr["G"][it["group"]]["co_name"] for it in items):
        chips.append(("複数の会社のまとめ", NEUTRAL))
    for nm, col in chips:
        wchip = 28 + text_w(nm, 8) + 14
        if xx + wchip > right and row_items:
            co_rows.append(row_items); row_items = []; xx = X0 + LABEL_W + 4 + 30
        row_items.append((nm, col, xx))
        xx += wchip
    if row_items:
        co_rows.append(row_items)
    H = legend_top + 24 + 18 + 18 * len(co_rows) + 10

    fig = plt.figure(figsize=(W / 72, H / 72), dpi=150)
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_xlim(0, W)
    ax.set_ylim(H, 0)
    ax.set_aspect("equal")
    ax.axis("off")

    def X(d):
        return X0 + LABEL_W + (d - vs) * DAY_W

    def Y(L):
        return lanes_top + L * LANE_H + ARROW_OFF

    has_hol = any(is_hol(vs + k) for k in range(ncols))
    ax.text(X0, row_title + 13, f"{title}　ネットワーク工程表{caption}", fontsize=14, weight="bold", va="center")
    ax.text(X0, row_sub + 8, "第1版　作成日：2026/11/05　目盛り：日単位（網掛け：灰色＝土日"
            + ("、赤＝休工日" if has_hol else "") + "）", fontsize=9, va="center", color="#333333")
    week = "日月火水木金土"
    prev_month = None
    for k in range(ncols):
        d = dt.date.fromordinal(vs + k)
        x = X(vs + k)
        wd = (d.weekday() + 1) % 7
        if is_hol(vs + k):
            ax.add_patch(Rectangle((x, row_day), DAY_W, chart_bottom - row_day, facecolor=HOL_FILL,
                                   edgecolor="none", zorder=0))
        elif wd in (0, 6):
            ax.add_patch(Rectangle((x, row_day), DAY_W, chart_bottom - row_day, facecolor="#F2F2F2",
                                   edgecolor="none", zorder=0))
        if d.month != prev_month:
            ax.text(x + 2, row_month + 8, f"{d.year}/{d.month}", fontsize=8, va="center")
            ax.plot([x, x], [row_month, chart_bottom], color="#969696", lw=0.8, zorder=1)
            prev_month = d.month
        ax.text(x + DAY_W / 2, row_day + 7, str(d.day), fontsize=7, ha="center", va="center")
        ax.text(x + DAY_W / 2, row_week + 7, week[wd], fontsize=7, ha="center", va="center",
                color=RED if wd == 0 else "#000000")
    ax.plot([X0, right], [lanes_top, lanes_top], color="#7F7F7F", lw=0.8)
    ax.plot([X0 + LABEL_W, X0 + LABEL_W], [row_month, chart_bottom], color="#7F7F7F", lw=0.8)
    ax.text(X0 + 4, row_day + 7, "工事", fontsize=8, va="center")

    if levels:
        ax.text(X0 + 4, row_phase + 8, "期", fontsize=9, va="center")
        for p, nm in enumerate(ph["names"], start=1):
            x1, x2 = X(ph["minS"][p]), X(ph["maxF"][p])
            t = row_phase + (levels[p] - 1) * 16 + 2
            ax.add_patch(Rectangle((x1, t), max(x2 - x1, 6), 12, facecolor="#D9E1F2", edgecolor="none", zorder=1))
            a, b = dt.date.fromordinal(ph["minS"][p]), dt.date.fromordinal(ph["last"][p])
            ax.text(x1 + 3, t + 6, f"{nm}（{a.month}/{a.day}～{b.month}/{b.day}）", fontsize=8,
                    weight="bold", color="#1F3864", va="center", zorder=2)

    for b, name in enumerate(sh["bands"]):
        t = lanes_top + sh["band_first"][b] * LANE_H
        btm = lanes_top + (sh["band_last"][b] + 1) * LANE_H
        ax.text(X0 + 6, (t + btm) / 2, name, fontsize=9, va="center")
        ax.plot([X0, right], [btm, btm], color="#BFBFBF", lw=0.8, zorder=1)
    for rng, txt in ((sh["from_rng"], "他の期から"), (sh["to_rng"], "他の期へ")):
        if rng is not None:
            t = lanes_top + rng[0] * LANE_H
            btm = lanes_top + (rng[1] + 1) * LANE_H
            ax.text(X0 + 6, (t + btm) / 2, txt, fontsize=9, va="center", color=GHOST_C, style="italic")
            ax.plot([X0, right], [btm, btm], color="#BFBFBF", lw=0.8, zorder=1)

    def arrow(x1, y1, x2, y2, color, lw, ls="-", head=True):
        if abs(x2 - x1) < 0.5 and abs(y2 - y1) < 0.5:
            return
        if head:
            ax.annotate("", xy=(x2, y2), xytext=(x1, y1), zorder=3,
                        arrowprops=dict(arrowstyle="-|>", color=color, lw=lw, linestyle=ls,
                                        shrinkA=0, shrinkB=0, mutation_scale=7))
        else:
            ax.plot([x1, x2], [y1, y2], color=color, lw=lw, linestyle=ls, zorder=3)

    def band(x1, y1, x2, y2):
        ax.plot([x1, x2], [y1, y2], color=BAND, alpha=0.45, lw=9, solid_capstyle="butt", zorder=2)

    nr, sr = {}, {}
    for k in order:
        it = items[k]
        L = it["lane"]
        yy = Y(L)
        xs, xf = X(it["s"]), X(it["f"])
        cx = (xs + xf) / 2
        is_item = it["kind"] in (KIND_ACT, KIND_SUM)
        col = item_color(acts, gr, colors, it)
        cr = is_item and icrit(it)
        if is_item:
            # 矢線（まとめは途中の結合点で区切り、区間ごとにクリティカルか判定）
            ev = [it["s"]] + [d for d in range(it["s"] + 1, it["f"]) if (L, d) in sh["nodes"]] + [it["f"]]
            if it["kind"] == KIND_ACT:
                ev = [it["s"], it["f"]]
            for e1, e2 in zip(ev, ev[1:]):
                sc = sh["seg_crit"](it, e1, e2)
                if sc and e2 > e1:
                    band(X(e1), yy, X(e2), yy)
                if X(e2) - X(e1) > NODE_D + 1:
                    arrow(X(e1) + R, yy, X(e2) - R, yy, col, 2.0 if sc else 1.1)
        txt = name_label(acts, ph, gr, it)
        w = text_w(txt, 8) + 6
        left = max(cx - w / 2 if is_item else xs + R + 2, nr.get(L, -1e9) + 2)
        nr[L] = left + w
        ax.text(left + w / 2 if is_item else left, yy - R - 1, txt, fontsize=8, ha="center" if is_item else "left",
                va="bottom", color=col, weight="bold" if cr else "normal", zorder=5)
        sub = sub_label(acts, c, gr, it)
        if sub:
            w = text_w(sub, 7) + 6
            left = max(cx - w / 2 if is_item else xs + R + 2, sr.get(L, -1e9) + 2)
            sr[L] = left + w
            ax.text(left + w / 2 if is_item else left, yy + R + 0.5, sub, fontsize=7,
                    ha="center" if is_item else "left", va="top", color=col, zorder=5)

    for lk in sh["prims"]:
        kind = lk[0]
        if kind == "vert":
            _, d, La, Lb, fn, tn, colr = lk
            y1, y2 = Y(La), Y(Lb)
            sg = 1 if y2 >= y1 else -1
            if fn:
                y1 += sg * R
            if tn:
                y2 -= sg * R
            if colr == "red":
                band(X(d), Y(La), X(d), Y(Lb))
            arrow(X(d), y1, X(d), y2, RED if colr == "red" else GRAY, 1.3 if colr == "red" else 0.9, "--")
        elif kind == "float":
            _, fa, sb, L, fn, tn = lk
            arrow(X(fa) + (R if fn else 0), Y(L), X(sb) - (R if tn else 0), Y(L), GRAY, 0.9, ":", head=False)
        elif kind == "bridge":
            _, fa, sb, L = lk
            yb = Y(L) - R - 8
            arrow(X(fa), Y(L) - R, X(fa), yb, GRAY, 0.9, "--", False)
            arrow(X(fa), yb, X(sb), yb, GRAY, 0.9, "--", False)
            arrow(X(sb), yb, X(sb), Y(L) - R, GRAY, 0.9, "--")
        elif kind == "diag":
            _, fa, La, sb, Lb, colr = lk
            x1, y1, x2, y2 = X(fa), Y(La), X(sb), Y(Lb)
            ln = ((x2 - x1) ** 2 + (y2 - y1) ** 2) ** 0.5
            ux, uy = (x2 - x1) / ln, (y2 - y1) / ln
            arrow(x1 + ux * R, y1 + uy * R, x2 - ux * R, y2 - uy * R, RED if colr == "red" else GRAY, 0.9, "--")

    num = 0
    for key in sorted(sh["nodes"], key=lambda k: (k[1], k[0])):
        L, d = key
        nd = sh["nodes"][key]
        if nd["ghost"]:
            ax.add_patch(Circle((X(d), Y(L)), R, facecolor="white", edgecolor=GHOST_C, lw=0.9,
                                linestyle=(0, (1.5, 1.5)), zorder=6))
        else:
            num += 1
            ax.add_patch(Circle((X(d), Y(L)), R, facecolor="white", edgecolor=RED if nd["crit"] else NEUTRAL,
                                lw=1.3 if nd["crit"] else 0.9, zorder=6))
            ax.text(X(d), Y(L), str(num), fontsize=6 if num < 100 else 5, ha="center", va="center", zorder=7)

    ly = legend_top + 12
    ax.text(X0 + 4, ly, "凡例", fontsize=8, va="center")
    x = X0 + LABEL_W + 4
    arrow(x, ly, x + 36, ly, NEUTRAL, 1.1)
    ax.text(x + 42, ly, "作業（矢線）", fontsize=8, va="center")
    x += 130
    ax.add_patch(Circle((x + R, ly), R, facecolor="white", edgecolor=NEUTRAL, lw=0.9, zorder=6))
    ax.text(x + NODE_D + 6, ly, "結合点（番号付きの○）", fontsize=8, va="center")
    x += 150
    arrow(x, ly, x + 36, ly, GRAY, 0.9, "--")
    ax.text(x + 42, ly, "ダミー（他の工事とのつながり）", fontsize=8, va="center")
    x += 190
    band(x, ly, x + 36, ly)
    arrow(x, ly, x + 36, ly, NEUTRAL, 2.0)
    ax.text(x + 42, ly, "クリティカルパス（赤い帯）", fontsize=8, va="center")
    x += 170
    arrow(x, ly, x + 36, ly, GRAY, 0.9, ":", head=False)
    ax.text(x + 42, ly, "余裕日数", fontsize=8, va="center")
    ly2 = legend_top + 24 + 9
    x = X0 + LABEL_W + 4
    if sh["mode"] > 0:
        ax.add_patch(Circle((x + R, ly2), R, facecolor="white", edgecolor=GHOST_C, lw=0.9,
                            linestyle=(0, (1.5, 1.5)), zorder=6))
        ax.text(x + NODE_D + 6, ly2, "他の期の作業（点線の○）", fontsize=8, va="center")
        x += 170
    note = "表記：矢線の上＝作業ID・作業名" + ("（全体工程のまとめは名前）" if sh["overall"] and gr["G"] else "") + \
           "、矢線の下＝会社・日数（進捗があれば％も）"
    ax.text(x, ly2, note, fontsize=8, va="center", color=NEUTRAL)
    for r_i, row in enumerate(co_rows):
        yrow = legend_top + 24 + 18 + 18 * r_i + 9
        if r_i == 0:
            ax.text(X0 + LABEL_W + 4, yrow, "会社：", fontsize=8, va="center")
        for nm, col, xx in row:
            ax.plot([xx, xx + 24], [yrow, yrow], color=col, lw=2.2)
            ax.text(xx + 28, yrow, nm, fontsize=8, va="center", color=col)
    fig.savefig(path, dpi=150, facecolor="white")
    plt.close(fig)
    return num


if __name__ == "__main__":
    outdir = sys.argv[1]
    HOL.update(d.toordinal() for d, _ in HOLIDAYS)
    # 休工日の計算の確認
    assert add_work(dt.date(2026, 12, 28).toordinal(), 3) == dt.date(2027, 1, 6).toordinal()
    assert start_from_end(dt.date(2027, 1, 5).toordinal(), 3) == dt.date(2026, 12, 28).toordinal()
    assert next_work(dt.date(2026, 12, 30).toordinal()) == dt.date(2027, 1, 4).toordinal()
    assert add_work(dt.date(2027, 1, 4).toordinal(), 0) == dt.date(2027, 1, 4).toordinal()

    acts = make_sample_dates()
    for a in acts:
        a["phase"] = PHASE_OF[a["id"]]
        a["group"] = GROUP_OF.get(a["id"], "")
    c = compute(acts)
    print("date checks:", check_dates(acts, c) or "none")
    ph = collect_phases(acts, c)
    gr = collect_groups(acts, c)
    for g, G in gr["G"].items():
        print(f"  group {g}: {G['name']} {dt.date.fromordinal(G['S'])}〜{dt.date.fromordinal(G['F'] - 1)} "
              f"crit={G['crit']} co={G['co_label']}")
    colors, co_order = company_colors(acts)
    title = "○○設備更新工事（サンプル）"
    sheets = [(0, "全体工程", "（全体）")] + [(p, nm, f"（{nm}）") for p, nm in enumerate(ph["names"], start=1)]
    for mode, nm, cap in sheets:
        sh = build_sheet(acts, c, ph, gr, mode)
        nn = render(acts, c, ph, gr, sh, colors, co_order, f"{outdir}/preview3_{nm}.png", title, cap)
        print(f"{nm}: items={len(sh['items'])} lanes={sh['nlane']} nodes={nn} links={len(sh['links'])} "
              f"prims={[p[0] for p in sh['prims']]}")
    with open(f"{outdir}/sample3.json", "w", encoding="utf-8") as fp:
        json.dump([dict(a, start=a["start"].isoformat(), end=a["end"].isoformat()) for a in acts],
                  fp, ensure_ascii=False, indent=1)
