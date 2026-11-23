"""NetworkSchedule.bas（工区 対応版）と同じ手順を再現し、完成イメージを描く。
縦軸は 工区 → 工事 の2段。まとめは工区ごと。まとめの中で作業のない期間は点線。
（目的の段は collect_zones・build_sheet だけが扱う。目的の列を描くのは network_logic5 の render）"""
import datetime as dt
import json
import sys

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Circle, Rectangle

from network_logic import compute, parse_preds, text_w
import network_logic3 as v3
from network_logic3 import (BAND, GHOST_C, GRAY, HOL_FILL, HOLIDAYS, KIND_ACT, KIND_FROM, KIND_SUM, KIND_TO,
                            NEUTRAL, RED, check_dates, collect_phases, company_colors, is_hol, item_color,
                            last_day, phase_levels, sub_label)

# 日本語フォント：入っているものを使う（Linux は Noto、Windows は游ゴシック・メイリオ）
from matplotlib import font_manager as _fm
_HAVE = {f.name for f in _fm.fontManager.ttflist}
plt.rcParams["font.family"] = next((f for f in ["Noto Sans CJK JP", "Yu Gothic", "Meiryo", "MS Gothic"]
                                    if f in _HAVE), "sans-serif")

COMMON = "共通"
ZONE_FILL = "#EDF2F9"
BASE = dt.date(2027, 1, 12)
SAMPLE4 = [  # 工区, 工事, 期, まとめ, 作業ID, 作業名, 会社, 日数, 先行作業
    ("共通", "準備・仮設工事", "1期", "準備・仮設一式", "A-010", "現地調査・墨出し", "○○建設", 3, ""),
    ("共通", "準備・仮設工事", "1期", "準備・仮設一式", "A-020", "仮設足場・養生", "○○建設", 3, "A-010"),
    ("第1工区", "撤去工事", "1期", "撤去一式", "B1-010", "停電・安全措置", "○○電設", 1, "A-020"),
    ("第1工区", "撤去工事", "1期", "撤去一式", "B1-020", "既設機器撤去", "○○建設", 4, "B1-010"),
    ("第1工区", "基礎・架台工事", "1期", "", "C1-010", "基礎補修・架台設置", "○○建設", 5, "B1-020"),
    ("第1工区", "機器据付工事", "2期", "", "D1-010", "機器搬入・据付", "○○電機", 4, "C1-010"),
    ("第1工区", "電気工事", "1期", "電気一式", "E1-010", "配線ルート整備・配線", "○○電設", 5, "B1-020"),
    ("第1工区", "電気工事", "2期", "電気一式", "E1-020", "結線", "○○電設", 2, "D1-010,E1-010"),
    ("第1工区", "試験・調整", "2期", "", "F1-010", "単体試験", "○○電機", 2, "E1-020"),
    ("第2工区", "撤去工事", "1期", "撤去一式", "B2-010", "停電・安全措置", "○○電設", 1, "B1-020"),
    ("第2工区", "撤去工事", "1期", "撤去一式", "B2-020", "既設機器撤去", "○○建設", 4, "B2-010"),
    ("第2工区", "基礎・架台工事", "1期", "", "C2-010", "基礎補修・架台設置", "○○建設", 5, "B2-020,C1-010"),
    ("第2工区", "機器据付工事", "2期", "", "D2-010", "機器搬入・据付", "○○電機", 4, "C2-010,D1-010"),
    ("第2工区", "電気工事", "1期", "電気一式", "E2-010", "配線ルート整備・配線", "○○電設", 5, "B2-020,E1-010"),
    ("第2工区", "電気工事", "2期", "電気一式", "E2-020", "結線", "○○電設", 2, "D2-010,E2-010,E1-020"),
    ("第2工区", "試験・調整", "2期", "", "F2-010", "単体試験", "○○電機", 2, "E2-020,F1-010"),
    ("共通", "試験・調整", "2期", "総合試運転・検査", "F-020", "総合試運転", "○○電機", 3, "F1-010,F2-010"),
    ("共通", "準備・仮設工事", "2期", "", "A-030", "足場解体・片付け", "○○建設", 2, "F-020"),
    ("共通", "試験・調整", "2期", "総合試運転・検査", "F-030", "竣工検査", "○○建設", 1, "F-020,A-030"),
    ("共通", "試験・調整", "2期", "", "F-040", "引渡し", "○○建設", 0, "F-030"),
]


def make_acts(rows=SAMPLE4):
    """サンプル用：先行作業と日数から最早の開始・終了日を入れる（取引先が日付を入れて返してきた想定）"""
    fin, out = {}, [None] * len(rows)
    remaining = list(range(len(rows)))
    while remaining:
        moved = False
        for i in list(remaining):
            z, k, p, g, aid, nm, co, d, pr = rows[i]
            ps = parse_preds(pr)
            if all(q in fin for q in ps):
                s = max([fin[q] for q in ps], default=BASE.toordinal())
                fin[aid] = s if d == 0 else s + d
                out[i] = dict(zone=z, kouji=k, phase=p, group=g, id=aid, name=nm, co=co, dur=d, pred=pr,
                              start=dt.date.fromordinal(s), end=dt.date.fromordinal(s if d == 0 else s + d - 1),
                              prog=0.0)
                remaining.remove(i)
                moved = True
        assert moved, "循環あり"
    return out


def collect_zones(acts):
    """工区と目的の一覧（出てきた順、空欄は「共通」）と、目的・工区・工事の組の順位（出てきた順）"""
    use = any(a.get("zone") for a in acts)
    puse = any(a.get("purpose") for a in acts)
    names, idx, pnames, pidx, rank = [], [], [], [], {}
    for a in acts:
        z = a.get("zone") or COMMON
        if z not in names:
            names.append(z)
        zi = names.index(z) + 1
        idx.append(zi)
        p = a.get("purpose") or COMMON
        if p not in pnames:
            pnames.append(p)
        pi = pnames.index(p) + 1
        pidx.append(pi)
        if (pi, zi, a["kouji"]) not in rank:
            rank[(pi, zi, a["kouji"])] = len(rank) + 1
    return dict(use=use, names=names, idx=idx, puse=puse, pnames=pnames, pidx=pidx, rank=rank)


def collect_groups(acts, c, zn):
    """まとめは目的・工区ごとに分ける（同じ名前でも目的か工区が違えば別のまとめ）"""
    keys, gidx = [], []
    for i, a in enumerate(acts):
        g = a.get("group", "")
        if g:
            key = (zn["pidx"][i], zn["idx"][i], g)
            if key not in keys:
                keys.append(key)
            gidx.append(keys.index(key) + 1)
        else:
            gidx.append(0)
    G = {}
    for g in range(1, len(keys) + 1):
        mem = [i for i in range(len(acts)) if gidx[i] == g]
        cos = []
        for i in mem:
            if acts[i]["co"] and acts[i]["co"] not in cos:
                cos.append(acts[i]["co"])
        w = sum(c["F"][i] - c["S"][i] for i in mem)
        p = sum(acts[i]["prog"] * (c["F"][i] - c["S"][i]) for i in mem)
        G[g] = dict(name=keys[g - 1][2], zone=keys[g - 1][1], purpose=keys[g - 1][0], kouji=acts[mem[0]]["kouji"],
                    members=mem,
                    S=min(c["S"][i] for i in mem), F=max(c.get("FD", c["F"])[i] for i in mem),
                    crit=any(c["crit"][i] for i in mem),
                    co_label=(cos[0] if len(cos) == 1 else (f"{cos[0]}ほか{len(cos) - 1}社" if cos else "")),
                    co_name=(cos[0] if len(cos) == 1 else ""),
                    prog=(p / w if w else 0))
    return dict(idx=gidx, G=G)


def phase_label(ph, i):
    p = ph["act_phase"][i]
    return ph["names"][p - 1] if p else "期未入力"


def name_label(acts, ph, gr, zn, it):
    if it["kind"] == KIND_SUM:
        return gr["G"][it["group"]]["name"]
    a = acts[it["act"]]
    t = a["id"] + (" " + a["name"] if a["name"] else "")
    if it["kind"] != KIND_ACT:
        tag = phase_label(ph, it["act"])
        if zn["use"]:
            tag += "・" + zn["names"][zn["idx"][it["act"]] - 1]
        t += "（" + tag + "）"
    return t


def build_sheet(acts, c, ph, gr, zn, mode, day_w=24.0, before_links=None, week=False):
    """mode：0＝全体工程、-1＝詳細工程、1以上＝その期"""
    S, crit, preds, succs = c["S"], c["crit"], c["preds"], c["succs"]
    F = c.get("FD", c["F"])  # 矢線の終点（VBA の actFd。休工日が続くときは次の工事する日）
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
    if week:  # 週単位の目盛りは月曜から
        vs -= dt.date.fromordinal(vs).weekday()

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

    links, pair = [], {}

    def add_link(fk, tk, fa, sb, cr):
        key = (fk, tk)
        if key in pair:
            L0 = links[pair[key]]
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

    order_all = sorted(range(len(items)), key=lambda k: (items[k]["s"], items[k]["f"]))

    def ghost_w(it):
        return max(text_w(name_label(acts, ph, gr, zn, it), 8) + 6, text_w(sub_label(acts, c, gr, it), 7) + 6)

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

    def izone(it):
        return gr["G"][it["group"]]["zone"] if it["kind"] == KIND_SUM else zn["idx"][it["act"]]

    def ipurp(it):
        return gr["G"][it["group"]]["purpose"] if it["kind"] == KIND_SUM else zn["pidx"][it["act"]]

    def ikouji(it):
        return gr["G"][it["group"]]["kouji"] if it["kind"] == KIND_SUM else acts[it["act"]]["kouji"]

    # 縦軸の帯：目的・工区・工事の組。並びは（目的の順, 工区の順, 工程データに出てきた順）
    bands, band_of = [], {}
    for k, it in enumerate(items):
        if it["kind"] in (KIND_ACT, KIND_SUM):
            key = (ipurp(it), izone(it), ikouji(it))
            if key not in bands:
                bands.append(key)
            band_of[k] = bands.index(key)
    band_order = sorted(range(len(bands)), key=lambda b: (bands[b][0], bands[b][1], zn["rank"][bands[b]]))

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

    band_first, band_last = {}, {}
    for b in band_order:
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
        band_first[b] = first
        band_last[b] = len(occ) - 1
    # 縦軸の工区・目的の欄：上下に続く帯をまとめる（同じ工区でも目的が違えば別の欄）
    zone_runs, purp_runs = [], []  # [(目的, 工区), 最初の段, 最後の段]、[目的, 最初の段, 最後の段]
    for b in band_order:
        pz = bands[b][:2]
        if zone_runs and zone_runs[-1][0] == pz:
            zone_runs[-1][2] = band_last[b]
        else:
            zone_runs.append([pz, band_first[b], band_last[b]])
        if purp_runs and purp_runs[-1][0] == pz[0]:
            purp_runs[-1][2] = band_last[b]
        else:
            purp_runs.append([pz[0], band_first[b], band_last[b]])
    first_to = len(occ)
    n_to = stack(KIND_TO, first_to)
    for _ in range(n_to):
        occ.append([]); lastF.append(None); lastK.append(None)

    def seg_busy(it, e1, e2):
        """まとめの区間（e1〜e2）に作業があるか（作業のない区間は点線で描く）"""
        if it["kind"] != KIND_SUM:
            return True
        return any(S[i] < e2 and F[i] > e1 for i in gr["G"][it["group"]]["members"])

    def seg_crit(it, e1, e2):
        if it["kind"] == KIND_ACT:
            return crit[it["act"]]
        return any(crit[i] and S[i] < e2 and F[i] > e1 for i in gr["G"][it["group"]]["members"])

    def label_span(it):
        """ラベルを置く範囲：まとめは作業のある期間のうち最も長いところ。ほかは項目の始点〜終点"""
        if it["kind"] != KIND_SUM:
            return it["s"], it["f"]
        best, run_s = None, None
        for d in range(it["s"], it["f"] + 1):
            if d < it["f"] and seg_busy(it, d, d + 1):
                if run_s is None:
                    run_s = d
            elif run_s is not None:
                if best is None or d - run_s > best[1] - best[0]:
                    best = (run_s, d)
                run_s = None
        return best or (it["s"], it["f"])

    # 結合点（項目の始点・終点、まとめのすき間の両端、つながりの付く点）
    nodes = {}

    def node(L, d, cr, g):
        if (L, d) in nodes:
            nodes[(L, d)]["crit"] = nodes[(L, d)]["crit"] or cr
        else:
            nodes[(L, d)] = dict(crit=cr, ghost=g)

    def sum_end_crit(it, at_start):
        mem = gr["G"][it["group"]]["members"]
        if at_start:
            return any(crit[i] and S[i] == it["s"] for i in mem)
        return any(crit[i] and F[i] == it["f"] for i in mem)

    for it in items:
        g = it["kind"] in (KIND_FROM, KIND_TO)
        if it["kind"] == KIND_SUM:
            node(it["lane"], it["s"], sum_end_crit(it, True), False)
            node(it["lane"], it["f"], sum_end_crit(it, False), False)
            for i in gr["G"][it["group"]]["members"]:
                d = F[i]
                if it["s"] < d < it["f"] and not seg_busy(it, d, d + 1):
                    node(it["lane"], d, crit[i], False)
                d = S[i]
                if it["s"] < d < it["f"] and not seg_busy(it, d - 1, d):
                    node(it["lane"], d, crit[i], False)
            continue
        cr = icrit(it) and not g
        node(it["lane"], it["s"], cr, g)
        node(it["lane"], it["f"], cr, g)
        if g:
            occ[it["lane"]].append((it["s"], it["f"]))
    for L0 in links:
        node(items[L0["f"]]["lane"], L0["a"], L0["crit"], False)
        node(items[L0["t"]]["lane"], L0["b"], L0["crit"], False)

    # ラベルの置き方を先に決める（マイルストーンの横のラベルの範囲には余裕の点線を引かない）
    if before_links is not None:
        pre = dict(items=items, nodes=nodes, nlane=len(occ), label_span=label_span, view_start=vs, view_end=ve)
        for L, a, b in before_links(pre):
            occ[L].append((a, b))

    # 線の引き方（VBAの LinkItems と同じ判定）
    def free(L, a, b):
        return not any(v0 < b and a < v1 for v0, v1 in occ[L])

    def vclear(d, La, Lb):
        lo, hi = min(La, Lb), max(La, Lb)
        return not any((L, d) in nodes for L in range(lo + 1, hi))

    def side_of(d, La, Lb):
        """途中の段の結合点の横を通すとき、右（+1）と左（-1）のどちらを通るか。
        右にあるもの（出ていく矢線、他の期の作業のラベル）と左にあるもの（入ってくる矢線）を数え、右が多ければ左"""
        lo, hi = min(La, Lb), max(La, Lb)
        nl = nr = 0
        for it in items:
            L = it["lane"]
            if lo < L < hi:
                if it["kind"] in (KIND_FROM, KIND_TO):
                    if it["s"] == d:
                        nr += 1
                else:
                    if it["s"] == d:
                        nr += 1
                    if it["f"] == d:
                        nl += 1
                    if it["s"] < d < it["f"] and (L, d) in nodes:
                        nl += 1
                        nr += 1
        for L in (La, Lb):  # 両端が他の期の作業なら、○の右にラベルがある
            if nodes[(L, d)]["ghost"]:
                nr += 2
        return -1 if nr > nl else 1

    sides = []

    def side_slot(d, sd, La, Lb):
        """横を通す線の位置（結合点からの距離の段階）。同じ結合点から出る線・同じ結合点へ入る線は同じ位置を共有"""
        lo, hi = min(La, Lb), max(La, Lb)
        slot = 0
        while any(q["d"] == d and q["side"] == sd and q["slot"] == slot and q["lo"] <= hi and lo <= q["hi"]
                  and q["from"] != La and q["to"] != Lb for q in sides):
            slot += 1
        sides.append(dict(d=d, side=sd, lo=lo, hi=hi, slot=slot, **{"from": La, "to": Lb}))
        return slot

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
                if vclear(fa, La, Lb):
                    prims.append(("vert", fa, La, Lb, True, True, col))
                else:
                    sd = side_of(fa, La, Lb)
                    prims.append(("side", fa, La, Lb, sd, side_slot(fa, sd, La, Lb), col))
            elif free(La, fa, sb) and (La, sb) not in nodes and vclear(sb, La, Lb):
                prims.append(("float", fa, sb, La, True, False)); occ[La].append((fa, sb))
                prims.append(("vert", sb, La, Lb, False, True, "gray"))
            elif free(Lb, fa, sb) and (Lb, fa) not in nodes and vclear(fa, La, Lb):
                prims.append(("vert", fa, La, Lb, True, False, "gray"))
                prims.append(("float", fa, sb, Lb, False, True)); occ[Lb].append((fa, sb))
            else:
                prims.append(("diag", fa, La, sb, Lb, "gray"))

    return dict(seg_crit=seg_crit, seg_busy=seg_busy, label_span=label_span, mode=mode, overall=overall, use_sum=use_sum, items=items,
                bands=bands, band_order=band_order, band_first=band_first, band_last=band_last, zone_runs=zone_runs,
                purp_runs=purp_runs,
                nlane=len(occ), from_rng=(0, n_from - 1) if n_from else None,
                to_rng=(first_to, first_to + n_to - 1) if n_to else None,
                nodes=nodes, prims=prims, links=links, view_start=vs, view_end=ve, icrit=icrit)


def render(acts, c, ph, gr, zn, sh, colors, co_order, path, title, caption):
    DAY_W, LANE_H, ARROW_OFF, NODE_D = 24.0, 44.0, 24.0, 12.0
    R = NODE_D / 2
    ZONE_W = 70.0 if zn["use"] else 0.0
    LABEL_W, X0 = ZONE_W + 110.0, 10.0
    items = sh["items"]
    vs = sh["view_start"]
    icrit = sh["icrit"]
    order = sorted(range(len(items)), key=lambda k: (items[k]["s"], items[k]["f"]))

    def nlabel(it):
        return name_label(acts, ph, gr, zn, it)

    def label_base(it, cx_days, w_days):
        return cx_days - w_days / 2 if it["kind"] in (KIND_ACT, KIND_SUM) else cx_days + (R + 2) / DAY_W

    ncols = max(sh["view_end"] - vs + 1, 40)
    nr, sr, mx = {}, {}, 0.0
    for k in order:
        it = items[k]
        L = it["lane"]
        e1, e2 = sh["label_span"](it)
        cx = ((e1 - vs) + (e2 - vs)) / 2
        w = (text_w(nlabel(it), 8) + 6) / DAY_W
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
    if zn["use"]:
        ax.text(X0 + 4, row_day + 7, "工区", fontsize=8, va="center")
        ax.text(X0 + ZONE_W + 4, row_day + 7, "工事", fontsize=8, va="center")
    else:
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

    # 縦軸：工事（工区があればB列）と、その下の細線
    for b in sh["band_order"]:
        t = lanes_top + sh["band_first"][b] * LANE_H
        btm = lanes_top + (sh["band_last"][b] + 1) * LANE_H
        ax.text(X0 + ZONE_W + 6, (t + btm) / 2, sh["bands"][b][2], fontsize=9, va="center")
        ax.plot([X0 + ZONE_W, right], [btm, btm], color="#BFBFBF", lw=0.8, zorder=1)
    # 縦軸：工区（A列を結合して塗る。工区の下は太線）
    if zn["use"]:
        for (_, z), l1, l2 in sh["zone_runs"]:
            t = lanes_top + l1 * LANE_H
            btm = lanes_top + (l2 + 1) * LANE_H
            ax.add_patch(Rectangle((X0, t), ZONE_W, btm - t, facecolor=ZONE_FILL, edgecolor="none", zorder=0.5))
            ax.text(X0 + ZONE_W / 2, (t + btm) / 2, zn["names"][z - 1], fontsize=9, weight="bold",
                    ha="center", va="center", zorder=2)
            ax.plot([X0 + ZONE_W, X0 + ZONE_W], [t, btm], color="#BFBFBF", lw=0.8, zorder=1)
            ax.plot([X0, right], [btm, btm], color="#595959", lw=1.5, zorder=1.5)
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
        xs = X(it["s"])
        e1, e2 = sh["label_span"](it)
        cx = (X(e1) + X(e2)) / 2
        is_item = it["kind"] in (KIND_ACT, KIND_SUM)
        col = item_color(acts, gr, colors, it)
        cr = is_item and icrit(it)
        if is_item:
            # 矢線（まとめは途中の結合点で区切る。作業のある区間は実線、作業のない区間は点線）
            if it["kind"] == KIND_ACT:
                ev = [it["s"], it["f"]]
            else:
                ev = [it["s"]] + [d for d in range(it["s"] + 1, it["f"]) if (L, d) in sh["nodes"]] + [it["f"]]
            for e1, e2 in zip(ev, ev[1:]):
                if sh["seg_busy"](it, e1, e2):
                    sc = sh["seg_crit"](it, e1, e2)
                    if sc and e2 > e1:
                        band(X(e1), yy, X(e2), yy)
                    if X(e2) - X(e1) > NODE_D + 1:
                        arrow(X(e1) + R, yy, X(e2) - R, yy, col, 2.0 if sc else 1.1)
                elif X(e2) - X(e1) > NODE_D + 1:
                    arrow(X(e1) + R, yy, X(e2) - R, yy, GRAY, 0.9, ":", head=False)
        txt = nlabel(it)
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
        elif kind == "side":
            # 同じ日のダミーで、途中の段の結合点を通るもの：結合点の横（右か左）を通す
            _, d, La, Lb, sd, slot, colr = lk
            x, ya, yb = X(d), Y(La), Y(Lb)
            xs2 = x + sd * (R + 5 + 4 * min(slot, 2))
            sg = 1 if yb > ya else -1
            y2, y3 = ya + sg * (R + 8), yb - sg * (R + 8)
            c_, w_ = (RED, 1.3) if colr == "red" else (GRAY, 0.9)
            if colr == "red":
                band(x, ya, xs2, y2)
                band(xs2, y2, xs2, y3)
                band(xs2, y3, x, yb)
            ln = ((xs2 - x) ** 2 + (y2 - ya) ** 2) ** 0.5
            arrow(x + (xs2 - x) / ln * R, ya + (y2 - ya) / ln * R, xs2, y2, c_, w_, "--", False)
            arrow(xs2, y2, xs2, y3, c_, w_, "--", False)
            ln = ((x - xs2) ** 2 + (yb - y3) ** 2) ** 0.5
            arrow(xs2, y3, x - (x - xs2) / ln * R, yb - (yb - y3) / ln * R, c_, w_, "--")
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
    ax.text(x + 42, ly, "余裕日数・まとめの待ち" if sh["use_sum"] else "余裕日数", fontsize=8, va="center")
    ly2 = legend_top + 24 + 9
    x = X0 + LABEL_W + 4
    if sh["mode"] > 0:
        ax.add_patch(Circle((x + R, ly2), R, facecolor="white", edgecolor=GHOST_C, lw=0.9,
                            linestyle=(0, (1.5, 1.5)), zorder=6))
        ax.text(x + NODE_D + 6, ly2, "他の期の作業（点線の○）", fontsize=8, va="center")
        x += 170
    note = "表記：矢線の上＝作業ID・作業名" + ("（まとめは名前）" if sh["use_sum"] else "") + \
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


def fmt(d):
    x = dt.date.fromordinal(d)
    return f"{x.month}/{x.day}"


if __name__ == "__main__":
    outdir = sys.argv[1]
    v3.HOL.update(d.toordinal() for d, _ in HOLIDAYS)

    acts = make_acts()
    c = compute(acts)
    expect = {
        "A-010": ("1/12", "1/14"), "A-020": ("1/15", "1/17"), "B1-010": ("1/18", "1/18"), "B1-020": ("1/19", "1/22"),
        "C1-010": ("1/23", "1/27"), "E1-010": ("1/23", "1/27"), "B2-010": ("1/23", "1/23"), "B2-020": ("1/24", "1/27"),
        "D1-010": ("1/28", "1/31"), "C2-010": ("1/28", "2/1"), "E2-010": ("1/28", "2/1"), "E1-020": ("2/1", "2/2"),
        "D2-010": ("2/2", "2/5"), "F1-010": ("2/3", "2/4"), "E2-020": ("2/6", "2/7"), "F2-010": ("2/8", "2/9"),
        "F-020": ("2/10", "2/12"), "A-030": ("2/13", "2/14"), "F-030": ("2/15", "2/15"), "F-040": ("2/16", "2/16"),
    }
    for i, a in enumerate(acts):
        got = (fmt(c["S"][i]), fmt(last_day(c, i)))
        assert got == expect[a["id"]], (a["id"], got, expect[a["id"]])
    crit_ids = [a["id"] for i, a in enumerate(acts) if c["crit"][i]]
    print("critical:", crit_ids)
    print("date checks:", check_dates(acts, c) or "none")
    assert not any(c["msgs"])

    ph = collect_phases(acts, c)
    zn = collect_zones(acts)
    gr = collect_groups(acts, c, zn)
    print("zones:", zn["names"], "rank:", {f"{zn['names'][k[1] - 1]}|{k[2]}": v for k, v in zn["rank"].items()})
    for g, G in gr["G"].items():
        print(f"  group {g}: {zn['names'][G['zone'] - 1]} / {G['name']} {fmt(G['S'])}〜{fmt(G['F'] - 1)} "
              f"crit={G['crit']} co={G['co_label']} members={[acts[i]['id'] for i in G['members']]}")
    colors, co_order = company_colors(acts)
    title = "○○設備更新工事（サンプル）"
    sheets = [(0, "全体工程", "（全体）")] + [(p, nm, f"（{nm}）") for p, nm in enumerate(ph["names"], start=1)]
    for mode, nm, cap in sheets:
        sh = build_sheet(acts, c, ph, gr, zn, mode)
        nn = render(acts, c, ph, gr, zn, sh, colors, co_order, f"{outdir}/preview4_{nm}.png", title, cap)
        print(f"{nm}: items={len(sh['items'])} lanes={sh['nlane']} nodes={nn} links={len(sh['links'])} "
              f"bands={[(zn['names'][sh['bands'][b][1] - 1], sh['bands'][b][2], sh['band_first'][b], sh['band_last'][b]) for b in sh['band_order']]}")
        print("   prims:", [p[0] for p in sh["prims"]])
    with open(f"{outdir}/sample4.json", "w", encoding="utf-8") as fp:
        json.dump([dict(a, start=a["start"].isoformat(), end=a["end"].isoformat()) for a in acts],
                  fp, ensure_ascii=False, indent=1)
