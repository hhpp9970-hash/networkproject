"""実績（進捗）とイナズマ線。VBA の CalcProgress / DrawProgressLine と同じ手順で計算し、完成イメージを描く。

決めごと
・基準日 T は「その日の作業が終わった時点」。時間軸の上では T の列の右端（Tb＝T＋1）
・完了＝実績終了あり（または進捗100%）、着手＝実績開始あり（または進捗＞0）、それ以外は未着手
・進捗の位置 xp：着手中は予定の開始から稼働日で「日数×進捗」進んだ位置。未着手は予定の開始（まだ予定前なら Tb）、
  完了は予定の終了（もう過ぎていれば Tb）
・進捗の遅れ＝Tb − xp（暦日。正＝遅れ、負＝進み）
・予測：完了は実績のまま。着手中は基準日の翌日から残り（日数×（1−進捗）、切り上げ）。
  未着手は「予定の開始・基準日の翌日・前の作業の予測終了」の遅い方から日数どおり（予定より前倒しはしない）
"""
import datetime as dt
import math
import sys

import network_logic3 as v3
import network_logic4 as v4
import network_logic5 as v5
from network_logic import compute
from network_logic3 import (HOLIDAYS, KIND_ACT, KIND_FROM, KIND_SUM, KIND_TO, add_work, collect_phases, company_colors, is_hol,
                            next_work)
from network_logic4 import collect_groups, collect_zones

STATUS = dt.date(2027, 1, 26)
ACTUALS = {  # 作業ID: (実績開始, 実績終了, 進捗)  土曜は休み（工事カレンダー）
    "A-010": ("1/12", "1/14", None),
    "A-020": ("1/15", "1/19", None),   # 1日遅れて完了（予定 1/15～1/18）
    "B1-010": ("1/20", "1/20", None),  # 1日遅れて完了（予定 1/19）
    "B1-020": ("1/21", None, 0.75),    # 撤去に手間取っている（予定 1/20～1/24）
    "B2-010": ("1/24", "1/24", None),  # 前倒しで完了（予定 1/25）
    "B2-020": ("1/25", "1/26", None),  # 予定（1/26～1/29）より早く完了
}
ORIG_SUB = v3.sub_label
PROG = {}


def md(d):
    d = dt.date.fromordinal(int(d))
    return f"{d.month}/{d.day}"


def rnd(x):
    """四捨五入（0から遠い方へ）。VBA の Round は銀行丸めなので使わない"""
    return int(x + 0.5) if x >= 0 else -int(-x + 0.5)


def work_days(s, f):
    return sum(1 for d in range(s, f) if not is_hol(d))


def pos_after(s, e):
    """s から稼働日で e 日（小数可）進んだ位置"""
    k = int(e + 1e-9)
    fr = e - k
    d = add_work(s, k) if k > 0 else s
    if fr > 1e-9:
        d = next_work(d) + fr
    return d


def apply_actuals(acts, actuals=ACTUALS, year=2027):
    for a in acts:
        a["act_s"] = a["act_e"] = None
        if a["id"] in actuals:
            s, e, p = actuals[a["id"]]
            conv = (lambda t: dt.date(year, *map(int, t.split("/"))) if t else None)
            a["act_s"], a["act_e"] = conv(s), conv(e)
            if p is not None:
                a["prog"] = p


def calc_progress(acts, c, gr, T):
    n = len(acts)
    S, F, preds = c["S"], c["F"], c["preds"]
    Tb = T + 1
    st, p, xp, dl, plan, wd = [0] * n, [0.0] * n, [0.0] * n, [0.0] * n, [0.0] * n, [0] * n
    use = False
    for i, a in enumerate(acts):
        a_s, a_e, pr = a.get("act_s"), a.get("act_e"), a.get("prog") or 0.0
        if a_s or a_e or pr > 0:
            use = True
        wd[i] = work_days(S[i], F[i])
        mile = F[i] <= S[i]
        if a_e or pr >= 1 or (mile and a_s):
            st[i], p[i] = 2, 1.0
        elif (a_s or pr > 0) and not mile:
            st[i], p[i] = 1, min(max(pr, 0.0), 1.0)
        # 予定進捗（基準日の時点で、予定ではどこまで進んでいるはずか）
        if mile:
            plan[i] = 1.0 if S[i] < Tb else 0.0
        else:
            plan[i] = work_days(S[i], min(F[i], max(Tb, S[i]))) / wd[i] if wd[i] else (1.0 if F[i] <= Tb else 0.0)
        # 進捗の位置
        if st[i] == 2:
            xp[i] = F[i] if F[i] > Tb else Tb
        elif st[i] == 1 and p[i] > 0 and wd[i] > 0:
            xp[i] = pos_after(S[i], p[i] * wd[i])
        else:
            xp[i] = S[i] if S[i] < Tb else Tb
        dl[i] = Tb - xp[i]

    # 予測（前から順に）
    indeg = [len(preds[i]) for i in range(n)]
    topo = [i for i in range(n) if indeg[i] == 0]
    h = 0
    while h < len(topo):
        i = topo[h]
        h += 1
        for j in c["succs"][i]:
            indeg[j] -= 1
            if indeg[j] == 0:
                topo.append(j)
    FS, FF = [0] * n, [0] * n
    for i in topo:
        a = acts[i]
        mile = F[i] <= S[i]
        if st[i] == 2:
            if mile:
                d = a["act_e"] or a["act_s"]
                FS[i] = FF[i] = d.toordinal() if d else min(S[i], Tb)
            else:
                FS[i] = a["act_s"].toordinal() if a["act_s"] else S[i]
                FF[i] = a["act_e"].toordinal() + 1 if a["act_e"] else min(F[i], Tb)
        elif st[i] == 1:
            FS[i] = a["act_s"].toordinal() if a["act_s"] else min(S[i], T)
            rem = max(1, math.ceil(wd[i] * (1 - p[i]) - 1e-6))
            FF[i] = add_work(Tb, rem)
        else:
            es = max([S[i], Tb] + [FF[j] for j in preds[i]])
            FS[i] = next_work(es)
            FF[i] = FS[i] if mile else add_work(FS[i], max(wd[i], 1))
    last_plan = [S[i] if F[i] <= S[i] else F[i] - 1 for i in range(n)]
    last_fc = [FF[i] if F[i] <= S[i] else FF[i] - 1 for i in range(n)]

    # まとめ：進捗は稼働日の重み付き平均、遅れは一番遅れている作業（なければ一番進んでいる作業）
    gp, gd = {}, {}
    for g, G in gr["G"].items():
        w = sum(max(wd[i], 0) for i in G["members"])
        gp[g] = (sum(p[i] * wd[i] for i in G["members"]) / w) if w else \
            (1.0 if all(st[i] == 2 for i in G["members"]) else 0.0)
        gd[g] = pick_delay([dl[i] for i in G["members"]])
    return dict(use=use, T=T, Tb=Tb, st=st, p=p, xp=xp, dl=dl, plan=plan, FS=FS, FF=FF, last_plan=last_plan,
                last_fc=last_fc, plan_end=max(last_plan), fc_end=max(last_fc), gp=gp, gd=gd)


def pick_delay(ds):
    """一番遅れているもの。遅れがなければ一番進んでいるもの。どちらもなければ 0"""
    pos = [d for d in ds if d > 1e-6]
    if pos:
        return max(pos)
    neg = [d for d in ds if d < -1e-6]
    return min(neg) if neg else 0.0


def item_delay(P, it):
    if it["kind"] == KIND_SUM:
        return P["gd"][it["group"]]
    if it["kind"] == KIND_ACT:
        return P["dl"][it["act"]]
    return None


def delay_text(d):
    r = rnd(d)
    return f"{r}日遅れ" if r > 0 else (f"{-r}日進み" if r < 0 else "")


def sub_label(acts, c, gr, it):
    """矢線の下：会社・日数・進捗（完了／％）・遅れ（進み）"""
    P = PROG
    if not P or not P["use"]:
        return ORIG_SUB(acts, c, gr, it)
    if it["kind"] == KIND_FROM:  # 他の期の前の作業：終わる日
        i = it["act"]
        if P["st"][i] == 2:
            return f"{md(P['last_fc'][i])}完了"
        t = f"{md(P['last_plan'][i])}終了"
        return t + (f"→予測{md(P['last_fc'][i])}" if P["last_fc"][i] != P["last_plan"][i] else "")
    if it["kind"] == KIND_TO:  # 他の期の後の作業：始まる日
        i = it["act"]
        if P["st"][i] == 2:
            return "完了"
        if P["st"][i] == 1:
            return f"{md(P['FS'][i])}着手済"
        t = f"{md(c['S'][i])}開始"
        return t + (f"→予測{md(P['FS'][i])}" if P["FS"][i] != c["S"][i] else "")
    if it["kind"] == KIND_SUM:
        G = gr["G"][it["group"]]
        t = G["co_label"]
        if G["F"] > G["S"]:
            t += ("・" if t else "") + f"{work_days(G['S'], G['F'])}日"
        q = P["gp"][it["group"]]
    else:
        i = it["act"]
        a = acts[i]
        t = a["co"]
        if c["F"][i] > c["S"][i]:
            t += ("・" if t else "") + f"{work_days(c['S'][i], c['F'][i])}日"
        q = 1.0 if P["st"][i] == 2 else P["p"][i]
    if q >= 0.999999:
        t += ("・" if t else "") + "完了"
    elif q > 0:
        t += ("・" if t else "") + f"{q:.0%}"
    dtx = delay_text(item_delay(P, it))
    if dtx:
        t += ("・" if t else "") + dtx
    return t


def lane_x(sh, L):
    ds = [item_delay(PROG, it) for it in sh["items"] if it["lane"] == L and it["kind"] in (KIND_ACT, KIND_SUM)]
    if not ds:
        return None
    return PROG["Tb"] - pick_delay(ds)


def summary():
    P = PROG
    d = P["fc_end"] - P["plan_end"]
    head = f"基準日 {md(P['T'])}：全体 " + (f"{d}日遅れ" if d > 0 else (f"{-d}日進み" if d < 0 else "予定どおり"))
    return head + f"（完了予定 {md(P['plan_end'])} → 予測 {md(P['fc_end'])}）"


def prog_for_render():
    return dict(T=PROG["T"], Tb=PROG["Tb"], summary=summary(), legend=f"イナズマ線（{md(PROG['T'])}の作業終了時点の進み・遅れ）",
                lane_x=lane_x, accent=lambda it: delay_text(item_delay(PROG, it)) if it["kind"] in (KIND_ACT, KIND_SUM)
                else "")


def install():
    v3.sub_label = v4.sub_label = v5.sub_label = sub_label


if __name__ == "__main__":
    outdir = sys.argv[1]
    week = "--week" in sys.argv[2:]  # 週単位の目盛りで描く（ファイル名の末尾に _週）
    dw, tail = (v5.WEEK_DAY_W, "_週") if week else (v5.DAY_W, "")
    v3.set_calendar(v3.load_sample_calendar())  # 工事カレンダー（土曜は休み、例外の日）
    acts = v5.make_acts()
    apply_actuals(acts)
    c = compute(acts, is_hol=v3.is_hol)
    ph = collect_phases(acts, c)
    zn = collect_zones(acts)
    gr = collect_groups(acts, c, zn)
    colors, co_order = company_colors(acts)
    PROG.update(calc_progress(acts, c, gr, STATUS.toordinal()))
    install()
    P = PROG
    print(summary())
    for i, a in enumerate(acts):
        print(f"  {a['id']:7s} 予定 {md(c['S'][i])}～{md(P['last_plan'][i])}  状態={P['st'][i]} 進捗={P['p'][i]:.0%} "
              f"予定進捗={P['plan'][i]:.0%} 遅れ={P['dl'][i]:+.1f}  予測 {md(P['FS'][i])}～{md(P['last_fc'][i])} "
              f"（終了 {P['last_fc'][i] - P['last_plan'][i]:+d}日）")
    for g, G in gr["G"].items():
        print(f"  まとめ {G['name']}({g}) 進捗={P['gp'][g]:.0%} 遅れ={P['gd'][g]:+.1f}")
    title = "○○設備更新工事（サンプル）"
    sheets = [(0, "全体工程", "（全体）")] + [(p, nm, f"（{nm}）") for p, nm in enumerate(ph["names"], start=1)]
    for mode, nm, cap in sheets:
        holder = {}

        def hook(pre):
            holder["lay"] = v5.layout_labels(acts, c, ph, gr, zn, pre, day_w=dw)
            return holder["lay"]["occ"]

        sh = v4.build_sheet(acts, c, ph, gr, zn, mode, day_w=dw, before_links=hook, week=week)
        lay = holder["lay"]
        v5.render(acts, c, ph, gr, zn, sh, lay, colors, co_order, f"{outdir}/preview6_{nm}{tail}.png", title, cap,
                  prog=prog_for_render(), day_w=dw, week=week)
        bad, pushed, stacked = v5.overlaps(sh, lay, dw)
        print(f"{nm}: overlaps={bad} pushed={pushed} stacked={stacked} "
              f"lane_x={[None if lane_x(sh, L) is None else round(lane_x(sh, L) - P['Tb'], 2) for L in range(sh['nlane'])]}")
