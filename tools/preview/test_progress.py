"""VBA の CalcProgress を1行ずつ写した関数と、プレビュー用の network_logic6 の計算を、
乱数で作った工程・休工日・実績で突き合わせる（VBA の書き間違いを見つけるため）"""
import datetime as dt
import math
import random

import network_logic3 as v3
import network_logic6 as v6


def vba_int(x):
    return math.floor(x)


def is_hol(d):
    return d in v3.HOL


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


def work_days_in(s, f):
    return sum(1 for d in range(s, f) if not is_hol(d))


def pos_after_work(s, e):
    k = vba_int(e + 0.000000001)
    fr = e - k
    d = add_work(s, k) if k > 0 else s
    return next_work(d) + fr if fr > 0.000000001 else d


def vba_calc(n, S, F, preds, topo, AS, AE, prog, T, gidx, ngroup):
    """VBA の CalcProgress をそのまま写したもの（1始まりの配列は0始まりに読み替え）"""
    Tb = T + 1
    st, p, delay, plan, wd, FS, FF, lastfc = [0] * n, [0.0] * n, [0.0] * n, [0.0] * n, [0] * n, [0] * n, [0] * n, [0] * n

    def last_day(i):
        return F[i] - 1 if F[i] > S[i] else S[i]

    for i in range(n):
        is_mile = F[i] <= S[i]
        has_as, has_ae = AS[i] is not None, AE[i] is not None
        wd[i] = work_days_in(S[i], F[i])
        if has_ae or prog[i] >= 1 or (is_mile and has_as):
            st[i], p[i] = 2, 1.0
        elif (has_as or prog[i] > 0) and not is_mile:
            st[i], p[i] = 1, prog[i]
        else:
            st[i], p[i] = 0, 0.0
        if is_mile:
            plan[i] = 1 if S[i] < Tb else 0
        elif wd[i] > 0:
            plan[i] = work_days_in(S[i], min(F[i], max(Tb, S[i]))) / wd[i]
        else:
            plan[i] = 1 if F[i] <= Tb else 0
        # ProgressPos
        if st[i] == 2:
            xp = F[i] if F[i] > Tb else Tb
        elif st[i] == 1 and p[i] > 0 and wd[i] > 0:
            xp = pos_after_work(S[i], p[i] * wd[i])
        else:
            xp = S[i] if S[i] < Tb else Tb
        delay[i] = Tb - xp
    for i in topo:
        is_mile = F[i] <= S[i]
        if st[i] == 2:
            if is_mile:
                if AE[i] is not None:
                    FF[i] = AE[i]
                elif AS[i] is not None:
                    FF[i] = AS[i]
                else:
                    FF[i] = min(S[i], Tb)
                FS[i] = FF[i]
            else:
                FS[i] = AS[i] if AS[i] is not None else S[i]
                FF[i] = AE[i] + 1 if AE[i] is not None else min(F[i], Tb)
        elif st[i] == 1:
            FS[i] = AS[i] if AS[i] is not None else min(S[i], T)
            rem = -vba_int(-(wd[i] * (1 - p[i]) - 0.000001))
            if rem < 1:
                rem = 1
            FF[i] = add_work(Tb, rem)
        else:
            es = max(S[i], Tb)
            for v in preds[i]:
                if FF[v] > es:
                    es = FF[v]
            FS[i] = next_work(es)
            FF[i] = FS[i] if is_mile else add_work(FS[i], max(wd[i], 1))
        lastfc[i] = FF[i] if is_mile else FF[i] - 1
    plan_end = max(last_day(i) for i in range(n))
    fc_end = max(lastfc)
    gp, gd = {}, {}
    for g in range(1, ngroup + 1):
        wsum = psum = posD = negD = 0.0
        all_done = True
        for i in range(n):
            if gidx[i] == g:
                wsum += wd[i]
                psum += p[i] * wd[i]
                if st[i] != 2:
                    all_done = False
                posD = max(posD, delay[i])
                negD = min(negD, delay[i])
        gp[g] = psum / wsum if wsum > 0 else (1 if all_done else 0)
        gd[g] = posD if posD > 0.000001 else (negD if negD < -0.000001 else 0)
    return dict(st=st, p=p, dl=delay, plan=plan, FS=FS, FF=FF, last_fc=lastfc, plan_end=plan_end, fc_end=fc_end,
                gp=gp, gd=gd)


def vba_round(d):
    return vba_int(d + 0.5) if d >= 0 else -vba_int(-d + 0.5)


def run(seed):
    rnd = random.Random(seed)
    base = dt.date(2027, 3, 1).toordinal()
    v3.HOL.clear()
    v3.HOL.update(base + rnd.randrange(0, 60) for _ in range(rnd.randrange(0, 10)))
    n = rnd.randrange(3, 14)
    preds = [sorted(set(rnd.sample(range(i), rnd.randrange(0, min(i, 3) + 1)))) if i else [] for i in range(n)]
    succs = [[j for j in range(n) if i in preds[j]] for i in range(n)]
    dur = [rnd.choice([0, 1, 2, 3, 4, 5, 6]) for _ in range(n)]
    S, F = [0] * n, [0] * n
    for i in range(n):
        m = max([F[j] for j in preds[i]] + [base])
        S[i] = next_work(m)
        F[i] = S[i] if dur[i] == 0 else add_work(S[i], dur[i])
    T = base + rnd.randrange(-3, 50)
    AS, AE, prog = [None] * n, [None] * n, [0.0] * n
    for i in range(n):
        r = rnd.random()
        if r < 0.3:
            AS[i] = S[i] + rnd.randrange(-2, 3)
            AE[i] = AS[i] + max(dur[i] - 1, 0) + rnd.randrange(-1, 3)
        elif r < 0.5:
            AS[i] = S[i] + rnd.randrange(-2, 3)
            prog[i] = rnd.choice([0, 0.1, 0.25, 0.5, 0.6, 0.75, 0.9])
        elif r < 0.6:
            prog[i] = rnd.choice([0.2, 0.5, 1.0])
    ngroup = rnd.randrange(0, 3)
    gidx = [rnd.randrange(0, ngroup + 1) if ngroup else 0 for _ in range(n)]
    for g in range(1, ngroup + 1):  # まとめには少なくとも1つの作業
        if g not in gidx:
            gidx[rnd.randrange(n)] = g

    # network_logic6（プレビュー）側の入力
    acts = []
    for i in range(n):
        acts.append(dict(id=f"X{i}", prog=prog[i],
                         act_s=dt.date.fromordinal(AS[i]) if AS[i] is not None else None,
                         act_e=dt.date.fromordinal(AE[i]) if AE[i] is not None else None))
    c = dict(S=S, F=F, preds=preds, succs=succs)
    G = {g: dict(members=[i for i in range(n) if gidx[i] == g]) for g in range(1, ngroup + 1)}
    gr = dict(idx=gidx, G=G)
    P = v6.calc_progress(acts, c, gr, T)
    topo = []
    seen = set()
    while len(topo) < n:  # Kahn（VBA の SortByLinks と同じく、先行がすべて済んだものから）
        for i in range(n):
            if i not in seen and all(j in seen for j in preds[i]):
                topo.append(i)
                seen.add(i)
                break
    V = vba_calc(n, S, F, preds, topo, AS, AE, prog, T, gidx, ngroup)
    bad = []
    for key in ("st", "FS", "FF", "last_fc"):
        if list(P[key]) != list(V[key]):
            bad.append(key)
    for key in ("p", "dl", "plan"):
        if any(abs(a - b) > 1e-9 for a, b in zip(P[key], V[key])):
            bad.append(key)
    if P["plan_end"] != V["plan_end"] or P["fc_end"] != V["fc_end"]:
        bad.append("ends")
    for g in G:
        if abs(P["gp"][g] - V["gp"][g]) > 1e-9 or abs(P["gd"][g] - V["gd"][g]) > 1e-9:
            bad.append(f"group{g}")
    # 表記（遅れの丸め）
    for i in range(n):
        if v6.rnd(P["dl"][i]) != vba_round(V["dl"][i]):
            bad.append("round")
    return bad


if __name__ == "__main__":
    fails = 0
    for seed in range(3000):
        b = run(seed)
        if b:
            fails += 1
            if fails <= 5:
                print("seed", seed, b)
    print("cases: 3000  mismatches:", fails)
