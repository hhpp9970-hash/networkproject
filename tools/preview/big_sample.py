"""2026〜2029年・4期・約150作業の工程を作り、プレビュー用の計算（VBA と同じ手順）でシートごとの
項目・結合点・つなぎ線の数と、計算にかかる時間を測る。Excel が重くなる原因の見積もりに使う。"""
import datetime as dt
import random
import sys
import time

import network_logic3 as v3
import network_logic4 as v4
import network_logic5 as v5
from network_logic import compute
from network_logic3 import collect_phases, company_colors
from network_logic4 import collect_groups, collect_zones

KOUJI = [("撤去工事", "撤去一式", ["停電・安全措置", "既設機器の解体撤去・搬出"], "○○建設"),
         ("基礎・架台工事", "", ["既設基礎はつり補修", "新設架台据付"], "○○建設"),
         ("機器据付工事", "据付一式", ["新設機器搬入", "機器据付・芯出し", "配管接続"], "○○機工"),
         ("電気工事", "電気一式", ["ケーブルラック新設", "幹線配線・端末処理", "制御盤改造"], "○○電設"),
         ("試験・調整", "", ["単体試験", "組合せ試験"], "○○電設")]


def make_rows(seed=1, use_groups=True):
    rnd = random.Random(seed)
    rows = []
    prev_phase_end = ""
    for ph in range(1, 5):
        pname = f"{ph}期"
        a = f"A{ph}-010"
        rows.append(("共通", "準備・仮設工事", pname, "準備一式" if use_groups else "", a, f"{pname}準備・仮設", "",
                     "○○建設", 10, prev_phase_end))
        last_of_zone = []
        prev_rm = a
        for z in range(1, 4):
            zone = f"第{(ph - 1) * 3 + z}工区"
            prev = None
            for k, (kouji, grp, names, co) in enumerate(KOUJI):
                for j, nm in enumerate(names):
                    aid = f"{'BCDEF'[k]}{(ph - 1) * 3 + z}-{(j + 1) * 10:03d}"
                    if prev is None:
                        pred = prev_rm  # 撤去は工区の順に（班が移る）
                    else:
                        pred = prev
                    if k == 3 and j == 0:  # 電気は据付の始めの方から並行で
                        pred = f"D{(ph - 1) * 3 + z}-010"
                    if k == 4 and j == 0:
                        pred = f"D{(ph - 1) * 3 + z}-030,E{(ph - 1) * 3 + z}-030"
                    rows.append((zone, kouji, pname, (grp if use_groups else ""), aid, nm, "", co,
                                 rnd.choice([5, 8, 10, 12, 15, 20]), pred))
                    prev = aid
                    if k == 0 and j == len(names) - 1:
                        prev_rm = aid
            last_of_zone.append(prev)
        insp = f"F{ph}-010"
        rows.append(("共通", "試験・調整", pname, "", insp, f"{pname}総合試運転・検査", "", "○○電設", 10,
                     ",".join(last_of_zone)))
        rows.append(("共通", "試験・調整", pname, "", f"F{ph}-020", f"{pname}引渡し", "", "", 0, insp))
        prev_phase_end = f"F{ph}-020"
    return rows


def make_acts(rows, starts):
    """期の始まり（4/1）を守りつつ前から日付を決める（休工日なし）"""
    fin, out = {}, [None] * len(rows)
    remaining = list(range(len(rows)))
    while remaining:
        for i in list(remaining):
            z, k, p, g, aid, nm, disp, co, d, pr = rows[i]
            ps = [q for q in pr.split(",") if q]
            if all(q in fin for q in ps):
                s = max([fin[q] for q in ps] + [starts[p]])
                fin[aid] = s if d == 0 else s + d
                out[i] = dict(zone=z, kouji=k, phase=p, group=g, id=aid, name=nm, disp=disp, co=co, dur=d, pred=pr,
                              start=dt.date.fromordinal(s), end=dt.date.fromordinal(s if d == 0 else s + d - 1),
                              prog=0.0)
                remaining.remove(i)
    return out


def measure(use_groups):
    rows = make_rows(use_groups=use_groups)
    starts = {f"{p}期": dt.date(2025 + p, 4, 1).toordinal() for p in range(1, 5)}
    acts = make_acts(rows, starts)
    c = compute(acts)
    ph = collect_phases(acts, c)
    zn = collect_zones(acts)
    gr = collect_groups(acts, c, zn)
    span = (max(c["F"]) - min(c["S"]))
    print(f"--- まとめ{'あり' if use_groups else 'なし'}：作業 {len(acts)} 件、期間 {dt.date.fromordinal(min(c['S']))}"
          f"～{dt.date.fromordinal(max(c['F']) - 1)}（{span} 日）")
    tot_shapes = 0
    for mode in range(0, len(ph["names"]) + 1):
        holder = {}

        def hook(pre):
            t0 = time.perf_counter()
            holder["lay"] = v5.layout_labels(acts, c, ph, gr, zn, pre)
            holder["t_lay"] = time.perf_counter() - t0
            return holder["lay"]["occ"]

        t0 = time.perf_counter()
        sh = v4.build_sheet(acts, c, ph, gr, zn, mode, before_links=hook)
        t = time.perf_counter() - t0
        items = sh["items"]
        n_act = sum(1 for it in items if it["kind"] in (v3.KIND_ACT, v3.KIND_SUM))
        n_ghost = len(items) - n_act
        n_nodes = len(sh["nodes"])
        n_links = len(sh.get("links", []))
        days = sh["view_end"] - sh["view_start"] + 1 if "view_end" in sh else None
        # VBA が作る図形のおおよその数：作業ごとに 矢線・上下の文字・クリック用の枠 の4つ、結合点、つなぎ線（点線）、
        # 余裕の点線（半分くらい）、クリティカルの帯（3割くらい）、他の期の○
        est = n_act * 4 + n_nodes + n_links + n_act // 2 + n_act * 3 // 10 + n_ghost * 2
        tot_shapes += est
        print(f"  {('全体' if mode == 0 else ph['names'][mode - 1]):4s} 項目 {n_act:3d}（他の期 {n_ghost:2d}） 段 {sh['nlane']:3d} "
              f"結合点 {n_nodes:3d} つなぎ線 {n_links:3d} 日数 {days}  図形およそ {est:4d}  計算 {t:5.2f}秒（うち文字の配置 {holder['t_lay']:5.2f}秒）")
    print(f"  図形の合計およそ {tot_shapes}")


if __name__ == "__main__":
    measure(True)
    measure(False)
