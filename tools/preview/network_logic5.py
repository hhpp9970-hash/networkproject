"""NetworkSchedule.bas（表示名・ラベルの折り返し・日付の罫線）と同じ手順を再現し、完成イメージを描く。
ラベルは自分の矢線の中央に置き、隣を押しのけない。使える幅に入らなければ折り返し、その段を高くする。
縦軸は 目的 → 工区 → 工事 の3段（目的・工区が無ければ、その列は描かない）。"""
import datetime as dt
import json
import sys

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Circle, Rectangle

from network_logic import compute, parse_preds
import network_logic3 as v3
import network_logic4 as v4
from network_logic3 import (BAND, GHOST_C, GRAY, HOL_FILL, HOLIDAYS, KIND_ACT, KIND_FROM, KIND_SUM, KIND_TO,
                            NEUTRAL, RED, check_dates, collect_phases, company_colors, is_hol, item_color,
                            last_day, phase_levels, sub_label)
from network_logic4 import ZONE_FILL, collect_groups, collect_zones, phase_label

# 日本語フォント：入っているものを使う（Linux は Noto、Windows は游ゴシック・メイリオ）
from matplotlib import font_manager as _fm
_HAVE = {f.name for f in _fm.fontManager.ttflist}
plt.rcParams["font.family"] = next((f for f in ["Noto Sans CJK JP", "Yu Gothic", "Meiryo", "MS Gothic"]
                                    if f in _HAVE), "sans-serif")

NAME_FONT, SUB_FONT, GAP, LINE_PITCH = 8.0, 7.0, 2.0, 1.75
INAZUMA = "#E6007E"  # イナズマ線（クリティカルパスの赤と見分ける）
PURP_FILL, PURP_TXT, PURP_LINE = "#D9E1F2", "#1F3864", "#404040"  # 目的の欄の塗り・文字、目的の境目の線
SAT_FILL, SUN_FILL, HOL_HATCH, SAT_TXT = "#DDEBF7", "#FCE4EC", "#A6A6A6", "#0070C0"  # 土・日の網掛け、休工日の斜線、土の文字
EXTRA_FILL = "#C6EFCE"  # 休みの曜日だけど工事する日
DAY_W, LANE_H, ARROW_OFF, NODE_D = 24.0, 44.0, 24.0, 12.0
WEEK_DAY_W = 6.75  # 週単位の目盛りの1日の幅（VBA は列幅 0.7＝9px）
R = NODE_D / 2
BASE = dt.date(2027, 1, 12)
SAMPLE5 = [  # 目的, 工区, 工事, 期, まとめ, 作業ID, 作業名, 表示名, 会社, 日数, 先行作業
    ("工事の準備", "共通", "準備・仮設工事", "1期", "準備・仮設一式", "A-010", "現地調査・墨出し", "", "○○建設", 3, ""),
    ("工事の準備", "共通", "準備・仮設工事", "1期", "準備・仮設一式", "A-020", "仮設足場組立・養生シート張り", "", "○○建設", 3, "A-010"),
    ("老朽設備の更新", "第1工区", "撤去工事", "1期", "撤去一式", "B1-010", "停電・安全措置・作業区画の養生", "", "○○電設", 1, "A-020"),
    ("老朽設備の更新", "第1工区", "撤去工事", "1期", "撤去一式", "B1-020", "既設機器の解体撤去・搬出", "", "○○建設", 4, "B1-010"),
    ("老朽設備の更新", "第1工区", "基礎・架台工事", "1期", "", "C1-010", "既設基礎はつり補修・新設架台据付", "", "○○建設", 5, "B1-020"),
    ("老朽設備の更新", "第1工区", "機器据付工事", "2期", "", "D1-010", "新設盤の搬入・据付・レベル調整", "", "○○電機", 4, "C1-010"),
    ("老朽設備の更新", "第1工区", "電気工事", "1期", "電気一式", "E1-010", "ケーブルラック新設・幹線ケーブル配線", "", "○○電設", 5, "B1-020"),
    ("老朽設備の更新", "第1工区", "電気工事", "2期", "電気一式", "E1-020", "端末処理・結線", "", "○○電設", 2, "D1-010,E1-010"),
    ("老朽設備の更新", "第1工区", "試験・調整", "2期", "", "F1-010", "単体試験", "", "○○電機", 2, "E1-020"),
    ("老朽設備の更新", "第2工区", "撤去工事", "1期", "撤去一式", "B2-010", "停電・安全措置・作業区画の養生", "停電・安全措置", "○○電設", 1, "B1-020"),
    ("老朽設備の更新", "第2工区", "撤去工事", "1期", "撤去一式", "B2-020", "既設機器の解体撤去・搬出", "機器撤去", "○○建設", 4, "B2-010"),
    ("老朽設備の更新", "第2工区", "基礎・架台工事", "1期", "", "C2-010", "既設基礎はつり補修・新設架台据付", "基礎補修\n架台据付", "○○建設", 5,
     "B2-020,C1-010"),
    ("老朽設備の更新", "第2工区", "機器据付工事", "2期", "", "D2-010", "新設盤の搬入・据付・レベル調整", "盤搬入・据付", "○○電機", 4, "C2-010,D1-010"),
    ("老朽設備の更新", "第2工区", "電気工事", "1期", "電気一式", "E2-010", "ケーブルラック新設・幹線ケーブル配線", "ラック・幹線配線", "○○電設", 5,
     "B2-020,E1-010"),
    ("老朽設備の更新", "第2工区", "電気工事", "2期", "電気一式", "E2-020", "端末処理・結線", "結線", "○○電設", 2, "D2-010,E2-010,E1-020"),
    ("老朽設備の更新", "第2工区", "試験・調整", "2期", "", "F2-010", "単体試験", "", "○○電機", 2, "E2-020,F1-010"),
    ("竣工・引渡し", "共通", "試験・調整", "2期", "総合試運転・検査", "F-020", "総合試運転・各種性能確認", "", "○○電機", 3, "F1-010,F2-010"),
    ("竣工・引渡し", "共通", "準備・仮設工事", "2期", "", "A-030", "足場解体・片付け・清掃", "", "○○建設", 2, "F-020"),
    ("竣工・引渡し", "共通", "試験・調整", "2期", "総合試運転・検査", "F-030", "竣工検査", "", "○○建設", 1, "F-020,A-030"),
    ("竣工・引渡し", "共通", "試験・調整", "2期", "", "F-040", "引渡し", "", "○○建設", 0, "F-030"),
]


def make_acts(rows=SAMPLE5):
    """先行作業と日数（工事する日の数）から開始・終了日を入れる。休工日は飛ばす（VBA の CalcDates と同じ置き方）"""
    fin, out = {}, [None] * len(rows)
    remaining = list(range(len(rows)))
    while remaining:
        moved = False
        for i in list(remaining):
            pu, z, k, p, g, aid, nm, disp, co, d, pr = rows[i]
            ps = parse_preds(pr)
            if all(q in fin for q in ps):
                s = v3.next_work(max([fin[q] for q in ps], default=BASE.toordinal()))
                f = s if d == 0 else v3.add_work(s, d)
                fin[aid] = f
                out[i] = dict(purpose=pu, zone=z, kouji=k, phase=p, group=g, id=aid, name=nm, disp=disp, co=co, dur=d,
                              pred=pr,
                              start=dt.date.fromordinal(s), end=dt.date.fromordinal(s if d == 0 else f - 1),
                              prog=0.0)
                remaining.remove(i)
                moved = True
        assert moved
    return out


# ---------------------------------------------------------------- 文字の幅と折り返し（VBAと同じ規則）
NO_START = set("・、。，．）」』】ー～…ぁぃぅぇぉっゃゅょァィゥェォッャュョ々.,)]")
NO_END = set("（「『【([")
GOOD_AFTER = set("・、，。 　/／）」』】,")
GOOD_BEFORE = set("（「『【(")


def cw(ch, fs):
    o = ord(ch)
    return fs * 0.6 if (o < 0x2000 or 0xFF61 <= o <= 0xFF9F) else fs


def tw(s, fs):
    return sum(cw(ch, fs) for ch in s)


def trim_sp(s):
    return s.strip(" 　")


def ok_break(p, j):
    """j 文字目のあとで切ってよいか（行頭禁則・行末禁則）"""
    return 1 <= j < len(p) and p[j] not in NO_START and p[j - 1] not in NO_END


def good_break(p, j):
    return ok_break(p, j) and (p[j - 1] in GOOD_AFTER or p[j] in GOOD_BEFORE)


def break_at(p, cnt):
    for j in range(cnt, (cnt + 1) // 2 - 1, -1):
        if good_break(p, j):
            return j
    j = cnt
    while j > 1 and not ok_break(p, j):
        j -= 1
    return j if ok_break(p, j) else cnt


def wrap_para(p, maxw, fs):
    """1段落を幅 maxw に折り返す。行数が最少 → 切れ目（・、空白、かっこ）で切れる → いちばん短い行が長い、の順に選ぶ"""
    m = len(p)
    bl, bu, bs, nx = [0] * (m + 1), [0] * (m + 1), [0.0] * (m + 1), [m] * (m + 1)
    bs[m] = 1e30
    for i in range(m - 1, -1, -1):
        best, w, jf = None, 0.0, i + 1
        for j in range(i + 1, m + 1):
            w += cw(p[j - 1], fs)
            if w > maxw and j > i + 1:
                break
            jf = j
            if j == m or ok_break(p, j):
                lw = tw(trim_sp(p[i:j]), fs)
                cand = (bl[j] + 1, bu[j] + (0 if j == m or good_break(p, j) else 1), -min(lw, bs[j]))
                if best is None or cand < best[0]:
                    best = (cand, j)
        if best is None:  # 禁則で切れる位置がないときは、幅に入るところで切る
            bl[i], bu[i], bs[i], nx[i] = bl[jf] + 1, bu[jf] + 1, 0.0, jf
        else:
            bl[i], bu[i], bs[i], nx[i] = best[0][0], best[0][1], -best[0][2], best[1]
    lines, i = [], 0
    while i < m:
        ln = trim_sp(p[i:nx[i]])
        if ln:
            lines.append(ln)
        i = nx[i]
    return lines


def one_line(id_, body):
    body = body.replace("\r\n", " ").replace("\n", " ").replace("\r", " ")
    if not id_:
        return body
    if not body:
        return id_
    return id_ + " " + body


def wrap_label(id_, body, maxw, fs):
    if not body:
        return [id_]
    paras = [trim_sp(p) for p in body.replace("\r\n", "\n").replace("\r", "\n").split("\n")]
    first = one_line(id_, paras[0])
    if tw(first, fs) <= maxw and all(tw(p, fs) <= maxw for p in paras[1:]):
        return ([first] if first else []) + [p for p in paras[1:] if p]
    lines = [id_] if id_ else []
    for p in paras:
        if p:
            lines += wrap_para(p, maxw, fs)
    return lines


# ---------------------------------------------------------------- ラベルの中身
def label_parts(acts, gr, it):
    if it["kind"] == KIND_SUM:
        return "", gr["G"][it["group"]]["name"]
    a = acts[it["act"]]
    return a["id"], (a.get("disp") or a["name"])


def name_label(acts, ph, gr, zn, it):
    """1行の名前（他の期の作業は期・工区も）"""
    t = one_line(*label_parts(acts, gr, it))
    if it["kind"] in (KIND_FROM, KIND_TO):
        tag = phase_label(ph, it["act"])
        if zn["use"]:
            tag += "・" + zn["names"][zn["idx"][it["act"]] - 1]
        t += "（" + tag + "）"
    return t


v4.name_label = name_label  # build_sheet の他の期の作業の段分けも表示名で


VGAP = 1.0


def box_h(n, fs):
    """n 行のラベルの枠の高さ（1行目は余白込みで 1.9 倍、2行目から行の間隔ずつ）"""
    return fs * (1.9 + LINE_PITCH * (n - 1))


def sub_wrap(text, maxw):
    """会社・日数は「・」の位置だけで折り返す（会社名の途中では切らない）"""
    parts = text.split("・")
    lines, cur = [], parts[0]
    for part in parts[1:]:
        if tw(cur + "・" + part, SUB_FONT) <= maxw:
            cur += "・" + part
        else:
            lines.append(cur)
            cur = part
    lines.append(cur)
    return lines


def stack_offset(placed, x1, x2, h):
    """重なる相手がいれば、その上（会社・日数は下）へ積む。いちばん矢線に近い空いた位置"""
    over = [(o, ph) for (px1, px2, o, ph) in placed if x1 < px2 + GAP and px1 < x2 + GAP]
    best = None
    for cnd in [0.0] + [o + ph + VGAP for (o, ph) in over]:
        if best is not None and cnd >= best:
            continue
        if all(cnd + h + VGAP <= o or o + ph + VGAP <= cnd for (o, ph) in over):
            best = cnd
    return best


def layout_labels(acts, c, ph, gr, zn, sh, day_w=DAY_W):
    """VBA の LayoutLabels と同じ。位置は日付目盛りの左端からの pt
    ・作業とまとめのラベルは自分の矢線の中央に置く（横へはずらさない）
    ・使える幅は、前のラベルの右端から、次のラベル（前後と半分ずつ分けたときの幅）の左端まで。入らなければ折り返す
    ・それでも重なるときは、上（会社・日数は下）の段に積む。段の高さはラベルに合わせて増やす"""
    items, vs = sh["items"], sh["view_start"]
    order = sorted(range(len(items)), key=lambda k: (items[k]["s"], items[k]["f"]))
    ctr, next_k, prev_k, last_k = {}, {}, {}, {}
    for k in order:
        e1, e2 = sh["label_span"](items[k])
        ctr[k] = ((e1 - vs) + (e2 - vs)) / 2 * day_w
        L = items[k]["lane"]
        if L in last_k:
            next_k[last_k[L]] = k
            prev_k[k] = last_k[L]
        last_k[L] = k
    nlane = sh["nlane"]

    def texts(k, is_name):
        if is_name:
            return label_parts(acts, gr, items[k])
        return "", sub_label(acts, c, gr, items[k])

    def fit(k, is_name, W):
        """幅 W に入れば1行、入らなければ折り返す。戻り値（行, 幅）"""
        id_, body = texts(k, is_name)
        if is_name:
            one = one_line(id_, body)
            need = tw(one, NAME_FONT) + 6
            if need <= W and "\n" not in body and "\r" not in body:
                return [one], need
            W = max(W, max(4 * NAME_FONT, tw(id_, NAME_FONT)) + 6)
            lines = wrap_label(id_, body, W - 6, NAME_FONT)
            return lines, max(tw(t, NAME_FONT) for t in lines) + 6
        need = tw(body, SUB_FONT) + 6
        if need <= W:
            return [body], need
        lines = sub_wrap(body, max(W, 4 * SUB_FONT + 6) - 6)
        return lines, max(tw(t, SUB_FONT) for t in lines) + 6

    def empty(k, is_name):
        id_, body = texts(k, is_name)
        return not id_ and not body

    # マイルストーン（日数0）：前後の矢線が○に付いていない側に、矢線の高さで1行に書く（会社は括弧書き）
    out = {k: {} for k in order}
    occ_add, mx = [], 0.0
    mid = {}
    for k in order:
        it = items[k]
        if it["kind"] != KIND_ACT or it["s"] != it["f"]:
            continue
        id_, body = texts(k, True)
        co = sub_label(acts, c, gr, it)
        txt = one_line(id_, body) + (f"（{co}）" if co else "")
        w = tw(txt, NAME_FONT) + 6
        span = int(-(-(R + 2 + w) // day_w))  # 切り上げ
        p, q = prev_k.get(k), next_k.get(k)
        if q is None or items[q]["s"] > it["s"]:
            room = (items[q]["s"] - it["s"]) * day_w - 2 * (R + 2) if q is not None else 1e30
            if w <= room:
                mid[k] = ([txt], ctr[k] + R + 2, w, 1)
                occ_add.append((it["lane"], it["s"], it["s"] + span))
                mx = max(mx, ctr[k] + R + 2 + w)
                continue
        if p is None or items[p]["f"] < it["s"]:
            room = (it["s"] - items[p]["f"]) * day_w - 2 * (R + 2) if p is not None else ctr[k] - R - 2
            if w <= room:
                mid[k] = ([txt], ctr[k] - R - 2 - w, w, 2)
                occ_add.append((it["lane"], it["s"] - span, it["s"]))

    # 上下のラベルどうしの前後（横に書いたマイルストーンは除く）
    nxt, prv, last2 = {}, {}, {}
    for k in order:
        if k in mid:
            continue
        L = items[k]["lane"]
        if L in last2:
            nxt[last2[L]] = k
            prv[k] = last2[L]
        last2[L] = k

    # 1回目：前後の項目と間を半分ずつ分けたときの幅（次のラベルのために残す幅の目安）
    w1 = {}
    for k in order:
        if items[k]["kind"] in (KIND_ACT, KIND_SUM):
            for is_name in (True, False):
                if empty(k, is_name) or k in mid:
                    w1[(k, is_name)] = 0.0
                    continue
                p, q = prv.get(k), nxt.get(k)
                room_l = (ctr[k] - ctr[p]) / 2 - GAP / 2 if p is not None else ctr[k]
                room_r = (ctr[q] - ctr[k]) / 2 - GAP / 2 if q is not None else 1e30
                w1[(k, is_name)] = fit(k, is_name, 2 * min(room_l, room_r))[1]

    # 2回目：左から順に置く
    need_top, need_btm = [0.0] * nlane, [0.0] * nlane
    for is_name, fs in ((True, NAME_FONT), (False, SUB_FONT)):
        placed = [[] for _ in range(nlane)]
        right0 = [-1e5] * nlane  # 矢線のすぐ上（下）に置いたラベルの右端
        for k in order:
            it = items[k]
            L = it["lane"]
            if it["kind"] in (KIND_FROM, KIND_TO):
                txt = name_label(acts, ph, gr, zn, it) if is_name else sub_label(acts, c, gr, it)
                lines, w = [txt], tw(txt, fs) + 6
                x1 = ctr[k] + R + 2
            else:
                if empty(k, is_name) or k in mid:
                    out[k][is_name] = ([], 0.0, 0.0, 0.0)
                    continue
                room_l = min(ctr[k] - right0[L] - GAP, ctr[k])
                q = nxt.get(k)
                room_r = (ctr[q] - w1[(q, is_name)] / 2) - ctr[k] - GAP if q is not None else 1e30
                lines, w = fit(k, is_name, 2 * min(room_l, room_r))
                x1 = ctr[k] - w / 2
            h = box_h(len(lines), fs)
            off = stack_offset(placed[L], x1, x1 + w, h)
            if off > 0 and it["kind"] in (KIND_ACT, KIND_SUM):
                # 積むことになるなら、折り返さない形と比べて低く収まる方にする
                lines2, w2 = fit(k, is_name, 1e30)
                x2 = ctr[k] - w2 / 2
                h2 = box_h(len(lines2), fs)
                off2 = stack_offset(placed[L], x2, x2 + w2, h2)
                if off2 + h2 < off + h:
                    lines, w, x1, h, off = lines2, w2, x2, h2, off2
            placed[L].append((x1, x1 + w, off, h))
            if off == 0:
                right0[L] = max(right0[L], x1 + w)
            out[k][is_name] = (lines, x1, w, off)
            if is_name:
                need_top[L] = max(need_top[L], off + h)
            else:
                need_btm[L] = max(need_btm[L], off + h)
            mx = max(mx, x1 + w)
    top_extra = [max(0.0, t - box_h(1, NAME_FONT)) for t in need_top]
    btm_extra = [max(0.0, b - box_h(1, SUB_FONT)) for b in need_btm]
    lane_h = [LANE_H + top_extra[L] + btm_extra[L] for L in range(nlane)]
    arrow_off = [ARROW_OFF + top_extra[L] for L in range(nlane)]
    return dict(labels=out, mid=mid, occ=occ_add, lane_h=lane_h, arrow_off=arrow_off, right_days=mx / day_w)


def render(acts, c, ph, gr, zn, sh, lay, colors, co_order, path, title, caption, grid=True, button=False, marks=(),
           prog=None, day_w=DAY_W, week=False):
    """week：週単位の目盛り（1日1列の細い列。週の境目＝実線、日の境目＝点線、日付の行に月曜の日付）"""
    PURP_W = 84.0 if zn["puse"] else 0.0
    ZONE_W = 70.0 if zn["use"] else 0.0
    LABEL_W, X0 = PURP_W + ZONE_W + 110.0, 10.0
    XZ, XK = X0 + PURP_W, X0 + PURP_W + ZONE_W  # 工区の列・工事の列の左端
    items = sh["items"]
    vs = sh["view_start"]
    icrit = sh["icrit"]
    order = sorted(range(len(items)), key=lambda k: (items[k]["s"], items[k]["f"]))
    if week:  # 週単位は 18 週以上、週の切れ目までの列
        ncols = max(sh["view_end"] - vs + 1, 18 * 7, int(lay["right_days"]) + 2)
        ncols = (ncols + 6) // 7 * 7
    else:
        ncols = max(sh["view_end"] - vs + 1, 40, int(lay["right_days"]) + 2)

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
    lane_top = [lanes_top]
    for h in lay["lane_h"]:
        lane_top.append(lane_top[-1] + h)
    chart_bottom = lane_top[-1]
    legend_top = chart_bottom + 15
    W = X0 + LABEL_W + ncols * day_w + 20
    right = X0 + LABEL_W + ncols * day_w
    date_left = X0 + LABEL_W

    co_rows, row_items, xx = [], [], date_left + 4 + 30
    chips = [(nm, colors[nm]) for nm in co_order]
    if sh["overall"] and any(it["kind"] == KIND_SUM and not gr["G"][it["group"]]["co_name"] for it in items):
        chips.append(("複数の会社のまとめ", NEUTRAL))
    for nm, col in chips:
        wchip = 28 + tw(nm, 8) + 14
        if xx + wchip > right and row_items:
            co_rows.append(row_items); row_items = []; xx = date_left + 4 + 30
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
        return date_left + (d - vs) * day_w

    renderer = fig.canvas.get_renderer()

    def pt_w(t, fs, weight="normal"):
        """実際に描いたときの文字列の幅（pt）"""
        o = ax.text(0, 0, t, fontsize=fs, weight=weight)
        w = o.get_window_extent(renderer).width * 72 / fig.dpi
        o.remove()
        return w

    def rich_lines(x, y, lines, accent, fs, center, color):
        """最後の行の末尾 accent だけ色を変えて、1行ずつ描く（行の間隔は複数行の text と同じ）"""
        one = ax.text(0, 0, "国", fontsize=fs, linespacing=1.45)
        two = ax.text(0, 0, "国\n国", fontsize=fs, linespacing=1.45)
        pitch = (two.get_window_extent(renderer).height - one.get_window_extent(renderer).height) * 72 / fig.dpi
        one.remove(); two.remove()
        for j, t in enumerate(lines):
            parts = [(t, color, "normal")]
            if j == len(lines) - 1 and accent and t.endswith(accent):
                parts = [(t[:-len(accent)], color, "normal"), (accent, INAZUMA, "bold")]
            widths = [pt_w(q, fs, wt) if q else 0 for q, _, wt in parts]
            xx = x - sum(widths) / 2 if center else x
            for (q, col_, wt), wq in zip(parts, widths):
                if q:
                    ax.text(xx, y + j * pitch, q, fontsize=fs, ha="left", va="top", color=col_, weight=wt, zorder=5)
                xx += wq

    def Y(L):
        return lane_top[L] + lay["arrow_off"][L]

    has_hol = any(is_hol(vs + k) for k in range(ncols))
    has_extra = any(v3.is_extra_work(vs + k) for k in range(ncols))
    ax.text(X0, row_title + 13, f"{title}　ネットワーク工程表{caption}", fontsize=14, weight="bold", va="center")
    sub_txt = ("第1版　作成日：2026/11/05　目盛り：" + ("週単位（月曜始まり。点線は日の区切り。" if week else "日単位（") +
               "網掛け：青＝土、赤＝日" + ("、緑＝休みの曜日の工事日" if has_extra else "") +
               ("、斜線＝休工日" if has_hol else "") + "）")
    ax.text(X0, row_sub + 8, sub_txt, fontsize=9, va="center", color="#333333")
    if prog:
        ax.text(X0 + pt_w(sub_txt, 9) + 9, row_sub + 8, prog["summary"], fontsize=9, va="center", color=INAZUMA,
                weight="bold")
    wdays = "日月火水木金土"
    prev_month = None
    for k in range(ncols):
        d = dt.date.fromordinal(vs + k)
        x = X(vs + k)
        wd = (d.weekday() + 1) % 7
        fill = EXTRA_FILL if v3.is_extra_work(vs + k) else (SAT_FILL if wd == 6 else (SUN_FILL if wd == 0 else None))
        if fill:
            ax.add_patch(Rectangle((x, row_day), day_w, chart_bottom - row_day, facecolor=fill, edgecolor="none",
                                   zorder=0))
        if is_hol(vs + k):
            ax.add_patch(Rectangle((x, lanes_top), day_w, chart_bottom - lanes_top, facecolor=fill or "white",
                                   edgecolor=HOL_HATCH, hatch="////", lw=0, zorder=0.05))
        if week:  # 週単位：日の境目は点線、週の境目（月曜）は実線、月の境目は濃い線
            if d.month != prev_month:
                ax.text(x + 2, row_month + 8, f"{d.year}/{d.month}", fontsize=8, va="center")
                if k > 0:
                    ax.plot([x, x], [row_month, chart_bottom], color="#969696", lw=0.8, zorder=1)
                prev_month = d.month
            elif wd == 1:
                ax.plot([x, x], [row_day, chart_bottom], color="#A6A6A6", lw=0.6, zorder=0.6)
            elif k > 0:
                ax.plot([x, x], [row_day, chart_bottom], color="#BFBFBF", lw=0.5, linestyle=(0, (1, 1.5)), zorder=0.6)
            if wd == 1:
                ax.text(x + 2, row_day + 7, f"{d.month}/{d.day}", fontsize=7, ha="left", va="center")
            if bool(prog) and vs + k == prog["T"]:  # 基準日の印は曜日の行だけ
                ax.add_patch(Rectangle((x, row_week), day_w, 14, facecolor=INAZUMA, edgecolor="none", zorder=0.7))
            continue
        if grid:  # 日の境目の罫線（薄い灰色の細線）
            ax.plot([x, x], [row_day, chart_bottom], color="#D9D9D9", lw=0.5, zorder=0.6)
        if d.month != prev_month:
            ax.text(x + 2, row_month + 8, f"{d.year}/{d.month}", fontsize=8, va="center")
            ax.plot([x, x], [row_month, chart_bottom], color="#969696", lw=0.8, zorder=1)
            prev_month = d.month
        is_t = bool(prog) and vs + k == prog["T"]
        if is_t:  # 基準日の目盛りに印
            ax.add_patch(Rectangle((x, row_day), day_w, row_week + 14 - row_day, facecolor=INAZUMA,
                                   edgecolor="none", zorder=0.7))
        ax.text(x + day_w / 2, row_day + 7, str(d.day), fontsize=7, ha="center", va="center",
                color="white" if is_t else "black", weight="bold" if is_t else "normal")
        ax.text(x + day_w / 2, row_week + 7, wdays[wd], fontsize=7, ha="center", va="center",
                color="white" if is_t else (RED if wd == 0 else (SAT_TXT if wd == 6 else "#000000")),
                weight="bold" if is_t else "normal")
    ax.plot([X0, right], [lanes_top, lanes_top], color="#7F7F7F", lw=0.8)
    ax.plot([date_left, date_left], [row_month, chart_bottom], color="#7F7F7F", lw=0.8)
    if zn["puse"]:
        ax.text(X0 + 4, row_day + 7, "目的", fontsize=8, va="center")
    if zn["use"]:
        ax.text(XZ + 4, row_day + 7, "工区", fontsize=8, va="center")
    ax.text(XK + 4, row_day + 7, "工事", fontsize=8, va="center")

    if levels:
        ax.text(X0 + 4, row_phase + 8, "期", fontsize=9, va="center")
        for p, nm in enumerate(ph["names"], start=1):
            x1, x2 = X(ph["minS"][p]), X(ph["maxF"][p])
            t = row_phase + (levels[p] - 1) * 16 + 2
            ax.add_patch(Rectangle((x1, t), max(x2 - x1, 6), 12, facecolor="#D9E1F2", edgecolor="none", zorder=1))
            a, b = dt.date.fromordinal(ph["minS"][p]), dt.date.fromordinal(ph["last"][p])
            ax.text(x1 + 3, t + 6, f"{nm}（{a.month}/{a.day}～{b.month}/{b.day}）", fontsize=8,
                    weight="bold", color="#1F3864", va="center", zorder=2)

    for b in sh["band_order"]:
        t, btm = lane_top[sh["band_first"][b]], lane_top[sh["band_last"][b] + 1]
        ax.text(XK + 6, (t + btm) / 2, sh["bands"][b][2], fontsize=9, va="center")
        ax.plot([XK, right], [btm, btm], color="#BFBFBF", lw=0.8, zorder=1)
    # 工区と目的の欄は、上下に続く帯ごと（同じ工区でも目的が違えば別の欄）。境目の線は工区が太線、目的はもっと太い線
    if zn["use"]:
        for (_, z), l1, l2 in sh["zone_runs"]:
            t, btm = lane_top[l1], lane_top[l2 + 1]
            ax.add_patch(Rectangle((XZ, t), ZONE_W, btm - t, facecolor=ZONE_FILL, edgecolor="none", zorder=0.5))
            ax.text(XZ + ZONE_W / 2, (t + btm) / 2, zn["names"][z - 1], fontsize=9, weight="bold",
                    ha="center", va="center", zorder=2)
            ax.plot([XK, XK], [t, btm], color="#BFBFBF", lw=0.8, zorder=1)
            ax.plot([XZ, right], [btm, btm], color="#595959", lw=1.5, zorder=1.5)
    if zn["puse"]:
        for pu, l1, l2 in sh["purp_runs"]:
            t, btm = lane_top[l1], lane_top[l2 + 1]
            ax.add_patch(Rectangle((X0, t), PURP_W, btm - t, facecolor=PURP_FILL, edgecolor="none", zorder=0.5))
            ax.text(X0 + PURP_W / 2, (t + btm) / 2, "\n".join(wrap_para(zn["pnames"][pu - 1], PURP_W - 8, 9)),
                    fontsize=9, weight="bold", color=PURP_TXT, ha="center", va="center", multialignment="center",
                    linespacing=1.3, zorder=2)
            ax.plot([XZ, XZ], [t, btm], color="#BFBFBF", lw=0.8, zorder=1)
            ax.plot([X0, right], [btm, btm], color=PURP_LINE, lw=2.4, zorder=1.6)
    for rng, txt in ((sh["from_rng"], "他の期から"), (sh["to_rng"], "他の期へ")):
        if rng is not None:
            t, btm = lane_top[rng[0]], lane_top[rng[1] + 1]
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

    for k in order:
        it = items[k]
        L = it["lane"]
        yy = Y(L)
        is_item = it["kind"] in (KIND_ACT, KIND_SUM)
        col = item_color(acts, gr, colors, it)
        cr = is_item and icrit(it)
        if is_item:
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
        if k in lay["mid"]:
            lines, left, w, align = lay["mid"][k]
            ax.text(date_left + left + (0 if align == 1 else w), yy, lines[0], fontsize=8,
                    ha="left" if align == 1 else "right", va="center", color=col,
                    weight="bold" if cr else "normal", zorder=5)
        lab = lay["labels"][k]
        lines, left, w, off = lab[True]
        if lines:
            ax.text(date_left + left + (w / 2 if is_item else 0), yy - R - 1 - off, "\n".join(lines), fontsize=8,
                    ha="center" if is_item else "left", va="bottom", color=col, weight="bold" if cr else "normal",
                    linespacing=1.45, zorder=5)
        lines, left, w, off = lab[False]
        acc = prog["accent"](it) if prog else ""
        if lines and acc and lines[-1].endswith(acc):
            rich_lines(date_left + left + (w / 2 if is_item else 0), yy + R + 0.5 + off, lines, acc, 7, is_item, col)
        elif lines:
            ax.text(date_left + left + (w / 2 if is_item else 0), yy + R + 0.5 + off, "\n".join(lines), fontsize=7,
                    ha="center" if is_item else "left", va="top", color=col, linespacing=1.45, zorder=5)

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

    # イナズマ線：基準日の縦の線から、段ごとに進み・遅れの点へ折れる
    if prog and vs <= prog["Tb"] <= vs + ncols:
        xt = X(prog["Tb"])  # 基準日の作業が終わった時点＝基準日の列の右端
        pts = [(xt, lanes_top)]
        for L in range(nlane):
            xd = prog["lane_x"](sh, L)
            xd = prog["Tb"] if xd is None else min(max(xd, vs), vs + ncols)
            pts.append((X(xd), Y(L)))
        pts.append((xt, chart_bottom))
        ax.plot([q[0] for q in pts], [q[1] for q in pts], color=INAZUMA, lw=1.8, zorder=1.8,  # 矢線・文字の下
                solid_joinstyle="miter")

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

    # 画面での操作の目印（印刷はしない）：再作図ボタン、選んだ前の作業、未反映のつなぎ
    if button:
        from matplotlib.patches import FancyBboxPatch
        bw = LABEL_W - 4
        ax.add_patch(FancyBboxPatch((X0 + 2, row_month + 1), bw, 14, boxstyle="round,pad=0,rounding_size=4",
                                    facecolor="#1F3864", edgecolor="none", zorder=9))
        ax.text(X0 + 2 + bw / 2, row_month + 8, "再作図（つないだ内容を反映）" if bw >= 125 else "再作図（反映）",
                fontsize=8, weight="bold", color="white", ha="center", va="center", zorder=10)

    def item_by_id(aid):
        for it in items:
            if it["kind"] in (KIND_ACT, KIND_FROM, KIND_TO) and acts[it["act"]]["id"] == aid:
                return it

    for mk in marks:
        if mk[0] == "pick":
            it = item_by_id(mk[1])
            ax.add_patch(Circle((X(it["f"]), Y(it["lane"])), 11, facecolor="#FFC000", alpha=0.55,
                                edgecolor="#ED7D31", lw=2, zorder=11))
        elif mk[0] in ("add", "remove"):
            a, b = item_by_id(mk[1]), item_by_id(mk[2])
            xa, ya, xb, yb = X(a["f"]), Y(a["lane"]), X(b["s"]), Y(b["lane"])
            colr = "#ED7D31" if mk[0] == "add" else "#808080"
            ax.annotate("", xy=(xb, yb), xytext=(xa, ya), zorder=11,
                        arrowprops=dict(arrowstyle="-|>" if mk[0] == "add" else "-", color=colr, lw=2.2,
                                        linestyle="--", shrinkA=0, shrinkB=0, mutation_scale=10))
            ax.text((xa + xb) / 2, (ya + yb) / 2 - 8, "追加（未反映）" if mk[0] == "add" else "外す（未反映）",
                    fontsize=8, weight="bold", color=colr, ha="center", va="center", zorder=12,
                    bbox=dict(facecolor="white", edgecolor="none", alpha=0.9, pad=1))

    ly = legend_top + 12
    ax.text(X0 + 4, ly, "凡例", fontsize=8, va="center")
    x = date_left + 4
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
    x = date_left + 4
    if sh["mode"] > 0:
        ax.add_patch(Circle((x + R, ly2), R, facecolor="white", edgecolor=GHOST_C, lw=0.9,
                            linestyle=(0, (1.5, 1.5)), zorder=6))
        ax.text(x + NODE_D + 6, ly2, "他の期の作業（点線の○）", fontsize=8, va="center")
        x += 170
    if prog and vs <= prog["Tb"] <= vs + ncols:  # イナズマ線を描いたシートだけ
        ax.plot([x, x + 12, x + 24, x + 36], [ly2 + 5, ly2 - 5, ly2 + 5, ly2 - 5], color=INAZUMA, lw=1.8)
        ax.text(x + 42, ly2, prog["legend"], fontsize=8, va="center")
        x += 42 + tw(prog["legend"], 8) + 20
    note = "表記：矢線の上＝作業ID・作業名" + ("（まとめは名前）" if sh["use_sum"] else "") + \
           ("、矢線の下＝会社・日数・進捗・遅れ（進み）" if prog else "、矢線の下＝会社・日数（進捗があれば％も）")
    ax.text(x, ly2, note, fontsize=8, va="center", color=NEUTRAL)
    for r_i, row in enumerate(co_rows):
        yrow = legend_top + 24 + 18 + 18 * r_i + 9
        if r_i == 0:
            ax.text(date_left + 4, yrow, "会社：", fontsize=8, va="center")
        for nm, col, xx in row:
            ax.plot([xx, xx + 24], [yrow, yrow], color=col, lw=2.2)
            ax.text(xx + 28, yrow, nm, fontsize=8, va="center", color=col)
    fig.savefig(path, dpi=150, facecolor="white")
    plt.close(fig)
    return num


def overlaps(sh, lay, day_w=DAY_W):
    """同じ段で重なっているラベル（縦横とも）、矢線の中央から外れたラベル、積んだラベルの数"""
    items = sh["items"]
    bad = pushed = stacked = 0
    for is_name, fs in ((True, NAME_FONT), (False, SUB_FONT)):
        by_lane = {}
        for k, lab in lay["labels"].items():
            if is_name not in lab:
                continue
            lines, left, w, off = lab[is_name]
            if lines:
                by_lane.setdefault(items[k]["lane"], []).append((left, left + w, off, box_h(len(lines), fs)))
                if off > 0:
                    stacked += 1
        for L, xs in by_lane.items():
            for i in range(len(xs)):
                for j in range(i + 1, len(xs)):
                    a, b = xs[i], xs[j]
                    if a[0] < b[1] and b[0] < a[1] and a[2] < b[2] + b[3] and b[2] < a[2] + a[3]:
                        bad += 1
    for k, lab in lay["labels"].items():
        it = items[k]
        if it["kind"] in (KIND_ACT, KIND_SUM):
            e1, e2 = sh["label_span"](it)
            cx = ((e1 - sh["view_start"]) + (e2 - sh["view_start"])) / 2 * day_w
            for is_name in (True, False):
                lines, left, w, off = lab[is_name]
                if lines and abs(left + w / 2 - cx) > 0.01:
                    pushed += 1
    return bad, pushed, stacked


if __name__ == "__main__":
    outdir = sys.argv[1]
    week = "--week" in sys.argv[2:]  # 週単位の目盛りで描く（ファイル名の末尾に _週）
    dw, tail = (WEEK_DAY_W, "_週") if week else (DAY_W, "")
    v3.set_calendar(v3.load_sample_calendar())  # 工事カレンダー（土曜は休み、例外の日）
    acts = make_acts()
    c = compute(acts, is_hol=v3.is_hol)
    ph = collect_phases(acts, c)
    zn = collect_zones(acts)
    gr = collect_groups(acts, c, zn)
    colors, co_order = company_colors(acts)
    title = "○○設備更新工事（サンプル）"
    sheets = [(0, "全体工程", "（全体）")] + [(p, nm, f"（{nm}）") for p, nm in enumerate(ph["names"], start=1)]
    for mode, nm, cap in sheets:
        holder = {}

        def hook(pre):
            holder["lay"] = layout_labels(acts, c, ph, gr, zn, pre, day_w=dw)
            return holder["lay"]["occ"]

        sh = v4.build_sheet(acts, c, ph, gr, zn, mode, day_w=dw, before_links=hook, week=week)
        lay = holder["lay"]
        render(acts, c, ph, gr, zn, sh, lay, colors, co_order, f"{outdir}/preview5_{nm}{tail}.png", title, cap,
               day_w=dw, week=week)
        bad, pushed, stacked = overlaps(sh, lay, dw)
        multi = {items_k: lab[True][0] for items_k, lab in lay["labels"].items() if len(lab[True][0]) > 1}
        print(f"{nm}: overlaps={bad} pushed={pushed} stacked={stacked} mid={[lay['mid'][k][0][0] for k in lay['mid']]} "
              f"lane_h={[round(h) for h in lay['lane_h']]}")
        for k, lines in multi.items():
            print("   ", " / ".join(lines))
    with open(f"{outdir}/sample5.json", "w", encoding="utf-8") as fp:
        json.dump([dict(a, start=a["start"].isoformat(), end=a["end"].isoformat()) for a in acts],
                  fp, ensure_ascii=False, indent=1)
