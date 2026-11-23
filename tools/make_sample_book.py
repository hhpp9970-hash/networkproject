"""サンプルのブック（ネットワーク工程表_サンプル.xlsx）を作る。
使い方：python tools/make_sample_book.py [出力先]（省略するとリポジトリ直下）
入力の作業は tools/sample_tasks.json（tools/preview/network_logic5.py が書き出すものと同じ）。要 openpyxl"""
import datetime as dt
import json
import math

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.worksheet.table import Table, TableStyleInfo

import pathlib
import sys

HERE = pathlib.Path(__file__).resolve().parent
OUT = sys.argv[1] if len(sys.argv) > 1 else str(HERE.parent / "ネットワーク工程表_サンプル.xlsx")
FONT = "游ゴシック"
NAVY = "1F3864"
GRAY_TXT = "595959"
thin = Side(style="thin", color="BFBFBF")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)


def F(size=11, bold=False, color="000000"):
    return Font(name=FONT, size=size, bold=bold, color=color)


def fill(c):
    return PatternFill("solid", fgColor=c)


def units(s):
    return sum(1.1 if ord(ch) < 0x2000 else 2.1 for ch in s)


def nlines(text, width):
    cap = max(width - 1.0, 1.0)
    return sum(max(1, math.ceil(units(p) / cap)) for p in str(text).split("\n"))


def base(ws, widths):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    ws.sheet_view.showGridLines = False
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0


def header_row(ws, r, labels, out_from=None):
    for c, t in enumerate(labels, 1):
        cell = ws.cell(r, c, t)
        out = out_from is not None and c >= out_from
        cell.font = F(11, True, "000000" if out else "FFFFFF")
        cell.fill = fill("D9D9D9" if out else NAVY)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BOX
    ws.row_dimensions[r].height = 22


wb = Workbook()

# ------------------------------------------------------------------ 使い方
ws = wb.active
ws.title = "使い方"
base(ws, {"A": 4, "B": 104})
ws["A1"] = "ネットワーク工程表　作図マクロ（サンプル）"
ws["A1"].font = F(14, True, NAVY)
ws.row_dimensions[1].height = 28
blocks = [
    ("準備（最初に1回だけ）", [
        "1. このブックを開き、Alt＋F11 でVBAの画面を開く",
        "2. メニューの「ファイル」→「ファイルのインポート」で NetworkSchedule.bas を選ぶ",
        "3. VBAの画面を閉じ、「名前を付けて保存」で種類を「Excel マクロ有効ブック（*.xlsm）」にして保存する",
        "※前の版のマクロを入れたブックを使い続けるときは、VBAの画面の左で古い「NetworkSchedule」を右クリック →"
        "「NetworkSchedule の解放」（エクスポートは「いいえ」）をしてから、新しい NetworkSchedule.bas をインポートする。"
        "「工程データ」には見出し「目的」「工区」「表示名」「実績開始」「実績終了」の列を足す（列の位置はどこでもよい。"
        "目的・表示名・実績の列は無くてもよい）。実績を使うときは「設定」に「進捗の基準日」「進捗の表示」の行を足す",
    ]),
    ("使い方", [
        "4. 「工程データ」シートに作業を入力する。取引先から返ってきた入力表の列に、縦軸の「目的」（任意）「工区」「工事」、"
        "シートを分ける「期」、全体工程でまとめる「まとめ」、工程表に出す短い名前「表示名」（任意）の6列を加えたもの",
        "・「工程データ」の表は Excel のテーブル（名前は「工程データ表」）にしてある。表のすぐ下の行に入力すると、表が自動で"
        "広がる。見出しの ▼ で絞り込み・並べ替えができる。絞り込んだままでも、作図はすべての作業で行う。"
        "並べ替えると、工程表の縦軸の並び（工程データに出てきた順）も変わる。"
        "自分のブックの表をテーブルにするときは、表の中を選んで Ctrl＋T（「先頭行をテーブルの見出しとして使用する」にチェック）",
        "5. Alt＋F8 →「DrawNetworkSchedule」を実行すると、次のシートが作り直される\n"
        "　・「全体工程」…まとめごとに1本の矢線（まとめが空欄の作業はそのまま）\n"
        "　・期ごとのシート（「1期」「2期」…）…作業を1つずつ\n"
        "　・「詳細工程」…「期」を使っていないときだけ、作業を1つずつ",
        "6. Alt＋F8 →「ExportSchedulePdf」を実行し、1つのPDFにまとめるか、シートごとに分けるかを選ぶ（A3横・横1ページに収める設定）",
    ]),
    ("工程表の上で先行作業をつなぐ（作業IDを打たなくてよい）", [
        "1. 作図したシートで、前の作業の矢線か名前をクリックする（その作業の終わりの○にオレンジの目印が付く）",
        "2. 後の作業の矢線か名前をクリックする。前の作業の作業IDが、後の作業の「先行作業」に書き足され、"
        "オレンジの点線矢印「追加（未反映）」が付く。続けて何組でもつなげる",
        "3. 終わったら左上の「再作図（変更を反映）」ボタンを押す。工程表に反映され、同じシートの同じ表示位置に戻る",
        "・つながっている組をもう一度同じ順にクリックすると、外すかどうか聞かれる（外すと灰色の点線「外す（未反映）」）",
        "・同じ作業を2回続けてクリックすると、選んだのを取り消す。循環になる向き（後の作業が前の作業より前にある"
        "つながり）にはつながない",
        "・まとめの矢線はクリックしてもつながらない（期のシートか「詳細工程」でつなぐ）。他の期の作業（点線の○の"
        "名前）はクリックしてつなげる",
        "・後の作業が今は前の作業より左にあっても、つなげば再作図で後ろへずれる（オレンジの矢印が左向きになる）",
        "・「再作図」のボタンと未反映の目印は印刷・PDFに出ない。未反映のままPDFにしようとすると、先に再作図する",
        "・図形を選んで動かしたいときは Ctrl を押しながらクリックする。マクロの一覧の NS_ClickItem はクリックで"
        "呼ばれるもの（直接は実行しない）",
        "・作業IDを手で入れたときは、「工程データ」の「先行作業の名前」列（マクロが書き込む）で合っているか確かめられる",
    ]),
    ("工程表の上で○を動かして日付を変える", [
        "1. 作図したシートで、○（結合点）を左右にドラッグする（いくつ動かしてもよい。縦の動きと、半日分より小さい動きは見ない）",
        "2. 左上の「再作図」ボタンを押す。その○で始まる作業の「開始」、その○で終わる作業の「終了」と「日数」を書き込む"
        "一覧が出るので、「はい」で書き込んで作図する（「いいえ」なら書き込まずに作図し、○は元の位置に戻る）",
        "3. 動かした作業の後に続く作業が、先行作業の終わりより前に始まることになったら、日数を変えずに後ろへずらすか"
        "聞かれる（「いいえ」ならずらさず、チェック列に指摘が出る）",
        "・同じ段で続く作業は○を共有している。その○を動かすと、前の作業の終わりと次の作業の始まりが一緒に動く",
        "・日数を変えずに作業ごとずらしたいときは、始まりと終わりの○を同じだけ動かす",
        "・まとめと他の期の作業の○は、動かしても反映しない（期のシートか「詳細工程」で動かす）。同じ作業を別のシートで"
        "違う日に動かしたときも反映しない",
        "・開始・終了が空欄で自動で置かれていた作業も、○を動かすと日付が書き込まれ、その日に決まる",
        "・動かした○が残ったままPDFにしようとすると、先に再作図する",
    ]),
    ("実績を入れて遅れを見る（イナズマ線）", [
        "1. 「工程データ」の「実績開始」「実績終了」に、始めた日・終えた日を入れる。途中の作業は「進捗」に何%終わったかを入れる"
        "（50 と入れても 50% とみなす。実績終了を入れた作業は100%として扱う）",
        "2. 「設定」の「進捗の基準日」に、いつの時点の実績か（その日の作業が終わった時点）を入れる。空欄なら作図した日",
        "3. 「DrawNetworkSchedule」で作図すると、基準日の縦線から段ごとに折れ曲がる赤紫の線（イナズマ線）が描かれる。"
        "左へ折れた段は遅れ、右へ折れた段は進み。基準日の日付の目盛りも赤紫になる",
        "・矢線の下に「完了」か進捗の％と、「n日遅れ」「n日進み」（赤紫の太字）が付く。表題の下には"
        "「基準日 1/26：全体 5日遅れ（完了予定 2/16 → 予測 2/21）」のように、全体の見通しが出る",
        "・遅れの日数は、基準日の時点で予定ならどこまで進んでいるはずかと、実際の進み具合との差（暦日）。"
        "まだ始まっていない作業は、予定の開始を過ぎていれば、そこから遅れとして数える。段の中に作業が複数あるときは、"
        "一番遅れている作業（遅れが無ければ一番進んでいる作業）の位置へ線を折る",
        "・予測は、着手中の作業は基準日の翌日から残りの日数（日数×（1−進捗））、まだの作業は前の作業の予測の終わりに続けて"
        "日数どおりに進むとして計算する（予定より前倒しはしない）。作業ごとの結果は「工程データ」の右の"
        "「予定進捗」「進捗の遅れ」「予測終了」「終了の遅れ」に出る",
        "・期のシートでは、他の期の作業の○の下に「1/27終了→予測2/1」のように予測が付く（次の期にどれだけ響くかが分かる）",
        "・予定（開始・終了・日数）は変えずに、実績だけを入れていく。工程を組み直したときは予定を直して版数を上げる"
        "（それ以後は直した予定との差になる）",
        "・実績の日付が基準日より後、実績終了が実績開始より前、日付でない文字、などはチェック列に出る",
        "・予定だけのPDFを出すときは、「設定」の「進捗の表示」を「しない」にして作図する（工程データの計算は行う）",
    ]),
    ("大きな工程表（作図が重いとき）", [
        "・作図中は、画面のいちばん下（ステータスバー）に、どのシートの何を描いているかと経過秒数が出る。Esc キーで中止できる"
        "（作りかけのシートは、次に作図すると作り直される）",
        "・作図が終わると、シートごとにかかった時間と図形の数がメッセージに出る。どこが重いかの目安にする",
        "・何年にもわたる工程は、目盛りを「自動」か「週」にする。日単位だと横に長くなりすぎて、重くて読めない表になる"
        "（全体の目盛りを「日」にしていても、全体工程は期間が2年を超えると週単位で描く。期ごとの目盛りの「日」は、"
        "期間が長くても日単位のまま）",
        "・全体工程は「まとめ」を使うと矢線と図形が減って軽くなる（作業ごとの矢線は期のシートで見る）",
        "・ブックが OneDrive・SharePoint にあって自動保存がオンのときは、作図の間だけ自動保存をオフにすると止まりにくい",
    ]),
    ("色と記号", [
        "・矢線と文字は会社ごとの色。「設定」の「会社の色」の表で、会社名の右のセルの塗りつぶし色を変えると、その色になる"
        "（表にない会社は自動で色が付く。複数の会社が入ったまとめは灰色）",
        "・クリティカルパスは、矢線の下の赤い帯・太線・太字・赤い○。まとめの矢線は、クリティカルな作業が重なる区間だけ帯が付く",
        "・作業＝矢線、作業の区切り＝番号付きの○、他の工事とのつながり＝点線矢印（ダミー）、余裕日数・まとめの待ち＝点線",
        "・矢線の上に作業IDと作業名（全体工程のまとめは名前）、下に会社と日数。マイルストーン（日数0）は○の横に"
        "「作業ID 名前（会社）」と1行で書く",
        "・日の境目に薄い縦線を引く。○は日の境目の線の上に乗る（線の右の日に始まり、左の日に終わる点）",
        "・週単位の目盛りでも日付の列は1日1列（細い列）。週の境目（月曜）を実線、日の境目を点線で引き、日付の行に月曜の"
        "日付を書く。○が何日にあるかは、月曜の線から点線を数えて読む（網掛けは日単位のときだけ）",
        "・網掛けは、土日が灰色、「休工日」シートの日が赤",
    ]),
    ("作図の決まりごと", [
        "・縦軸は左から「目的」→「工区」→「工事」の3段。並びはどれも工程データに最初に出てきた順。目的・工区が空欄の作業は"
        "「共通」にまとめる（「目的」「工区」の列が無いか、すべて空欄なら、その段は出さない）",
        "・同じ工区でも目的が違えば、それぞれの目的の中に分けて置く（工区の欄が目的ごとに分かれる）。目的の境目は工区より太い線",
        "・まとめは目的・工区ごとに分ける（同じ名前でも目的か工区が違えば別のまとめ）。まとめは、その最初の作業の"
        "目的・工区・工事の行に置く",
        "・まとめの中で作業のない期間（ほかの作業を待っている期間）は点線で描き、両端に結合点を置く。"
        "まとめの名前は、作業のある期間のうち最も長いところの上に書く",
        "・まとめの中の作業から他へつながるときは、まとめの矢線の途中に結合点を置き、実際の日付の位置から点線矢印を出す",
        "・名前は自分の矢線の中央に置き、隣の名前を押しのけない。前後の作業との間に入らないときは折り返し（作業IDは1行目。"
        "会社・日数は「・」の位置で）、それでも重なるときは上（会社・日数は下）に積む。そのぶん段が高くなる",
        "・「表示名」に名前を書くと、作業名の代わりにそちらを工程表に出す（作業名はそのまま残る）。セルの中で Alt＋Enter で"
        "改行すると、その位置で折り返す。名前が長くて段が高くなりすぎる作業や、週目盛りで使うときに短い名前を入れる",
        "・同じ日の点線矢印（ダミー）が途中の段の結合点を通ってしまうときは、結合点の横を通して描く（別の作業とのつながりに見えないように）",
        "・期のシートでは、他の期の作業とのつながりを、上の段「他の期から」・下の段「他の期へ」に点線の○で示す（期・工区も表示）",
        "・余裕日数とクリティカルパスは全体で計算するので、どのシートの赤い帯も同じ意味になる",
        "・日数は暦日で計算する。「休工日」シートの日は作業しない日として、開始・終了が空欄の作業を自動で置くときに飛ばす",
        "・チェック列に出るもの：休工日にかかっている作業、「日数」と開始〜終了の日数が合わない作業、先行作業の終了前に始まる作業 など",
        "・先行作業は「終わってから開始」の関係として扱う。日数0の作業はマイルストーン（○だけ）",
        "・「工程データ」の灰色の見出しの列（計算開始〜先行作業の名前）はマクロが毎回書き込む（入力不要）。"
        "予定進捗〜終了の遅れは、実績が入っているときだけ書き込む",
        "・工事名・版数・作成日・目盛り（自動／日／週、全体と期ごとで別々）は「設定」シートで変える",
    ]),
    ("社外に出すもの", [
        "・このブック（.xlsm）は社内用。取引先にはPDFを渡す",
    ]),
]
r = 3
for head, items in blocks:
    ws.cell(r, 1, head).font = F(11, True, NAVY)
    for c in (1, 2):
        ws.cell(r, c).fill = fill("D9E1F2")
    ws.row_dimensions[r].height = 22
    r += 1
    for t in items:
        cell = ws.cell(r, 2, t)
        cell.font = F(11)
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = nlines(t, 104) * 18.75 + 3
        r += 1
    r += 1

# ------------------------------------------------------------------ 設定
ws = wb.create_sheet("設定")
base(ws, {"A": 18, "B": 34, "C": 70})
ws["A1"] = "設定"
ws["A1"].font = F(14, True, NAVY)
ws.row_dimensions[1].height = 28
header_row(ws, 3, ["項目", "値", "説明"])
settings = [
    ("工事名", "○○設備更新工事（サンプル）", "工程表の表題とPDFのファイル名に使う"),
    ("版数", 1, "表題とPDFのファイル名に「第○版」として入る"),
    ("作成日", dt.date(2026, 11, 5), "表題とPDFのファイル名に入る（空欄なら当日）"),
    ("全体の目盛り", "自動", "「全体工程」シートの目盛り。「自動」「日」「週」。自動（空欄も同じ）は期間が半年までなら日、"
                         "それより長ければ週。「日」でも期間が2年を超えると週で描く"),
    ("期ごとの目盛り", "自動", "期ごとのシートと「詳細工程」シートの目盛り。「自動」「日」「週」。自動は、一番長い期が半年までなら日、"
                          "それより長ければ週（どの期も同じ目盛りになる）。「日」にすると、期間が長くても日で描く"),
    ("着工日", dt.date(2027, 1, 12), "開始日が空欄で、先行作業もない作業の開始日"),
    ("表示開始日", None, "全体工程の目盛りの開始日。空欄なら最初の作業の2日前から"),
    ("進捗の基準日", dt.date(2027, 1, 26), "実績・進捗を、この日の作業が終わった時点のものとして、イナズマ線と遅れ・予測を計算する"
                                        "（空欄なら作図した日）"),
    ("進捗の表示", "する", "「しない」にすると、実績が入っていても工程表にイナズマ線・遅れを描かない（予定だけのPDFを出すとき）"),
]
for i, (k, v, note) in enumerate(settings, start=4):
    a = ws.cell(i, 1, k)
    a.font = F(11, True)
    b = ws.cell(i, 2, v)
    b.font = F(11)
    b.fill = fill("FFF2CC")
    if isinstance(v, dt.date):
        b.number_format = "yyyy/mm/dd"
    b.alignment = Alignment(horizontal="left", vertical="center")
    c = ws.cell(i, 3, note)
    c.font = F(10, color=GRAY_TXT)
    for cell in (a, b, c):
        cell.border = BOX
    a.alignment = Alignment(vertical="center")
    c.alignment = Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[i].height = max(22, nlines(note, 70) * 16 + 4)
dv = DataValidation(type="list", formula1='"自動,日,週"', allow_blank=True, showErrorMessage=True,
                    errorTitle="入力エラー", error="「自動」「日」「週」から選んでください")
ws.add_data_validation(dv)
dv.add("B7:B8")
dv2 = DataValidation(type="list", formula1='"する,しない"', allow_blank=True, showErrorMessage=True,
                     errorTitle="入力エラー", error="「する」または「しない」を選んでください")
ws.add_data_validation(dv2)
dv2.add("B12")

# 会社の色
R0 = 15
cell = ws.cell(R0, 1, "会社の色")
cell.font = F(11, True, "FFFFFF")
cell.fill = fill(NAVY)
cell.border = BOX
cell.alignment = Alignment(vertical="center")
cell = ws.cell(R0, 2, "色（セルの塗りつぶし）")
cell.font = F(11, True, "FFFFFF")
cell.fill = fill(NAVY)
cell.border = BOX
cell.alignment = Alignment(horizontal="center", vertical="center")
note = ws.cell(R0, 3, "会社名の右のセルの塗りつぶし色が、その会社の矢線と文字の色になります。"
                      "会社を足すときは、下の行に続けて書いてください（空いた行で表は終わり）。"
                      "表にない会社や、塗りつぶしのない会社には自動で色が付きます。")
note.font = F(10, color=GRAY_TXT)
note.alignment = Alignment(vertical="center", wrap_text=True)
note.border = BOX
ws.row_dimensions[R0].height = max(22, nlines(note.value, 70) * 16 + 4)
for j, (name, color) in enumerate([("○○建設", "2F5597"), ("○○電設", "38761D"), ("○○電機", "7030A0")], start=1):
    a = ws.cell(R0 + j, 1, name)
    a.font = F(11)
    a.border = BOX
    a.alignment = Alignment(vertical="center")
    b = ws.cell(R0 + j, 2)
    b.fill = fill(color)
    b.border = BOX
    ws.row_dimensions[R0 + j].height = 22
ws.cell(R0 + 5, 1, "黄色のセルと、会社の色のセルを書き換えてください。").font = F(10, color=GRAY_TXT)

# ------------------------------------------------------------------ 工程データ
ws = wb.create_sheet("工程データ")
cols = ["目的", "工区", "工事", "期", "まとめ", "作業ID", "作業名", "表示名", "会社", "日数", "先行作業", "開始", "終了", "進捗",
        "実績開始", "実績終了",
        "計算開始", "計算終了", "全余裕(日)", "クリティカル", "予定進捗", "進捗の遅れ", "予測終了", "終了の遅れ", "チェック",
        "先行作業の名前"]
widths = [16, 10, 16, 7, 16, 10, 30, 16, 12, 7, 20, 13, 13, 8, 13, 13, 13, 13, 11, 11, 10, 12, 13, 12, 40, 50]
base(ws, {chr(65 + i): w for i, w in enumerate(widths)})
ws["A1"] = "工程データ"
ws["A1"].font = F(14, True, NAVY)
ws.row_dimensions[1].height = 28
ws["A2"] = ("取引先から返ってきた入力表の8列に、縦軸の「目的」（任意）「工区」「工事」、シートを分ける「期」、"
            "全体工程でまとめる「まとめ」、工程表に出す短い名前「表示名」（任意）の6列を加えたものです。"
            "表は Excel のテーブルで、すぐ下の行に入力すると広がります。目的・工区が空欄の作業は「共通」になります。"
            "実績は「進捗」「実績開始」「実績終了」に入れます。Q〜Z列（灰色の見出し）はマクロが書き込みます"
            "（予定進捗〜終了の遅れは、実績が入っているときだけ）。")
ws["A2"].font = F(10, color=GRAY_TXT)
HDR, NOUT = 4, 17
header_row(ws, HDR, cols, out_from=NOUT)
acts = json.load(open(HERE / "sample_tasks.json", encoding="utf-8"))
ACTUALS = {  # 作業ID: (実績開始, 実績終了, 進捗)  基準日 2027/1/26 の時点
    "A-010": (dt.date(2027, 1, 12), dt.date(2027, 1, 14), None),
    "A-020": (dt.date(2027, 1, 15), dt.date(2027, 1, 18), None),
    "B1-010": (dt.date(2027, 1, 19), dt.date(2027, 1, 19), None),
    "B1-020": (dt.date(2027, 1, 20), None, 0.75),
    "B2-010": (dt.date(2027, 1, 22), dt.date(2027, 1, 22), None),
    "B2-020": (dt.date(2027, 1, 23), dt.date(2027, 1, 26), None),
}
center_cols = {4, 6, 10, 12, 13, 14, 15, 16}
wrap_cols = {7, 8}
for i, a in enumerate(acts, start=HDR + 1):
    vals = [a.get("purpose") or None, a["zone"] or None, a["kouji"], a["phase"], a["group"] or None, a["id"], a["name"],
            a["disp"] or None, a["co"], a["dur"], a["pred"] or None, dt.date.fromisoformat(a["start"]),
            dt.date.fromisoformat(a["end"]),
            ACTUALS.get(a["id"], (None, None, None))[2], ACTUALS.get(a["id"], (None, None, None))[0],
            ACTUALS.get(a["id"], (None, None, None))[1]]
    fmts = [None, None, None, None, None, None, None, None, None, "0", None, "yyyy/mm/dd", "yyyy/mm/dd", "0%",
            "yyyy/mm/dd", "yyyy/mm/dd"]
    for c, (v, fm) in enumerate(zip(vals, fmts), 1):
        cell = ws.cell(i, c, v)
        cell.font = F(11)
        cell.border = BOX
        cell.alignment = Alignment(horizontal="center" if c in center_cols else "left", vertical="center",
                                   wrap_text=c in wrap_cols)
        if fm:
            cell.number_format = fm
    for c in range(NOUT, NOUT + 10):
        cell = ws.cell(i, c)
        cell.border = BOX
        cell.font = F(11)
        cell.alignment = Alignment(horizontal="center" if c < NOUT + 8 else "left", vertical="center")
        if c in (NOUT, NOUT + 1, NOUT + 6):
            cell.number_format = "yyyy/mm/dd"
    ws.row_dimensions[i].height = max(20, max(nlines(a["name"], 30), nlines(a["disp"] or "", 16)) * 16 + 4)
ws.freeze_panes = f"G{HDR + 1}"
# 工程データの表を Excel のテーブルにしておく（行を足すと広がる。見出しの ▼ で絞り込み・並べ替え）
tab = Table(displayName="工程データ表", ref=f"A{HDR}:{get_column_letter(len(cols))}{HDR + len(acts)}")
tab.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True)
ws.add_table(tab)

# ------------------------------------------------------------------ 休工日
ws = wb.create_sheet("休工日")
base(ws, {"A": 16, "B": 40})
ws["A1"] = "休工日"
ws["A1"].font = F(14, True, NAVY)
ws.row_dimensions[1].height = 28
ws["A2"] = ("作業しない日（年末年始・お盆・停電できない日・工場の操業日など、例外の日）だけをA列に書きます。土日は作業日のままです。"
            "工程表では赤で網掛けし、開始・終了が空欄の作業を自動で置くときは飛ばします。開始・終了が入力された作業が"
            "かかっているときは、工程データのチェック列に出ます。")
ws["A2"].font = F(10, color=GRAY_TXT)
ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
ws.merge_cells("A2:B2")
ws.row_dimensions[2].height = nlines(ws["A2"].value, 56) * 16 + 4
header_row(ws, 4, ["日付", "内容"])
hol = [(dt.date(2026, 12, 29) + dt.timedelta(days=k), "年末年始") for k in range(6)] + \
      [(dt.date(2027, 1, 11), "成人の日（例）")]
for i, (d, t) in enumerate(hol, start=5):
    a = ws.cell(i, 1, d)
    a.number_format = "yyyy/mm/dd"
    a.font = F(11)
    a.fill = fill("FFF2CC")
    a.alignment = Alignment(horizontal="center", vertical="center")
    b = ws.cell(i, 2, t)
    b.font = F(11)
    b.fill = fill("FFF2CC")
    for cell in (a, b):
        cell.border = BOX
    ws.row_dimensions[i].height = 20

wb.active = 0
wb.save(OUT)
print("saved", OUT, "rows:", len(acts))
