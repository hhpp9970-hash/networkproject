# 工程表マクロ（NetworkSchedule）開発メモ

Excel の VBA で、工事のネットワーク工程表（矢線式・時間軸つき）を描くマクロ。横軸＝日程、縦軸＝目的→工区→工事。
ユーザーとのやりとりは日本語で、短く。

## 変えてはいけない条件（ユーザーの要件）

- 取引先と共有する。取引先には PDF で渡し、取引先の入力はマクロなしの xlsx
- シェアウェアは使えない（Excel と VBA だけで動くこと）
- バーチャートは不可。横軸＝日程・縦軸＝工事の**ネットワーク工程表**であること
- 想定する規模：数年・数期にわたり、作業 150 件くらい

## ファイル

| パス | 中身 |
|---|---|
| `src/NetworkSchedule.vba` | **編集するのはこれ**（UTF-8・LF）。マクロの全文 |
| `NetworkSchedule.bas` | ユーザーが Excel に取り込む配布物（Shift_JIS・CRLF）。`tools/build_bas.py` で作る。**直接編集しない** |
| `ネットワーク工程表_サンプル.xlsx` | 入力のサンプル。`tools/make_sample_book.py` で作る。工程データは Excel のテーブル（工程データ表） |
| `images/` | README 用の完成イメージ |
| `tools/build_bas.py` | チェック3つ → `.bas` を書く。`--verify` で src と `.bas` が同じか確かめる |
| `tools/check_vba*.py` | VBA の静的チェック（下の「チェック」） |
| `tools/make_sample_book.py` | サンプルのブックを作る（openpyxl）。作業は `tools/sample_tasks.json`、工事カレンダーは `tools/sample_calendar.json` |
| `tools/preview/` | マクロと同じ計算を Python でなぞり、完成イメージを描く（matplotlib） |

`.gitattributes` で `*.bas` の改行変換を止めている。`.bas` を Windows で clone しても CRLF のまま。

## 変更の手順

1. `src/NetworkSchedule.vba` を編集する
2. `python tools/build_bas.py` を実行する。チェックで止まったら直す
3. 表示や計算を変えたときは、`tools/preview` の Python 版も同じように直し、画像を描いて目で確かめる
4. ブランチを切ってコミットし、push して PR を作る（**main へ直接 push しない**。ユーザーがマージして取り込む）
5. サンプルの入力や「使い方」シートを変えたら `python tools/make_sample_book.py` も実行する

**Excel はこの環境では動かせない。** VBA が本当に動くかはユーザーが Excel で確かめる。変更を渡すときは、
試していないことをはっきり伝え、エラーが出たら「メッセージと止まった行」を教えてもらう。
コンパイル（デバッグ →「VBAProject のコンパイル」）も頼むとよい。

## チェック（build_bas.py が走らせる）

- `check_vba.py`：If/For/With/Select の閉じ忘れ、宣言のない代入、定義のない呼び出し、Shift_JIS にない文字
- `check_vba2.py`：式の中の、宣言も定義もない名前。Excel/VBA の組み込み（`xlWait` など）は `build_bas.py` の `KNOWN` に
  載せたものだけ通す。新しい定数を使ったら、組み込みであることを確かめてから `KNOWN` に足す
- `check_vba3.py`：予約語と同じ名前、同じ範囲での名前の重複、ByRef 引数に型の違う変数を渡していないか

## VBA の落とし穴（実際に起きたものを含む）

- **予約語を名前にしない。** 大文字・小文字を区別しないので `eNum` は `Enum` と同じ（v9 で構文エラーになった）
- `And` / `Or` は短絡評価しない。両辺を評価するので、範囲外の添字や型変換エラーは入れ子の `If` で避ける
- `Round` は偶数への丸め。日数の四捨五入は `RoundDays`（0 から遠い方へ）を使う
- 引数は既定で ByRef。ByRef の引数に型の違う変数を渡すとコンパイルエラー。新しい引数には `ByVal` を明示する
- `On Error` 文を実行すると `Err` が消える。エラー処理では最初に `Err.Number` / `Err.Description` を退避する
- Shift_JIS にない文字は使わない。`〜`（U+301C）は取り込みで `～`（U+FF5E）に変わるので、最初から `～` を使う
- 行の高さは 409pt まで。`PageSetup` は `Application.PrintCommunication = False` の間に設定する
- 図形の作成と、セル・罫線への書き込みは遅い。列ごと・セルごとに Excel を呼ばず、配列やまとめた Range で書く
- 作図中は `SpeedOn` / `SpeedOff`（再計算・イベント停止、Esc で中止）。長い処理のループには `Progress` を入れる
- `End(xlUp)` はフィルターで隠れた行を飛ばす。工程データの最後の行は `DataLastRow`（テーブルならテーブルの最後の行も見る）で求める
- `InStr(s, "")` は 1 を返す（空の文字で探すと「見つかった」になる）。空欄のセルの文字で探すときは先に長さを確かめる

## マクロの構造

公開している入口：`DrawNetworkSchedule`（作図）、`ExportSchedulePdf`（PDF）、`NS_ClickItem`（図形のクリック。OnAction）。

```
DrawNetworkSchedule
  SpeedOn → ApplyMovedNodes（動かした○を工程データへ）→ LoadData → LoadHolidays → BuildLinks → SortByLinks
  → CalcDates →（PushSuccessors）→ CalcFloat → CheckLogic → CheckDates → CollectPhases → CollectZones
  → CollectGroups → LoadCompanyColors → CalcProgress → WriteResults → DeleteGeneratedSheets
  → DrawSheet(0：全体) →［DrawSheet(-1：詳細。期がなくまとめがあるとき)］→ DrawSheet(1..期の数) → SpeedOff → メッセージ

DrawSheet
  SetViewRange → AddGhosts → BuildSheetLinks → CreateChartSheet → AssignLanes → RegisterNodes
  → FitColumns（LayoutLabels）→ SetRowHeights → DrawGrid → DrawPhaseBars → DrawProgressLine
  → DrawItems → DrawLinks → DrawNodes → DrawLegend → FinishSheet
```

- 作業は `act*` の配列（1〜nAct）。`actS` ＝開始日、`actF` ＝**終了日の翌日**（矢線の終点）。日付はシリアル値（Long）
- 1枚のシートに描くものは `it*`（項目：作業・まとめ・他の期の作業）、段は `lane*`、結合点は `node*`、ラベルは `lb*`
- 位置：`XOf(日)`（日の左端の x）、`YOf(段)`（矢線の y）。どちらも v9 からキャッシュ（列幅・行高を変えたら無効にする）
- 工程データの列は見出しで探す（`FindCol` / `EnsureCol`）。列の位置は自由

## 仕様の決めごと（変えるときはユーザーに確かめる）

- **矢線式**。結合点（○）は日の境目にある。日数が分かっていて開始・終了が空の作業は、先行作業のあと休工日を飛ばして置く。
  開始・終了を書いた作業はその日のまま
- 全体工程と期ごとのシート。全体工程は「まとめ」を1本の矢線にする。期のシートには他の期の前後の作業を薄く出す
- 縦軸は目的 → 工区 → 工事の3段（v10。目的・工区は列が無いか全部空欄なら出さない。空欄は「共通」）。
  同じ工区でも目的が違えば、目的ごとに分けて置く（工区の欄も分かれる）。まとめも目的・工区ごと。
  クリティカルパスは赤の帯、余裕は点線
- **文字の置き方（v5：D+A）**：作業名・会社は自分の矢線の中央に置き、横へ押し出さない。幅が足りなければ折り返す
  （禁則あり。会社名などは「・」でだけ折る）。日付の点線
- **クリックでつなぐ（v6）**：図形の代替テキスト `NWCLICK|…`。先行作業を工程データに書き、作り直すまで印を付ける
- **○を動かす（v7）**：○の代替テキスト `NWNODE…`。次の作図のはじめに、動かした日数を開始・終了・日数へ反映
- **実績（v8）**：基準日＝その日の作業が終わった時点（`progTb`＝基準日＋1）。遅れは暦日。完了＝実績終了あり、
  着手＝実績開始か進捗 > 0。予測は予定より前倒ししない。イナズマ線の色 RGB(230, 0, 126)。段の中では一番遅れている作業へ折る
- **速さ（v9）**：目盛り「自動」は期間 184 日までが日単位、それより長いと週単位（期のシートは一番長い期で決める）。
  全体工程は「日」でも 731 日を超えると週単位（v11 から、期のシート・詳細工程は「日」なら長くても日単位）。
  終わりのメッセージにシートごとの秒数と図形の数
- **印刷（v11）**：A3横・余白0・すべての行を1ページ（`FitToPagesTall = 1`、横は何ページでも）。タイトル列＝目的・工区・工事の列、
  タイトル行＝表題～曜日の行。余白が0なのでページ番号（フッター）は入れない
- **週単位の目盛り（v11）**：日付の列は週単位でも1日1列（列幅 0.7 の細い列）。週の境目（月曜）を実線、日の境目を点線にし、
  ○が何日にあるか数えて読めるようにする。日付の行に月曜の日付。`XOf` は日単位と同じ計算
- **工事カレンダー（v11）**：「工事カレンダー」シート（無ければ前の版の「休工日」シート＝A列の日付がすべて休み）。
  右に「曜日」の表（月～日の ○×）、左に例外の日の表（開始日・終了日・工事・内容。終了日が空欄なら1日、工事が空欄なら休み、
  下の行が優先）。`IsHoliday` は例外の日に書いてあればそれ、なければ曜日の決まり（全部 × なら曜日の決まりは使わない）
  - 日数・余裕日数は工事する日の数（`WorkDaysIn`・`StartFromEnd`）。チェックは「休工日に始まる・終わる」と「日数と開始～終了の
    工事する日の数が合わない」。矢線の下の「n日」も工事する日の数
  - 矢線の終点 `actFd`＝終了日の翌日から休工日を飛ばした日（まとめの終点・つながり・○・他の期の作業もこれ）。
    休工日をはさんで続く作業と○を共有し、余裕に見えないようにする。工程データに書く日付（計算終了）は今までどおり
  - ○を動かしたとき：始まりは次の工事する日へ、終わりは前の工事する日へ寄せ、日数は工事する日の数を書く
  - 網掛け：土＝薄い青、日＝薄い赤（日単位も週単位も）、休工日＝斜線（段の行だけ）。曜日の文字は土＝青、日＝赤

## Python のプレビュー（tools/preview）

VBA を Excel で動かせないので、計算と配置を Python でなぞって確かめてきた。ファイルの番号は版の名残で、
`network_logic3` → `4` → `5` → `6` の順に上へ重ねている（前の関数を差し替えて使う）。

- `python tools/preview/network_logic6.py 出力先` … 実績入り（基準日 1/26）の3シート。`--week` を付けると週単位の目盛り
- `python tools/preview/network_logic5.py 出力先` … 実績なしの3シートと `sample5.json`（＝`tools/sample_tasks.json`）。
  サンプルの日程は `tools/sample_calendar.json`（`network_logic3.set_calendar`）の休工日を飛ばして組む
- `python tools/preview/test_progress.py` … 進捗の計算を、VBA を写した関数と 3000 通り突き合わせる
- `python tools/preview/big_sample.py` … 4期・156件の工程で、シートごとの項目・結合点・図形のおおよその数
- 必要なもの：Python 3.9 以上、matplotlib、openpyxl、日本語フォント（Noto Sans CJK JP、なければ 游ゴシック／メイリオ）

## 残っていること・次の候補

- v9～v11 は Excel ではまだ試していない（v9 は構文エラーを直した直後）。かかった時間の表示を見て、重いシートに絞って直す
- 候補：開いているシートだけ作り直す（他のシートの○の扱いに注意）、何年にもわたる全体工程の月単位の目盛り、
  図形を減らす（クリック用の透明な枠など）
- 既知の見た目の問題：段に1つだけの作業の文字が矢線より長いと、端の○から下りる縦の帯と重なることがある

## 公開リポジトリ

このリポジトリは公開されている。実際の工事の工程データ（会社名・日程など）や、個人・会社が分かる情報は入れない。
