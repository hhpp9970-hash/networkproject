Attribute VB_Name = "NetworkSchedule"
'==============================================================================
'  ネットワーク工程表（横軸＝日程・縦軸＝目的・工区・工事）自動作図マクロ  サンプル
'
'  「工程データ」シートの表から、次のシートを作り直して作図します。
'    ・「全体工程」 … すべての作業。「まとめ」列があれば、まとめごとに1本の矢線にして
'                     名前を表示します。期があれば目盛りの上に期の帯を描きます
'    ・期ごとのシート（「1期」「2期」…） … 「期」列で分けた作業を1つずつ
'    ・「詳細工程」 … 「まとめ」を使っていて「期」が無いときだけ、すべての作業を1つずつ
'
'  実行するマクロ（Alt＋F8 から実行）
'    DrawNetworkSchedule … 工程表を作図する
'                           （計算結果とチェックは「工程データ」の右側の列に書き込みます）
'    ExportSchedulePdf   … 作図したシートをPDFに保存する（1つにまとめる／シートごと）
'
'  工程表の上で先行作業をつなぐ
'    ・作図したシートで、前の作業の矢線か名前をクリック → 後の作業の矢線か名前をクリックすると、
'      前の作業の作業IDを、後の作業の「先行作業」に書き足します（工程データが変わります）
'    ・すでにつながっている組をもう一度同じ順にクリックすると、外すかどうか聞きます
'    ・同じ作業を2回続けてクリックすると、選んだのを取り消します。循環になる向きにはつなぎません
'    ・つないだところには、オレンジの点線矢印（未反映）を描きます。左上の「再作図」で工程表に
'      反映します（もとのシート・表示位置に戻ります）。「再作図」のボタンと未反映の目印は印刷しません
'    ・まとめの矢線はクリックしてもつなぎません（期のシートか「詳細工程」でつなぎます）
'    ・NS_ClickItem は作業をクリックしたときに呼ばれるマクロです（直接は実行しません）
'
'  工程表の上で○（結合点）を動かして日付を変える
'    ・作図したシートで○を左右にドラッグしてから「再作図」を押すと、動かした日数を読み取り、
'      その○で始まる作業の「開始」、その○で終わる作業の「終了」と「日数」を工程データに書き込みます
'      （書き込む前に一覧を出して確かめます。半日分より小さい動きと縦の動きは見ません）
'    ・動かした作業の後に続く作業が、先行作業の終わりより前に始まることになったら、日数を変えずに
'      後ろへずらすか聞きます（開始・終了が空欄の作業は、もともと先行作業に合わせて置かれます）
'    ・まとめと他の期の作業の○は動かしても反映しません。同じ作業を別のシートで違う日に動かしたときも
'      反映しません
'
'  実績と遅れ（イナズマ線）  ※処理はモジュールの最後
'    ・工程データの「実績開始」「実績終了」「進捗」に実績を入れ、「設定」の「進捗の基準日」に
'      いつの時点の実績か（その日の作業が終わった時点）を入れます（空欄なら当日）
'    ・基準日の縦線から、段ごとに進み・遅れの位置へ折れるイナズマ線を描きます（左へ折れたら遅れ、
'      右へ折れたら進み）。矢線の下に「完了」か％と「n日遅れ」「n日進み」、表題の下に全体の予測を書きます
'    ・工程データの右に「予定進捗」「進捗の遅れ」「予測終了」「終了の遅れ」を書き込みます
'    ・「設定」の「進捗の表示」を「しない」にすると、実績が入っていても工程表には描きません
'
'  大きな工程表（重さ）
'    ・作図中はステータスバー（画面のいちばん下）に進み具合を出します。Esc キーで中止できます
'    ・目盛りの「自動」（空欄も同じ）は、期間が半年までのシートは日単位、それより長いシートは週単位にします。
'      全体工程は「日」にしていても、期間が2年を超えると週単位で描きます（横に長くなりすぎて読めないため）。
'      期ごとのシートと「詳細工程」は、「日」にすれば長くても日単位で描きます
'    ・作図が終わったら、シートごとにかかった時間と図形の数を表示します
'
'  縦軸
'    ・「目的」→「工区」→「工事」の3段です（「目的」「工区」の列が無いか、すべて空欄なら、その段は出しません）
'    ・並びは目的 → 工区 → 工事の順で、どれも工程データに出てきた順です。同じ工区でも目的が違えば、
'      それぞれの目的の中に分けて置きます
'    ・目的・工区が空欄の作業は「共通」にまとめます
'
'  文字（ラベル）
'    ・「表示名」の列に名前があれば、作業名の代わりにそれを書きます（セルの中の改行もそのまま）
'    ・ラベルは自分の矢線の中央に置き、隣のラベルを押しのけません。前後の作業との間に入らない
'      ときは折り返し（作業IDは1行目、会社・日数は「・」の位置で）、それでも重なるときは上
'      （会社・日数は下）に積みます。段の高さはラベルに合わせて増やします
'    ・マイルストーン（日数0）は、○の横（矢線の高さ）に「作業ID 名前（会社）」と1行で書きます
'
'  目盛り
'    ・日の境目に薄い縦線を引きます。結合点（○）はこの線の上に乗ります（線の右の日に始まり、左の日に終わる点）
'    ・週単位の目盛りでも列は1日ずつ（細い列）で、週の境目（月曜）を実線、日の境目を点線にします。
'      日付の行には月曜の日付を書きます。○が何日にあるかは、月曜の線から点線を数えて読みます
'
'  色と記号
'    ・矢線と文字は会社ごとの色です。「設定」の「会社の色」の表で、会社名の右のセル（B列）の
'      塗りつぶし色を使います。表にない会社には自動で色を付けます
'    ・クリティカルパスは、矢線の下の赤い帯・太線・太字・赤い○で示します
'    ・「休工日」シートの日は赤、土日は灰色で網掛けします
'
'  前提・簡略化しているところ
'    ・日数は暦日で数えます。ただし「休工日」シートの日は作業しない日として扱い、
'      開始・終了が空欄の作業を自動で置くときは飛ばします
'    ・先行作業は「先行作業が終わってから開始する」関係として扱います
'    ・日数が0の作業はマイルストーン（結合点だけ）として描きます
'    ・同じ工事の中で期間が重なる作業は、段を分けて描きます
'    ・余裕日数とクリティカルパスは全体で計算します（どのシートも同じ結果）
'    ・期ごとのシートでは、他の期の作業とのつながりを、上の段「他の期から」・下の段
'      「他の期へ」に点線の○で示します（ラベルは○の右に書き、重なるときは段を増やします）
'    ・まとめは目的・工区ごとに分けて、最初の作業の目的・工区・工事の行に置きます。まとめの中で
'      作業のない期間（待ち）は点線で描きます
'    ・同じ日のダミーが途中の段の結合点を通ってしまうときは、結合点の横を通して描きます
'
'  表示：矢線の上に「作業ID 作業名（表示名）」（全体工程のまとめは名前）、矢線の下に「会社・日数（・進捗）」
'==============================================================================
Option Explicit

' ---- シート名など ----
Private Const DATA_SHEET As String = "工程データ"
Private Const SETTING_SHEET As String = "設定"
Private Const HOLIDAY_SHEET As String = "休工日"
Private Const OVERALL_SHEET As String = "全体工程"
Private Const DETAIL_SHEET As String = "詳細工程"
Private Const COMMON_ZONE As String = "共通"         ' 工区（目的）が空欄の作業をまとめる名前
Private Const MARK_NAME As String = "NW_GENERATED"   ' マクロが作ったシートの目印（シート内の名前）
Private Const CLICK_TAG As String = "NWCLICK"        ' クリックでつなぐ図形の目印（代替テキストの先頭）
Private Const PICK_MARK As String = "NW_PICK"        ' 選んだ前の作業の目印の図形名
Private Const TMP_PREFIX As String = "NW_TMP_"       ' 未反映のつなぎの目印の図形名の先頭
Private Const REDRAW_BTN As String = "NW_REDRAW"     ' 再作図ボタンの図形名
Private Const NODE_TAG As String = "NWNODE"          ' 動かせる○の目印（代替テキストの先頭）

' ---- レイアウト ----
Private Const LABEL_COL_WIDTH As Double = 16    ' 工事名の列の幅
Private Const DAY_COL_WIDTH As Double = 3.4     ' 日単位のときの1日分の列幅
Private Const WEEK_DAY_COL_WIDTH As Double = 0.7 ' 週単位のときの1日分の列幅（1週でおよそ7文字分）
Private Const ZONE_COL_WIDTH As Double = 10     ' 工区名の列の幅（工区があるとき）
Private Const PURP_COL_WIDTH As Double = 12     ' 目的の列（A列）の幅（目的があるとき）
Private Const LANE_HEIGHT As Double = 44        ' 1段の高さ（pt）
Private Const ARROW_OFFSET As Double = 24       ' 段の上端から矢線までの距離（pt）
Private Const NODE_D As Double = 12             ' 結合点（○）の直径（pt）
Private Const MIN_DAY_COLS As Long = 40         ' 日単位の最小列数（凡例が収まる幅）
Private Const MIN_WEEK_COLS As Long = 18        ' 週単位の最小の週数
Private Const AUTO_DAY_SPAN As Long = 184       ' 目盛りが「自動」のとき、日単位にする期間の上限（日）
Private Const MAX_DAY_SPAN As Long = 731        ' 「日」にしていても、これより長い期間は週単位にする（日）
Private Const NAME_FONT As Double = 8           ' 矢線の上（作業ID・作業名）の文字サイズ
Private Const SUB_FONT As Double = 7            ' 矢線の下（会社・日数）の文字サイズ
Private Const LABEL_GAP As Double = 2           ' 同じ段のラベル同士のすき間（pt）
Private Const LINE_PITCH As Double = 1.75       ' ラベルの行の間隔（文字サイズの何倍か）
Private Const VGAP As Double = 1                ' 積んだラベル同士のすき間（pt）
' 折り返しの決まり：行頭に来てはいけない文字・行末に来てはいけない文字・切れ目として良い位置
Private Const NO_START As String = "・、。，．）」』】ー～…ぁぃぅぇぉっゃゅょァィゥェォッャュョ々.,)]"
Private Const NO_END As String = "（「『【(["
Private Const GOOD_AFTER As String = "・、，。 　/／）」』】,"
Private Const GOOD_BEFORE As String = "（「『【("
Private Const PHASE_ROW_HEIGHT As Double = 16   ' 全体工程の期の帯の高さ（pt）

' ---- シートに描く項目の種類 ----
Private Const KIND_ACT As Long = 0              ' 作業
Private Const KIND_FROM As Long = 1             ' 他の期の先行作業（上の段「他の期から」）
Private Const KIND_TO As Long = 2               ' 他の期の後続作業（下の段「他の期へ」）
Private Const KIND_SUM As Long = 3              ' まとめ（全体工程）

' ---- 作業データ（全体） ----
Private nAct As Long
Private actKouji() As String, actPhase() As String, actGroup() As String, actZone() As String, actPurp() As String
Private actId() As String, actName() As String, actDisp() As String, actCo() As String
Private actDur() As Long, actDurKnown() As Boolean, actPredText() As String
Private actStartIn() As Variant, actEndIn() As Variant, actProg() As Double, actRow() As Long
Private actS() As Long          ' 開始日（日付のシリアル値）
Private actF() As Long          ' 終了日の翌日（シリアル値）＝矢線の終点の位置
Private actTF() As Long         ' 全余裕（日）
Private actCrit() As Boolean    ' クリティカルパス上の作業か
Private actMsg() As String      ' チェック結果
Private preds() As Collection   ' 先行作業（添字）
Private succs() As Collection   ' 後続作業（添字）
Private topo() As Long          ' 先行関係を守った並び順

' ---- 期 ----
Private nPhase As Long, phaseName() As String, actPhaseIdx() As Long
Private phaseMinS() As Long, phaseMaxF() As Long, phaseLastDay() As Long, phaseLevel() As Long

' ---- 目的・工区 ----
Private nZone As Long, zoneName() As String, actZoneIdx() As Long, useZones As Boolean
Private nPurp As Long, purpName() As String, actPurpIdx() As Long, usePurps As Boolean
Private pairRank As Object      ' "目的|工区|工事" → 工程データに出てきた順

' ---- まとめ ----
Private nGroup As Long, actGroupIdx() As Long
Private gName() As String, gKouji() As String, gZone() As Long, gPurp() As Long
Private gS() As Long, gF() As Long, gCrit() As Boolean
Private gCoLabel() As String, gCoName() As String, gProg() As Double

' ---- 会社の色・休工日 ----
Private coDict As Object, nCo As Long, coOrder() As String
Private holDict As Object

' ---- 1枚のシートに描くもの（シートごとに作り直す） ----
Private nItem As Long
Private itAct() As Long, itKind() As Long, itGroup() As Long, itS() As Long, itF() As Long, itLane() As Long
Private actItem() As Long       ' 作業 → このシートの項目（0＝このシートに無い）
Private ghostDict As Object     ' "種類|作業" → 項目
Private nLink As Long, lkFrom() As Long, lkTo() As Long, lkA() As Long, lkB() As Long, lkCrit() As Boolean
Private nBand As Long, bandName() As String, bandZone() As Long, bandPurp() As Long
Private bandFirstLane() As Long, bandLastLane() As Long, bandSeq() As Long   ' bandSeq：帯を上から並べた順
Private nLane As Long, laneLastF() As Long, laneLastItem() As Long, laneOcc() As Collection
' ---- ラベルの置き方（LayoutLabels で決め、DrawItems で使う。位置は日付目盛りの左端からの pt）----
Private lbNameTxt() As String, lbNameN() As Long, lbNameL() As Double, lbNameW() As Double, lbNameOff() As Double
Private lbSubTxt() As String, lbSubN() As Long, lbSubL() As Double, lbSubW() As Double, lbSubOff() As Double
Private lbMid() As Long         ' マイルストーンを矢線の高さで書くとき 1：○の右、2：○の左（0：矢線の上下）
Private laneArrowOff() As Double, laneExtra() As Double     ' 段の上端から矢線まで、段の高さの増分
Private nPl As Long, plLane() As Long, plIsName() As Boolean, plX1() As Double, plX2() As Double
Private plOff() As Double, plH() As Double                   ' 置いたラベル（重なりの判定用）
Private ghostFromFirst As Long, ghostFromLast As Long, ghostToFirst As Long, ghostToLast As Long
Private nodeDict As Object, nNode As Long
Private nodeLane() As Long, nodeDay() As Long, nodeCrit() As Boolean, nodeGhost() As Boolean
Private nSide As Long, sdDay() As Long, sdSide() As Long, sdLo() As Long, sdHi() As Long   ' 結合点の横を通すダミー
Private sdFrom() As Long, sdTo() As Long, sdSlot() As Long

' ---- 描画 ----
Private wsC As Worksheet
Private sheetMode As Long, isOverall As Boolean, useSummary As Boolean, sheetCaption As String
Private viewStart As Long, viewEnd As Long, nCols As Long, dayW As Double
Private weekScale As Boolean    ' 週単位の目盛り（列は1日ずつのまま細くし、週の境目を実線・日の境目を点線にする）
Private labelCols As Long, dateCol As Long      ' 縦軸の列数（目的・工区・工事のうち使うもの）と、日付目盛りの開始列
Private zoneCol As Long                         ' 工区の列（工区が無ければ 0）。目的があれば A列、工事は labelCols 列
Private nPhaseRows As Long, rowMonth As Long, rowDay As Long, rowWeekday As Long, firstLaneRow As Long
Private legendRow As Long, legendLastRow As Long, shapeSeq As Long
Private lastShape As Shape      ' 最後に作った図形（クリックでつなぐ印を付けるため）

' ---- 工程表の上でつなぐ（クリックとクリックの間で覚えておくもの）----
Private pickId As String, pickSheet As String, pickX As Double, pickY As Double, pendingCount As Long
' ---- 工程表の上で動かした○ ----
Private calcBaseDay As Long, movedApplied As Long, pushedCount As Long
' ---- 実績と遅れ（イナズマ線）v8 ----
Private Const PROG_COLOR As Long = 8257766          ' イナズマ線の色 RGB(230, 0, 126)
Private actAS() As Variant, actAE() As Variant      ' 実績開始・実績終了（入力のまま）
Private useProg As Boolean, showProg As Boolean, progT As Long, progTb As Long
Private prgState() As Long, prgP() As Double, prgDelay() As Double, prgPlan() As Double, prgWd() As Long
Private prgFS() As Long, prgFF() As Long, prgLastFc() As Long, progPlanEnd As Long, progFcEnd As Long
Private gPrgP() As Double, gDelay() As Double
' ---- 速さ（v9）----
Private xCacheOK As Boolean, xLeft0 As Double, xColW As Double   ' XOf 用：日付の最初の列の左端と、列の幅
Private yCacheOK As Boolean, yTop() As Double                    ' YOf 用：段の行の上端
Private spdOn As Boolean, spdCalc As Long, spdEvents As Boolean, spdT0 As Single, spdLast As Single
Private progMsg As String, curSheetName As String, sheetNo As Long, nSheetsTotal As Long
Private timeNote As String, autoWeekNote As String
Private drawOK As Boolean       ' 最後の作図が最後まで終わったか（PDF にする前の再作図で中止されたら PDF にしない）


'==============================================================================
'  工程表を作図する
'==============================================================================
Public Sub DrawNetworkSchedule()
    Dim wsD As Worksheet, i As Long, p As Long, cnt As Long, txt As String, firstName As String
    Dim keepSheet As String, keepRow As Long, keepCol As Long, movedIds As Collection
    Dim errNo As Long, eDesc As String

    drawOK = False
    If spdOn Then Exit Sub      ' 作図の途中（進み具合を出している間）に、もう一度押されたときは何もしない

    On Error Resume Next
    Set wsD = ThisWorkbook.Worksheets(DATA_SHEET)
    On Error GoTo 0
    If wsD Is Nothing Then
        MsgBox "「" & DATA_SHEET & "」シートが見つかりません。", vbExclamation
        Exit Sub
    End If

    ' 作図したシートを見ていたら、作り直したあと同じシート・同じ表示位置に戻る
    If ActiveWorkbook Is ThisWorkbook Then
        If TypeName(ActiveSheet) = "Worksheet" Then
            If IsGeneratedSheet(ActiveSheet) Then
                keepSheet = ActiveSheet.Name
                keepRow = ActiveWindow.ScrollRow
                keepCol = ActiveWindow.ScrollColumn
            End If
        End If
    End If
    pickId = ""
    pickSheet = ""
    pendingCount = 0
    Application.StatusBar = False

    SpeedOn
    timeNote = ""
    autoWeekNote = ""
    progMsg = ""
    On Error GoTo ErrHandler
    Progress "工程データを読み込み中", True

    ' 工程表の上で動かした○を、工程データの開始・終了・日数に書き込む（シートを作り直す前に読む）
    movedApplied = 0
    pushedCount = 0
    Set movedIds = New Collection
    If Not ApplyMovedNodes(wsD, movedIds) Then GoTo CleanUp

    If Not LoadData(wsD) Then GoTo CleanUp
    LoadHolidays
    BuildLinks
    If Not SortByLinks() Then GoTo CleanUp
    CalcDates
    If movedIds.Count > 0 Then
        ' 後に続く作業を後ろへずらしたら、工程データを読み直して計算し直す
        If PushSuccessors(wsD, movedIds) Then
            If Not LoadData(wsD) Then GoTo CleanUp
            LoadHolidays
            BuildLinks
            If Not SortByLinks() Then GoTo CleanUp
            CalcDates
        End If
    End If
    CalcFloat
    CheckLogic
    CheckDates
    CollectPhases
    CollectZones
    CollectGroups
    LoadCompanyColors
    CalcProgress
    WriteResults wsD

    timeNote = "計算 " & Format(Elapsed(), "0.0") & "秒"
    sheetNo = 0
    nSheetsTotal = 1 + nPhase
    If nGroup > 0 And nPhase = 0 Then nSheetsTotal = nSheetsTotal + 1
    Progress "前に作ったシートを削除中", True
    DeleteGeneratedSheets
    firstName = DrawSheet(0)
    If nGroup > 0 And nPhase = 0 Then Call DrawSheet(-1)
    For p = 1 To nPhase
        Call DrawSheet(p)
    Next p
    ThisWorkbook.Worksheets(firstName).Activate
    If Len(keepSheet) > 0 Then
        If SheetExists(keepSheet) Then
            ThisWorkbook.Worksheets(keepSheet).Activate
            On Error Resume Next
            ActiveWindow.ScrollRow = keepRow
            ActiveWindow.ScrollColumn = keepCol
            On Error GoTo ErrHandler
        End If
    End If

    For i = 1 To nAct
        If Len(actMsg(i)) > 0 Then cnt = cnt + 1
    Next i
    SpeedOff
    drawOK = True
    txt = "工程表を作成しました（「" & firstName & "」"
    If nGroup > 0 And nPhase = 0 Then txt = txt & "と「" & DETAIL_SHEET & "」"
    If nPhase > 0 Then txt = txt & "と期ごとの " & nPhase & " シート"
    txt = txt & "）。"
    If movedApplied > 0 Then txt = txt & vbCrLf & "動かした○を " & movedApplied & " 件の作業の日付に反映しました。"
    If pushedCount > 0 Then txt = txt & vbCrLf & "後に続く " & pushedCount & " 件の作業を後ろへずらしました。"
    If useProg Then txt = txt & ProgressMsg()
    If Len(autoWeekNote) > 0 Then txt = txt & vbCrLf & autoWeekNote & "は、期間が2年を超えるので週単位の目盛りで描きました。"
    txt = txt & vbCrLf & "かかった時間：" & timeNote & "　（合計 " & Format(Elapsed(), "0.0") & "秒）"
    If cnt > 0 Then
        MsgBox txt & vbCrLf & "ただし " & cnt & " 件の作業に指摘があります。「" & DATA_SHEET & _
               "」のチェック列を確認してください。", vbExclamation
    Else
        MsgBox txt, vbInformation
    End If

CleanUp:
    SpeedOff
    Exit Sub

ErrHandler:
    errNo = Err.Number
    eDesc = Err.Description
    SpeedOff
    If errNo = 18 Then
        MsgBox "作図を中止しました。作りかけのシートは、次に作図すると作り直されます。", vbInformation
    Else
        MsgBox "作図中にエラーが発生しました。" & vbCrLf & eDesc, vbCritical
    End If
End Sub


'==============================================================================
'  作図したシートをPDFに保存する
'==============================================================================
Public Sub ExportSchedulePdf()
    Dim ws As Worksheet, sel() As Variant, n As Long, i As Long, ans As VbMsgBoxResult
    Dim baseName As String, tail As String, fName As Variant, folder As String
    Dim sv As Variant, created As Variant

    ' 工程表の上でつないだ内容が未反映なら、先に再作図する（シートの名前を集める前に）
    If HasPendingMarks() Or HasMovedNodes() Then
        If MsgBox("工程表の上でつないだ（外した）内容や、動かした○が、まだ工程表に反映されていません。" & vbCrLf & _
                  "再作図してからPDFにします。よろしいですか？", vbOKCancel + vbQuestion) = vbCancel Then Exit Sub
        DrawNetworkSchedule
        If Not drawOK Then Exit Sub
    End If

    For Each ws In ThisWorkbook.Worksheets
        If IsGeneratedSheet(ws) Then
            n = n + 1
            ReDim Preserve sel(1 To n)
            sel(n) = ws.Name
        End If
    Next ws
    If n = 0 Then
        MsgBox "先に DrawNetworkSchedule で工程表を作成してください。", vbExclamation
        Exit Sub
    End If
    baseName = CellText(GetSettingValue("工事名"))
    If Len(baseName) = 0 Then baseName = "工事"
    baseName = baseName & "_ネットワーク工程表"
    sv = GetSettingValue("版数")
    If Len(CellText(sv)) > 0 Then tail = "_第" & CellText(sv) & "版"
    created = GetSettingValue("作成日")
    If Not IsDate(created) Then created = Date
    tail = tail & "_" & Format(created, "yyyymmdd")

    If n = 1 Then
        ans = vbYes
    Else
        ans = MsgBox("作図した " & n & " シートを1つのPDFにまとめますか？" & vbCrLf & _
                     "「いいえ」を選ぶと、シートごとに別のPDFにします。", vbYesNoCancel + vbQuestion)
        If ans = vbCancel Then Exit Sub
    End If

    On Error GoTo ErrHandler
    If ans = vbYes Then
        fName = Application.GetSaveAsFilename(InitialFileName:=SafeFileName(baseName & tail) & ".pdf", _
                                              FileFilter:="PDF ファイル (*.pdf),*.pdf")
        If VarType(fName) = vbBoolean Then Exit Sub      ' キャンセル
        If n = 1 Then
            ThisWorkbook.Worksheets(sel(1)).ExportAsFixedFormat Type:=xlTypePDF, Filename:=CStr(fName), _
                Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
                OpenAfterPublish:=True
        Else
            ThisWorkbook.Activate
            ThisWorkbook.Worksheets(sel).Select          ' 複数シートを選んだ状態で出すと1つのPDFになる
            ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=CStr(fName), _
                Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
                OpenAfterPublish:=True
            ThisWorkbook.Worksheets(sel(1)).Select
        End If
    Else
        With Application.FileDialog(msoFileDialogFolderPicker)
            .Title = "PDFを保存するフォルダを選んでください"
            If .Show <> -1 Then Exit Sub
            folder = .SelectedItems(1)
        End With
        For i = 1 To n
            ThisWorkbook.Worksheets(sel(i)).ExportAsFixedFormat Type:=xlTypePDF, _
                Filename:=folder & Application.PathSeparator & SafeFileName(baseName & "_" & sel(i) & tail) & ".pdf", _
                Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
                OpenAfterPublish:=False
        Next i
        MsgBox n & " 個のPDFを保存しました。" & vbCrLf & folder, vbInformation
    End If
    Exit Sub

ErrHandler:
    MsgBox "PDFの保存に失敗しました。" & vbCrLf & Err.Description, vbCritical
End Sub


'==============================================================================
'  データの読み込みと計算（全体で1回）
'==============================================================================
Private Function LoadData(ByVal ws As Worksheet) As Boolean
    Dim hdr As Long, r As Long, k As Long, lastRow As Long, v As Variant
    Dim cKouji As Long, cPhase As Long, cGroup As Long, cZone As Long, cId As Long, cName As Long, cCo As Long
    Dim cDisp As Long, cPurp As Long
    Dim cDur As Long, cPred As Long, cStart As Long, cEnd As Long, cProg As Long, cAS As Long, cAE As Long

    hdr = FindHeaderRow(ws, "作業ID")
    If hdr = 0 Then
        MsgBox "「" & ws.Name & "」シートに見出し「作業ID」が見つかりません。", vbExclamation
        Exit Function
    End If
    cKouji = FindCol(ws, hdr, "工事")
    cPhase = FindCol(ws, hdr, "期")
    cGroup = FindCol(ws, hdr, "まとめ")
    cZone = FindCol(ws, hdr, "工区")
    cPurp = FindCol(ws, hdr, "目的")
    cDisp = FindCol(ws, hdr, "表示名")
    cId = FindCol(ws, hdr, "作業ID")
    cName = FindCol(ws, hdr, "作業名")
    cCo = FindCol(ws, hdr, "会社")
    cDur = FindCol(ws, hdr, "日数")
    cPred = FindCol(ws, hdr, "先行作業")
    cStart = FindCol(ws, hdr, "開始")
    cEnd = FindCol(ws, hdr, "終了")
    cProg = FindCol(ws, hdr, "進捗")
    cAS = FindCol(ws, hdr, "実績開始")
    cAE = FindCol(ws, hdr, "実績終了")
    If cKouji = 0 Or cName = 0 Or cDur = 0 Or cPred = 0 Or cStart = 0 Or cEnd = 0 Then
        MsgBox "見出しに「工事」「作業ID」「作業名」「日数」「先行作業」「開始」「終了」が必要です。", vbExclamation
        Exit Function
    End If

    lastRow = DataLastRow(ws, hdr, cId)
    nAct = 0
    For r = hdr + 1 To lastRow
        If Len(CellText(ws.Cells(r, cId).Value)) > 0 Then nAct = nAct + 1
    Next r
    If nAct = 0 Then
        MsgBox "作業データがありません。", vbExclamation
        Exit Function
    End If

    ReDim actKouji(1 To nAct): ReDim actPhase(1 To nAct): ReDim actGroup(1 To nAct): ReDim actZone(1 To nAct)
    ReDim actPurp(1 To nAct)
    ReDim actId(1 To nAct): ReDim actName(1 To nAct): ReDim actDisp(1 To nAct): ReDim actCo(1 To nAct)
    ReDim actDur(1 To nAct): ReDim actDurKnown(1 To nAct): ReDim actPredText(1 To nAct)
    ReDim actStartIn(1 To nAct): ReDim actEndIn(1 To nAct): ReDim actProg(1 To nAct): ReDim actRow(1 To nAct)
    ReDim actS(1 To nAct): ReDim actF(1 To nAct): ReDim actTF(1 To nAct): ReDim actCrit(1 To nAct)
    ReDim actMsg(1 To nAct)
    ReDim actAS(1 To nAct): ReDim actAE(1 To nAct)

    k = 0
    For r = hdr + 1 To lastRow
        If Len(CellText(ws.Cells(r, cId).Value)) > 0 Then
            k = k + 1
            actRow(k) = r
            actId(k) = NormalizeId(CellText(ws.Cells(r, cId).Value))
            actKouji(k) = CellText(ws.Cells(r, cKouji).Value)
            If Len(actKouji(k)) = 0 Then actKouji(k) = "（工事未入力）"
            If cPhase > 0 Then actPhase(k) = CellText(ws.Cells(r, cPhase).Value)
            If cGroup > 0 Then actGroup(k) = CellText(ws.Cells(r, cGroup).Value)
            If cZone > 0 Then actZone(k) = CellText(ws.Cells(r, cZone).Value)
            If cPurp > 0 Then actPurp(k) = CellText(ws.Cells(r, cPurp).Value)
            actName(k) = CellText(ws.Cells(r, cName).Value)
            If cDisp > 0 Then actDisp(k) = CellText(ws.Cells(r, cDisp).Value)
            If cCo > 0 Then actCo(k) = CellText(ws.Cells(r, cCo).Value)
            v = ws.Cells(r, cDur).Value
            If IsNum(v) Then
                actDur(k) = CLng(v)
                actDurKnown(k) = True
                If actDur(k) < 0 Then
                    actDur(k) = 0
                    AddMsg k, "日数がマイナスです"
                End If
            End If
            actPredText(k) = CellText(ws.Cells(r, cPred).Value)
            actStartIn(k) = ws.Cells(r, cStart).Value
            actEndIn(k) = ws.Cells(r, cEnd).Value
            If cProg > 0 Then
                v = ws.Cells(r, cProg).Value
                If IsNum(v) Then
                    actProg(k) = CDbl(v)
                    If actProg(k) > 1 Then actProg(k) = actProg(k) / 100   ' 25 と入っていても 25% とみなす
                End If
            End If
            If cAS > 0 Then actAS(k) = ws.Cells(r, cAS).Value
            If cAE > 0 Then actAE(k) = ws.Cells(r, cAE).Value
        End If
    Next r
    LoadData = True
End Function

' 「休工日」シートのA列の日付を読む（シートが無ければ休工日なし）
Private Sub LoadHolidays()
    Dim ws As Worksheet, r As Long, lastRow As Long, v As Variant, key As String

    Set holDict = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(HOLIDAY_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 1 To lastRow
        v = ws.Cells(r, 1).Value
        If Not IsError(v) Then
            If IsDate(v) Then
                key = CStr(DaySerial(v))
                If Not holDict.Exists(key) Then holDict.Add key, True
            End If
        End If
    Next r
End Sub

' 先行作業の文字列を、作業の添字のつながりに直す
Private Sub BuildLinks()
    Dim idDict As Object, i As Long, j As Long, t As Variant, s As String

    Set idDict = CreateObject("Scripting.Dictionary")
    ReDim preds(1 To nAct): ReDim succs(1 To nAct)
    For i = 1 To nAct
        Set preds(i) = New Collection
        Set succs(i) = New Collection
        If idDict.Exists(actId(i)) Then
            AddMsg i, "作業IDが重複しています"
        Else
            idDict.Add actId(i), i
        End If
    Next i

    For i = 1 To nAct
        s = Replace(actPredText(i), "、", ",")
        s = NormalizeId(s)              ' 全角の英数字・カンマ・空白を半角にして空白を除く
        For Each t In Split(s, ",")
            If Len(t) > 0 Then
                If idDict.Exists(CStr(t)) Then
                    j = idDict(CStr(t))
                    If j <> i Then
                        preds(i).Add j
                        succs(j).Add i
                    End If
                Else
                    AddMsg i, "先行作業「" & t & "」が見つかりません"
                End If
            End If
        Next t
    Next i
End Sub

' 先行関係を守った並び順を作る（循環があれば中止）
Private Function SortByLinks() As Boolean
    Dim indeg() As Long, i As Long, j As Long, qh As Long, qt As Long, v As Variant, ids As String

    ReDim indeg(1 To nAct): ReDim topo(1 To nAct)
    For i = 1 To nAct
        indeg(i) = preds(i).Count
        If indeg(i) = 0 Then
            qt = qt + 1
            topo(qt) = i
        End If
    Next i
    qh = 1
    Do While qh <= qt
        i = topo(qh)
        qh = qh + 1
        For Each v In succs(i)
            j = v
            indeg(j) = indeg(j) - 1
            If indeg(j) = 0 Then
                qt = qt + 1
                topo(qt) = j
            End If
        Next v
    Loop

    If qt < nAct Then
        For i = 1 To nAct
            If indeg(i) > 0 Then ids = ids & IIf(Len(ids) > 0, "、", "") & actId(i)
        Next i
        MsgBox "先行作業の関係が循環しています。次の作業の先行作業を確認してください。" & vbCrLf & ids, vbExclamation
        Exit Function
    End If
    SortByLinks = True
End Function

' 開始日・終了日を決める（空欄は先行作業から求める。休工日は飛ばす）
Private Sub CalcDates()
    Dim k As Long, i As Long, baseDay As Long, m As Long, v As Variant, sv As Variant, found As Boolean

    sv = GetSettingValue("着工日")
    If IsDate(sv) Then
        baseDay = DaySerial(sv)
    Else
        For i = 1 To nAct
            If IsDate(actStartIn(i)) Then
                If Not found Then
                    baseDay = DaySerial(actStartIn(i))
                    found = True
                ElseIf DaySerial(actStartIn(i)) < baseDay Then
                    baseDay = DaySerial(actStartIn(i))
                End If
            End If
        Next i
        If Not found Then baseDay = DaySerial(Date)
    End If
    calcBaseDay = baseDay

    For k = 1 To nAct
        i = topo(k)
        ' 開始日
        If IsDate(actStartIn(i)) Then
            actS(i) = DaySerial(actStartIn(i))
        ElseIf IsDate(actEndIn(i)) And actDurKnown(i) Then
            actS(i) = StartFromEnd(DaySerial(actEndIn(i)), actDur(i))
        Else
            m = baseDay
            For Each v In preds(i)
                If actF(v) > m Then m = actF(v)
            Next v
            actS(i) = NextWorkDay(m)
        End If
        ' 終了日の翌日
        actF(i) = FinishFor(i, actS(i))
        If Not actDurKnown(i) And Not IsDate(actEndIn(i)) Then AddMsg i, "日数・終了日が空欄のため1日として描きました"
        If actF(i) < actS(i) Then
            AddMsg i, "終了日が開始日より前です"
            actF(i) = actS(i)
        End If
    Next k
End Sub

' 全余裕とクリティカルパス（後ろから計算）
Private Sub CalcFloat()
    Dim k As Long, i As Long, pe As Long, lf As Long, ls As Long, v As Variant
    Dim lateF() As Long

    ReDim lateF(1 To nAct)
    pe = actF(1)
    For i = 2 To nAct
        If actF(i) > pe Then pe = actF(i)
    Next i
    For k = nAct To 1 Step -1
        i = topo(k)
        If succs(i).Count = 0 Then
            lf = pe
        Else
            lf = 2147483647
            For Each v In succs(i)
                ls = lateF(v) - (actF(v) - actS(v))
                If ls < lf Then lf = ls
            Next v
        End If
        lateF(i) = lf
        actTF(i) = lf - actF(i)
        actCrit(i) = (actTF(i) <= 0)
    Next k
End Sub

' 先行作業の終了前に始まっている作業を指摘する
Private Sub CheckLogic()
    Dim i As Long, v As Variant
    For i = 1 To nAct
        For Each v In preds(i)
            If actS(i) < actF(v) Then AddMsg i, "先行作業「" & actId(v) & "」の終了前に開始しています"
        Next v
    Next i
End Sub

' 日付のチェック（開始・終了が入力されている作業だけ）
'   ・休工日にかかっていないか
'   ・「日数」と、開始～終了の日数（暦日）が合うか
Private Sub CheckDates()
    Dim i As Long, d As Long, n As Long, hits As String
    For i = 1 To nAct
        If IsDate(actStartIn(i)) And IsDate(actEndIn(i)) Then
            If actF(i) > actS(i) Then
                hits = ""
                n = 0
                For d = actS(i) To actF(i) - 1
                    If IsHoliday(d) Then
                        n = n + 1
                        If n <= 3 Then hits = hits & IIf(Len(hits) > 0, "・", "") & Format(CDate(d), "m/d")
                    End If
                Next d
                If n > 0 Then AddMsg i, "休工日（" & hits & IIf(n > 3, "ほか", "") & "）にかかっています"
                If actDurKnown(i) And actDur(i) > 0 And (actF(i) - actS(i)) <> actDur(i) Then
                    AddMsg i, "日数（" & actDur(i) & "日）と開始～終了（" & (actF(i) - actS(i)) & "日）が合いません"
                End If
            End If
        End If
    Next i
End Sub

' 期の一覧（開始の早い順）と、期ごとの期間
Private Sub CollectPhases()
    Dim dict As Object, i As Long, p As Long, k As Long, j As Long
    Dim phNames() As String, mins() As Long, ord() As Long

    Set dict = CreateObject("Scripting.Dictionary")
    nPhase = 0
    ReDim actPhaseIdx(1 To nAct)
    ReDim phNames(1 To nAct): ReDim mins(1 To nAct)
    For i = 1 To nAct
        If Len(actPhase(i)) > 0 Then
            If Not dict.Exists(actPhase(i)) Then
                nPhase = nPhase + 1
                dict.Add actPhase(i), nPhase
                phNames(nPhase) = actPhase(i)
                mins(nPhase) = actS(i)
            ElseIf actS(i) < mins(dict(actPhase(i))) Then
                mins(dict(actPhase(i))) = actS(i)
            End If
        End If
    Next i
    If nPhase = 0 Then Exit Sub

    ReDim ord(1 To nPhase)
    For p = 1 To nPhase
        ord(p) = p
        k = p
        Do While k > 1
            If mins(ord(k)) < mins(ord(k - 1)) Then
                j = ord(k - 1): ord(k - 1) = ord(k): ord(k) = j
                k = k - 1
            Else
                Exit Do
            End If
        Loop
    Next p

    ReDim phaseName(1 To nPhase): ReDim phaseMinS(1 To nPhase)
    ReDim phaseMaxF(1 To nPhase): ReDim phaseLastDay(1 To nPhase)
    dict.RemoveAll
    For p = 1 To nPhase
        phaseName(p) = phNames(ord(p))
        dict.Add phaseName(p), p
    Next p
    For i = 1 To nAct
        If Len(actPhase(i)) > 0 Then
            p = dict(actPhase(i))
            actPhaseIdx(i) = p
            If phaseMinS(p) = 0 Or actS(i) < phaseMinS(p) Then phaseMinS(p) = actS(i)
            If actF(i) > phaseMaxF(p) Then phaseMaxF(p) = actF(i)
            If LastDay(i) > phaseLastDay(p) Then phaseLastDay(p) = LastDay(i)
        Else
            AddMsg i, "期が空欄です（全体工程にだけ描きます）"
        End If
    Next i
End Sub

' 工区と目的の一覧（工程データに出てきた順。空欄は「共通」）と、目的・工区・工事の組の順位（出てきた順）
Private Sub CollectZones()
    Dim dict As Object, pDict As Object, i As Long, z As String, kk As String

    Set dict = CreateObject("Scripting.Dictionary")
    Set pDict = CreateObject("Scripting.Dictionary")
    Set pairRank = CreateObject("Scripting.Dictionary")
    nZone = 0
    nPurp = 0
    useZones = False
    usePurps = False
    ReDim actZoneIdx(1 To nAct): ReDim zoneName(1 To nAct)
    ReDim actPurpIdx(1 To nAct): ReDim purpName(1 To nAct)
    For i = 1 To nAct
        If Len(actZone(i)) > 0 Then useZones = True
        If Len(actPurp(i)) > 0 Then usePurps = True
    Next i
    For i = 1 To nAct
        z = actZone(i)
        If Len(z) = 0 Then z = COMMON_ZONE
        If Not dict.Exists(z) Then
            nZone = nZone + 1
            dict.Add z, nZone
            zoneName(nZone) = z
        End If
        actZoneIdx(i) = dict(z)
        z = actPurp(i)
        If Len(z) = 0 Then z = COMMON_ZONE
        If Not pDict.Exists(z) Then
            nPurp = nPurp + 1
            pDict.Add z, nPurp
            purpName(nPurp) = z
        End If
        actPurpIdx(i) = pDict(z)
        kk = actPurpIdx(i) & "|" & actZoneIdx(i) & "|" & actKouji(i)
        If Not pairRank.Exists(kk) Then pairRank.Add kk, pairRank.Count + 1
    Next i
End Sub

' まとめ（入れ子）の一覧と、まとめごとの期間・会社・進捗（同じ名前でも目的か工区が違えば別のまとめ）
Private Sub CollectGroups()
    Dim dict As Object, coSet As Object, i As Long, g As Long, wsum As Double, psum As Double, key As String

    Set dict = CreateObject("Scripting.Dictionary")
    nGroup = 0
    ReDim actGroupIdx(1 To nAct)
    ReDim gName(1 To nAct): ReDim gKouji(1 To nAct): ReDim gS(1 To nAct): ReDim gF(1 To nAct)
    ReDim gCrit(1 To nAct): ReDim gCoLabel(1 To nAct): ReDim gCoName(1 To nAct): ReDim gProg(1 To nAct)
    ReDim gZone(1 To nAct): ReDim gPurp(1 To nAct)
    For i = 1 To nAct
        If Len(actGroup(i)) > 0 Then
            key = actPurpIdx(i) & "|" & actZoneIdx(i) & "|" & actGroup(i)
            If Not dict.Exists(key) Then
                nGroup = nGroup + 1
                dict.Add key, nGroup
                gName(nGroup) = actGroup(i)
                gKouji(nGroup) = actKouji(i)
                gZone(nGroup) = actZoneIdx(i)
                gPurp(nGroup) = actPurpIdx(i)
                gS(nGroup) = actS(i)
                gF(nGroup) = actF(i)
            End If
            g = dict(key)
            actGroupIdx(i) = g
            If actS(i) < gS(g) Then gS(g) = actS(i)
            If actF(i) > gF(g) Then gF(g) = actF(i)
            If actCrit(i) Then gCrit(g) = True
        End If
    Next i

    For g = 1 To nGroup
        Set coSet = CreateObject("Scripting.Dictionary")
        wsum = 0
        psum = 0
        For i = 1 To nAct
            If actGroupIdx(i) = g Then
                If Len(actCo(i)) > 0 Then
                    If Not coSet.Exists(actCo(i)) Then coSet.Add actCo(i), True
                End If
                wsum = wsum + (actF(i) - actS(i))
                psum = psum + actProg(i) * (actF(i) - actS(i))
            End If
        Next i
        If coSet.Count = 1 Then
            gCoName(g) = coSet.Keys()(0)
            gCoLabel(g) = gCoName(g)
        ElseIf coSet.Count > 1 Then
            gCoLabel(g) = coSet.Keys()(0) & "ほか" & (coSet.Count - 1) & "社"
        End If
        If wsum > 0 Then gProg(g) = psum / wsum
    Next g
End Sub

' 会社の色：「設定」の「会社の色」の表（A列＝会社名、B列の塗りつぶし色）。無い会社は自動で割り当てる
Private Sub LoadCompanyColors()
    Dim ws As Worksheet, seen As Object, r As Long, r0 As Long, i As Long, used As Long, nm As String

    Set coDict = CreateObject("Scripting.Dictionary")
    Set seen = CreateObject("Scripting.Dictionary")
    nCo = 0
    ReDim coOrder(1 To nAct)
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SETTING_SHEET)
    On Error GoTo 0
    If Not ws Is Nothing Then
        For r = 1 To 200
            If CellText(ws.Cells(r, 1).Value) = "会社の色" Then
                r0 = r
                Exit For
            End If
        Next r
        If r0 > 0 Then
            For r = r0 + 1 To r0 + 100
                nm = CellText(ws.Cells(r, 1).Value)
                If Len(nm) = 0 Then Exit For
                If ws.Cells(r, 2).Interior.ColorIndex <> xlColorIndexNone Then
                    If Not coDict.Exists(nm) Then coDict.Add nm, CLng(ws.Cells(r, 2).Interior.Color)
                End If
            Next r
        End If
    End If
    ' 表にない会社は、工程データに出てきた順に自動で色を付ける（凡例もこの順）
    For i = 1 To nAct
        nm = actCo(i)
        If Len(nm) > 0 Then
            If Not coDict.Exists(nm) Then
                coDict.Add nm, AutoColor(used)
                used = used + 1
            End If
            If Not seen.Exists(nm) Then
                seen.Add nm, True
                nCo = nCo + 1
                coOrder(nCo) = nm
            End If
        End If
    Next i
End Sub

Private Function AutoColor(ByVal n As Long) As Long
    Select Case n Mod 8
        Case 0: AutoColor = RGB(47, 85, 151)
        Case 1: AutoColor = RGB(56, 118, 29)
        Case 2: AutoColor = RGB(112, 48, 160)
        Case 3: AutoColor = RGB(140, 90, 43)
        Case 4: AutoColor = RGB(27, 127, 134)
        Case 5: AutoColor = RGB(180, 64, 126)
        Case 6: AutoColor = RGB(31, 56, 100)
        Case Else: AutoColor = RGB(127, 127, 0)
    End Select
End Function

Private Function CompanyColor(ByVal nm As String) As Long
    If Len(nm) > 0 Then
        If coDict.Exists(nm) Then
            CompanyColor = coDict(nm)
            Exit Function
        End If
    End If
    CompanyColor = RGB(64, 64, 64)
End Function

' 計算結果を「工程データ」の右側に書き込む
Private Sub WriteResults(ByVal ws As Worksheet)
    Dim hdr As Long, cols(1 To 6) As Long, heads As Variant, i As Long, k As Long, lastRow As Long

    heads = Array("計算開始", "計算終了", "全余裕(日)", "クリティカル", "チェック", "先行作業の名前")
    hdr = FindHeaderRow(ws, "作業ID")
    lastRow = DataLastRow(ws, hdr, FindCol(ws, hdr, "作業ID"))
    For k = 1 To 6
        cols(k) = EnsureCol(ws, hdr, CStr(heads(k - 1)))
        If lastRow > hdr Then ws.Range(ws.Cells(hdr + 1, cols(k)), ws.Cells(lastRow, cols(k))).ClearContents
    Next k
    For i = 1 To nAct
        ws.Cells(actRow(i), cols(1)).Value = CDate(actS(i))
        ws.Cells(actRow(i), cols(2)).Value = CDate(LastDay(i))
        ws.Cells(actRow(i), cols(1)).NumberFormat = "yyyy/mm/dd"
        ws.Cells(actRow(i), cols(2)).NumberFormat = "yyyy/mm/dd"
        ws.Cells(actRow(i), cols(3)).Value = actTF(i)
        ws.Cells(actRow(i), cols(4)).Value = IIf(actCrit(i), "●", "")
        ws.Cells(actRow(i), cols(5)).Value = actMsg(i)
        ws.Cells(actRow(i), cols(6)).Value = PredNames(i)
    Next i
    WriteProgressCols ws, hdr, lastRow
End Sub

' 先行作業の作業IDと作業名（入力した作業IDが合っているか確かめるため）。「／」区切り
Private Function PredNames(ByVal i As Long) As String
    Dim v As Variant, t As String
    For Each v In preds(i)
        If Len(t) > 0 Then t = t & "／"
        t = t & actId(v) & " " & OneLine("", actName(v))
    Next v
    PredNames = t
End Function

' 作業の最終日（マイルストーンは開始日）
' 作業 i を開始日 s に置いたときの終了日の翌日（日数0は開始日、終了日があればそれ、なければ日数の作業日）
Private Function FinishFor(ByVal i As Long, ByVal s As Long) As Long
    If actDurKnown(i) And actDur(i) = 0 Then
        FinishFor = s
    ElseIf IsDate(actEndIn(i)) Then
        FinishFor = DaySerial(actEndIn(i)) + 1
    ElseIf actDurKnown(i) Then
        FinishFor = AddWorkDays(s, actDur(i))
    Else
        FinishFor = AddWorkDays(s, 1)
    End If
End Function

Private Function LastDay(ByVal i As Long) As Long
    If actF(i) > actS(i) Then LastDay = actF(i) - 1 Else LastDay = actS(i)
End Function


'==============================================================================
'  1枚のシートを作図する。戻り値はシート名
'    mode＝0：全体工程（まとめがあれば、まとめ単位）
'    mode＝-1：詳細工程（すべての作業を1つずつ）
'    mode＞0：その期のシート
'==============================================================================
Private Function DrawSheet(ByVal mode As Long) As String
    Dim i As Long, g As Long, nm As String, gItem() As Long, t0 As Single

    sheetMode = mode
    isOverall = (mode = 0)
    useSummary = (isOverall And nGroup > 0)
    ReDim actItem(1 To nAct)
    ReDim itAct(1 To 3 * nAct): ReDim itKind(1 To 3 * nAct): ReDim itGroup(1 To 3 * nAct)
    ReDim itS(1 To 3 * nAct): ReDim itF(1 To 3 * nAct): ReDim itLane(1 To 3 * nAct)
    If nGroup > 0 Then ReDim gItem(1 To nGroup)
    nItem = 0
    For i = 1 To nAct
        If mode <= 0 Or actPhaseIdx(i) = mode Then
            If useSummary And actGroupIdx(i) > 0 Then
                g = actGroupIdx(i)
                If gItem(g) = 0 Then
                    nItem = nItem + 1
                    itAct(nItem) = i
                    itKind(nItem) = KIND_SUM
                    itGroup(nItem) = g
                    itS(nItem) = gS(g)
                    itF(nItem) = gF(g)
                    gItem(g) = nItem
                End If
                actItem(i) = gItem(g)
            Else
                nItem = nItem + 1
                itAct(nItem) = i
                itKind(nItem) = KIND_ACT
                itS(nItem) = actS(i)
                itF(nItem) = actF(i)
                actItem(i) = nItem
            End If
        End If
    Next i

    Select Case mode
        Case 0
            nm = UniqueSheetName(OVERALL_SHEET)
            sheetCaption = "（全体）"
        Case -1
            nm = UniqueSheetName(DETAIL_SHEET)
            sheetCaption = "（詳細）"
        Case Else
            nm = UniqueSheetName(SafeSheetName(phaseName(mode)))
            sheetCaption = "（" & phaseName(mode) & "）"
    End Select

    curSheetName = nm
    sheetNo = sheetNo + 1
    progMsg = "「" & nm & "」（" & sheetNo & "/" & nSheetsTotal & "）"
    t0 = Elapsed()
    Progress "配置を計算中", True
    SetViewRange
    AddGhosts
    BuildSheetLinks
    CreateChartSheet nm
    AssignLanes
    RegisterNodes
    FitColumns
    SetRowHeights
    DrawGrid
    If isOverall And nPhase > 0 Then DrawPhaseBars
    DrawProgressLine
    DrawItems
    DrawLinks
    DrawNodes
    DrawLegend
    FinishSheet
    timeNote = timeNote & "、" & nm & " " & Format(Elapsed() - t0, "0.0") & "秒（図形 " & wsC.Shapes.Count & "）"
    DrawSheet = nm
End Function

' 目盛りの単位と表示する期間（このシートの項目から決める）
Private Sub SetViewRange()
    Dim k As Long, minS As Long, maxF As Long, sv As Variant, sc As String

    minS = itS(1): maxF = itF(1)
    For k = 2 To nItem
        If itS(k) < minS Then minS = itS(k)
        If itF(k) > maxF Then maxF = itF(k)
    Next k
    viewStart = minS - 2
    If isOverall Then
        sv = GetSettingValue("表示開始日")
        If IsDate(sv) Then
            viewStart = DaySerial(sv)
            If viewStart > minS Then viewStart = minS
        End If
    End If
    viewEnd = maxF + 2

    ' 目盛り：「週」「日」「自動」（空欄も自動）。自動は期間が AUTO_DAY_SPAN 日までなら日単位。
    ' 全体工程は「日」でも期間が MAX_DAY_SPAN 日を超えるときは、横に長くなりすぎるので週単位にする
    ' （期のシート・詳細工程は「日」なら、長くても日単位のまま）
    If isOverall Then sv = GetSettingValue("全体の目盛り") Else sv = GetSettingValue("期ごとの目盛り")
    If Len(CellText(sv)) = 0 Then sv = GetSettingValue("目盛り")
    sc = CellText(sv)
    weekScale = False
    If InStr(sc, "週") > 0 Then
        weekScale = True
    ElseIf InStr(sc, "日") > 0 Then
        If isOverall And viewEnd - viewStart > MAX_DAY_SPAN Then
            weekScale = True
            If Len(autoWeekNote) > 0 Then autoWeekNote = autoWeekNote & "・"
            autoWeekNote = autoWeekNote & "「" & curSheetName & "」"
        End If
    ElseIf AutoSpan() > AUTO_DAY_SPAN Then
        weekScale = True
    End If
    If weekScale Then viewStart = viewStart - (Weekday(CDate(viewStart), vbMonday) - 1)   ' 月曜から
End Sub

' 目盛りが「自動」のときに見る期間。期のシートは、どの期も同じ目盛りになるよう一番長い期で決める
Private Function AutoSpan() As Long
    Dim p As Long
    AutoSpan = viewEnd - viewStart
    If sheetMode >= 1 Then
        For p = 1 To nPhase
            If phaseMaxF(p) - phaseMinS(p) + 4 > AutoSpan Then AutoSpan = phaseMaxF(p) - phaseMinS(p) + 4
        Next p
    End If
End Function

' 期のシートだけ：他の期の先行作業・後続作業を「他の期から」「他の期へ」の項目として加える
Private Sub AddGhosts()
    Dim k As Long, n0 As Long, a As Long, v As Variant, g As Long

    Set ghostDict = CreateObject("Scripting.Dictionary")
    If sheetMode <= 0 Then Exit Sub
    n0 = nItem
    For k = 1 To n0
        a = itAct(k)
        For Each v In preds(a)
            If actItem(v) = 0 Then g = GhostItem(KIND_FROM, CLng(v), ClampDay(actF(v)))
        Next v
        For Each v In succs(a)
            If actItem(v) = 0 Then g = GhostItem(KIND_TO, CLng(v), ClampDay(actS(v)))
        Next v
    Next k
End Sub

Private Function GhostItem(ByVal kind As Long, ByVal a As Long, ByVal d As Long) As Long
    Dim key As String
    key = kind & "|" & a
    If ghostDict.Exists(key) Then
        GhostItem = ghostDict(key)
    Else
        nItem = nItem + 1
        itAct(nItem) = a
        itKind(nItem) = kind
        itS(nItem) = d
        itF(nItem) = d
        ghostDict.Add key, nItem
        GhostItem = nItem
    End If
End Function

' 表示する期間の外の日付は、端に寄せる（左端は枠線と重ならないよう1日内側）
Private Function ClampDay(ByVal d As Long) As Long
    If d < viewStart + 1 Then d = viewStart + 1
    If d > viewEnd Then d = viewEnd
    ClampDay = d
End Function

' このシートのつながり（先行→後続）。全体工程のまとめどうしは、余裕の最も少ない1本にまとめる
Private Sub BuildSheetLinks()
    Dim a As Long, v As Variant, fromK As Long, toK As Long, gk As Long, fa As Long, cnt As Long
    Dim pairDict As Object

    cnt = 0
    For a = 1 To nAct
        cnt = cnt + preds(a).Count
    Next a
    nLink = 0
    ReDim lkFrom(1 To 2 * cnt + 2): ReDim lkTo(1 To 2 * cnt + 2): ReDim lkA(1 To 2 * cnt + 2)
    ReDim lkB(1 To 2 * cnt + 2): ReDim lkCrit(1 To 2 * cnt + 2)
    Set pairDict = CreateObject("Scripting.Dictionary")

    For a = 1 To nAct
        If actItem(a) > 0 Then
            toK = actItem(a)
            For Each v In preds(a)
                If actItem(v) > 0 Then
                    fromK = actItem(v)
                    fa = actF(v)
                Else
                    fromK = ghostDict(KIND_FROM & "|" & v)
                    fa = itF(fromK)
                End If
                If fromK <> toK Then
                    AddLink fromK, toK, fa, actS(a), (actCrit(v) And actCrit(a) And actF(v) = actS(a)), pairDict
                End If
            Next v
            For Each v In succs(a)
                If actItem(v) = 0 Then
                    gk = ghostDict(KIND_TO & "|" & v)
                    AddLink toK, gk, actF(a), itS(gk), (actCrit(a) And actCrit(v) And actF(a) = actS(v)), pairDict
                End If
            Next v
        End If
    Next a
End Sub

Private Sub AddLink(ByVal fromK As Long, ByVal toK As Long, ByVal fa As Long, ByVal sb As Long, _
                    ByVal cr As Boolean, ByVal pairDict As Object)
    Dim key As String, j As Long
    key = fromK & "|" & toK
    If pairDict.Exists(key) Then
        j = pairDict(key)
        If (sb - fa) < (lkB(j) - lkA(j)) Or ((sb - fa) = (lkB(j) - lkA(j)) And fa > lkA(j)) Then
            lkA(j) = fa
            lkB(j) = sb
            lkCrit(j) = cr
        End If
    Else
        nLink = nLink + 1
        lkFrom(nLink) = fromK
        lkTo(nLink) = toK
        lkA(nLink) = fa
        lkB(nLink) = sb
        lkCrit(nLink) = cr
        pairDict.Add key, nLink
    End If
End Sub


'==============================================================================
'  段の割り当てと結合点
'==============================================================================
' 目的・工区・工事ごとに（並びは目的の順 → 工区の順 → 工程データに出てきた順）、開始日の早い順に段へ置く。
' 重なる項目は段を分ける。
' 前の項目の終点と次の項目の始点が同じ日でも、その2つが直接つながっていなければ結合点を共有しない。
' 期のシートでは、いちばん上に「他の期から」、いちばん下に「他の期へ」の段を置く
' （他の期の作業のラベルが重なるときは、その中で段を増やす）。
Private Sub AssignLanes()
    Dim bandDict As Object, itBand() As Long, members() As Long, kk As String
    Dim bandKey() As Double, bi As Long
    Dim k As Long, b As Long, m As Long, j As Long, L As Long, cnt As Long, n As Long
    Dim placed As Boolean

    Set bandDict = CreateObject("Scripting.Dictionary")
    ReDim bandName(1 To nItem): ReDim bandZone(1 To nItem): ReDim bandPurp(1 To nItem): ReDim bandKey(1 To nItem)
    ReDim bandSeq(1 To nItem): ReDim itBand(1 To nItem)
    nBand = 0
    For k = 1 To nItem
        If itKind(k) = KIND_ACT Or itKind(k) = KIND_SUM Then
            kk = ItemPurp(k) & "|" & ItemZone(k) & "|" & ItemKouji(k)
            If Not bandDict.Exists(kk) Then
                nBand = nBand + 1
                bandDict.Add kk, nBand
                bandName(nBand) = ItemKouji(k)
                bandZone(nBand) = ItemZone(k)
                bandPurp(nBand) = ItemPurp(k)
                ' 並べる順の値：目的の順 → 工区の順 → 目的・工区・工事の組が工程データに出てきた順
                bandKey(nBand) = (CDbl(bandPurp(nBand)) * (nZone + 1) + bandZone(nBand)) * (pairRank.Count + 1) _
                                 + pairRank(kk)
            End If
            itBand(k) = bandDict(kk)
        End If
    Next k

    ' 帯（目的・工区・工事）の並び
    For b = 1 To nBand
        m = b
        Do While m > 1
            j = bandSeq(m - 1)
            If bandKey(b) < bandKey(j) Then
                bandSeq(m) = j
                m = m - 1
            Else
                Exit Do
            End If
        Loop
        bandSeq(m) = b
    Next b

    ReDim bandFirstLane(1 To nBand): ReDim bandLastLane(1 To nBand)
    ReDim laneLastF(1 To nItem + 2): ReDim laneLastItem(1 To nItem + 2): ReDim laneOcc(1 To nItem + 2)
    ReDim members(1 To nItem)
    nLane = 0

    ' 他の期から（いちばん上）
    ghostFromFirst = 0
    ghostFromLast = 0
    n = StackGhosts(KIND_FROM, 1)
    If n > 0 Then
        ghostFromFirst = 1
        ghostFromLast = n
        For L = 1 To n
            Set laneOcc(L) = New Collection
        Next L
        nLane = n
    End If

    For bi = 1 To nBand
        b = bandSeq(bi)
        cnt = 0
        For k = 1 To nItem
            If itKind(k) = KIND_ACT Or itKind(k) = KIND_SUM Then
                If itBand(k) = b Then
                    cnt = cnt + 1
                    members(cnt) = k
                    m = cnt
                    Do While m > 1
                        If ComesBefore(members(m), members(m - 1)) Then
                            j = members(m - 1): members(m - 1) = members(m): members(m) = j
                            m = m - 1
                        Else
                            Exit Do
                        End If
                    Loop
                End If
            End If
        Next k

        bandFirstLane(b) = nLane + 1
        For m = 1 To cnt
            k = members(m)
            placed = False
            For L = bandFirstLane(b) To nLane
                If laneLastF(L) < itS(k) Then
                    placed = True
                ElseIf laneLastF(L) = itS(k) Then
                    If SharesNode(laneLastItem(L), k) Then placed = True
                End If
                If placed Then
                    itLane(k) = L
                    laneLastF(L) = itF(k)
                    laneLastItem(L) = k
                    laneOcc(L).Add Array(itS(k), itF(k))
                    Exit For
                End If
            Next L
            If Not placed Then
                nLane = nLane + 1
                laneLastF(nLane) = itF(k)
                laneLastItem(nLane) = k
                Set laneOcc(nLane) = New Collection
                laneOcc(nLane).Add Array(itS(k), itF(k))
                itLane(k) = nLane
            End If
        Next m
        bandLastLane(b) = nLane
    Next bi

    ' 他の期へ（いちばん下）
    ghostToFirst = 0
    ghostToLast = 0
    n = StackGhosts(KIND_TO, nLane + 1)
    If n > 0 Then
        ghostToFirst = nLane + 1
        ghostToLast = nLane + n
        For L = ghostToFirst To ghostToLast
            Set laneOcc(L) = New Collection
        Next L
        nLane = ghostToLast
    End If
End Sub

' 他の期の作業（kind）を、ラベルが重ならないように段に積む。firstLane から順に使い、使った段数を返す
Private Function StackGhosts(ByVal kind As Long, ByVal firstLane As Long) As Long
    Dim seq() As Long, laneRight() As Double, m As Long, k As Long, L As Long, nUsed As Long
    Dim lft As Double, placed As Boolean

    ReDim laneRight(1 To nItem)
    seq = OrderByStart()
    For m = 1 To nItem
        k = seq(m)
        If itKind(k) = kind Then
            lft = (itS(k) - viewStart) * dayW - NODE_D / 2          ' ○の左端
            placed = False
            For L = 1 To nUsed
                If laneRight(L) + LABEL_GAP <= lft Then
                    placed = True
                    Exit For
                End If
            Next L
            If Not placed Then
                nUsed = nUsed + 1
                L = nUsed
            End If
            itLane(k) = firstLane + L - 1
            laneRight(L) = lft + NODE_D + 2 + GhostLabelWidth(k)   ' ラベルは○の右から
        End If
    Next m
    StackGhosts = nUsed
End Function

' 他の期の作業のラベル（上下2行のうち長いほう）の幅
Private Function GhostLabelWidth(ByVal k As Long) As Double
    Dim w1 As Double, w2 As Double
    w1 = TextWidth(NameLabel(k), NAME_FONT) + 6
    w2 = TextWidth(SubLabel(k), SUB_FONT) + 6
    If w1 > w2 Then GhostLabelWidth = w1 Else GhostLabelWidth = w2
End Function

' 段に置く順：開始が早い → クリティカルパス優先 → 終了が早い
Private Function ComesBefore(ByVal a As Long, ByVal b As Long) As Boolean
    If itS(a) <> itS(b) Then
        ComesBefore = (itS(a) < itS(b))
    ElseIf ItemCrit(a) <> ItemCrit(b) Then
        ComesBefore = ItemCrit(a)
    Else
        ComesBefore = (itF(a) < itF(b))
    End If
End Function

' lastK の終点と k の始点が、つながりで直接結ばれているか（同じ段で結合点を共有してよいか）
Private Function SharesNode(ByVal lastK As Long, ByVal k As Long) As Boolean
    Dim j As Long
    For j = 1 To nLink
        If lkFrom(j) = lastK And lkTo(j) = k Then
            If lkA(j) = itF(lastK) And lkB(j) = itS(k) Then
                SharesNode = True
                Exit Function
            End If
        End If
    Next j
End Function

Private Sub RegisterNodes()
    Dim k As Long, j As Long, n As Long, sz As Long, ghost As Boolean, crit As Boolean, i As Long, d As Long

    Set nodeDict = CreateObject("Scripting.Dictionary")
    nNode = 0
    sz = 2 * nItem + 2 * nLink + 2 * nAct + 2
    ReDim nodeLane(1 To sz): ReDim nodeDay(1 To sz): ReDim nodeCrit(1 To sz): ReDim nodeGhost(1 To sz)
    For k = 1 To nItem
        ghost = (itKind(k) = KIND_FROM Or itKind(k) = KIND_TO)
        If itKind(k) = KIND_SUM Then
            n = NodeIndex(itLane(k), itS(k), SumEndCrit(k, True), False)
            n = NodeIndex(itLane(k), itF(k), SumEndCrit(k, False), False)
            ' まとめの中で作業のない期間（待ち）の両端にも結合点を置く
            For i = 1 To nAct
                If actGroupIdx(i) = itGroup(k) Then
                    d = actF(i)
                    If d > itS(k) And d < itF(k) Then
                        If Not SegBusy(k, d, d + 1) Then n = NodeIndex(itLane(k), d, actCrit(i), False)
                    End If
                    d = actS(i)
                    If d > itS(k) And d < itF(k) Then
                        If Not SegBusy(k, d - 1, d) Then n = NodeIndex(itLane(k), d, actCrit(i), False)
                    End If
                End If
            Next i
        Else
            crit = ItemCrit(k) And Not ghost
            n = NodeIndex(itLane(k), itS(k), crit, ghost)
            n = NodeIndex(itLane(k), itF(k), crit, ghost)
            If ghost Then laneOcc(itLane(k)).Add Array(itS(k), itF(k))
        End If
    Next k
    ' つながりの付く点（まとめの矢線の途中など）にも結合点を置く
    For j = 1 To nLink
        n = NodeIndex(itLane(lkFrom(j)), lkA(j), lkCrit(j), False)
        n = NodeIndex(itLane(lkTo(j)), lkB(j), lkCrit(j), False)
    Next j
End Sub

' まとめの始点（終点）の結合点を赤にするか：クリティカルな作業がそこで始まる（終わる）とき
Private Function SumEndCrit(ByVal k As Long, ByVal atStart As Boolean) As Boolean
    Dim i As Long, g As Long
    g = itGroup(k)
    For i = 1 To nAct
        If actGroupIdx(i) = g And actCrit(i) Then
            If atStart Then
                If actS(i) = itS(k) Then SumEndCrit = True
            Else
                If actF(i) = itF(k) Then SumEndCrit = True
            End If
            If SumEndCrit Then Exit Function
        End If
    Next i
End Function

Private Function NodeIndex(ByVal L As Long, ByVal d As Long, ByVal crit As Boolean, ByVal ghost As Boolean) As Long
    Dim key As String, k As Long
    key = L & "|" & d
    If nodeDict.Exists(key) Then
        k = nodeDict(key)
        If crit Then nodeCrit(k) = True
    Else
        nNode = nNode + 1
        k = nNode
        nodeLane(k) = L
        nodeDay(k) = d
        nodeCrit(k) = crit
        nodeGhost(k) = ghost
        nodeDict.Add key, k
    End If
    NodeIndex = k
End Function

Private Function NodeExists(ByVal L As Long, ByVal d As Long) As Boolean
    NodeExists = nodeDict.Exists(L & "|" & d)
End Function


'==============================================================================
'  項目の情報
'==============================================================================
Private Function ItemCrit(ByVal k As Long) As Boolean
    Select Case itKind(k)
        Case KIND_SUM: ItemCrit = gCrit(itGroup(k))
        Case KIND_ACT: ItemCrit = actCrit(itAct(k))
        Case Else: ItemCrit = False
    End Select
End Function

Private Function ItemKouji(ByVal k As Long) As String
    If itKind(k) = KIND_SUM Then ItemKouji = gKouji(itGroup(k)) Else ItemKouji = actKouji(itAct(k))
End Function

Private Function ItemZone(ByVal k As Long) As Long
    If itKind(k) = KIND_SUM Then ItemZone = gZone(itGroup(k)) Else ItemZone = actZoneIdx(itAct(k))
End Function

Private Function ItemPurp(ByVal k As Long) As Long
    If itKind(k) = KIND_SUM Then ItemPurp = gPurp(itGroup(k)) Else ItemPurp = actPurpIdx(itAct(k))
End Function

Private Function ItemColor(ByVal k As Long) As Long
    Select Case itKind(k)
        Case KIND_SUM
            ItemColor = CompanyColor(gCoName(itGroup(k)))   ' 複数の会社のまとめは灰色
        Case KIND_ACT
            ItemColor = CompanyColor(actCo(itAct(k)))
        Case Else
            ItemColor = RGB(118, 118, 118)
    End Select
End Function

' まとめの矢線の区間（e1～e2）がクリティカルか：その区間に重なるクリティカルな作業があるとき
Private Function SegCrit(ByVal k As Long, ByVal e1 As Long, ByVal e2 As Long) As Boolean
    Dim i As Long, g As Long
    If itKind(k) <> KIND_SUM Then
        SegCrit = ItemCrit(k)
        Exit Function
    End If
    g = itGroup(k)
    For i = 1 To nAct
        If actGroupIdx(i) = g And actCrit(i) Then
            If actS(i) < e2 And actF(i) > e1 Then
                SegCrit = True
                Exit Function
            End If
        End If
    Next i
End Function

' まとめの区間（e1～e2）に作業があるか。作業のない区間（待ち）は点線で描く。まとめ以外は常に True
Private Function SegBusy(ByVal k As Long, ByVal e1 As Long, ByVal e2 As Long) As Boolean
    Dim i As Long, g As Long
    If itKind(k) <> KIND_SUM Then
        SegBusy = True
        Exit Function
    End If
    g = itGroup(k)
    For i = 1 To nAct
        If actGroupIdx(i) = g Then
            If actS(i) < e2 And actF(i) > e1 Then
                SegBusy = True
                Exit Function
            End If
        End If
    Next i
End Function

' ラベルを置く範囲（e1～e2）：まとめは作業のある期間のうち最も長いところ。ほかは項目の始点～終点
Private Sub LabelSpan(ByVal k As Long, ByRef e1 As Long, ByRef e2 As Long)
    Dim d As Long, runS As Long, best As Long, inRun As Boolean
    e1 = itS(k)
    e2 = itF(k)
    If itKind(k) <> KIND_SUM Then Exit Sub
    best = -1
    For d = itS(k) To itF(k)
        If d < itF(k) And SegBusy(k, d, d + 1) Then
            If Not inRun Then
                runS = d
                inRun = True
            End If
        ElseIf inRun Then
            If d - runS > best Then
                best = d - runS
                e1 = runS
                e2 = d
            End If
            inRun = False
        End If
    Next d
End Sub

' 矢線の上の文字（1行）：作業ID 作業名（表示名があればそれ。まとめは名前、他の期の作業は期・工区も）
Private Function NameLabel(ByVal k As Long) As String
    Dim id As String, body As String
    LabelParts k, id, body
    NameLabel = OneLine(id, body)
    If itKind(k) = KIND_FROM Or itKind(k) = KIND_TO Then NameLabel = NameLabel & "（" & GhostTag(itAct(k)) & "）"
End Function

' ラベルの作業IDと名前（名前は「表示名」があればそれ。まとめは名前だけ）
Private Sub LabelParts(ByVal k As Long, ByRef id As String, ByRef body As String)
    Dim a As Long
    If itKind(k) = KIND_SUM Then
        id = ""
        body = gName(itGroup(k))
    Else
        a = itAct(k)
        id = actId(a)
        If Len(actDisp(a)) > 0 Then body = actDisp(a) Else body = actName(a)
    End If
End Sub

' 作業IDと名前を1行に（名前の中の改行は空白に）
Private Function OneLine(ByVal id As String, ByVal body As String) As String
    body = Replace(Replace(Replace(body, vbCrLf, " "), vbLf, " "), vbCr, " ")
    If Len(id) = 0 Then
        OneLine = body
    ElseIf Len(body) = 0 Then
        OneLine = id
    Else
        OneLine = id & " " & body
    End If
End Function

' 矢線の下の文字：会社・日数・進捗（他の期の作業は終了日／開始日）
Private Function SubLabel(ByVal k As Long) As String
    Dim a As Long, g As Long, t As String
    a = itAct(k)
    Select Case itKind(k)
        Case KIND_FROM
            t = Format(CDate(LastDay(a)), "m/d") & "終了"
        Case KIND_TO
            t = Format(CDate(actS(a)), "m/d") & "開始"
        Case KIND_SUM
            g = itGroup(k)
            t = gCoLabel(g)
            If gF(g) > gS(g) Then
                If Len(t) > 0 Then t = t & "・"
                t = t & (gF(g) - gS(g)) & "日"
            End If
            If gProg(g) > 0 And Not showProg Then
                If Len(t) > 0 Then t = t & "・"
                t = t & Format(gProg(g), "0%")
            End If
        Case Else
            t = actCo(a)
            If actF(a) > actS(a) Then
                If Len(t) > 0 Then t = t & "・"
                t = t & (actF(a) - actS(a)) & "日"
            End If
            If actProg(a) > 0 And Not showProg Then
                If Len(t) > 0 Then t = t & "・"
                t = t & Format(actProg(a), "0%")
            End If
    End Select
    If showProg Then t = ProgLabel(k, t)
    SubLabel = t
End Function

Private Function PhaseLabel(ByVal a As Long) As String
    If actPhaseIdx(a) > 0 Then PhaseLabel = phaseName(actPhaseIdx(a)) Else PhaseLabel = "期未入力"
End Function

' 他の期の作業に添える期（工区があれば工区も）
Private Function GhostTag(ByVal a As Long) As String
    GhostTag = PhaseLabel(a)
    If useZones Then GhostTag = GhostTag & "・" & zoneName(actZoneIdx(a))
End Function

' 開始日の順に並べた項目の添字
Private Function OrderByStart() As Long()
    Dim seq() As Long, i As Long, k As Long, j As Long
    ReDim seq(1 To nItem)
    For i = 1 To nItem
        seq(i) = i
        k = i
        Do While k > 1
            If StartsBefore(seq(k), seq(k - 1)) Then
                j = seq(k - 1): seq(k - 1) = seq(k): seq(k) = j
                k = k - 1
            Else
                Exit Do
            End If
        Loop
    Next i
    OrderByStart = seq
End Function

Private Function StartsBefore(ByVal a As Long, ByVal b As Long) As Boolean
    If itS(a) <> itS(b) Then
        StartsBefore = (itS(a) < itS(b))
    Else
        StartsBefore = (itF(a) < itF(b))
    End If
End Function


'==============================================================================
'  作図
'==============================================================================
Private Sub CreateChartSheet(ByVal sheetName As String)
    Set wsC = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsC.Name = sheetName
    wsC.Names.Add Name:=MARK_NAME, RefersTo:="=TRUE", Visible:=False
    shapeSeq = 0
    xCacheOK = False
    yCacheOK = False

    ' 行の配置（全体工程で期があるときは、目盛りの上に期の帯の行を入れる）
    nPhaseRows = 0
    If isOverall And nPhase > 0 Then nPhaseRows = PhaseLevels()
    rowMonth = 4 + nPhaseRows
    rowDay = rowMonth + 1
    rowWeekday = rowMonth + 2
    firstLaneRow = rowMonth + 3

    nCols = viewEnd - viewStart + 1                 ' 日付の列は1日1列（週単位も）
    If weekScale Then
        nCols = WholeWeeks(MaxL(nCols, MIN_WEEK_COLS * 7))
    ElseIf nCols < MIN_DAY_COLS Then
        nCols = MIN_DAY_COLS
    End If

    With wsC.Cells
        .Font.Name = "游ゴシック"
        .Font.Size = 9
        .VerticalAlignment = xlCenter
    End With
    ' 縦軸の列：左から目的・工区・工事（目的・工区が無ければ、その列は作らない）
    labelCols = 1
    If usePurps Then labelCols = labelCols + 1
    If useZones Then labelCols = labelCols + 1
    zoneCol = 0
    If useZones Then zoneCol = labelCols - 1
    If usePurps Then wsC.Columns(1).ColumnWidth = PURP_COL_WIDTH
    If useZones Then wsC.Columns(zoneCol).ColumnWidth = ZONE_COL_WIDTH
    wsC.Columns(labelCols).ColumnWidth = LABEL_COL_WIDTH
    dateCol = labelCols + 1
    SetDateColumns
    dayW = wsC.Cells(1, dateCol).Width                  ' 1日分の幅（pt）
End Sub

Private Sub SetDateColumns()
    wsC.Range(wsC.Columns(dateCol), wsC.Columns(dateCol + nCols - 1)).ColumnWidth = _
        IIf(weekScale, WEEK_DAY_COL_WIDTH, DAY_COL_WIDTH)
    xCacheOK = False
End Sub

' 週単位のときは、日付の列の数を週の切れ目（7日ごと）にそろえる
Private Function WholeWeeks(ByVal n As Long) As Long
    If weekScale Then WholeWeeks = ((n + 6) \ 7) * 7 Else WholeWeeks = n
End Function

' ラベルが右端からはみ出さないよう、必要なら日付の列を足す
Private Sub FitColumns()
    Dim needCols As Long
    needCols = WholeWeeks(Int(LayoutLabels()) + 2)
    If needCols > nCols Then
        nCols = needCols
        SetDateColumns
    End If
End Sub

Private Sub SetRowHeights()
    Dim r As Long, L As Long, h As Double
    wsC.Rows(1).RowHeight = 24
    wsC.Rows(2).RowHeight = 16
    wsC.Rows(3).RowHeight = 8
    For r = 4 To rowMonth - 1
        wsC.Rows(r).RowHeight = PHASE_ROW_HEIGHT
    Next r
    wsC.Rows(rowMonth).RowHeight = 16
    wsC.Rows(rowDay).RowHeight = 14
    wsC.Rows(rowWeekday).RowHeight = 14
    ' 段の高さは、折り返したり積んだりしたラベルに合わせて増やす
    For L = 1 To nLane
        h = LANE_HEIGHT + laneExtra(L)
        If h > 409 Then h = 409
        wsC.Rows(firstLaneRow + L - 1).RowHeight = h
        yCacheOK = False
    Next L
End Sub

' 期の帯を重ならないように何段に分けるか（全体工程の目盛りの上に描く）
Private Function PhaseLevels() As Long
    Dim p As Long, q As Long, lv As Long, mx As Long, ok As Boolean
    ReDim phaseLevel(1 To nPhase)
    For p = 1 To nPhase
        lv = 1
        Do
            ok = True
            For q = 1 To p - 1
                If phaseLevel(q) = lv Then
                    If phaseMinS(q) < phaseMaxF(p) And phaseMinS(p) < phaseMaxF(q) Then ok = False
                End If
            Next q
            If ok Then Exit Do
            lv = lv + 1
        Loop
        phaseLevel(p) = lv
        If lv > mx Then mx = lv
    Next p
    PhaseLevels = mx
End Function

' 表題・日付目盛り・網掛け（土日・休工日）・工事名
Private Sub DrawGrid()
    Dim c As Long, d As Long, col As Long, lastCol As Long, lastLaneRow As Long
    Dim b As Long, r1 As Long, r2 As Long, prevMonth As Long, wd As Long, hasHol As Boolean
    Dim bi As Long, b1 As Long, runEnd As Boolean
    Dim sv As Variant, txt As String, created As Variant
    Dim vMon() As Variant, vDay() As Variant, vWk() As Variant, clr As Long, runClr As Long, runCol As Long

    lastCol = dateCol + nCols - 1
    lastLaneRow = firstLaneRow + nLane - 1

    ' 表題
    txt = CellText(GetSettingValue("工事名"))
    If Len(txt) = 0 Then txt = "工事"
    With wsC.Cells(1, 1)
        .Value = txt & "　ネットワーク工程表" & sheetCaption
        .Font.Size = 14
        .Font.Bold = True
    End With

    ' 日付目盛りと網掛け
    If usePurps Then wsC.Cells(rowDay, 1).Value = "目的"
    If useZones Then wsC.Cells(rowDay, zoneCol).Value = "工区"
    wsC.Cells(rowDay, labelCols).Value = "工事"
    ' 日の境目の罫線：日単位は薄い灰色の細線、週単位は点線（週の境目は下で実線にする）。結合点（○）はこの線の上に乗る
    ' （列ごとではなく、まとめて引く）
    With wsC.Range(wsC.Cells(rowDay, dateCol), wsC.Cells(lastLaneRow, lastCol)).Borders(xlInsideVertical)
        .LineStyle = IIf(weekScale, xlDot, xlContinuous)
        .Weight = xlThin
        .Color = IIf(weekScale, RGB(191, 191, 191), RGB(217, 217, 217))
    End With
    ' 目盛りの文字は配列にためて1回で書き、網掛けは同じ色が続く列をまとめて塗る
    ReDim vMon(1 To 1, 1 To nCols): ReDim vDay(1 To 1, 1 To nCols): ReDim vWk(1 To 1, 1 To nCols)
    For c = 0 To nCols - 1
        d = viewStart + c
        col = dateCol + c
        wd = Weekday(CDate(d), vbSunday)
        If c = 0 Or Month(CDate(d)) <> prevMonth Then
            vMon(1, c + 1) = "'" & Format(CDate(d), "yyyy/m")
            If c > 0 Then       ' 月の境目の濃い線（最初の列の左は、あとでまとめて引く）
                With wsC.Range(wsC.Cells(rowMonth, col), wsC.Cells(lastLaneRow, col)).Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                    .Color = RGB(150, 150, 150)
                End With
            End If
            prevMonth = Month(CDate(d))
        ElseIf weekScale And wd = vbMonday Then        ' 週の境目（月曜の左）の実線
            With wsC.Range(wsC.Cells(rowDay, col), wsC.Cells(lastLaneRow, col)).Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .Color = RGB(166, 166, 166)
            End With
        End If
        clr = 0
        If weekScale Then
            If wd = vbMonday Then vDay(1, c + 1) = "'" & Format(CDate(d), "m/d")   ' 週の初め（月曜）の日付
        Else
            vDay(1, c + 1) = Day(CDate(d))
            vWk(1, c + 1) = Mid$("日月火水木金土", wd, 1)
            If IsHoliday(d) Then
                clr = RGB(252, 228, 214)
                hasHol = True
            ElseIf wd = vbSunday Or wd = vbSaturday Then
                clr = RGB(242, 242, 242)
            End If
            If wd = vbSunday Then wsC.Cells(rowWeekday, col).Font.Color = RGB(192, 0, 0)
        End If
        If clr <> runClr Then
            If runClr <> 0 Then wsC.Range(wsC.Cells(rowDay, runCol), wsC.Cells(lastLaneRow, col - 1)).Interior.Color = runClr
            runCol = col
            runClr = clr
        End If
    Next c
    If runClr <> 0 Then wsC.Range(wsC.Cells(rowDay, runCol), wsC.Cells(lastLaneRow, lastCol)).Interior.Color = runClr
    wsC.Range(wsC.Cells(rowMonth, dateCol), wsC.Cells(rowMonth, lastCol)).Value = vMon
    wsC.Range(wsC.Cells(rowDay, dateCol), wsC.Cells(rowDay, lastCol)).Value = vDay
    If Not weekScale Then wsC.Range(wsC.Cells(rowWeekday, dateCol), wsC.Cells(rowWeekday, lastCol)).Value = vWk
    If showProg Then
        If progT >= viewStart And progT < viewStart + nCols Then MarkStatusDay dateCol + progT - viewStart
    End If
    With wsC.Range(wsC.Cells(rowDay, dateCol), wsC.Cells(rowWeekday, lastCol))
        .HorizontalAlignment = IIf(weekScale, xlLeft, xlCenter)     ' 週単位の日付は月曜の列から右へはみ出して書く
        .Font.Size = 8
    End With
    wsC.Range(wsC.Cells(rowMonth, dateCol), wsC.Cells(rowMonth, lastCol)).Font.Size = 8
    With wsC.Range(wsC.Cells(rowWeekday, 1), wsC.Cells(rowWeekday, lastCol)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = RGB(127, 127, 127)
    End With
    With wsC.Range(wsC.Cells(rowMonth, dateCol), wsC.Cells(lastLaneRow, dateCol)).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = RGB(127, 127, 127)
    End With

    ' 2行目：版数・作成日・目盛り
    sv = GetSettingValue("版数")
    created = GetSettingValue("作成日")
    If Not IsDate(created) Then created = Date
    txt = ""
    If Len(CellText(sv)) > 0 Then txt = "第" & CellText(sv) & "版　"
    txt = txt & "作成日：" & Format(created, "yyyy/mm/dd") & "　目盛り：" & IIf(weekScale, "週", "日") & "単位"
    If Not weekScale Then
        If hasHol Then
            txt = txt & "（網掛け：灰色＝土日、赤＝休工日）"
        Else
            txt = txt & "（網掛けは土日）"
        End If
    Else
        txt = txt & "（月曜始まり。点線は日の区切り）"
    End If
    wsC.Cells(2, 1).Value = txt
    If showProg Then AddProgressSummary txt

    ' 工事名（縦軸のいちばん右の列）
    For b = 1 To nBand
        r1 = firstLaneRow + bandFirstLane(b) - 1
        r2 = firstLaneRow + bandLastLane(b) - 1
        With wsC.Range(wsC.Cells(r1, labelCols), wsC.Cells(r2, labelCols))
            If r2 > r1 Then .Merge
            .WrapText = True
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .IndentLevel = 1
        End With
        wsC.Cells(r1, labelCols).Value = "'" & bandName(b)      ' 「1-2」などが日付にならないよう文字列で
        With wsC.Range(wsC.Cells(r2, labelCols), wsC.Cells(r2, lastCol)).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = RGB(191, 191, 191)
        End With
    Next b
    ' 工区名と目的：上下に続く帯の段をまとめて塗る。境目の線は工区が太線、目的はもっと太い線
    ' （工区は目的の中で分ける。同じ工区でも目的が違えば別の欄。線が重なる目的の境目は、あとから引いて上書きする）
    If useZones Then
        b1 = 1
        For bi = 1 To nBand
            b = bandSeq(bi)
            If bi = nBand Then
                runEnd = True
            Else
                runEnd = (bandZone(bandSeq(bi + 1)) <> bandZone(b) Or bandPurp(bandSeq(bi + 1)) <> bandPurp(b))
            End If
            If runEnd Then
                LabelRun zoneCol, bandFirstLane(bandSeq(b1)), bandLastLane(b), zoneName(bandZone(b)), _
                         RGB(237, 242, 249), RGB(0, 0, 0), xlMedium, RGB(89, 89, 89), lastCol
                b1 = bi + 1
            End If
        Next bi
    End If
    If usePurps Then
        b1 = 1
        For bi = 1 To nBand
            b = bandSeq(bi)
            If bi = nBand Then
                runEnd = True
            Else
                runEnd = (bandPurp(bandSeq(bi + 1)) <> bandPurp(b))
            End If
            If runEnd Then
                LabelRun 1, bandFirstLane(bandSeq(b1)), bandLastLane(b), purpName(bandPurp(b)), _
                         RGB(217, 225, 242), RGB(31, 56, 100), xlThick, RGB(64, 64, 64), lastCol
                b1 = bi + 1
            End If
        Next bi
    End If
    If ghostFromFirst > 0 Then LabelGhostLanes ghostFromFirst, ghostFromLast, "他の期から", lastCol
    If ghostToFirst > 0 Then LabelGhostLanes ghostToFirst, ghostToLast, "他の期へ", lastCol
End Sub

' 縦軸の工区・目的の欄：col 列の段 L1～L2 を結合して塗り、名前を書き、欄の下に境目の線を引く（col 列から右端まで）
Private Sub LabelRun(ByVal col As Long, ByVal L1 As Long, ByVal L2 As Long, ByVal txt As String, _
                     ByVal fillClr As Long, ByVal txtClr As Long, ByVal lineWt As Long, ByVal lineClr As Long, _
                     ByVal lastCol As Long)
    Dim r1 As Long, r2 As Long
    r1 = firstLaneRow + L1 - 1
    r2 = firstLaneRow + L2 - 1
    With wsC.Range(wsC.Cells(r1, col), wsC.Cells(r2, col))
        If r2 > r1 Then .Merge
        .WrapText = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Font.Color = txtClr
        .Interior.Color = fillClr
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Color = RGB(191, 191, 191)
        End With
    End With
    wsC.Cells(r1, col).Value = "'" & txt          ' 「1-2」などが日付にならないよう文字列で
    With wsC.Range(wsC.Cells(r2, col), wsC.Cells(r2, lastCol)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = lineWt
        .Color = lineClr
    End With
End Sub

Private Sub LabelGhostLanes(ByVal L1 As Long, ByVal L2 As Long, ByVal txt As String, ByVal lastCol As Long)
    Dim r1 As Long, r2 As Long
    r1 = firstLaneRow + L1 - 1
    r2 = firstLaneRow + L2 - 1
    With wsC.Range(wsC.Cells(r1, 1), wsC.Cells(r2, labelCols))
        If r2 > r1 Or labelCols > 1 Then .Merge
        .Font.Color = RGB(118, 118, 118)
        .Font.Italic = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
    End With
    wsC.Cells(r1, 1).Value = txt
    With wsC.Range(wsC.Cells(r2, 1), wsC.Cells(r2, lastCol)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = RGB(191, 191, 191)
    End With
End Sub

' 全体工程：目盛りの上に期の帯（期の名前と期間）
Private Sub DrawPhaseBars()
    Dim p As Long, s As Shape, x1 As Double, x2 As Double, t As Double, w As Double

    wsC.Cells(4, 1).Value = "期"
    wsC.Cells(4, 1).IndentLevel = 1
    For p = 1 To nPhase
        x1 = XOf(phaseMinS(p))
        x2 = XOf(phaseMaxF(p))
        w = x2 - x1
        If w < 6 Then w = 6
        t = wsC.Cells(3 + phaseLevel(p), 1).Top + 2
        Set s = wsC.Shapes.AddShape(msoShapeRectangle, x1, t, w, PHASE_ROW_HEIGHT - 4)
        s.Fill.Visible = msoTrue
        s.Fill.Solid
        s.Fill.ForeColor.RGB = RGB(217, 225, 242)
        s.Line.Visible = msoFalse
        s.Shadow.Visible = msoFalse
        With s.TextFrame2
            .MarginLeft = 3
            .MarginRight = 0
            .MarginTop = 0
            .MarginBottom = 0
            .WordWrap = msoFalse
            .VerticalAnchor = msoAnchorMiddle
            With .TextRange
                .Text = phaseName(p) & "（" & Format(CDate(phaseMinS(p)), "m/d") & "～" & _
                        Format(CDate(phaseLastDay(p)), "m/d") & "）"
                .Font.Size = 8
                .Font.Bold = msoTrue
                .Font.Fill.ForeColor.RGB = RGB(31, 56, 100)
                .ParagraphFormat.Alignment = msoAlignLeft
            End With
        End With
        NameShape s, "P"
    Next p
End Sub

' 項目の矢線とラベル（色は会社ごと。クリティカルパスは赤い帯と太線・太字。他の期の作業はラベルだけ）
Private Sub DrawItems()
    Dim seq() As Long, s As Shape
    Dim m As Long, k As Long, L As Long, d As Long, prevE As Long
    Dim y As Double, r As Double, x0 As Double, info As String
    Dim clr As Long, cp As Boolean, sc As Boolean, isItem As Boolean, al As Long

    r = NODE_D / 2
    x0 = XOf(viewStart)
    seq = OrderByStart()
    For m = 1 To nItem
        Progress "矢線と文字 " & m & "/" & nItem
        k = seq(m)
        L = itLane(k)
        y = YOf(L)
        isItem = (itKind(k) = KIND_ACT Or itKind(k) = KIND_SUM)
        clr = ItemColor(k)
        cp = isItem And ItemCrit(k)
        info = ClickInfo(k)

        If isItem Then
            ' 矢線。まとめは途中の結合点で区切り、区間ごとにクリティカルか判定する。
            ' まとめの中で作業のない区間（待ち）は点線
            prevE = itS(k)
            For d = itS(k) + 1 To itF(k)
                If d = itF(k) Or (itKind(k) = KIND_SUM And NodeExists(L, d)) Then
                    If SegBusy(k, prevE, d) Then
                        sc = SegCrit(k, prevE, d)
                        If sc Then
                            Set lastShape = Nothing
                            DrawBand XOf(prevE), y, XOf(d), y
                            MarkClick lastShape, info
                        End If
                        If XOf(d) - XOf(prevE) > NODE_D + 1 Then
                            Set lastShape = Nothing
                            DrawLine XOf(prevE) + r, y, XOf(d) - r, y, clr, IIf(sc, 2.25, 1.25), msoLineSolid, True
                            MarkClick lastShape, info
                        End If
                    ElseIf XOf(d) - XOf(prevE) > NODE_D + 1 Then
                        Set lastShape = Nothing
                        DrawLine XOf(prevE) + r, y, XOf(d) - r, y, RGB(89, 89, 89), 1, msoLineRoundDot, False
                        MarkClick lastShape, info
                    End If
                    prevE = d
                End If
            Next d
            ' 細い矢線でもクリックしやすいよう、矢線の上に透明な当たり判定を重ねる
            If XOf(itF(k)) - XOf(itS(k)) > NODE_D + 1 Then
                AddHitArea XOf(itS(k)) + r, y, XOf(itF(k)) - r, info
            End If
        End If

        ' ラベル（位置・折り返し・積み上げは LayoutLabels で決めたもの）
        If isItem Then al = 0 Else al = 1          ' 作業とまとめは中央揃え、他の期の作業は○の右から
        If lbMid(k) > 0 Then
            ' マイルストーン：矢線の高さで○の横に1行
            Set s = AddLabelBox(x0 + lbNameL(k), y, lbNameW(k), lbNameTxt(k), 1, NAME_FONT, clr, lbMid(k), 2, cp)
            MarkClick s, info
        Else
            ' 矢線の上：作業ID・作業名（まとめは名前）
            If lbNameN(k) > 0 Then
                Set s = AddLabelBox(x0 + lbNameL(k), y - r - 1 - lbNameOff(k), lbNameW(k), lbNameTxt(k), _
                                    lbNameN(k), NAME_FONT, clr, al, 0, cp)
                MarkClick s, info
            End If
            ' 矢線の下：会社・日数・進捗（他の期の作業は日付）
            If lbSubN(k) > 0 Then
                Set s = AddLabelBox(x0 + lbSubL(k), y + r + 0.5 + lbSubOff(k), lbSubW(k), lbSubTxt(k), _
                                    lbSubN(k), SUB_FONT, clr, al, 1)
                If showProg And isItem Then AccentTail s, lbSubTxt(k), DelayText(ItemDelay(k))
                MarkClick s, info
            End If
        End If
    Next m
End Sub

' ラベルの置き方を決める（位置は日付目盛りの左端からの pt。DrawItems はこの結果で描く）
'   ・作業とまとめのラベルは、自分の矢線の中央に置く（横へはずらさない）
'   ・使える幅は、前のラベルの右端から、次のラベル（前後と間を半分ずつ分けたときの幅）の左端まで。
'     入らなければ折り返す（名前は作業IDを1行目に、会社・日数は「・」の位置で）
'   ・それでも重なるときは、上（会社・日数は下）に積む（積むなら、折り返さない形と比べて低い方）。
'     段の高さはラベルに合わせて増やす
'   ・マイルストーン（日数0）は、前後の矢線が付いていない側に、矢線の高さで1行に書く
'     （その範囲には余裕の点線を引かないよう、段の使用中の範囲に加える）
'   ・他の期の作業は○のすぐ右から1行
' 戻り値：ラベルの右端が、表示開始日から何日分の位置まで届くか
Private Function LayoutLabels() As Double
    Dim seq() As Long, m As Long, k As Long, L As Long, p As Long, q As Long, pass As Long
    Dim ctr() As Double, nextK() As Long, prevK() As Long, lastK() As Long
    Dim nxt() As Long, prv() As Long, w1n() As Double, w1s() As Double
    Dim right0() As Double, needTop() As Double, needBtm() As Double
    Dim e1 As Long, e2 As Long, n As Long, span As Long
    Dim r As Double, mx As Double, w As Double, room As Double, roomL As Double, roomR As Double
    Dim x1 As Double, h As Double, off As Double, fs As Double, tE As Double, bE As Double
    Dim txt As String, id As String, body As String, co As String, isName As Boolean, has As Boolean
    Dim txt2 As String, n2 As Long, w2 As Double, x2 As Double, h2 As Double, off2 As Double

    r = NODE_D / 2
    ReDim lbNameTxt(1 To nItem): ReDim lbNameN(1 To nItem): ReDim lbNameL(1 To nItem)
    ReDim lbNameW(1 To nItem): ReDim lbNameOff(1 To nItem)
    ReDim lbSubTxt(1 To nItem): ReDim lbSubN(1 To nItem): ReDim lbSubL(1 To nItem)
    ReDim lbSubW(1 To nItem): ReDim lbSubOff(1 To nItem): ReDim lbMid(1 To nItem)
    ReDim laneArrowOff(1 To nLane): ReDim laneExtra(1 To nLane)
    ReDim plLane(1 To 2 * nItem + 1): ReDim plIsName(1 To 2 * nItem + 1): ReDim plX1(1 To 2 * nItem + 1)
    ReDim plX2(1 To 2 * nItem + 1): ReDim plOff(1 To 2 * nItem + 1): ReDim plH(1 To 2 * nItem + 1)
    nPl = 0
    ReDim ctr(1 To nItem): ReDim nextK(1 To nItem): ReDim prevK(1 To nItem): ReDim lastK(1 To nLane)
    ReDim nxt(1 To nItem): ReDim prv(1 To nItem): ReDim w1n(1 To nItem): ReDim w1s(1 To nItem)
    ReDim right0(1 To nLane): ReDim needTop(1 To nLane): ReDim needBtm(1 To nLane)

    ' 中央の位置（まとめは作業のある期間のうち最も長いところの中央）と、同じ段の前後の項目
    seq = OrderByStart()
    For m = 1 To nItem
        k = seq(m)
        LabelSpan k, e1, e2
        ctr(k) = ((e1 - viewStart) + (e2 - viewStart)) / 2 * dayW
        L = itLane(k)
        If lastK(L) > 0 Then
            nextK(lastK(L)) = k
            prevK(k) = lastK(L)
        End If
        lastK(L) = k
    Next m

    ' マイルストーン：前後の矢線が○に付いていない側（右が先）に、入れば矢線の高さで1行に書く
    For m = 1 To nItem
        k = seq(m)
        If itKind(k) = KIND_ACT And itS(k) = itF(k) Then
            LabelParts k, id, body
            co = SubLabel(k)
            txt = OneLine(id, body)
            If Len(co) > 0 Then txt = txt & "（" & co & "）"
            w = TextWidth(txt, NAME_FONT) + 6
            span = -Int(-((r + 2 + w) / dayW))
            q = nextK(k)
            p = prevK(k)
            room = -1
            If q = 0 Then
                room = 1E+30
            ElseIf itS(q) > itS(k) Then
                room = (itS(q) - itS(k)) * dayW - 2 * (r + 2)
            End If
            If w <= room Then
                lbMid(k) = 1
                lbNameL(k) = ctr(k) + r + 2
                laneOcc(itLane(k)).Add Array(itS(k), itS(k) + span)
                If lbNameL(k) + w > mx Then mx = lbNameL(k) + w
            Else
                room = -1
                If p = 0 Then
                    room = ctr(k) - r - 2
                ElseIf itF(p) < itS(k) Then
                    room = (itS(k) - itF(p)) * dayW - 2 * (r + 2)
                End If
                If w <= room Then
                    lbMid(k) = 2
                    lbNameL(k) = ctr(k) - r - 2 - w
                    laneOcc(itLane(k)).Add Array(itS(k) - span, itS(k))
                End If
            End If
            If lbMid(k) > 0 Then
                lbNameTxt(k) = txt
                lbNameN(k) = 1
                lbNameW(k) = w
            End If
        End If
    Next m

    ' 矢線の上下のラベルどうしの前後（横に書いたマイルストーンは除く）
    ReDim lastK(1 To nLane)
    For m = 1 To nItem
        k = seq(m)
        If lbMid(k) = 0 Then
            L = itLane(k)
            If lastK(L) > 0 Then
                nxt(lastK(L)) = k
                prv(k) = lastK(L)
            End If
            lastK(L) = k
        End If
    Next m

    ' 1回目：前後の項目と間を半分ずつ分けたときの幅（次のラベルのために残す幅の目安）
    For m = 1 To nItem
        k = seq(m)
        If (itKind(k) = KIND_ACT Or itKind(k) = KIND_SUM) And lbMid(k) = 0 Then
            For pass = 1 To 2
                isName = (pass = 1)
                w = 0
                If Not LabelEmpty(k, isName) Then
                    p = prv(k)
                    q = nxt(k)
                    If p > 0 Then roomL = (ctr(k) - ctr(p)) / 2 - LABEL_GAP / 2 Else roomL = ctr(k)
                    If q > 0 Then roomR = (ctr(q) - ctr(k)) / 2 - LABEL_GAP / 2 Else roomR = 1E+30
                    If roomL < roomR Then room = 2 * roomL Else room = 2 * roomR
                    FitLabel k, isName, room, txt, n, w
                End If
                If isName Then w1n(k) = w Else w1s(k) = w
            Next pass
        End If
    Next m

    ' 2回目：左から順に置く（重なれば積む）
    For pass = 1 To 2
        isName = (pass = 1)
        If isName Then fs = NAME_FONT Else fs = SUB_FONT
        For L = 1 To nLane
            right0(L) = -100000#
        Next L
        For m = 1 To nItem
            k = seq(m)
            L = itLane(k)
            has = False
            If itKind(k) = KIND_FROM Or itKind(k) = KIND_TO Then
                If isName Then txt = NameLabel(k) Else txt = SubLabel(k)
                n = 1
                w = TextWidth(txt, fs) + 6
                x1 = ctr(k) + r + 2
                has = True
            ElseIf lbMid(k) = 0 Then
                If Not LabelEmpty(k, isName) Then
                    roomL = ctr(k) - right0(L) - LABEL_GAP
                    If roomL > ctr(k) Then roomL = ctr(k)
                    q = nxt(k)
                    If q > 0 Then
                        If isName Then roomR = ctr(q) - w1n(q) / 2 - ctr(k) - LABEL_GAP Else roomR = ctr(q) - w1s(q) / 2 - ctr(k) - LABEL_GAP
                    Else
                        roomR = 1E+30
                    End If
                    If roomL < roomR Then room = 2 * roomL Else room = 2 * roomR
                    FitLabel k, isName, room, txt, n, w
                    x1 = ctr(k) - w / 2
                    has = True
                End If
            End If
            If has Then
                h = BoxH(n, fs)
                off = StackOffset(L, isName, x1, x1 + w, h)
                If off > 0 And (itKind(k) = KIND_ACT Or itKind(k) = KIND_SUM) Then
                    ' 積むことになるなら、折り返さない形と比べて低く収まる方にする
                    FitLabel k, isName, 1E+30, txt2, n2, w2
                    x2 = ctr(k) - w2 / 2
                    h2 = BoxH(n2, fs)
                    off2 = StackOffset(L, isName, x2, x2 + w2, h2)
                    If off2 + h2 < off + h Then
                        txt = txt2
                        n = n2
                        w = w2
                        x1 = x2
                        h = h2
                        off = off2
                    End If
                End If
                nPl = nPl + 1
                plLane(nPl) = L
                plIsName(nPl) = isName
                plX1(nPl) = x1
                plX2(nPl) = x1 + w
                plOff(nPl) = off
                plH(nPl) = h
                If off = 0 And x1 + w > right0(L) Then right0(L) = x1 + w
                If isName Then
                    lbNameTxt(k) = txt
                    lbNameN(k) = n
                    lbNameL(k) = x1
                    lbNameW(k) = w
                    lbNameOff(k) = off
                    If off + h > needTop(L) Then needTop(L) = off + h
                Else
                    lbSubTxt(k) = txt
                    lbSubN(k) = n
                    lbSubL(k) = x1
                    lbSubW(k) = w
                    lbSubOff(k) = off
                    If off + h > needBtm(L) Then needBtm(L) = off + h
                End If
                If x1 + w > mx Then mx = x1 + w
            End If
        Next m
    Next pass

    ' 段ごとの、矢線の上下に足りない高さ
    For L = 1 To nLane
        tE = needTop(L) - BoxH(1, NAME_FONT)
        If tE < 0 Then tE = 0
        bE = needBtm(L) - BoxH(1, SUB_FONT)
        If bE < 0 Then bE = 0
        laneArrowOff(L) = ARROW_OFFSET + tE
        laneExtra(L) = tE + bE
    Next L
    LayoutLabels = mx / dayW
End Function

' 項目 k のラベルを幅 wAvail（pt）に収める（isName：矢線の上の名前／False：矢線の下の会社・日数）。
' 1行で入ればそのまま、入らなければ折り返す。txt（行は vbCr でつなぐ）・n（行数）・w（幅）を返す
Private Sub FitLabel(ByVal k As Long, ByVal isName As Boolean, ByVal wAvail As Double, ByRef txt As String, _
                     ByRef n As Long, ByRef w As Double)
    Dim id As String, body As String, one As String, need As Double, wMin As Double
    If isName Then
        LabelParts k, id, body
        one = OneLine(id, body)
        need = TextWidth(one, NAME_FONT) + 6
        If need <= wAvail And InStr(body, vbLf) = 0 And InStr(body, vbCr) = 0 Then
            txt = one
            n = 1
            w = need
            Exit Sub
        End If
        wMin = 4 * NAME_FONT
        If TextWidth(id, NAME_FONT) > wMin Then wMin = TextWidth(id, NAME_FONT)
        wMin = wMin + 6
        If wAvail < wMin Then wAvail = wMin
        txt = WrapLabel(id, body, wAvail - 6, NAME_FONT, n)
        w = MaxLineWidth(txt, NAME_FONT) + 6
    Else
        body = SubLabel(k)
        need = TextWidth(body, SUB_FONT) + 6
        If need <= wAvail Then
            txt = body
            n = 1
            w = need
            Exit Sub
        End If
        wMin = 4 * SUB_FONT + 6
        If wAvail < wMin Then wAvail = wMin
        txt = SubWrap(body, wAvail - 6, n)
        w = MaxLineWidth(txt, SUB_FONT) + 6
    End If
End Sub

' ラベルに書くものがないか
Private Function LabelEmpty(ByVal k As Long, ByVal isName As Boolean) As Boolean
    Dim id As String, body As String
    If isName Then
        LabelParts k, id, body
        LabelEmpty = (Len(id) = 0 And Len(body) = 0)
    Else
        LabelEmpty = (Len(SubLabel(k)) = 0)
    End If
End Function

' 名前を幅 maxW（pt）に収まるよう折り返す（行は vbCr でつなぐ。n に行数）
'   ・「作業ID 名前の1行目」と残りの行（名前の中の改行）がすべて幅に入れば、そのまま
'   ・入らなければ、作業IDを1行目に置き、名前を2行目以降に折り返す
Private Function WrapLabel(ByVal id As String, ByVal body As String, ByVal maxW As Double, ByVal fs As Double, _
                           ByRef n As Long) As String
    Dim paras() As String, i As Long, out As String, ok As Boolean, first As String
    n = 0
    If Len(body) = 0 Then
        n = 1
        WrapLabel = id
        Exit Function
    End If
    body = Replace(Replace(body, vbCrLf, vbLf), vbCr, vbLf)
    paras = Split(body, vbLf)
    For i = 0 To UBound(paras)
        paras(i) = TrimSp(paras(i))
    Next i
    first = OneLine(id, paras(0))
    ok = (TextWidth(first, fs) <= maxW)
    For i = 1 To UBound(paras)
        If TextWidth(paras(i), fs) > maxW Then ok = False
    Next i
    If ok Then
        If Len(first) > 0 Then
            out = first
            n = 1
        End If
        For i = 1 To UBound(paras)
            If Len(paras(i)) > 0 Then
                If Len(out) > 0 Then out = out & vbCr
                out = out & paras(i)
                n = n + 1
            End If
        Next i
    Else
        If Len(id) > 0 Then
            out = id
            n = 1
        End If
        For i = 0 To UBound(paras)
            If Len(paras(i)) > 0 Then WrapPara paras(i), maxW, fs, out, n
        Next i
    End If
    WrapLabel = out
End Function

' 1段落を幅 maxW に折り返して out に足す（行は vbCr でつなぐ。n に行数を足す）
'   行数がいちばん少ない切り方のうち、「・」などの切れ目で切れるもの、さらに短すぎる行がないものを選ぶ
Private Sub WrapPara(ByVal p As String, ByVal maxW As Double, ByVal fs As Double, ByRef out As String, ByRef n As Long)
    Dim m As Long, i As Long, j As Long, jf As Long, w As Double, lw As Double
    Dim cl As Long, cu As Long, cs As Double, found As Boolean, ln As String
    Dim bl() As Long, bu() As Long, bs() As Double, nx() As Long

    m = Len(p)
    If m = 0 Then Exit Sub
    ReDim bl(0 To m): ReDim bu(0 To m): ReDim bs(0 To m): ReDim nx(0 To m)
    bs(m) = 1E+30
    For i = m - 1 To 0 Step -1
        found = False
        w = 0
        jf = i + 1
        For j = i + 1 To m
            w = w + CharWidth(Mid$(p, j, 1), fs)
            If w > maxW And j > i + 1 Then Exit For
            jf = j
            If j = m Or OkBreak(p, j) Then
                lw = TextWidth(TrimSp(Mid$(p, i + 1, j - i)), fs)
                cl = bl(j) + 1
                cu = bu(j)
                If j < m Then
                    If Not GoodBreak(p, j) Then cu = cu + 1
                End If
                If lw < bs(j) Then cs = lw Else cs = bs(j)
                If Not found Then
                    found = True
                    bl(i) = cl: bu(i) = cu: bs(i) = cs: nx(i) = j
                ElseIf cl < bl(i) Or (cl = bl(i) And (cu < bu(i) Or (cu = bu(i) And cs > bs(i)))) Then
                    bl(i) = cl: bu(i) = cu: bs(i) = cs: nx(i) = j
                End If
            End If
        Next j
        If Not found Then
            ' 禁則で切れる位置がないときは、幅に入るところで切る
            bl(i) = bl(jf) + 1: bu(i) = bu(jf) + 1: bs(i) = 0: nx(i) = jf
        End If
    Next i
    i = 0
    Do While i < m
        j = nx(i)
        ln = TrimSp(Mid$(p, i + 1, j - i))
        If Len(ln) > 0 Then
            If Len(out) > 0 Then out = out & vbCr
            out = out & ln
            n = n + 1
        End If
        i = j
    Loop
End Sub

' p の j 文字目のあとで切ってよいか（行頭禁則・行末禁則）
Private Function OkBreak(ByVal p As String, ByVal j As Long) As Boolean
    If j < 1 Or j >= Len(p) Then Exit Function
    If InStr(NO_START, Mid$(p, j + 1, 1)) > 0 Then Exit Function
    If InStr(NO_END, Mid$(p, j, 1)) > 0 Then Exit Function
    OkBreak = True
End Function

' p の j 文字目のあとが、切れ目として良い位置か（「・」「、」空白のあと、かっこの前後）
Private Function GoodBreak(ByVal p As String, ByVal j As Long) As Boolean
    If Not OkBreak(p, j) Then Exit Function
    GoodBreak = (InStr(GOOD_AFTER, Mid$(p, j, 1)) > 0 Or InStr(GOOD_BEFORE, Mid$(p, j + 1, 1)) > 0)
End Function

' 会社・日数は「・」の位置だけで折り返す（会社名の途中では切らない）
Private Function SubWrap(ByVal txt As String, ByVal maxW As Double, ByRef n As Long) As String
    Dim parts() As String, i As Long, cur As String, out As String
    parts = Split(txt, "・")
    cur = parts(0)
    n = 0
    For i = 1 To UBound(parts)
        If TextWidth(cur & "・" & parts(i), SUB_FONT) <= maxW Then
            cur = cur & "・" & parts(i)
        Else
            If Len(out) > 0 Then out = out & vbCr
            out = out & cur
            n = n + 1
            cur = parts(i)
        End If
    Next i
    If Len(out) > 0 Then out = out & vbCr
    SubWrap = out & cur
    n = n + 1
End Function

' 同じ段のラベル（置いたもの）と重なるときは、その上（会社・日数は下）へ積む。
' 矢線にいちばん近い空いた位置（矢線からの距離 pt）を返す
Private Function StackOffset(ByVal L As Long, ByVal isName As Boolean, ByVal x1 As Double, ByVal x2 As Double, _
                             ByVal h As Double) As Double
    Dim q As Long, q2 As Long, cnd As Double, best As Double, found As Boolean, ok As Boolean, use As Boolean
    For q = 0 To nPl
        use = False
        If q = 0 Then
            cnd = 0
            use = True
        ElseIf plLane(q) = L And plIsName(q) = isName Then
            If x1 < plX2(q) + LABEL_GAP And plX1(q) < x2 + LABEL_GAP Then
                cnd = plOff(q) + plH(q) + VGAP
                use = True
            End If
        End If
        If use And found Then
            If cnd >= best Then use = False
        End If
        If use Then
            ok = True
            For q2 = 1 To nPl
                If plLane(q2) = L And plIsName(q2) = isName Then
                    If x1 < plX2(q2) + LABEL_GAP And plX1(q2) < x2 + LABEL_GAP Then
                        If cnd + h + VGAP > plOff(q2) And plOff(q2) + plH(q2) + VGAP > cnd Then ok = False
                    End If
                End If
            Next q2
            If ok Then
                best = cnd
                found = True
            End If
        End If
    Next q
    StackOffset = best
End Function

' n 行のラベルの枠の高さ（1行目は余白込みで文字サイズの 1.9 倍、2行目からは行の間隔ずつ）
Private Function BoxH(ByVal n As Long, ByVal fs As Double) As Double
    BoxH = fs * (1.9 + LINE_PITCH * (n - 1))
End Function

' 行（vbCr 区切り）のうち最も広い行の幅
Private Function MaxLineWidth(ByVal txt As String, ByVal fs As Double) As Double
    Dim ls() As String, i As Long, w As Double
    ls = Split(txt, vbCr)
    For i = 0 To UBound(ls)
        w = TextWidth(ls(i), fs)
        If w > MaxLineWidth Then MaxLineWidth = w
    Next i
End Function

' 前後の半角・全角の空白を取る
Private Function TrimSp(ByVal t As String) As String
    Do While Len(t) > 0
        If Left$(t, 1) = " " Or Left$(t, 1) = "　" Then t = Mid$(t, 2) Else Exit Do
    Loop
    Do While Len(t) > 0
        If Right$(t, 1) = " " Or Right$(t, 1) = "　" Then t = Left$(t, Len(t) - 1) Else Exit Do
    Loop
    TrimSp = t
End Function

' つながり（ダミー・余裕）の線
Private Sub DrawLinks()
    Dim j As Long
    nSide = 0
    ReDim sdDay(1 To nLink + 1): ReDim sdSide(1 To nLink + 1): ReDim sdLo(1 To nLink + 1): ReDim sdHi(1 To nLink + 1)
    ReDim sdFrom(1 To nLink + 1): ReDim sdTo(1 To nLink + 1): ReDim sdSlot(1 To nLink + 1)
    For j = 1 To nLink
        Progress "つなぎ線 " & j & "/" & nLink
        LinkItems j
    Next j
End Sub

Private Sub LinkItems(ByVal j As Long)
    Dim La As Long, Lb As Long, fa As Long, sb As Long
    Dim xa As Double, xb As Double, ya As Double, yb As Double, r As Double
    Dim clr As Long, gray As Long, red As Long

    r = NODE_D / 2
    gray = RGB(89, 89, 89)
    red = RGB(192, 0, 0)
    La = itLane(lkFrom(j)): Lb = itLane(lkTo(j))
    fa = lkA(j): sb = lkB(j)
    xa = XOf(fa): xb = XOf(sb)
    ya = YOf(La): yb = YOf(Lb)

    If fa > sb Then
        ' 先行作業の終了前に開始している（チェック列にも指摘あり）
        DrawLineBetweenNodes xa, ya, xb, yb, red, 1, msoLineDash, True

    ElseIf La = Lb Then
        ' 同じ段：同じ日なら結合点を共有しているので線は不要
        If fa < sb Then
            If LaneFree(La, fa, sb) Then
                DrawLine xa + r, ya, xb - r, yb, gray, 1, msoLineRoundDot, False
                laneOcc(La).Add Array(fa, sb)
            Else
                ' 間に別の項目があるときは上を回す
                DrawLine xa, ya - r, xa, ya - r - 8, gray, 1, msoLineDash, False
                DrawLine xa, ya - r - 8, xb, ya - r - 8, gray, 1, msoLineDash, False
                DrawLine xb, ya - r - 8, xb, yb - r, gray, 1, msoLineDash, True
            End If
        End If

    Else
        If lkCrit(j) And fa = sb Then clr = red Else clr = gray
        If fa = sb Then
            If VerticalClear(fa, La, Lb) Then
                ' 同じ日：縦の点線矢印（ダミー）。クリティカルパスは赤い帯も
                If clr = red Then DrawBand xa, ya, xa, yb
                DrawVertical xa, ya, yb, clr, True, True
            Else
                ' 途中の段の結合点を通ってしまうときは、結合点の横を通す
                DrawSideRoute fa, La, Lb, clr
            End If
        ElseIf LaneFree(La, fa, sb) And Not NodeExists(La, sb) And VerticalClear(sb, La, Lb) Then
            ' 先行側の段で余裕（点線）を引いてから、後続へダミー
            DrawLine xa + r, ya, xb, ya, gray, 1, msoLineRoundDot, False
            laneOcc(La).Add Array(fa, sb)
            DrawVertical xb, ya, yb, gray, False, True
        ElseIf LaneFree(Lb, fa, sb) And Not NodeExists(Lb, fa) And VerticalClear(fa, La, Lb) Then
            ' 後続側の段へダミーを下ろしてから、余裕（点線）
            DrawVertical xa, ya, yb, gray, True, False
            DrawLine xa, yb, xb - r, yb, gray, 1, msoLineRoundDot, False
            laneOcc(Lb).Add Array(fa, sb)
        Else
            ' どちらの段も空いていないときは斜めの点線矢印
            DrawLineBetweenNodes xa, ya, xb, yb, gray, 1, msoLineDash, True
        End If
    End If
End Sub

' 同じ日のダミーで、途中の段の結合点を通ってしまうもの：結合点の横（右か左）を通して描く
'   ・右と左のどちらを通るかは SideOf、結合点からの距離は SideSlot で決める
'   ・同じ結合点から出る線、同じ結合点へ入る線は、同じ位置を共有する
Private Sub DrawSideRoute(ByVal d As Long, ByVal La As Long, ByVal Lb As Long, ByVal clr As Long)
    Dim sd As Long, slot As Long, x As Double, xs As Double, ya As Double, yb As Double
    Dim y2 As Double, y3 As Double, sg As Double, r As Double, ln As Double, wt As Double

    r = NODE_D / 2
    sd = SideOf(d, La, Lb)
    slot = SideSlot(d, sd, La, Lb)
    If slot > 2 Then slot = 2
    x = XOf(d)
    xs = x + sd * (r + 5 + 4 * slot)
    ya = YOf(La)
    yb = YOf(Lb)
    If yb > ya Then sg = 1 Else sg = -1
    y2 = ya + sg * (r + 8)
    y3 = yb - sg * (r + 8)
    If clr = RGB(192, 0, 0) Then
        wt = 1.5
        DrawBand x, ya, xs, y2
        DrawBand xs, y2, xs, y3
        DrawBand xs, y3, x, yb
    Else
        wt = 1
    End If
    ln = Sqr((xs - x) ^ 2 + (y2 - ya) ^ 2)
    DrawLine x + (xs - x) / ln * r, ya + (y2 - ya) / ln * r, xs, y2, clr, wt, msoLineDash, False
    DrawLine xs, y2, xs, y3, clr, wt, msoLineDash, False
    ln = Sqr((x - xs) ^ 2 + (yb - y3) ^ 2)
    DrawLine xs, y3, x - (x - xs) / ln * r, yb - (yb - y3) / ln * r, clr, wt, msoLineDash, True
End Sub

' 結合点の横を右（+1）と左（-1）のどちらで通るか。途中の段の、右にあるもの（出ていく矢線・
' 他の期の作業のラベル）と左にあるもの（入ってくる矢線）を数え、右が多ければ左
Private Function SideOf(ByVal d As Long, ByVal La As Long, ByVal Lb As Long) As Long
    Dim k As Long, L1 As Long, L2 As Long, nL As Long, nR As Long
    If La < Lb Then
        L1 = La
        L2 = Lb
    Else
        L1 = Lb
        L2 = La
    End If
    For k = 1 To nItem
        If itLane(k) > L1 And itLane(k) < L2 Then
            If itKind(k) = KIND_FROM Or itKind(k) = KIND_TO Then
                If itS(k) = d Then nR = nR + 1
            Else
                If itS(k) = d Then nR = nR + 1
                If itF(k) = d Then nL = nL + 1
                If itS(k) < d And d < itF(k) Then
                    If NodeExists(itLane(k), d) Then
                        nL = nL + 1
                        nR = nR + 1
                    End If
                End If
            End If
        End If
    Next k
    ' 両端が他の期の作業なら、○の右にラベルがある
    If NodeIsGhost(La, d) Then nR = nR + 2
    If NodeIsGhost(Lb, d) Then nR = nR + 2
    If nR > nL Then SideOf = -1 Else SideOf = 1
End Function

' 結合点からの距離の段階（0, 1, 2…）。同じ日・同じ側で段の範囲が重なる線とは別の段階にする
' （同じ結合点から出る線・同じ結合点へ入る線どうしは共有してよい）
Private Function SideSlot(ByVal d As Long, ByVal sd As Long, ByVal La As Long, ByVal Lb As Long) As Long
    Dim q As Long, slot As Long, lo As Long, hi As Long, ok As Boolean
    If La < Lb Then
        lo = La
        hi = Lb
    Else
        lo = Lb
        hi = La
    End If
    Do
        ok = True
        For q = 1 To nSide
            If sdDay(q) = d And sdSide(q) = sd And sdSlot(q) = slot Then
                If sdLo(q) <= hi And lo <= sdHi(q) Then
                    If sdFrom(q) <> La And sdTo(q) <> Lb Then ok = False
                End If
            End If
        Next q
        If ok Then Exit Do
        slot = slot + 1
    Loop
    nSide = nSide + 1
    sdDay(nSide) = d
    sdSide(nSide) = sd
    sdLo(nSide) = lo
    sdHi(nSide) = hi
    sdFrom(nSide) = La
    sdTo(nSide) = Lb
    sdSlot(nSide) = slot
    SideSlot = slot
End Function

Private Function NodeIsGhost(ByVal L As Long, ByVal d As Long) As Boolean
    Dim key As String
    key = L & "|" & d
    If nodeDict.Exists(key) Then NodeIsGhost = nodeGhost(nodeDict(key))
End Function

' 縦線が途中の段の結合点を通らないか
Private Function VerticalClear(ByVal d As Long, ByVal La As Long, ByVal Lb As Long) As Boolean
    Dim L As Long, L1 As Long, L2 As Long
    If La < Lb Then
        L1 = La
        L2 = Lb
    Else
        L1 = Lb
        L2 = La
    End If
    For L = L1 + 1 To L2 - 1
        If NodeExists(L, d) Then Exit Function
    Next L
    VerticalClear = True
End Function

Private Function LaneFree(ByVal L As Long, ByVal a As Long, ByVal b As Long) As Boolean
    Dim v As Variant
    For Each v In laneOcc(L)
        If v(0) < b And a < v(1) Then Exit Function
    Next v
    LaneFree = True
End Function

' 結合点（番号は日付の早い順、同じ日は上の段から。他の期の作業は番号なしの点線の○）
Private Sub DrawNodes()
    Dim idx() As Long, i As Long, k As Long, j As Long, num As Long

    ReDim idx(1 To nNode)
    For i = 1 To nNode
        idx(i) = i
        k = i
        Do While k > 1
            If NodeBefore(idx(k), idx(k - 1)) Then
                j = idx(k - 1): idx(k - 1) = idx(k): idx(k) = j
                k = k - 1
            Else
                Exit Do
            End If
        Loop
    Next i
    For k = 1 To nNode
        Progress "結合点 " & k & "/" & nNode
        i = idx(k)
        If nodeGhost(i) Then
            DrawGhostNode XOf(nodeDay(i)), YOf(nodeLane(i))
        Else
            num = num + 1
            DrawNode XOf(nodeDay(i)), YOf(nodeLane(i)), CStr(num), nodeCrit(i)
            lastShape.AlternativeText = NodeInfo(i)
        End If
    Next k
End Sub

Private Function NodeBefore(ByVal a As Long, ByVal b As Long) As Boolean
    If nodeDay(a) <> nodeDay(b) Then
        NodeBefore = (nodeDay(a) < nodeDay(b))
    Else
        NodeBefore = (nodeLane(a) < nodeLane(b))
    End If
End Function

Private Sub DrawLegend()
    Dim x As Double, y As Double, r As Double, xStart As Double, xMax As Double
    Dim i As Long, k As Long, rowN As Long, hasMulti As Boolean, txt As String

    r = NODE_D / 2
    legendRow = firstLaneRow + nLane + 1
    wsC.Rows(legendRow).RowHeight = 24
    wsC.Rows(legendRow + 1).RowHeight = 18
    wsC.Cells(legendRow, 1).Value = "凡例"
    wsC.Cells(legendRow, 1).IndentLevel = 1

    ' 1行目：記号
    y = wsC.Cells(legendRow, 1).Top + 12
    x = wsC.Cells(legendRow, dateCol).Left + 4
    DrawLine x, y, x + 36, y, RGB(64, 64, 64), 1.25, msoLineSolid, True
    AddText x + 42, y, "作業（矢線）", 8, RGB(0, 0, 0), False, 2
    x = x + 130
    DrawNode x + r, y, "", False
    AddText x + NODE_D + 6, y, "結合点（番号付きの○）", 8, RGB(0, 0, 0), False, 2
    x = x + 150
    DrawLine x, y, x + 36, y, RGB(89, 89, 89), 1, msoLineDash, True
    AddText x + 42, y, "ダミー（他の工事とのつながり）", 8, RGB(0, 0, 0), False, 2
    x = x + 190
    DrawBand x, y, x + 36, y
    DrawLine x, y, x + 36, y, RGB(64, 64, 64), 2.25, msoLineSolid, True
    AddText x + 42, y, "クリティカルパス（赤い帯）", 8, RGB(0, 0, 0), False, 2
    x = x + 170
    DrawLine x, y, x + 36, y, RGB(89, 89, 89), 1, msoLineRoundDot, False
    AddText x + 42, y, IIf(useSummary, "余裕日数・まとめの待ち", "余裕日数"), 8, RGB(0, 0, 0), False, 2

    ' 2行目：他の期の作業（期のシートだけ）と、文字の表記
    y = wsC.Cells(legendRow + 1, 1).Top + 9
    x = wsC.Cells(legendRow + 1, dateCol).Left + 4
    If sheetMode > 0 Then
        DrawGhostNode x + r, y
        AddText x + NODE_D + 6, y, "他の期の作業（点線の○）", 8, RGB(0, 0, 0), False, 2
        x = x + 170
    End If
    If showProg Then LegendProgress x, y
    txt = "表記：矢線の上＝作業ID・作業名"
    If useSummary Then txt = txt & "（まとめは名前）"
    If showProg Then txt = txt & "、矢線の下＝会社・日数・進捗・遅れ（進み）" Else txt = txt & "、矢線の下＝会社・日数（進捗があれば％も）"
    AddText x, y, txt, 8, RGB(64, 64, 64), False, 2
    legendLastRow = legendRow + 1

    ' 3行目から：会社の色（右端で折り返す）
    If useSummary Then
        For k = 1 To nItem
            If itKind(k) = KIND_SUM Then
                If Len(gCoName(itGroup(k))) = 0 Then hasMulti = True
            End If
        Next k
    End If
    If nCo = 0 And Not hasMulti Then Exit Sub
    rowN = legendRow + 2
    wsC.Rows(rowN).RowHeight = 18
    y = wsC.Cells(rowN, 1).Top + 9
    xStart = wsC.Cells(rowN, dateCol).Left + 4 + 30
    xMax = wsC.Cells(1, dateCol + nCols - 1).Left + wsC.Cells(1, dateCol + nCols - 1).Width
    AddText xStart - 30, y, "会社：", 8, RGB(0, 0, 0), False, 2
    x = xStart
    For i = 1 To nCo
        LegendChip coOrder(i), CompanyColor(coOrder(i)), x, y, rowN, xStart, xMax
    Next i
    If hasMulti Then LegendChip "複数の会社のまとめ", RGB(64, 64, 64), x, y, rowN, xStart, xMax
    legendLastRow = rowN
End Sub

' 凡例の会社1つ分（右端を超えるときは次の行へ）
Private Sub LegendChip(ByVal txt As String, ByVal clr As Long, ByRef x As Double, ByRef y As Double, _
                       ByRef rowN As Long, ByVal xStart As Double, ByVal xMax As Double)
    Dim w As Double
    w = 28 + TextWidth(txt, 8) + 14
    If x + w > xMax And x > xStart Then
        rowN = rowN + 1
        wsC.Rows(rowN).RowHeight = 18
        y = wsC.Cells(rowN, 1).Top + 9
        x = xStart
    End If
    DrawLine x, y, x + 24, y, clr, 2.25, msoLineSolid, False
    AddText x + 28, y, txt, 8, clr, False, 2
    x = x + w
End Sub

' 枠線を消し、見出しを固定し、A3横・横1ページで印刷できるようにする
Private Sub FinishSheet()
    Dim lastCol As Long

    lastCol = dateCol + nCols - 1
    AddRedrawButton
    wsC.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.FreezePanes = False
    wsC.Cells(firstLaneRow, dateCol).Select
    ActiveWindow.FreezePanes = True
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1

    On Error Resume Next    ' プリンターによっては用紙サイズを設定できないことがあるため
    Application.PrintCommunication = False
    With wsC.PageSetup
        .PrintArea = wsC.Range(wsC.Cells(1, 1), wsC.Cells(legendLastRow + 1, lastCol)).Address
        .Orientation = xlLandscape
        .PaperSize = xlPaperA3
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .CenterHorizontally = True
        .LeftMargin = Application.CentimetersToPoints(1)
        .RightMargin = Application.CentimetersToPoints(1)
        .TopMargin = Application.CentimetersToPoints(1.2)
        .BottomMargin = Application.CentimetersToPoints(1.2)
        .CenterFooter = "&P / &N"
    End With
    Application.PrintCommunication = True
    On Error GoTo 0
End Sub


'==============================================================================
'  図形を描く部品
'==============================================================================
' 日 d の左端の x 座標。日付の列は1日1列ですべて同じ幅なので、最初の列の位置と幅を一度だけ読んで計算する
' （毎回セルに問い合わせると、項目が多いときに遅い）
Private Function XOf(ByVal d As Long) As Double
    If Not xCacheOK Then
        xLeft0 = wsC.Cells(rowDay, dateCol).Left
        xColW = wsC.Cells(rowDay, dateCol).Width
        xCacheOK = True
    End If
    XOf = xLeft0 + (d - viewStart) * xColW
End Function

' 段 L の矢線の y 座標（段の行の上端は、行の高さを決めたあとに一度だけ読む）
Private Function YOf(ByVal L As Long) As Double
    Dim i As Long
    If Not yCacheOK Then
        ReDim yTop(1 To nLane + 1)
        For i = 1 To nLane + 1
            yTop(i) = wsC.Cells(firstLaneRow + i - 1, 1).Top
        Next i
        yCacheOK = True
    End If
    YOf = yTop(L) + laneArrowOff(L)
End Function

Private Sub DrawLine(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, _
                     ByVal clr As Long, ByVal wt As Double, ByVal dash As Long, ByVal arrow As Boolean)
    Dim s As Shape
    If Abs(x2 - x1) < 0.5 And Abs(y2 - y1) < 0.5 Then Exit Sub
    Set s = wsC.Shapes.AddLine(x1, y1, x2, y2)
    With s.Line
        .ForeColor.RGB = clr
        .Weight = wt
        .DashStyle = dash
        If arrow Then
            .EndArrowheadStyle = msoArrowheadTriangle
            .EndArrowheadLength = msoArrowheadShort
            .EndArrowheadWidth = msoArrowheadNarrow
        End If
    End With
    NameShape s, "L"
End Sub

' クリティカルパスの赤い帯（矢線の下に敷く、半透明の太い線）
Private Sub DrawBand(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double)
    Dim s As Shape
    If Abs(x2 - x1) < 0.5 And Abs(y2 - y1) < 0.5 Then Exit Sub
    Set s = wsC.Shapes.AddLine(x1, y1, x2, y2)
    With s.Line
        .ForeColor.RGB = RGB(255, 124, 128)
        .Weight = 9
        .Transparency = 0.55
        .DashStyle = msoLineSolid
    End With
    NameShape s, "B"
End Sub

' 結合点の円にかからないよう、両端を半径ぶん縮めて引く
Private Sub DrawLineBetweenNodes(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, _
                                 ByVal clr As Long, ByVal wt As Double, ByVal dash As Long, ByVal arrow As Boolean)
    Dim dx As Double, dy As Double, ln As Double, r As Double
    r = NODE_D / 2
    dx = x2 - x1
    dy = y2 - y1
    ln = Sqr(dx * dx + dy * dy)
    If ln <= NODE_D Then Exit Sub
    DrawLine x1 + dx / ln * r, y1 + dy / ln * r, x2 - dx / ln * r, y2 - dy / ln * r, clr, wt, dash, arrow
End Sub

' 縦のダミー（点線矢印）。fromNode/toNode が True の端は結合点の半径ぶん縮める
Private Sub DrawVertical(ByVal x As Double, ByVal y1 As Double, ByVal y2 As Double, ByVal clr As Long, _
                         ByVal fromNode As Boolean, ByVal toNode As Boolean)
    Dim dirSign As Double, r As Double
    r = NODE_D / 2
    If y2 >= y1 Then dirSign = 1 Else dirSign = -1
    If fromNode Then y1 = y1 + dirSign * r
    If toNode Then y2 = y2 - dirSign * r
    DrawLine x, y1, x, y2, clr, IIf(clr = RGB(192, 0, 0), 1.5, 1), msoLineDash, True
End Sub

Private Sub DrawNode(ByVal x As Double, ByVal y As Double, ByVal numText As String, ByVal crit As Boolean)
    Dim s As Shape
    Set s = wsC.Shapes.AddShape(msoShapeOval, x - NODE_D / 2, y - NODE_D / 2, NODE_D, NODE_D)
    s.Fill.Visible = msoTrue
    s.Fill.Solid
    s.Fill.ForeColor.RGB = RGB(255, 255, 255)
    s.Line.Visible = msoTrue
    s.Line.ForeColor.RGB = IIf(crit, RGB(192, 0, 0), RGB(64, 64, 64))
    s.Line.Weight = IIf(crit, 1.5, 1)
    s.Shadow.Visible = msoFalse
    With s.TextFrame2
        .MarginLeft = 0
        .MarginRight = 0
        .MarginTop = 0
        .MarginBottom = 0
        .WordWrap = msoFalse
        .VerticalAnchor = msoAnchorMiddle
        With .TextRange
            .Text = numText
            .Font.Size = IIf(Len(numText) >= 3, 5, 6)
            .Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
            .ParagraphFormat.Alignment = msoAlignCenter
        End With
    End With
    NameShape s, "N"
End Sub

' 他の期の作業を表す点線の○（番号なし）
Private Sub DrawGhostNode(ByVal x As Double, ByVal y As Double)
    Dim s As Shape
    Set s = wsC.Shapes.AddShape(msoShapeOval, x - NODE_D / 2, y - NODE_D / 2, NODE_D, NODE_D)
    s.Fill.Visible = msoTrue
    s.Fill.Solid
    s.Fill.ForeColor.RGB = RGB(255, 255, 255)
    s.Line.Visible = msoTrue
    s.Line.ForeColor.RGB = RGB(118, 118, 118)
    s.Line.Weight = 1
    s.Line.DashStyle = msoLineSquareDot
    s.Shadow.Visible = msoFalse
    NameShape s, "G"
End Sub

' 文字のテキストボックス
'   centered：True なら x を中心に、False なら x を左端に置く
'   vPos：0＝yRef の上、1＝yRef の下、2＝yRef に中央を合わせる
Private Function AddText(ByVal x As Double, ByVal yRef As Double, ByVal txt As String, ByVal fontSize As Double, _
                         ByVal clr As Long, ByVal centered As Boolean, ByVal vPos As Long, _
                         Optional ByVal isBold As Boolean = False) As Shape
    Dim s As Shape, w As Double, h As Double, t As Double

    w = TextWidth(txt, fontSize) + 6
    h = fontSize * 1.9
    Select Case vPos
        Case 0: t = yRef - h
        Case 1: t = yRef
        Case Else: t = yRef - h / 2
    End Select
    If centered Then x = x - w / 2

    Set s = wsC.Shapes.AddTextbox(msoTextOrientationHorizontal, x, t, w, h)
    s.Fill.Visible = msoFalse
    s.Line.Visible = msoFalse
    With s.TextFrame2
        .MarginLeft = 0
        .MarginRight = 0
        .MarginTop = 0
        .MarginBottom = 0
        .WordWrap = msoFalse
        .AutoSize = msoAutoSizeNone
        Select Case vPos
            Case 0: .VerticalAnchor = msoAnchorBottom
            Case 1: .VerticalAnchor = msoAnchorTop
            Case Else: .VerticalAnchor = msoAnchorMiddle
        End Select
        With .TextRange
            .Text = txt
            .Font.Size = fontSize
            .Font.Bold = IIf(isBold, msoTrue, msoFalse)
            .Font.Fill.ForeColor.RGB = clr
            .ParagraphFormat.Alignment = IIf(centered, msoAlignCenter, msoAlignLeft)
        End With
    End With
    NameShape s, "T"
    Set AddText = s
End Function

' 文字列のおおよその幅（全角＝文字サイズ、半角＝文字サイズの0.6倍）
Private Function TextWidth(ByVal txt As String, ByVal fontSize As Double) As Double
    Dim i As Long, w As Double
    For i = 1 To Len(txt)
        w = w + CharWidth(Mid$(txt, i, 1), fontSize)
    Next i
    TextWidth = w
End Function

' 1文字の幅（半角は文字サイズの 0.6 倍、全角は 1 倍とみなす）
Private Function CharWidth(ByVal ch As String, ByVal fontSize As Double) As Double
    Dim code As Long
    code = AscW(ch)
    If code < 0 Then code = code + 65536
    If code < &H2000& Or (code >= &HFF61& And code <= &HFF9F&) Then
        CharWidth = fontSize * 0.6
    Else
        CharWidth = fontSize
    End If
End Function

' ラベルの文字枠（位置・幅・行数は LayoutLabels で決めたもの。行は vbCr 区切り）
'   align 0：中央揃え、1：左揃え、2：右揃え
'   vPos 0：yRef の上に下揃え、1：yRef の下に上揃え、2：yRef を中心に
Private Function AddLabelBox(ByVal lft As Double, ByVal yRef As Double, ByVal w As Double, ByVal txt As String, _
                             ByVal nLines As Long, ByVal fontSize As Double, ByVal clr As Long, ByVal align As Long, _
                             ByVal vPos As Long, Optional ByVal isBold As Boolean = False) As Shape
    Dim s As Shape, h As Double, t As Double

    h = BoxH(nLines, fontSize)
    Select Case vPos
        Case 0: t = yRef - h
        Case 1: t = yRef
        Case Else: t = yRef - h / 2
    End Select
    Set s = wsC.Shapes.AddTextbox(msoTextOrientationHorizontal, lft, t, w, h)
    s.Fill.Visible = msoFalse
    s.Line.Visible = msoFalse
    With s.TextFrame2
        .MarginLeft = 0
        .MarginRight = 0
        .MarginTop = 0
        .MarginBottom = 0
        .WordWrap = msoFalse
        .AutoSize = msoAutoSizeNone
        Select Case vPos
            Case 0: .VerticalAnchor = msoAnchorBottom
            Case 1: .VerticalAnchor = msoAnchorTop
            Case Else: .VerticalAnchor = msoAnchorMiddle
        End Select
        With .TextRange
            .Text = txt
            .Font.Size = fontSize
            .Font.Bold = IIf(isBold, msoTrue, msoFalse)
            .Font.Fill.ForeColor.RGB = clr
            Select Case align
                Case 1: .ParagraphFormat.Alignment = msoAlignLeft
                Case 2: .ParagraphFormat.Alignment = msoAlignRight
                Case Else: .ParagraphFormat.Alignment = msoAlignCenter
            End Select
        End With
    End With
    NameShape s, "T"
    Set AddLabelBox = s
End Function

Private Sub NameShape(ByVal s As Shape, ByVal kind As String)
    shapeSeq = shapeSeq + 1
    s.Name = "NW_" & kind & shapeSeq
    s.Placement = xlFreeFloating
    Set lastShape = s
End Sub


'==============================================================================
'  休工日
'==============================================================================
Private Function IsHoliday(ByVal d As Long) As Boolean
    If holDict Is Nothing Then Exit Function
    IsHoliday = holDict.Exists(CStr(d))
End Function

' d が休工日なら、休工日でない日まで進める
Private Function NextWorkDay(ByVal d As Long) As Long
    Do While IsHoliday(d)
        d = d + 1
    Loop
    NextWorkDay = d
End Function

' 開始日 s から、休工日を除いて dur 日数えた最終日の翌日
Private Function AddWorkDays(ByVal s As Long, ByVal dur As Long) As Long
    Dim d As Long, cnt As Long
    d = s
    Do While cnt < dur
        If Not IsHoliday(d) Then cnt = cnt + 1
        d = d + 1
    Loop
    AddWorkDays = d
End Function

' 終了日 e から、休工日を除いて dur 日さかのぼった開始日
Private Function StartFromEnd(ByVal e As Long, ByVal dur As Long) As Long
    Dim d As Long, cnt As Long
    If dur <= 0 Then
        StartFromEnd = e
        Exit Function
    End If
    d = e + 1
    Do While cnt < dur
        d = d - 1
        If Not IsHoliday(d) Then cnt = cnt + 1
    Loop
    StartFromEnd = d
End Function


'==============================================================================
'  シートの管理
'==============================================================================
' マクロが作ったシート（目印の名前があるもの）を消す。前の版で作った「工程表」シートも消す
Private Sub DeleteGeneratedSheets()
    Dim i As Long, ws As Worksheet
    Application.DisplayAlerts = False
    For i = ThisWorkbook.Worksheets.Count To 1 Step -1
        Set ws = ThisWorkbook.Worksheets(i)
        If IsGeneratedSheet(ws) And ThisWorkbook.Worksheets.Count > 1 Then ws.Delete
    Next i
    Application.DisplayAlerts = True
End Sub

Private Function IsGeneratedSheet(ByVal ws As Worksheet) As Boolean
    Dim nm As Name, shp As Shape
    For Each nm In ws.Names
        If nm.Name Like "*!" & MARK_NAME Then
            IsGeneratedSheet = True
            Exit Function
        End If
    Next nm
    If ws.Name = "工程表" Then
        For Each shp In ws.Shapes
            If Left$(shp.Name, 3) = "NW_" Then
                IsGeneratedSheet = True
                Exit Function
            End If
        Next shp
    End If
End Function

Private Function SheetExists(ByVal nm As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

' 同じ名前のシートがあれば (1)(2)… を付ける
Private Function UniqueSheetName(ByVal baseText As String) As String
    Dim nm As String, n As Long
    nm = baseText
    Do While SheetExists(nm)
        n = n + 1
        nm = Left$(baseText, 27) & "(" & n & ")"
    Loop
    UniqueSheetName = nm
End Function

' シート名に使えない文字を置き換える
Private Function SafeSheetName(ByVal s As String) As String
    Dim ch As Variant
    For Each ch In Array(":", "\", "/", "?", "*", "[", "]")
        s = Replace(s, ch, "_")
    Next ch
    If Len(s) > 31 Then s = Left$(s, 31)
    If Len(s) = 0 Then s = "期"
    SafeSheetName = s
End Function


'==============================================================================
'  小さな部品
'==============================================================================
Private Sub AddMsg(ByVal i As Long, ByVal msg As String)
    If Len(actMsg(i)) > 0 Then actMsg(i) = actMsg(i) & " / "
    actMsg(i) = actMsg(i) & msg
End Sub

Private Function CellText(ByVal v As Variant) As String
    If IsError(v) Then Exit Function
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    CellText = Trim$(CStr(v))
End Function

Private Function IsNum(ByVal v As Variant) As Boolean
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbString Then
        If Len(Trim$(v)) = 0 Then Exit Function
    End If
    IsNum = IsNumeric(v)
End Function

Private Function DaySerial(ByVal v As Variant) As Long
    DaySerial = CLng(Int(CDbl(CDate(v))))
End Function

' 作業IDの表記ゆれをそろえる（全角→半角、空白を除く、大文字に）
Private Function NormalizeId(ByVal s As String) As String
    s = StrConv(s, vbNarrow)
    s = Replace(s, " ", "")
    NormalizeId = UCase$(s)
End Function

Private Function GetSettingValue(ByVal key As String) As Variant
    Dim ws As Worksheet, r As Long
    GetSettingValue = Empty
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SETTING_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    For r = 1 To 50
        If CellText(ws.Cells(r, 1).Value) = key Then
            GetSettingValue = ws.Cells(r, 2).Value
            Exit Function
        End If
    Next r
End Function

Private Function FindHeaderRow(ByVal ws As Worksheet, ByVal key As String) As Long
    Dim r As Long, c As Long
    For r = 1 To 20
        For c = 1 To 40
            If CellText(ws.Cells(r, c).Value) = key Then
                FindHeaderRow = r
                Exit Function
            End If
        Next c
    Next r
End Function

Private Function FindCol(ByVal ws As Worksheet, ByVal hdr As Long, ByVal key As String) As Long
    Dim c As Long, lastCol As Long
    lastCol = ws.Cells(hdr, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        If CellText(ws.Cells(hdr, c).Value) = key Then
            FindCol = c
            Exit Function
        End If
    Next c
End Function

' 工程データの最後の行（作業IDの列で値のある、いちばん下の行）。End(xlUp) はフィルターで隠れた行を飛ばすので、
' 表がテーブルならテーブルの最後の行、シートで絞り込み中なら使っている範囲の最後の行とくらべて、大きいほうにする
Private Function DataLastRow(ByVal ws As Worksheet, ByVal hdr As Long, ByVal cId As Long) As Long
    Dim r As Long, lo As ListObject
    r = ws.Cells(ws.Rows.Count, cId).End(xlUp).Row
    Set lo = ws.Cells(hdr, cId).ListObject
    If Not lo Is Nothing Then r = MaxL(r, lo.Range.Row + lo.Range.Rows.Count - 1)
    If ws.FilterMode Then r = MaxL(r, ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1)
    DataLastRow = MaxL(r, hdr)
End Function

Private Function EnsureCol(ByVal ws As Worksheet, ByVal hdr As Long, ByVal key As String) As Long
    Dim c As Long
    c = FindCol(ws, hdr, key)
    If c = 0 Then
        c = ws.Cells(hdr, ws.Columns.Count).End(xlToLeft).Column + 1
        With ws.Cells(hdr, c)
            .Value = key
            .Font.Bold = True
            .Interior.Color = RGB(217, 217, 217)
        End With
    End If
    EnsureCol = c
End Function

Private Function SafeFileName(ByVal s As String) As String
    Dim ch As Variant
    For Each ch In Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        s = Replace(s, ch, "_")
    Next ch
    SafeFileName = s
End Function

' ==============================================================================
'  工程表の上でつなぐ（作業をクリック → 後につなぐ作業をクリック）
' ==============================================================================

' 作業の矢線・名前をクリックしたときに呼ばれる（作図のときに図形へ割り当てる。直接は実行しない）
'   1回目：前の作業として選ぶ／2回目：後の作業。前の作業の作業IDを、後の作業の「先行作業」に書き足す
'   すでにつながっている組を選ぶと、外すかどうか聞く。同じ作業を2回続けてクリックすると取り消し
Public Sub NS_ClickItem()
    Dim s As Shape, info() As String, kind As String, id As String, i As Long
    Dim x1 As Double, x2 As Double, y As Double

    If spdOn Then Exit Sub      ' 作図の途中のクリックは受け付けない
    If TypeName(Application.Caller) <> "String" Then
        MsgBox "工程表の作業（矢線や名前）をクリックしたときに使うマクロです。", vbInformation
        Exit Sub
    End If
    On Error GoTo Fail
    Set s = ActiveSheet.Shapes(CStr(Application.Caller))
    info = Split(s.AlternativeText, "|")
    If UBound(info) < 5 Then Exit Sub
    If info(0) <> CLICK_TAG Then Exit Sub
    kind = info(1)
    x1 = Val(info(2)) / 10
    y = Val(info(3)) / 10
    x2 = Val(info(4)) / 10
    id = info(5)
    For i = 6 To UBound(info)
        id = id & "|" & info(i)
    Next i

    If kind = "S" Then
        MsgBox "「" & id & "」はまとめの矢線です。まとめの中の作業は、期のシート" & vbCrLf & _
               "（期を使っていなければ「" & DETAIL_SHEET & "」）でつないでください。", vbInformation
        Exit Sub
    End If
    If Len(pickId) > 0 And pickSheet = ActiveSheet.Name Then
        If id = pickId Then
            ClearPick
            ShowPendingStatus
        Else
            LinkPicked id, x1, y
        End If
    Else
        StartPick id, x2, y
    End If
    Exit Sub

Fail:
    MsgBox "つなぐ操作に失敗しました。" & vbCrLf & Err.Description, vbExclamation
End Sub

' 前の作業を選ぶ（作業の終わりの○にオレンジの目印）
Private Sub StartPick(ByVal id As String, ByVal x As Double, ByVal y As Double)
    Dim s As Shape
    ClearPick
    pickId = id
    pickSheet = ActiveSheet.Name
    pickX = x
    pickY = y
    Set s = ActiveSheet.Shapes.AddShape(msoShapeOval, x - 11, y - 11, 22, 22)
    s.Fill.ForeColor.RGB = RGB(255, 192, 0)
    s.Fill.Transparency = 0.45
    s.Line.ForeColor.RGB = RGB(237, 125, 49)
    s.Line.Weight = 2
    s.Name = PICK_MARK
    NoPrint s
    Application.StatusBar = "つなぐ：前の作業「" & id & " " & ActNameById(id) & "」を選びました。" & _
                            "後につなぐ作業の矢線か名前をクリックしてください（同じ作業をもう一度クリックで取り消し）"
End Sub

' 選んだ前の作業と、クリックした後の作業をつなぐ（すでにつながっていれば外すか聞く）
Private Sub LinkPicked(ByVal succId As String, ByVal x As Double, ByVal y As Double)
    Dim ws As Worksheet, hdr As Long, cId As Long, cPred As Long, r As Long
    Dim cur As String, ids As Collection, v As Variant, found As Boolean, newTxt As String, predId As String

    predId = pickId
    Set ws = ThisWorkbook.Worksheets(DATA_SHEET)
    hdr = FindHeaderRow(ws, "作業ID")
    If hdr > 0 Then
        cId = FindCol(ws, hdr, "作業ID")
        cPred = FindCol(ws, hdr, "先行作業")
    End If
    If hdr = 0 Or cId = 0 Or cPred = 0 Then
        MsgBox "「" & DATA_SHEET & "」に「作業ID」と「先行作業」の列が見つかりません。", vbExclamation
        ClearPick
        Exit Sub
    End If
    r = FindActRow(ws, hdr, cId, succId)
    If r = 0 Then
        MsgBox "作業ID「" & succId & "」が「" & DATA_SHEET & "」に見つかりません。" & vbCrLf & _
               "作業IDを変えたときは、先に再作図してください。", vbExclamation
        ClearPick
        Exit Sub
    End If

    cur = CellText(ws.Cells(r, cPred).Value)
    Set ids = PredIdList(cur)
    For Each v In ids
        If v = predId Then found = True
    Next v

    If found Then
        If MsgBox("「" & predId & " " & ActNameById(predId) & "」→「" & succId & " " & ActNameById(succId) & _
                  "」はつながっています。" & vbCrLf & "このつながりを外しますか？", vbYesNo + vbQuestion) = vbYes Then
            For Each v In ids
                If v <> predId Then newTxt = JoinId(newTxt, CStr(v))
            Next v
            WritePredText ws.Cells(r, cPred), newTxt
            DrawPendingMark pickX, pickY, x, y, False
            pendingCount = pendingCount + 1
        End If
    ElseIf IsUpstream(ws, hdr, cId, cPred, succId, predId) Then
        MsgBox "「" & succId & "」は「" & predId & "」より前につながっている作業なので、" & vbCrLf & _
               "この向き（" & predId & " → " & succId & "）にはつなげません（循環になります）。", vbExclamation
    Else
        WritePredText ws.Cells(r, cPred), JoinId(cur, predId)
        DrawPendingMark pickX, pickY, x, y, True
        pendingCount = pendingCount + 1
    End If
    ClearPick
    ShowPendingStatus
End Sub

Private Sub ShowPendingStatus()
    If pendingCount > 0 Then
        Application.StatusBar = "つなぐ：まだ反映していない変更が " & pendingCount & " 件あります" & _
                                "（左上の「再作図」で工程表に反映します）"
    Else
        Application.StatusBar = False
    End If
End Sub

' 選んだ前の作業の目印を消す
Private Sub ClearPick()
    If Len(pickSheet) > 0 Then
        If SheetExists(pickSheet) Then
            On Error Resume Next
            ThisWorkbook.Worksheets(pickSheet).Shapes(PICK_MARK).Delete
            On Error GoTo 0
        End If
    End If
    pickId = ""
    pickSheet = ""
End Sub

' 未反映のつなぎの目印（追加：オレンジの点線矢印、外す：灰色の点線）。前の作業の終わり → 後の作業の始まり
Private Sub DrawPendingMark(ByVal xa As Double, ByVal ya As Double, ByVal xb As Double, ByVal yb As Double, _
                            ByVal isAdd As Boolean)
    Dim s As Shape, t As Shape, clr As Long, tag As String
    tag = TMP_PREFIX & ActiveSheet.Shapes.Count & "_" & CLng(Timer * 100)
    If isAdd Then clr = RGB(237, 125, 49) Else clr = RGB(128, 128, 128)
    Set s = ActiveSheet.Shapes.AddLine(xa, ya, xb, yb)
    With s.Line
        .ForeColor.RGB = clr
        .Weight = 2.25
        .DashStyle = msoLineDash
        If isAdd Then .EndArrowheadStyle = msoArrowheadTriangle
    End With
    s.Name = tag & "L"
    NoPrint s
    Set t = ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, (xa + xb) / 2 - 40, (ya + yb) / 2 - 15, 80, 13)
    t.Fill.ForeColor.RGB = RGB(255, 255, 255)
    t.Fill.Transparency = 0.1
    t.Line.Visible = msoFalse
    With t.TextFrame2
        .MarginLeft = 0
        .MarginRight = 0
        .MarginTop = 0
        .MarginBottom = 0
        .WordWrap = msoFalse
        .VerticalAnchor = msoAnchorMiddle
        With .TextRange
            .Text = IIf(isAdd, "追加（未反映）", "外す（未反映）")
            .Font.Size = 8
            .Font.Bold = msoTrue
            .Font.Fill.ForeColor.RGB = clr
            .ParagraphFormat.Alignment = msoAlignCenter
        End With
    End With
    t.Name = tag & "T"
    NoPrint t
End Sub

' 作図したシートに、未反映のつなぎの目印が残っているか
Private Function HasPendingMarks() As Boolean
    Dim ws As Worksheet, shp As Shape
    For Each ws In ThisWorkbook.Worksheets
        If IsGeneratedSheet(ws) Then
            For Each shp In ws.Shapes
                If Left$(shp.Name, Len(TMP_PREFIX)) = TMP_PREFIX Then
                    HasPendingMarks = True
                    Exit Function
                End If
            Next shp
        End If
    Next ws
End Function

' 先行作業の文字を書き込む（「1-2」などが日付にならないよう文字列の書式にしてから）
Private Sub WritePredText(ByVal cell As Range, ByVal txt As String)
    cell.NumberFormat = "@"
    cell.Value = txt
End Sub

' 先行作業の文字の後ろに作業IDを足す
Private Function JoinId(ByVal txt As String, ByVal id As String) As String
    If Len(Trim$(txt)) = 0 Then JoinId = id Else JoinId = txt & "," & id
End Function

' 先行作業の文字を作業IDの一覧にする（作図のときと同じ読み方）
Private Function PredIdList(ByVal txt As String) As Collection
    Dim c As New Collection, t As Variant
    For Each t In Split(NormalizeId(Replace(txt, "、", ",")), ",")
        If Len(t) > 0 Then c.Add CStr(t)
    Next t
    Set PredIdList = c
End Function

' 工程データで作業IDの行を探す（見つからなければ 0）
Private Function FindActRow(ByVal ws As Worksheet, ByVal hdr As Long, ByVal cId As Long, ByVal id As String) As Long
    Dim r As Long, lastRow As Long
    lastRow = DataLastRow(ws, hdr, cId)
    For r = hdr + 1 To lastRow
        If NormalizeId(CellText(ws.Cells(r, cId).Value)) = id Then
            FindActRow = r
            Exit Function
        End If
    Next r
End Function

' 作業IDから作業名（工程データの今の内容から）
Private Function ActNameById(ByVal id As String) As String
    Dim ws As Worksheet, hdr As Long, cId As Long, cName As Long, r As Long
    On Error GoTo Done
    Set ws = ThisWorkbook.Worksheets(DATA_SHEET)
    hdr = FindHeaderRow(ws, "作業ID")
    If hdr = 0 Then Exit Function
    cId = FindCol(ws, hdr, "作業ID")
    cName = FindCol(ws, hdr, "作業名")
    If cId = 0 Or cName = 0 Then Exit Function
    r = FindActRow(ws, hdr, cId, id)
    If r > 0 Then ActNameById = OneLine("", CellText(ws.Cells(r, cName).Value))
Done:
End Function

' start の先行作業を（工程データの今の内容で）さかのぼって target に着くか（着くなら target → start は循環）
Private Function IsUpstream(ByVal ws As Worksheet, ByVal hdr As Long, ByVal cId As Long, ByVal cPred As Long, _
                            ByVal target As String, ByVal start As String) As Boolean
    Dim dict As Object, seen As Object, todo As Collection, r As Long, lastRow As Long
    Dim cur As String, v As Variant

    Set dict = CreateObject("Scripting.Dictionary")
    lastRow = DataLastRow(ws, hdr, cId)
    For r = hdr + 1 To lastRow
        cur = NormalizeId(CellText(ws.Cells(r, cId).Value))
        If Len(cur) > 0 And Not dict.Exists(cur) Then dict.Add cur, CellText(ws.Cells(r, cPred).Value)
    Next r
    Set seen = CreateObject("Scripting.Dictionary")
    Set todo = New Collection
    todo.Add start
    Do While todo.Count > 0
        cur = todo(todo.Count)
        todo.Remove todo.Count
        If cur = target Then
            IsUpstream = True
            Exit Function
        End If
        If Not seen.Exists(cur) Then
            seen.Add cur, True
            If dict.Exists(cur) Then
                For Each v In PredIdList(dict(cur))
                    todo.Add CStr(v)
                Next v
            End If
        End If
    Loop
End Function

' クリックでつなぐ図形の情報（代替テキストに入れる）：目印|種類|始まりx|y|終わりx|作業ID（まとめは名前）
'   種類 A：作業、G：他の期の作業、S：まとめ。座標は 0.1pt 単位の整数
Private Function ClickInfo(ByVal k As Long) As String
    Dim kind As String, key As String
    Select Case itKind(k)
        Case KIND_SUM
            kind = "S"
            key = gName(itGroup(k))
        Case KIND_ACT
            kind = "A"
            key = actId(itAct(k))
        Case Else
            kind = "G"
            key = actId(itAct(k))
    End Select
    ClickInfo = CLICK_TAG & "|" & kind & "|" & CLng(XOf(itS(k)) * 10) & "|" & CLng(YOf(itLane(k)) * 10) & "|" & _
                CLng(XOf(itF(k)) * 10) & "|" & key
End Function

' 図形をクリックでつなげるようにする
Private Sub MarkClick(ByVal s As Shape, ByVal info As String)
    If s Is Nothing Then Exit Sub
    s.AlternativeText = info
    s.OnAction = "NS_ClickItem"
End Sub

' 矢線の上に重ねる、透明なクリックの当たり判定
Private Sub AddHitArea(ByVal x1 As Double, ByVal y As Double, ByVal x2 As Double, ByVal info As String)
    Dim s As Shape
    Set s = wsC.Shapes.AddShape(msoShapeRectangle, x1, y - 6, x2 - x1, 12)
    s.Fill.Visible = msoTrue
    s.Fill.ForeColor.RGB = RGB(255, 255, 255)
    s.Fill.Transparency = 1
    s.Line.Visible = msoFalse
    NameShape s, "H"
    NoPrint s
    MarkClick s, info
End Sub

' 再作図のボタン（つないだ内容を反映）。目盛りの左（固定した範囲）に置き、印刷・PDFには出さない
Private Sub AddRedrawButton()
    Dim s As Shape, w As Double
    w = wsC.Cells(rowMonth, dateCol).Left - 4
    Set s = wsC.Shapes.AddShape(msoShapeRoundedRectangle, 2, wsC.Rows(rowMonth).Top + 1, w, wsC.Rows(rowMonth).Height - 2)
    s.Fill.ForeColor.RGB = RGB(31, 56, 100)
    s.Line.Visible = msoFalse
    With s.TextFrame2
        .MarginLeft = 2
        .MarginRight = 2
        .MarginTop = 0
        .MarginBottom = 0
        .WordWrap = msoFalse
        .VerticalAnchor = msoAnchorMiddle
        With .TextRange
            If w >= 125 Then .Text = "再作図（変更を反映）" Else .Text = "再作図"
            .Font.Size = 8
            .Font.Bold = msoTrue
            .Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .ParagraphFormat.Alignment = msoAlignCenter
        End With
    End With
    NameShape s, "R"
    s.Name = REDRAW_BTN
    s.OnAction = "DrawNetworkSchedule"
    NoPrint s
End Sub

' 図形を印刷・PDFに出さない
Private Sub NoPrint(ByVal s As Shape)
    On Error Resume Next
    s.DrawingObject.PrintObject = False
    On Error GoTo 0
End Sub

' ==============================================================================
'  工程表の上で○（結合点）を動かして日付を変える
' ==============================================================================

' ○の代替テキスト：目印 タブ 日 タブ 元のx（0.1pt） タブ 1日の幅（0.001pt） タブ つながる作業
'   つながる作業は「S（ここで始まる）かF（ここで終わる）,開始,終了の翌日,作業ID」を改行でつなぐ（まとめは入れない）
Private Function NodeInfo(ByVal i As Long) As String
    Dim k As Long, ents As String
    For k = 1 To nItem
        If itKind(k) = KIND_ACT And itLane(k) = nodeLane(i) Then
            If itS(k) = nodeDay(i) Then
                If Len(ents) > 0 Then ents = ents & vbLf
                ents = ents & "S," & itS(k) & "," & itF(k) & "," & actId(itAct(k))
            End If
            If itF(k) = nodeDay(i) Then
                If Len(ents) > 0 Then ents = ents & vbLf
                ents = ents & "F," & itS(k) & "," & itF(k) & "," & actId(itAct(k))
            End If
        End If
    Next k
    NodeInfo = NODE_TAG & vbTab & nodeDay(i) & vbTab & CLng(XOf(nodeDay(i)) * 10) & vbTab & _
               CLng(dayW * 1000) & vbTab & ents
End Function

' ○を横に何日分動かしたか（○でなければ、または半日分より小さければ 0）。parts に代替テキストを分けて返す
Private Function NodeDelta(ByVal shp As Shape, ByRef parts() As String) As Long
    Dim alt As String, x0 As Double, dw As Double
    On Error GoTo Done
    alt = shp.AlternativeText
    If Left$(alt, Len(NODE_TAG)) <> NODE_TAG Then Exit Function
    parts = Split(alt, vbTab)
    If UBound(parts) < 4 Then Exit Function
    x0 = Val(parts(2)) / 10
    dw = Val(parts(3)) / 1000
    If dw <= 0 Then Exit Function
    NodeDelta = Int((shp.Left + shp.Width / 2 - x0) / dw + 0.5)
Done:
End Function

' 作図したシートに、動かした○があるか
Private Function HasMovedNodes() As Boolean
    Dim ws As Worksheet, shp As Shape, parts() As String
    For Each ws In ThisWorkbook.Worksheets
        If IsGeneratedSheet(ws) Then
            For Each shp In ws.Shapes
                If NodeDelta(shp, parts) <> 0 Then
                    HasMovedNodes = True
                    Exit Function
                End If
            Next shp
        End If
    Next ws
End Function

' 工程表の上で動かした○を読み取り、つながっている作業の開始・終了・日数を工程データに書き込む
'   ・その○で始まる作業は開始が、その○で終わる作業は終了が変わる（日数は開始～終了の暦日）
'   ・まとめと他の期の作業の○は反映しない。同じ作業を別のシートで違う日に動かしたときも反映しない
'   ・書き込んだ作業の作業IDを movedIds に入れる。戻り値 False はキャンセル（作図をやめる）
Private Function ApplyMovedNodes(ByVal wsD As Worksheet, ByVal movedIds As Collection) As Boolean
    Dim ws As Worksheet, shp As Shape, parts() As String, ents() As String, e() As String
    Dim newS As Object, newF As Object, oldS As Object, oldF As Object, bad As Object
    Dim order As New Collection, okIds As New Collection, v As Variant
    Dim day As Long, delta As Long, j As Long, nIgnored As Long, nOk As Long
    Dim s1 As Long, f1 As Long, s2 As Long, f2 As Long, id As String, msg As String, skipped As String
    Dim hdr As Long, cId As Long, cStart As Long, cEnd As Long, cDur As Long, r As Long
    Dim ans As VbMsgBoxResult

    ApplyMovedNodes = True
    Set newS = CreateObject("Scripting.Dictionary")
    Set newF = CreateObject("Scripting.Dictionary")
    Set oldS = CreateObject("Scripting.Dictionary")
    Set oldF = CreateObject("Scripting.Dictionary")
    Set bad = CreateObject("Scripting.Dictionary")

    For Each ws In ThisWorkbook.Worksheets
        If IsGeneratedSheet(ws) Then
            For Each shp In ws.Shapes
                delta = NodeDelta(shp, parts)
                If delta <> 0 Then
                    day = CLng(Val(parts(1)))
                    If Len(parts(4)) = 0 Then
                        nIgnored = nIgnored + 1
                    Else
                        ents = Split(parts(4), vbLf)
                        For j = 0 To UBound(ents)
                            e = Split(ents(j), ",", 4)
                            If UBound(e) = 3 Then
                                id = e(3)
                                If Not oldS.Exists(id) Then
                                    oldS.Add id, CLng(Val(e(1)))
                                    oldF.Add id, CLng(Val(e(2)))
                                    order.Add id
                                End If
                                If e(0) = "S" Then
                                    If newS.Exists(id) Then
                                        If newS(id) <> day + delta Then bad(id) = True
                                    Else
                                        newS.Add id, day + delta
                                    End If
                                Else
                                    If newF.Exists(id) Then
                                        If newF(id) <> day + delta Then bad(id) = True
                                    Else
                                        newF.Add id, day + delta
                                    End If
                                End If
                            End If
                        Next j
                    End If
                End If
            Next shp
        End If
    Next ws
    If order.Count = 0 Then
        If nIgnored > 0 Then
            MsgBox "動かした○のうち " & nIgnored & " 個は、まとめか他の期の作業の○なので反映しません" & _
                   "（元の位置に戻ります）。", vbInformation
        End If
        Exit Function
    End If

    ' 作業ごとに新しい開始・終了を決める
    For Each v In order
        id = CStr(v)
        s1 = oldS(id)
        f1 = oldF(id)
        If newS.Exists(id) Then s2 = newS(id) Else s2 = s1
        If newF.Exists(id) Then f2 = newF(id) Else f2 = f1
        If bad.Exists(id) Then
            skipped = skipped & vbCrLf & "・" & id & "：シートによって動かした日が違います"
        ElseIf f1 > s1 And f2 <= s2 Then
            skipped = skipped & vbCrLf & "・" & id & "：終わりが始まりと同じか、より前になります"
        ElseIf s2 <> s1 Or f2 <> f1 Then
            nOk = nOk + 1
            okIds.Add id
            newS(id) = s2
            newF(id) = f2
            If nOk <= 8 Then
                msg = msg & "・" & id & " " & ActNameById(id) & vbCrLf & "　　" & SpanText(s1, f1) & " → " & _
                      SpanText(s2, f2) & vbCrLf
            End If
        End If
    Next v
    If nOk > 8 Then msg = msg & "　ほか " & (nOk - 8) & " 件" & vbCrLf
    If Len(skipped) > 0 Then msg = msg & vbCrLf & "反映しないもの：" & skipped & vbCrLf
    If nIgnored > 0 Then msg = msg & vbCrLf & "（まとめ・他の期の作業の○ " & nIgnored & " 個は反映しません）" & vbCrLf
    If nOk = 0 Then
        MsgBox "動かした○を反映できる作業がありません。○は元の位置に戻ります。" & vbCrLf & msg, vbInformation
        Exit Function
    End If
    ans = MsgBox("工程表の上で動かした○を、「" & DATA_SHEET & "」の開始・終了・日数に次のとおり書き込みます。" & _
                 vbCrLf & vbCrLf & msg & vbCrLf & "はい：書き込んで作図　いいえ：書き込まずに作図（○は元に戻る）" & _
                 "　キャンセル：やめる", vbYesNoCancel + vbQuestion)
    If ans = vbCancel Then
        ApplyMovedNodes = False
        Exit Function
    End If
    If ans = vbNo Then Exit Function

    hdr = FindHeaderRow(wsD, "作業ID")
    If hdr = 0 Then Exit Function
    cId = FindCol(wsD, hdr, "作業ID")
    cStart = FindCol(wsD, hdr, "開始")
    cEnd = FindCol(wsD, hdr, "終了")
    cDur = FindCol(wsD, hdr, "日数")
    If cId = 0 Or cStart = 0 Or cEnd = 0 Or cDur = 0 Then Exit Function
    For Each v In okIds
        id = CStr(v)
        r = FindActRow(wsD, hdr, cId, id)
        If r > 0 Then
            s2 = newS(id)
            f2 = newF(id)
            WriteDate wsD.Cells(r, cStart), s2
            If f2 > s2 Then WriteDate wsD.Cells(r, cEnd), f2 - 1 Else WriteDate wsD.Cells(r, cEnd), s2
            wsD.Cells(r, cDur).Value = f2 - s2
            movedIds.Add id
            movedApplied = movedApplied + 1
        End If
    Next v
End Function

' 動かした作業の後に続く作業が、先行作業の終わりより前に始まることになったら、後ろへずらすか聞く
'   ・開始か終了が入っている作業は、日数を変えずに後ろへずらし、入っている方の欄を書き換える
'   ・開始・終了が空欄の作業は、もともと先行作業に合わせて置かれるので書き込まない（その先の判定には使う）
'   ・戻り値：書き込んだら True（呼び出し側で読み直して計算し直す）
Private Function PushSuccessors(ByVal wsD As Worksheet, ByVal movedIds As Collection) As Boolean
    Dim idx As Object, i As Long, k As Long, v As Variant, m As Long, mx As Long, d As Long, n As Long
    Dim s2() As Long, f2() As Long, isMoved() As Boolean, isChg() As Boolean, needW() As Boolean
    Dim hasChg As Boolean, fixedS As Boolean, ns As Long, nf As Long, msg As String
    Dim hdr As Long, cStart As Long, cEnd As Long, r As Long

    Set idx = CreateObject("Scripting.Dictionary")
    For i = 1 To nAct
        If Not idx.Exists(actId(i)) Then idx.Add actId(i), i
    Next i
    ReDim s2(1 To nAct): ReDim f2(1 To nAct)
    ReDim isMoved(1 To nAct): ReDim isChg(1 To nAct): ReDim needW(1 To nAct)
    For i = 1 To nAct
        s2(i) = actS(i)
        f2(i) = actF(i)
    Next i
    For Each v In movedIds
        If idx.Exists(CStr(v)) Then isMoved(idx(CStr(v))) = True
    Next v

    ' 先行関係の順に、変わった作業の後ろを確かめる
    For k = 1 To nAct
        i = topo(k)
        If Not isMoved(i) Then
            hasChg = False
            mx = -2147483647
            For Each v In preds(i)
                If f2(v) > mx Then mx = f2(v)
                If isMoved(v) Or isChg(v) Then hasChg = True
            Next v
            If hasChg Then
                fixedS = IsDate(actStartIn(i)) Or (IsDate(actEndIn(i)) And actDurKnown(i))
                If fixedS Then
                    If s2(i) < mx Then
                        d = NextWorkDay(mx) - s2(i)
                        s2(i) = s2(i) + d
                        If IsDate(actEndIn(i)) Then f2(i) = f2(i) + d Else f2(i) = FinishFor(i, s2(i))
                        If f2(i) < s2(i) Then f2(i) = s2(i)
                        isChg(i) = True
                        needW(i) = True
                        n = n + 1
                        If n <= 8 Then
                            msg = msg & "・" & actId(i) & " " & OneLine("", actName(i)) & vbCrLf & "　　" & _
                                  SpanText(actS(i), actF(i)) & " → " & SpanText(s2(i), f2(i)) & vbCrLf
                        End If
                    End If
                Else
                    m = calcBaseDay
                    If mx > m Then m = mx
                    ns = NextWorkDay(m)
                    nf = FinishFor(i, ns)
                    If nf < ns Then nf = ns
                    If ns <> s2(i) Or nf <> f2(i) Then
                        s2(i) = ns
                        f2(i) = nf
                        isChg(i) = True
                    End If
                End If
            End If
        End If
    Next k
    If n = 0 Then Exit Function
    If n > 8 Then msg = msg & "　ほか " & (n - 8) & " 件" & vbCrLf
    If MsgBox("動かした作業の後に続く " & n & " 件の作業が、先行作業の終わりより前に始まることになります。" & _
              vbCrLf & vbCrLf & msg & vbCrLf & "これらを日数を変えずに後ろへずらしますか？" & vbCrLf & _
              "（いいえ：ずらさない。チェック列に指摘が出ます）", vbYesNo + vbQuestion) = vbNo Then Exit Function

    hdr = FindHeaderRow(wsD, "作業ID")
    If hdr = 0 Then Exit Function
    cStart = FindCol(wsD, hdr, "開始")
    cEnd = FindCol(wsD, hdr, "終了")
    If cStart = 0 Or cEnd = 0 Then Exit Function
    For i = 1 To nAct
        If needW(i) Then
            r = actRow(i)
            If IsDate(actStartIn(i)) Then WriteDate wsD.Cells(r, cStart), s2(i)
            If IsDate(actEndIn(i)) Then
                If f2(i) > s2(i) Then WriteDate wsD.Cells(r, cEnd), f2(i) - 1 Else WriteDate wsD.Cells(r, cEnd), s2(i)
            End If
            pushedCount = pushedCount + 1
        End If
    Next i
    PushSuccessors = True
End Function

' 日付を書き込む（日付の書式で）
Private Sub WriteDate(ByVal cell As Range, ByVal d As Long)
    cell.NumberFormat = "yyyy/mm/dd"
    cell.Value = CDate(d)
End Sub

' 「1/19～1/22（4日）」のような期間の表記（f は終了日の翌日）
Private Function SpanText(ByVal st As Long, ByVal f As Long) As String
    If f > st Then
        SpanText = Format(CDate(st), "m/d") & "～" & Format(CDate(f - 1), "m/d") & "（" & (f - st) & "日）"
    Else
        SpanText = Format(CDate(st), "m/d") & "（0日）"
    End If
End Function


'==============================================================================
'  実績と遅れ（イナズマ線）  v8 で追加（ここから下はモジュールの最後）
'==============================================================================
' 実績から、基準日の時点の進み・遅れと、この先の予測を計算する
'   ・基準日は「その日の作業が終わった時点」。時間軸では基準日の翌日の左端（progTb）
'   ・完了＝実績終了あり（または進捗100%）、着手中＝実績開始あり（または進捗が0より大きい）
'   ・進捗の遅れ＝基準日 － 進捗の位置（暦日。正＝遅れ、負＝進み）
'   ・予測：着手中は基準日の翌日から残り（日数×（1－進捗）を切り上げ）。未着手は予定の開始・
'     基準日の翌日・先行作業の予測終了の遅い方から日数どおり（前倒しはしない）
Private Sub CalcProgress()
    Dim i As Long, k As Long, g As Long, v As Variant, sv As Variant
    Dim isMile As Boolean, hasAS As Boolean, hasAE As Boolean, late As Boolean, es As Long, remD As Long
    Dim wsum As Double, psum As Double, posD As Double, negD As Double, allDone As Boolean

    ReDim prgState(1 To nAct): ReDim prgP(1 To nAct): ReDim prgDelay(1 To nAct): ReDim prgPlan(1 To nAct)
    ReDim prgWd(1 To nAct): ReDim prgFS(1 To nAct): ReDim prgFF(1 To nAct): ReDim prgLastFc(1 To nAct)
    ReDim gPrgP(1 To nAct): ReDim gDelay(1 To nAct)

    sv = GetSettingValue("進捗の基準日")
    If IsDate(sv) Then progT = DaySerial(sv) Else progT = DaySerial(Date)
    progTb = progT + 1

    ' 実績の入力の確認
    useProg = False
    For i = 1 To nAct
        hasAS = IsDate(actAS(i))
        hasAE = IsDate(actAE(i))
        If Len(CellText(actAS(i))) > 0 And Not hasAS Then AddMsg i, "実績開始が日付ではありません"
        If Len(CellText(actAE(i))) > 0 And Not hasAE Then AddMsg i, "実績終了が日付ではありません"
        If hasAS And hasAE Then
            If DaySerial(actAE(i)) < DaySerial(actAS(i)) Then AddMsg i, "実績終了が実績開始より前です"
        End If
        late = False
        If hasAS Then
            If DaySerial(actAS(i)) > progT Then late = True
        End If
        If hasAE Then
            If DaySerial(actAE(i)) > progT Then late = True
        End If
        If late Then AddMsg i, "実績の日付が進捗の基準日（" & DateMD(progT) & "）より後です"
        If hasAS Or hasAE Or actProg(i) > 0 Then useProg = True
    Next i
    showProg = useProg And CellText(GetSettingValue("進捗の表示")) <> "しない"
    If Not useProg Then Exit Sub

    ' 作業ごとの状態・予定進捗・進捗の遅れ
    For i = 1 To nAct
        isMile = (actF(i) <= actS(i))
        hasAS = IsDate(actAS(i))
        hasAE = IsDate(actAE(i))
        prgWd(i) = WorkDaysIn(actS(i), actF(i))
        If hasAE Or actProg(i) >= 1 Or (isMile And hasAS) Then
            prgState(i) = 2
            prgP(i) = 1
        ElseIf (hasAS Or actProg(i) > 0) And Not isMile Then
            prgState(i) = 1
            prgP(i) = actProg(i)
            If prgP(i) < 0 Then prgP(i) = 0
        Else
            prgState(i) = 0
            prgP(i) = 0
        End If
        If isMile Then
            If actS(i) < progTb Then prgPlan(i) = 1 Else prgPlan(i) = 0
        ElseIf prgWd(i) > 0 Then
            prgPlan(i) = WorkDaysIn(actS(i), MinL(actF(i), MaxL(progTb, actS(i)))) / prgWd(i)
        Else
            If actF(i) <= progTb Then prgPlan(i) = 1 Else prgPlan(i) = 0
        End If
        prgDelay(i) = progTb - ProgressPos(i)
    Next i

    ' 予測（先行関係の順に）
    For k = 1 To nAct
        i = topo(k)
        isMile = (actF(i) <= actS(i))
        Select Case prgState(i)
            Case 2
                If isMile Then
                    If IsDate(actAE(i)) Then
                        prgFF(i) = DaySerial(actAE(i))
                    ElseIf IsDate(actAS(i)) Then
                        prgFF(i) = DaySerial(actAS(i))
                    Else
                        prgFF(i) = MinL(actS(i), progTb)
                    End If
                    prgFS(i) = prgFF(i)
                Else
                    If IsDate(actAS(i)) Then prgFS(i) = DaySerial(actAS(i)) Else prgFS(i) = actS(i)
                    If IsDate(actAE(i)) Then prgFF(i) = DaySerial(actAE(i)) + 1 Else prgFF(i) = MinL(actF(i), progTb)
                End If
            Case 1
                If IsDate(actAS(i)) Then prgFS(i) = DaySerial(actAS(i)) Else prgFS(i) = MinL(actS(i), progT)
                remD = -Int(-(prgWd(i) * (1 - prgP(i)) - 0.000001))      ' 残りの日数（切り上げ）
                If remD < 1 Then remD = 1
                prgFF(i) = AddWorkDays(progTb, remD)
            Case Else
                es = MaxL(actS(i), progTb)
                For Each v In preds(i)
                    If prgFF(v) > es Then es = prgFF(v)
                Next v
                prgFS(i) = NextWorkDay(es)
                If isMile Then prgFF(i) = prgFS(i) Else prgFF(i) = AddWorkDays(prgFS(i), MaxL(prgWd(i), 1))
        End Select
        If isMile Then prgLastFc(i) = prgFF(i) Else prgLastFc(i) = prgFF(i) - 1
    Next k
    progPlanEnd = LastDay(1)
    progFcEnd = prgLastFc(1)
    For i = 2 To nAct
        If LastDay(i) > progPlanEnd Then progPlanEnd = LastDay(i)
        If prgLastFc(i) > progFcEnd Then progFcEnd = prgLastFc(i)
    Next i

    ' まとめ：進捗は稼働日数の重み付き平均、遅れは一番遅れている作業（遅れが無ければ一番進んでいる作業）
    For g = 1 To nGroup
        wsum = 0: psum = 0: posD = 0: negD = 0: allDone = True
        For i = 1 To nAct
            If actGroupIdx(i) = g Then
                wsum = wsum + prgWd(i)
                psum = psum + prgP(i) * prgWd(i)
                If prgState(i) <> 2 Then allDone = False
                If prgDelay(i) > posD Then posD = prgDelay(i)
                If prgDelay(i) < negD Then negD = prgDelay(i)
            End If
        Next i
        If wsum > 0 Then
            gPrgP(g) = psum / wsum
        ElseIf allDone Then
            gPrgP(g) = 1
        End If
        gDelay(g) = PickDelay(posD, negD)
    Next g
End Sub

' 作業 i の進捗の位置（時間軸の上の日。小数あり）
Private Function ProgressPos(ByVal i As Long) As Double
    If prgState(i) = 2 Then
        If actF(i) > progTb Then ProgressPos = actF(i) Else ProgressPos = progTb
    ElseIf prgState(i) = 1 And prgP(i) > 0 And prgWd(i) > 0 Then
        ProgressPos = PosAfterWork(actS(i), prgP(i) * prgWd(i))
    Else
        If actS(i) < progTb Then ProgressPos = actS(i) Else ProgressPos = progTb
    End If
End Function

' 一番遅れているもの（遅れが無ければ一番進んでいるもの。どちらも無ければ 0）
Private Function PickDelay(ByVal posD As Double, ByVal negD As Double) As Double
    If posD > 0.000001 Then
        PickDelay = posD
    ElseIf negD < -0.000001 Then
        PickDelay = negD
    End If
End Function

' s から稼働日で e 日（小数あり）進んだ位置
Private Function PosAfterWork(ByVal s As Long, ByVal e As Double) As Double
    Dim k As Long, fr As Double, d As Long
    k = Int(e + 0.000000001)
    fr = e - k
    If k > 0 Then d = AddWorkDays(s, k) Else d = s
    If fr > 0.000000001 Then
        PosAfterWork = NextWorkDay(d) + fr
    Else
        PosAfterWork = d
    End If
End Function

' s から f の前日までの稼働日（休工日でない日）の数
Private Function WorkDaysIn(ByVal s As Long, ByVal f As Long) As Long
    Dim d As Long
    For d = s To f - 1
        If Not IsHoliday(d) Then WorkDaysIn = WorkDaysIn + 1
    Next d
End Function

' 項目の遅れ（作業とまとめ。他の期の作業は 0）
Private Function ItemDelay(ByVal k As Long) As Double
    Select Case itKind(k)
        Case KIND_ACT: ItemDelay = prgDelay(itAct(k))
        Case KIND_SUM: ItemDelay = gDelay(itGroup(k))
    End Select
End Function

' 「n日遅れ」「n日進み」（四捨五入して0なら空）
Private Function DelayText(ByVal d As Double) As String
    Dim r As Long
    r = RoundDays(d)
    If r > 0 Then
        DelayText = r & "日遅れ"
    ElseIf r < 0 Then
        DelayText = (-r) & "日進み"
    End If
End Function

' 四捨五入（0から遠い方へ。VBA の Round は偶数への丸めなので使わない）
Private Function RoundDays(ByVal d As Double) As Long
    If d >= 0 Then RoundDays = Int(d + 0.5) Else RoundDays = -Int(-d + 0.5)
End Function

' 進捗の表記（1なら「完了」、0なら空）
Private Function ProgText(ByVal q As Double) As String
    If q >= 0.999999 Then
        ProgText = "完了"
    ElseIf q > 0 Then
        ProgText = Format(q, "0%")
    End If
End Function

' 「・」でつなぐ（どちらかが空ならつながない）
Private Function JoinDot(ByVal t As String, ByVal u As String) As String
    If Len(u) = 0 Then
        JoinDot = t
    ElseIf Len(t) = 0 Then
        JoinDot = u
    Else
        JoinDot = t & "・" & u
    End If
End Function

' 実績があるときの矢線の下の表記（SubLabel の続き）
Private Function ProgLabel(ByVal k As Long, ByVal t As String) As String
    Dim a As Long
    a = itAct(k)
    Select Case itKind(k)
        Case KIND_FROM      ' 他の期の前の作業：終わる日
            If prgState(a) = 2 Then
                t = DateMD(prgLastFc(a)) & "完了"
            ElseIf prgLastFc(a) <> LastDay(a) Then
                t = t & "→予測" & DateMD(prgLastFc(a))
            End If
        Case KIND_TO        ' 他の期の後の作業：始まる日
            If prgState(a) = 2 Then
                t = "完了"
            ElseIf prgState(a) = 1 Then
                t = DateMD(prgFS(a)) & "着手済"
            ElseIf prgFS(a) <> actS(a) Then
                t = t & "→予測" & DateMD(prgFS(a))
            End If
        Case KIND_SUM
            t = JoinDot(JoinDot(t, ProgText(gPrgP(itGroup(k)))), DelayText(gDelay(itGroup(k))))
        Case Else
            t = JoinDot(JoinDot(t, ProgText(prgP(a))), DelayText(prgDelay(a)))
    End Select
    ProgLabel = t
End Function

' 全体の進み・遅れ（表題の下と終わりのメッセージ）
Private Function ProgressSummary() As String
    Dim d As Long, t As String
    d = progFcEnd - progPlanEnd
    t = "基準日 " & DateMD(progT) & "：全体 "
    If d > 0 Then
        t = t & d & "日遅れ"
    ElseIf d < 0 Then
        t = t & (-d) & "日進み"
    Else
        t = t & "予定どおり"
    End If
    ProgressSummary = t & "（完了予定 " & DateMD(progPlanEnd) & " → 予測 " & DateMD(progFcEnd) & "）"
End Function

Private Function ProgressMsg() As String
    ProgressMsg = vbCrLf & "進み・遅れ（" & ProgressSummary() & "）"
    If Not showProg Then ProgressMsg = ProgressMsg & vbCrLf & "（「設定」の「進捗の表示」が「しない」のため、工程表には描いていません）"
End Function

Private Function DateMD(ByVal d As Long) As String
    DateMD = Format(CDate(d), "m/d")
End Function

Private Function MinL(ByVal a As Long, ByVal b As Long) As Long
    If a < b Then MinL = a Else MinL = b
End Function

Private Function MaxL(ByVal a As Long, ByVal b As Long) As Long
    If a > b Then MaxL = a Else MaxL = b
End Function

' 工程データの右に、進み・遅れと予測の列を書く（実績が無ければ前の結果を消すだけ）
Private Sub WriteProgressCols(ByVal ws As Worksheet, ByVal hdr As Long, ByVal lastRow As Long)
    Dim c(1 To 4) As Long, heads As Variant, i As Long, k As Long
    heads = Array("予定進捗", "進捗の遅れ", "予測終了", "終了の遅れ")
    For k = 1 To 4
        If useProg Then c(k) = EnsureCol(ws, hdr, CStr(heads(k - 1))) Else c(k) = FindCol(ws, hdr, CStr(heads(k - 1)))
        If c(k) > 0 And lastRow > hdr Then ws.Range(ws.Cells(hdr + 1, c(k)), ws.Cells(lastRow, c(k))).ClearContents
    Next k
    If Not useProg Then Exit Sub
    For i = 1 To nAct
        ws.Cells(actRow(i), c(1)).Value = prgPlan(i)
        ws.Cells(actRow(i), c(2)).Value = RoundDays(prgDelay(i))
        ws.Cells(actRow(i), c(3)).Value = CDate(prgLastFc(i))
        ws.Cells(actRow(i), c(4)).Value = prgLastFc(i) - LastDay(i)
    Next i
    ws.Range(ws.Cells(hdr + 1, c(1)), ws.Cells(lastRow, c(1))).NumberFormat = "0%"
    ws.Range(ws.Cells(hdr + 1, c(2)), ws.Cells(lastRow, c(2))).NumberFormat = "[Red]0""日遅れ"";[Blue]0""日進み"";""－"""
    ws.Range(ws.Cells(hdr + 1, c(3)), ws.Cells(lastRow, c(3))).NumberFormat = "yyyy/mm/dd"
    ws.Range(ws.Cells(hdr + 1, c(4)), ws.Cells(lastRow, c(4))).NumberFormat = "[Red]0""日遅れ"";[Blue]0""日早い"";""－"""
End Sub

' 基準日の列の目盛りに印を付ける（週単位は、日付の行に月曜の日付を書くので、その下の行だけ）
Private Sub MarkStatusDay(ByVal col As Long)
    Dim r1 As Long
    If weekScale Then r1 = rowWeekday Else r1 = rowDay
    With wsC.Range(wsC.Cells(r1, col), wsC.Cells(rowWeekday, col))
        .Interior.Color = PROG_COLOR
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With
End Sub

' 表題の下の行に続けて、全体の進み・遅れを書く（イナズマ線の色の太字）
Private Sub AddProgressSummary(ByVal txt As String)
    Dim n As Long
    n = Len(txt)
    txt = txt & "　" & ProgressSummary()
    wsC.Cells(2, 1).Value = txt
    With wsC.Cells(2, 1).Characters(n + 2, Len(txt) - n - 1).Font
        .Color = PROG_COLOR
        .Bold = True
    End With
End Sub

' 凡例のイナズマ線（イナズマ線を描いたシートだけ）
Private Sub LegendProgress(ByRef x As Double, ByVal y As Double)
    Dim txt As String
    If progTb < viewStart Or progTb > viewStart + nCols Then Exit Sub
    DrawZigzag x, y
    txt = "イナズマ線（" & DateMD(progT) & "の作業終了時点の進み・遅れ）"
    AddText x + 42, y, txt, 8, RGB(0, 0, 0), False, 2
    x = x + 42 + TextWidth(txt, 8) + 20
End Sub

' イナズマ線：基準日の縦線から、段ごとに進み・遅れの位置へ折れる（矢線・文字の下に敷く）
'   段の位置は、その段の作業・まとめのうち一番遅れているもの（遅れが無ければ一番進んでいるもの）
Private Sub DrawProgressLine()
    Dim L As Long, k As Long, d As Double, posD As Double, negD As Double, has As Boolean
    Dim lo As Double, hi As Double, x As Double, xT As Double, fb As FreeformBuilder

    If Not showProg Then Exit Sub
    lo = viewStart
    hi = viewStart + nCols
    If progTb < lo Or progTb > hi Then Exit Sub
    xT = XOfDay(progTb)
    Set fb = wsC.Shapes.BuildFreeform(msoEditingAuto, xT, wsC.Cells(firstLaneRow, 1).Top)
    For L = 1 To nLane
        has = False
        posD = 0
        negD = 0
        For k = 1 To nItem
            If itLane(k) = L And (itKind(k) = KIND_ACT Or itKind(k) = KIND_SUM) Then
                has = True
                d = ItemDelay(k)
                If d > posD Then posD = d
                If d < negD Then negD = d
            End If
        Next k
        x = progTb
        If has Then x = progTb - PickDelay(posD, negD)
        If x < lo Then x = lo
        If x > hi Then x = hi
        fb.AddNodes msoSegmentLine, msoEditingAuto, XOfDay(x), YOf(L)
    Next L
    fb.AddNodes msoSegmentLine, msoEditingAuto, xT, wsC.Cells(firstLaneRow + nLane, 1).Top
    FinishProgressShape fb
End Sub

' 凡例のイナズマ線（短いジグザグ）
Private Sub DrawZigzag(ByVal x As Double, ByVal y As Double)
    Dim fb As FreeformBuilder
    Set fb = wsC.Shapes.BuildFreeform(msoEditingAuto, x, y + 5)
    fb.AddNodes msoSegmentLine, msoEditingAuto, x + 12, y - 5
    fb.AddNodes msoSegmentLine, msoEditingAuto, x + 24, y + 5
    fb.AddNodes msoSegmentLine, msoEditingAuto, x + 36, y - 5
    FinishProgressShape fb
End Sub

Private Sub FinishProgressShape(ByVal fb As FreeformBuilder)
    Dim s As Shape
    Set s = fb.ConvertToShape
    s.Fill.Visible = msoFalse
    With s.Line
        .ForeColor.RGB = PROG_COLOR
        .Weight = 1.75
        .DashStyle = msoLineSolid
    End With
    NameShape s, "P"
End Sub

' 日の位置（小数あり）の x 座標
Private Function XOfDay(ByVal x As Double) As Double
    Dim d As Long
    d = Int(x)
    XOfDay = XOf(d)
    If x > d Then XOfDay = XOfDay + (x - d) * (XOf(d + 1) - XOf(d))
End Function

' ラベルの末尾（「n日遅れ」「n日進み」）をイナズマ線の色の太字にする
Private Sub AccentTail(ByVal s As Shape, ByVal txt As String, ByVal tail As String)
    If Len(tail) = 0 Or Len(txt) < Len(tail) Then Exit Sub
    If Right$(txt, Len(tail)) <> tail Then Exit Sub
    On Error Resume Next
    With s.TextFrame2.TextRange.Characters(Len(txt) - Len(tail) + 1, Len(tail)).Font
        .Fill.ForeColor.RGB = PROG_COLOR
        .Bold = msoTrue
    End With
    On Error GoTo 0
End Sub


'==============================================================================
'  速さ（v9）：作図の間だけ Excel の再計算・イベントを止め、進み具合をステータスバーに出す
'==============================================================================
Private Sub SpeedOn()
    If spdOn Then Exit Sub
    spdOn = True
    spdT0 = Timer
    spdLast = -1
    On Error Resume Next
    spdCalc = Application.Calculation
    spdEvents = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.EnableCancelKey = xlErrorHandler    ' Esc で中止（エラー 18 として ErrHandler へ）
    Application.Cursor = xlWait
    On Error GoTo 0
End Sub

Private Sub SpeedOff()
    If Not spdOn Then Exit Sub
    spdOn = False
    On Error Resume Next
    Application.EnableCancelKey = xlInterrupt
    Application.Calculation = spdCalc
    Application.EnableEvents = spdEvents
    Application.Cursor = xlDefault
    Application.StatusBar = False
    Application.ScreenUpdating = True
    On Error GoTo 0
End Sub

' 進み具合（0.5秒ごと）。そのときに DoEvents をして、Excel が「応答なし」にならないようにする
Private Sub Progress(ByVal detail As String, Optional ByVal force As Boolean = False)
    Dim t As Single
    If Not spdOn Then Exit Sub
    t = Timer
    If t < spdLast Then spdLast = -1           ' 0時をまたいだとき
    If force Or spdLast < 0 Or t - spdLast >= 0.5 Then
        Application.StatusBar = "工程表を作成中… " & progMsg & "　" & detail & "　（" & Format(Elapsed(), "0") & _
                                "秒・Esc キーで中止）"
        DoEvents
        spdLast = t
    End If
End Sub

' 作図を始めてからの秒数
Private Function Elapsed() As Single
    Elapsed = Timer - spdT0
    If Elapsed < 0 Then Elapsed = Elapsed + 86400
End Function
