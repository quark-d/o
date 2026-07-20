Attribute VB_Name = "m_LedgerCore"
Option Explicit

'==============================================================================
' m_LedgerCore: UI 非依存の台帳ロジック (設計合意 2026-07-20)
'
' - c_DialogContext の構築 (選択肢+在庫スナップショット)
' - c_LedgerOperation の検証 (形式/業務整合/在庫負防止)
' - 操作 → 台帳行への展開 (例: 不良+補填 → 不良発生+充当 の 2 行)
' - tbl_ledger への書込・再計算・書込後の負在庫チェック (ロールバック可)
' - シート保護/解除・読み取り専用推奨フラグ付き保存
'
' UI (Forms/PS どちらの実装) もここを経由せずに台帳へ書き込まないこと。
'==============================================================================

Public Enum LedgerOpKind
    opCarryIn = 1       ' 搬入
    opProgress = 2      ' 工程進行 (発送含む)
    opDefect = 3        ' 不良+補填
    opException = 4     ' 例外操作 (廃棄/倉庫A戻し/余剰化/充当)
End Enum

' シート保護のパスワード (事故防止用。変更する場合はここを書き換えて再ビルド)
Public Const PROTECT_PWD As String = "ledger"

' シート名
Public Const SH_LEDGER As String = "T_台帳"
Public Const SH_PLAN As String = "T_納期計画"
Public Const SH_MASTER As String = "M_品番"
Public Const SH_LIST As String = "M_リスト"
Public Const SH_STOCK As String = "V_在庫内訳"

' イベント種類 (M_リスト A列と一致させること)
Public Const EV_CARRYIN As String = "搬入"
Public Const EV_SHIP As String = "発送"
Public Const EV_DEFECT As String = "不良発生"
Public Const EV_DISCARD As String = "廃棄"
Public Const EV_RETURN As String = "倉庫A戻し"
Public Const EV_SURPLUS As String = "余剰化"
Public Const EV_ALLOCATE As String = "充当"

' 状態 (M_リスト B列と一致させること)
Public Const ST_RAW As String = "未処理"
Public Const ST_DEFECT As String = "不良"
Public Const ST_READY As String = "発送準備完了"

'==============================================================================
' 工程パターン (状態連鎖) とイベント→状態の対応
'==============================================================================

' パターンの状態連鎖 (未処理 → … → 発送準備完了)。パターン外は空配列
Public Function PatternChain(ByVal pattern As Long) As Variant
    Dim tail As Variant: tail = Array("梱包中", "梱包済み", "発送準備中", ST_READY)
    Select Case pattern
        Case 1
            PatternChain = ConcatArrays(Array(ST_RAW), tail)
        Case 2
            PatternChain = ConcatArrays(Array(ST_RAW, "アセンブリ中", "アセンブリ済み"), tail)
        Case 3
            PatternChain = ConcatArrays(Array(ST_RAW, "アニール中", "アニール済み"), tail)
        Case 4
            PatternChain = ConcatArrays(Array(ST_RAW, "アニール中", "アニール済み", _
                                              "アセンブリ中", "アセンブリ済み"), tail)
        Case Else
            PatternChain = Array()
    End Select
End Function

Private Function ConcatArrays(ByVal a As Variant, ByVal b As Variant) As Variant
    Dim result() As Variant
    Dim n As Long, i As Long
    n = UBound(a) - LBound(a) + 1 + UBound(b) - LBound(b) + 1
    ReDim result(0 To n - 1)
    Dim k As Long
    For i = LBound(a) To UBound(a)
        result(k) = a(i): k = k + 1
    Next
    For i = LBound(b) To UBound(b)
        result(k) = b(i): k = k + 1
    Next
    ConcatArrays = result
End Function

' 工程イベント → 遷移先状態 (工程イベント以外は空文字)
Public Function EventToState(ByVal eventKind As String) As String
    Select Case eventKind
        Case "アニール開始":     EventToState = "アニール中"
        Case "アニール完了":     EventToState = "アニール済み"
        Case "アセンブリ開始":   EventToState = "アセンブリ中"
        Case "アセンブリ完了":   EventToState = "アセンブリ済み"
        Case "梱包開始":         EventToState = "梱包中"
        Case "梱包完了":         EventToState = "梱包済み"
        Case "発送準備開始":     EventToState = "発送準備中"
        Case "発送準備完了":     EventToState = ST_READY
    End Select
End Function

' 工程進行イベントの元状態を工程パターンから導出 (パターン外は空文字)
Public Function DeriveFromState(ByVal pattern As Long, ByVal eventKind As String) As String
    If eventKind = EV_SHIP Then
        DeriveFromState = ST_READY
        Exit Function
    End If
    Dim toState As String: toState = m_LedgerCore.EventToState(eventKind)
    If toState = "" Then Exit Function
    Dim chain As Variant: chain = m_LedgerCore.PatternChain(pattern)
    Dim i As Long
    For i = LBound(chain) + 1 To UBound(chain)
        If chain(i) = toState Then
            DeriveFromState = chain(i - 1)
            Exit Function
        End If
    Next
End Function

' パターンで許可される工程進行イベント (発送含む・連鎖順)
Public Function AllowedEvents(ByVal pattern As Long) As Collection
    Dim result As New Collection
    Dim chain As Variant: chain = m_LedgerCore.PatternChain(pattern)
    Dim i As Long
    Dim ev As Variant
    For i = LBound(chain) + 1 To UBound(chain)
        ' 遷移先状態 chain(i) に対応するイベント名を逆引き
        For Each ev In Array("アニール開始", "アニール完了", "アセンブリ開始", _
                             "アセンブリ完了", "梱包開始", "梱包完了", _
                             "発送準備開始", "発送準備完了")
            If m_LedgerCore.EventToState(CStr(ev)) = chain(i) Then
                result.Add CStr(ev)
                Exit For
            End If
        Next
    Next
    result.Add EV_SHIP
    Set AllowedEvents = result
End Function

'==============================================================================
' c_DialogContext の構築
'
' 呼び出し前に Application.CalculateFullRebuild を済ませておくこと。
'==============================================================================
Public Function BuildDialogContext(ByVal wb As Workbook) As c_DialogContext
    Dim ctx As New c_DialogContext

    ' --- M_品番 ---
    Dim lo As ListObject: Set lo = wb.Worksheets(SH_MASTER).ListObjects("tbl_master")
    Dim r As ListRow
    For Each r In lo.ListRows
        If Trim$(CStr(r.Range(1).Value & "")) <> "" Then
            Dim it As c_ItemInfo: Set it = New c_ItemInfo
            it.ItemNo = CStr(r.Range(1).Value)
            it.ItemName = CStr(r.Range(2).Value & "")
            it.Pattern = CLng(Val(r.Range(3).Value & ""))
            ctx.Items.Add it
        End If
    Next

    ' --- T_納期計画 (キャンセル含む全行。UI 側で用途に応じて絞る) ---
    Set lo = wb.Worksheets(SH_PLAN).ListObjects("tbl_plan")
    For Each r In lo.ListRows
        If Trim$(CStr(r.Range(1).Value & "")) <> "" Then
            Dim lt As c_LotInfo: Set lt = New c_LotInfo
            lt.LotId = CStr(r.Range(1).Value)
            lt.ItemNo = CStr(r.Range(2).Value & "")
            lt.ShipWeek = CDate(r.Range(3).Value)
            lt.Dest = CStr(r.Range(4).Value & "")
            lt.Required = CLng(Val(r.Range(5).Value & ""))
            lt.PlanStatus = CStr(r.Range(6).Value & "")
            ctx.Lots.Add lt
        End If
    Next

    ' --- V_在庫内訳 → 状態リスト + c_StockSnapshot ---
    Dim ws As Worksheet: Set ws = wb.Worksheets(SH_STOCK)
    ' ヘッダ D1～: 「合計」の手前までが状態列
    Dim col As Long: col = 4
    Do While CStr(ws.Cells(1, col).Value & "") <> "" And _
             CStr(ws.Cells(1, col).Value & "") <> "合計"
        ctx.States.Add CStr(ws.Cells(1, col).Value)
        col = col + 1
    Loop

    Dim snap As New c_StockSnapshot
    snap.Init ctx.States

    Dim rowNo As Long
    Dim kindCell As String
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row   ' C列 (品番) 基準
    For rowNo = 2 To lastRow
        kindCell = CStr(ws.Cells(rowNo, 1).Value & "")
        If kindCell = "ロット" Or kindCell = "未割当" Then
            Dim stIdx As Long
            Dim q As Long
            For stIdx = 1 To ctx.States.Count
                q = CLng(Val(ws.Cells(rowNo, 3 + stIdx).Value & ""))
                If q <> 0 Then
                    If kindCell = "ロット" Then
                        snap.SetLotQty CStr(ws.Cells(rowNo, 2).Value & ""), _
                                       ctx.States(stIdx), q
                    Else
                        snap.SetUnallocQty CStr(ws.Cells(rowNo, 3).Value & ""), _
                                           ctx.States(stIdx), q
                    End If
                End If
            Next
        End If
    Next
    Set ctx.Stock = snap

    ' --- M_リスト G列 (記録者) ---
    Set ws = wb.Worksheets(SH_LIST)
    rowNo = 2
    Do While CStr(ws.Cells(rowNo, 7).Value & "") <> ""
        ctx.Recorders.Add CStr(ws.Cells(rowNo, 7).Value)
        rowNo = rowNo + 1
    Loop

    ' --- 既定値 ---
    ctx.DefaultActionDate = Date
    Dim lastRec As String: lastRec = GetSetting("LedgerTool", "UI", "LastRecorder", "")
    Dim v As Variant
    For Each v In ctx.Recorders
        If CStr(v) = lastRec Then ctx.DefaultRecorder = lastRec
    Next

    Set BuildDialogContext = ctx
End Function

' 書込成功後に呼ぶ (次回の既定記録者を記憶)
Public Sub SaveLastRecorder(ByVal recorder As String)
    SaveSetting "LedgerTool", "UI", "LastRecorder", recorder
End Sub

'==============================================================================
' 検証 (エラー = 書込拒否 / 警告 = 確認のうえ続行可)
'
' 戻り値 True = エラーなし。op の導出フィールド (工程進行の FromState 等) は
' この関数が設定する。
'==============================================================================
Public Function ValidateOperation(ByVal op As c_LedgerOperation, _
                                  ByVal ctx As c_DialogContext, _
                                  ByRef errs As Collection, _
                                  ByRef warns As Collection) As Boolean
    Set errs = New Collection
    Set warns = New Collection

    ' --- 共通の形式チェック ---
    If op.Qty < 1 Then errs.Add "数量は 1 以上の整数で入力してください。"
    If Trim$(op.Recorder) = "" Then errs.Add "記録者を選択してください。"
    If op.ActionDate = 0 Then
        errs.Add "実施日を入力してください。"
    ElseIf op.ActionDate > Date Then
        warns.Add "実施日が未来日です (" & Format$(op.ActionDate, "yyyy/mm/dd") & ")。"
    End If

    ' --- 種別ごとの業務整合チェック ---
    Dim lt As c_LotInfo
    Dim it As c_ItemInfo

    Select Case op.Kind
        Case LedgerOpKind.opCarryIn
            op.EventKind = EV_CARRYIN
            op.LotId = ""
            op.FromState = ""
            If ctx.FindItem(op.ItemNo) Is Nothing Then
                errs.Add "品番が M_品番 にありません: " & op.ItemNo
            End If

        Case LedgerOpKind.opProgress
            Set lt = ctx.FindLot(op.LotId)
            If lt Is Nothing Then
                errs.Add "ロットが T_納期計画 にありません: " & op.LotId
            ElseIf lt.PlanStatus <> "有効" Then
                errs.Add "キャンセルされたロットには工程進行を記録できません" & _
                         " (例外操作で対応してください): " & op.LotId
            Else
                op.ItemNo = lt.ItemNo
                Set it = ctx.FindItem(lt.ItemNo)
                If it Is Nothing Then
                    errs.Add "ロットの品番が M_品番 にありません: " & lt.ItemNo
                Else
                    op.FromState = m_LedgerCore.DeriveFromState(it.Pattern, op.EventKind)
                    If op.FromState = "" Then
                        errs.Add "イベント「" & op.EventKind & "」は品番 " & it.ItemNo & _
                                 " の工程パターン " & it.Pattern & " では使用できません。"
                    End If
                End If
            End If

        Case LedgerOpKind.opDefect
            op.EventKind = EV_DEFECT
            Set lt = ctx.FindLot(op.LotId)
            If lt Is Nothing Then
                errs.Add "ロットが T_納期計画 にありません: " & op.LotId
            ElseIf lt.PlanStatus <> "有効" Then
                errs.Add "キャンセルされたロットには不良発生を記録できません: " & op.LotId
            Else
                op.ItemNo = lt.ItemNo
            End If
            If Trim$(op.FromState) = "" Then errs.Add "不良になった在庫の元状態を選択してください。"
            If op.Refill Then
                op.TargetLotId = op.LotId
                If Trim$(op.RefillFromState) = "" Then
                    errs.Add "補填元の状態を選択してください。"
                End If
                If op.RefillQty < 1 Then
                    errs.Add "補填数量は 1 以上で入力してください。"
                ElseIf op.RefillQty > op.Qty Then
                    ' 設計合意: 上限 = 不良数量に固定 (超過割当は例外操作の充当で行う)
                    errs.Add "補填数量 (" & op.RefillQty & ") は不良数量 (" & op.Qty & _
                             ") を超えられません。追加の割当は例外操作の「充当」で行ってください。"
                End If
            End If

        Case LedgerOpKind.opException
            Select Case op.EventKind
                Case EV_DISCARD
                    op.FromState = ST_DEFECT   ' 廃棄の元状態は常に不良
                    ValidateExceptionTarget op, ctx, errs, False
                Case EV_RETURN
                    If Trim$(op.FromState) = "" Then errs.Add "元状態を選択してください。"
                    ValidateExceptionTarget op, ctx, errs, False
                Case EV_SURPLUS
                    If Trim$(op.FromState) = "" Then errs.Add "元状態を選択してください。"
                    If Trim$(op.LotId) = "" Then
                        errs.Add "余剰化はロットの在庫を未割当へ戻す操作です。ロットを選択してください。"
                    Else
                        ValidateExceptionTarget op, ctx, errs, False
                    End If
                Case EV_ALLOCATE
                    If Trim$(op.FromState) = "" Then errs.Add "元状態を選択してください。"
                    ValidateExceptionTarget op, ctx, errs, False
                    Set lt = ctx.FindLot(op.TargetLotId)
                    If lt Is Nothing Then
                        errs.Add "充当先ロットが T_納期計画 にありません: " & op.TargetLotId
                    ElseIf lt.PlanStatus <> "有効" Then
                        errs.Add "キャンセルされたロットへは充当できません: " & op.TargetLotId
                    ElseIf op.TargetLotId = op.LotId Then
                        errs.Add "充当元と充当先が同じロットです。"
                    ElseIf op.ItemNo <> "" And lt.ItemNo <> op.ItemNo Then
                        errs.Add "充当元と充当先の品番が一致しません。"
                    End If
                Case Else
                    errs.Add "不明な例外操作です: " & op.EventKind
            End Select

        Case Else
            errs.Add "不明な操作種別です。"
    End Select

    ' --- 在庫負防止 (展開行のシミュレーション) ---
    If errs.Count = 0 Then
        Dim rows As Collection: Set rows = m_LedgerCore.ExpandOperation(op, ctx)
        m_LedgerCore.SimulateRows rows, ctx, errs
    End If

    ValidateOperation = (errs.Count = 0)
End Function

' 例外操作の対象 (ロット or 未割当) の存在チェック。ItemNo の補完も行う
Private Sub ValidateExceptionTarget(ByVal op As c_LedgerOperation, _
                                    ByVal ctx As c_DialogContext, _
                                    ByRef errs As Collection, _
                                    ByVal requireActiveLot As Boolean)
    If Trim$(op.LotId) <> "" Then
        Dim lt As c_LotInfo: Set lt = ctx.FindLot(op.LotId)
        If lt Is Nothing Then
            errs.Add "ロットが T_納期計画 にありません: " & op.LotId
        Else
            op.ItemNo = lt.ItemNo
            If requireActiveLot And lt.PlanStatus <> "有効" Then
                errs.Add "キャンセルされたロットは対象にできません: " & op.LotId
            End If
        End If
    Else
        If ctx.FindItem(op.ItemNo) Is Nothing Then
            errs.Add "未割当在庫の品番が M_品番 にありません: " & op.ItemNo
        End If
    End If
End Sub

'==============================================================================
' 操作 → 台帳行への展開
'
' 1 行 = Variant 配列 (0～9):
'   (0)実施日 (1)イベント種類 (2)品番 (3)ロットID (4)充当先ロットID
'   (5)元状態 (6)数量 (7)納期先 (8)記録者 (9)備考
' No 列は書込時に採番する。
'==============================================================================
Public Function ExpandOperation(ByVal op As c_LedgerOperation, _
                                ByVal ctx As c_DialogContext) As Collection
    Dim rows As New Collection
    Dim lt As c_LotInfo

    Select Case op.Kind
        Case LedgerOpKind.opCarryIn
            rows.Add MakeRow(op.ActionDate, EV_CARRYIN, op.ItemNo, "", "", "", _
                             op.Qty, "", op.Recorder, op.Note)

        Case LedgerOpKind.opProgress
            Dim dest As String
            If op.EventKind = EV_SHIP Then
                Set lt = ctx.FindLot(op.LotId)
                If Not lt Is Nothing Then dest = lt.Dest
            End If
            rows.Add MakeRow(op.ActionDate, op.EventKind, op.ItemNo, op.LotId, "", _
                             op.FromState, op.Qty, dest, op.Recorder, op.Note)

        Case LedgerOpKind.opDefect
            rows.Add MakeRow(op.ActionDate, EV_DEFECT, op.ItemNo, op.LotId, "", _
                             op.FromState, op.Qty, "", op.Recorder, op.Note)
            If op.Refill Then
                ' 補填 = 未割当 → 当該ロットへの充当 (状態維持)
                rows.Add MakeRow(op.ActionDate, EV_ALLOCATE, op.ItemNo, "", op.LotId, _
                                 op.RefillFromState, op.RefillQty, "", op.Recorder, _
                                 AppendNote(op.Note, "不良補填"))
            End If

        Case LedgerOpKind.opException
            rows.Add MakeRow(op.ActionDate, op.EventKind, op.ItemNo, op.LotId, _
                             op.TargetLotId, op.FromState, op.Qty, "", _
                             op.Recorder, op.Note)
    End Select

    Set ExpandOperation = rows
End Function

Private Function MakeRow(ByVal actionDate As Date, ByVal eventKind As String, _
                         ByVal itemNo As String, ByVal lotId As String, _
                         ByVal targetLotId As String, ByVal fromState As String, _
                         ByVal qty As Long, ByVal dest As String, _
                         ByVal recorder As String, ByVal note As String) As Variant
    MakeRow = Array(actionDate, eventKind, itemNo, lotId, targetLotId, fromState, _
                    qty, dest, recorder, note)
End Function

Private Function AppendNote(ByVal note As String, ByVal tag As String) As String
    If Trim$(note) = "" Then
        AppendNote = tag
    Else
        AppendNote = tag & " / " & note
    End If
End Function

'==============================================================================
' 在庫負防止のシミュレーション
'
' スナップショットの複製に展開行を順に適用し、消費が可用量 (負は 0 扱い) を
' 超えたらエラーを追加する。
'==============================================================================
Public Sub SimulateRows(ByVal rows As Collection, ByVal ctx As c_DialogContext, _
                        ByRef errs As Collection)
    Dim dict As Object: Set dict = ctx.Stock.CloneDict()

    Dim row As Variant
    For Each row In rows
        Dim eventKind As String, itemNo As String, lotId As String
        Dim targetLotId As String, fromState As String
        Dim qty As Long
        eventKind = row(1): itemNo = row(2): lotId = row(3)
        targetLotId = row(4): fromState = row(5): qty = row(6)

        ' "種別|ID|" まで組み立て、状態を後ろに連結して使う
        Dim bucket As String
        If lotId <> "" Then
            bucket = "L|" & lotId & "|"
        Else
            bucket = "U|" & itemNo & "|"
        End If

        Select Case eventKind
            Case EV_CARRYIN
                AddQty dict, bucket & ST_RAW, qty
            Case EV_SHIP, EV_DISCARD, EV_RETURN
                ConsumeQty dict, bucket & fromState, qty, DescribeBucket(lotId, itemNo), _
                           fromState, errs
            Case EV_DEFECT
                ConsumeQty dict, bucket & fromState, qty, DescribeBucket(lotId, itemNo), _
                           fromState, errs
                AddQty dict, bucket & ST_DEFECT, qty
            Case EV_SURPLUS
                ConsumeQty dict, bucket & fromState, qty, DescribeBucket(lotId, itemNo), _
                           fromState, errs
                AddQty dict, "U|" & itemNo & "|" & fromState, qty
            Case EV_ALLOCATE
                ConsumeQty dict, bucket & fromState, qty, DescribeBucket(lotId, itemNo), _
                           fromState, errs
                AddQty dict, "L|" & targetLotId & "|" & fromState, qty
            Case Else
                ' 工程イベント (開始/完了)
                Dim toState As String: toState = m_LedgerCore.EventToState(eventKind)
                ConsumeQty dict, bucket & fromState, qty, DescribeBucket(lotId, itemNo), _
                           fromState, errs
                If toState <> "" Then AddQty dict, bucket & toState, qty
        End Select
    Next
End Sub

Private Sub AddQty(ByVal dict As Object, ByVal key As String, ByVal qty As Long)
    If dict.Exists(key) Then
        dict(key) = dict(key) + qty
    Else
        dict(key) = qty
    End If
End Sub

Private Sub ConsumeQty(ByVal dict As Object, ByVal key As String, ByVal qty As Long, _
                       ByVal targetDesc As String, ByVal state As String, _
                       ByRef errs As Collection)
    Dim avail As Long
    If dict.Exists(key) Then avail = dict(key)
    If avail < 0 Then avail = 0    ' 既に負のバケットからは消費できない
    If qty > avail Then
        errs.Add targetDesc & " の「" & state & "」在庫が不足しています" & _
                 " (可用 " & avail & " に対して " & qty & ")。"
    End If
    If dict.Exists(key) Then
        dict(key) = dict(key) - qty
    Else
        dict(key) = -qty
    End If
End Sub

Private Function DescribeBucket(ByVal lotId As String, ByVal itemNo As String) As String
    If lotId <> "" Then
        DescribeBucket = "ロット " & lotId
    Else
        DescribeBucket = "未割当 (" & itemNo & ")"
    End If
End Function

'==============================================================================
' 書込・再計算・書込後チェック・保存
'==============================================================================

' 展開行を tbl_ledger へ追記し、再計算・負在庫チェックまで行う。
' 戻り値: 追記した No の一覧 (ロールバックした場合は空 Collection)
Public Function WriteOperation(ByVal wb As Workbook, ByVal rows As Collection, _
                               ByVal ctx As c_DialogContext) As Collection
    Dim addedNos As New Collection
    Dim ws As Worksheet: Set ws = wb.Worksheets(SH_LEDGER)
    Dim lo As ListObject: Set lo = ws.ListObjects("tbl_ledger")

    ' 書込前の負バケット (既知の不整合) を控えておく
    Dim negBefore As Collection: Set negBefore = ctx.Stock.NegativeBucketKeys()

    ws.Unprotect PROTECT_PWD

    Dim nextNo As Long: nextNo = MaxLedgerNo(lo) + 1

    Dim addedRows As New Collection
    Dim row As Variant
    For Each row In rows
        Dim lr As ListRow: Set lr = lo.ListRows.Add
        lr.Range(1).Value = nextNo
        lr.Range(2).Value = CDate(row(0))
        SetCell lr, 3, row(1)   ' イベント種類
        SetCell lr, 4, row(2)   ' 品番
        SetCell lr, 5, row(3)   ' ロットID
        SetCell lr, 6, row(4)   ' 充当先ロットID
        SetCell lr, 7, row(5)   ' 元状態
        lr.Range(8).Value = row(6)   ' 数量
        SetCell lr, 9, row(7)   ' 納期先
        SetCell lr, 10, row(8)  ' 記録者
        SetCell lr, 11, row(9)  ' 備考
        addedNos.Add nextNo
        addedRows.Add lr
        nextNo = nextNo + 1
    Next

    Application.CalculateFullRebuild

    ' 書込後の負在庫チェック (既知の負バケットは除外)
    Dim newNegatives As String: newNegatives = FindNewNegatives(wb, negBefore)
    If newNegatives <> "" Then
        Dim answer As VbMsgBoxResult
        answer = MsgBox("書込の結果、V_在庫内訳に新たなマイナスが発生しました:" & vbCrLf & _
                        newNegatives & vbCrLf & vbCrLf & _
                        "他の入力と競合した可能性があります。今回の書込を取り消しますか?", _
                        vbYesNo + vbExclamation, "負在庫の検出")
        If answer = vbYes Then
            Dim i As Long
            For i = addedRows.Count To 1 Step -1
                addedRows(i).Delete
            Next
            Application.CalculateFullRebuild
            Set addedNos = New Collection
        End If
    End If

    ws.Protect PROTECT_PWD
    Set WriteOperation = addedNos
End Function

' 空文字はセルに書かず空セルのまま残す (既存サンプル行と SUMIFS の挙動を揃える)
Private Sub SetCell(ByVal lr As ListRow, ByVal colIdx As Long, ByVal value As Variant)
    If Trim$(CStr(value & "")) <> "" Then lr.Range(colIdx).Value = value
End Sub

Private Function MaxLedgerNo(ByVal lo As ListObject) As Long
    Dim maxNo As Long
    Dim r As ListRow
    For Each r In lo.ListRows
        If IsNumeric(r.Range(1).Value) Then
            If CLng(r.Range(1).Value) > maxNo Then maxNo = CLng(r.Range(1).Value)
        End If
    Next
    MaxLedgerNo = maxNo
End Function

' 書込後の V_在庫内訳を走査し、「書込前には負でなかった」負バケットを列挙する
Private Function FindNewNegatives(ByVal wb As Workbook, _
                                  ByVal negBefore As Collection) As String
    Dim afterCtx As c_DialogContext: Set afterCtx = m_LedgerCore.BuildDialogContext(wb)

    Dim befDict As Object: Set befDict = CreateObject("Scripting.Dictionary")
    Dim v As Variant
    For Each v In negBefore
        befDict(CStr(v)) = True
    Next

    Dim result As String
    Dim afterDict As Object: Set afterDict = afterCtx.Stock.CloneDict()
    For Each v In afterDict.Keys
        If afterDict(v) < 0 And Not befDict.Exists(CStr(v)) Then
            result = result & "  " & CStr(v) & " = " & afterDict(v) & vbCrLf
        End If
    Next
    FindNewNegatives = result
End Function

'==============================================================================
' シート保護 / 解除 / 保存 (設計合意: 層2+層3)
'==============================================================================
Public Sub ProtectLedgerBook(ByVal wb As Workbook)
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        ws.Protect Password:=PROTECT_PWD
    Next
    wb.Protect Password:=PROTECT_PWD, Structure:=True
End Sub

Public Sub UnprotectLedgerBook(ByVal wb As Workbook)
    Dim ws As Worksheet
    On Error GoTo Fail
    wb.Unprotect Password:=PROTECT_PWD
    For Each ws In wb.Worksheets
        ws.Unprotect Password:=PROTECT_PWD
    Next
    Exit Sub
Fail:
    Err.Raise vbObjectError + 10, , _
        "保護を解除できませんでした (パスワード不一致の可能性): " & Err.Description
End Sub

' 読み取り専用推奨フラグ付きで上書き保存する
Public Sub SaveLedgerBook(ByVal wb As Workbook)
    Application.DisplayAlerts = False
    wb.SaveAs Filename:=wb.FullName, FileFormat:=wb.FileFormat, _
              ReadOnlyRecommended:=True
    Application.DisplayAlerts = True
End Sub
