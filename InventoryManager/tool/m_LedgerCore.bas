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
    opException = 4     ' 例外操作 (廃棄/倉庫戻し/余剰化/充当)
End Enum

' シート保護のパスワード (事故防止用。変更する場合はここを書き換えて再ビルド)
Public Const PROTECT_PWD As String = "ledger"

' シート名
Public Const SH_LEDGER As String = "T_台帳"
Public Const SH_PLAN As String = "T_納期計画"
Public Const SH_MASTER As String = "M_品番"
Public Const SH_LIST As String = "M_リスト"
Public Const SH_STOCK As String = "V_在庫内訳"
Public Const SH_CHECK As String = "V_充足確認"

' 列名 (ブックのヘッダーと一致させること。読み書きはすべて列名で行い、列の位置には依存しない)
' M_品番 (tbl_master)
Public Const MCOL_LOCALID As String = "localId"
Public Const MCOL_ITEMNO As String = "品番"
Public Const MCOL_ITEMNAME As String = "品名"
Public Const MCOL_PATTERN As String = "工程パターン"
' T_納期計画 (tbl_plan)
Public Const PCOL_LOTID As String = "ロットID"
Public Const PCOL_LOCALID As String = "localId"
Public Const PCOL_ITEMNO As String = "品番"
Public Const PCOL_SHIPWEEK As String = "出荷週"
Public Const PCOL_DEST As String = "納期先"
Public Const PCOL_REQUIRED As String = "必要数"
Public Const PCOL_STATUS As String = "計画状態"
' T_台帳 (tbl_ledger)
Public Const LCOL_NO As String = "No"
Public Const LCOL_DATE As String = "実施日"
Public Const LCOL_EVENT As String = "イベント種類"
Public Const LCOL_LOCALID As String = "localId"
Public Const LCOL_ITEMNO As String = "品番"
Public Const LCOL_LOTID As String = "ロットID"
Public Const LCOL_TARGETLOT As String = "充当先ロットID"
Public Const LCOL_FROMSTATE As String = "元状態"
Public Const LCOL_QTY As String = "数量"
Public Const LCOL_DEST As String = "納期先"
Public Const LCOL_RECORDER As String = "記録者"
Public Const LCOL_NOTE As String = "備考"
' V_在庫内訳 (両セクションのヘッダー行)
Public Const VCOL_KIND As String = "区分"
Public Const VCOL_LOTID As String = "ロットID"
Public Const VCOL_LOCALID As String = "localId"
Public Const VCOL_TOTAL As String = "合計"
Public Const VKIND_LOT As String = "ロット"
Public Const VKIND_UNALLOC As String = "未割当"
' M_リスト
Public Const LISTCOL_STATE As String = "状態"
Public Const LISTCOL_RECORDER As String = "記録者"

' イベント種類 (M_リスト A列と一致させること)
Public Const EV_CARRYIN As String = "搬入"
Public Const EV_SHIP As String = "発送"
Public Const EV_DEFECT As String = "不良発生"
Public Const EV_DISCARD As String = "廃棄"
Public Const EV_RETURN As String = "倉庫戻し"   ' 元の倉庫へ戻す (抽象名。2026-07-22 改名)
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
' 列はすべて列名 (定数) で引くため、列の追加・並べ替えに影響されない。
'==============================================================================
Public Function BuildDialogContext(ByVal wb As Workbook) As c_DialogContext
    Dim ctx As New c_DialogContext

    ' --- M_品番 ---
    Dim lo As ListObject: Set lo = wb.Worksheets(SH_MASTER).ListObjects("tbl_master")
    Dim mLocal As Long: mLocal = TblCol(lo, MCOL_LOCALID)
    Dim mItemNo As Long: mItemNo = TblCol(lo, MCOL_ITEMNO)
    Dim mName As Long: mName = TblCol(lo, MCOL_ITEMNAME)
    Dim mPattern As Long: mPattern = TblCol(lo, MCOL_PATTERN)
    Dim r As ListRow
    For Each r In lo.ListRows
        If Trim$(CStr(r.Range(mLocal).Value & "")) <> "" Then
            Dim it As c_ItemInfo: Set it = New c_ItemInfo
            it.LocalId = CStr(r.Range(mLocal).Value)
            it.ItemNo = CStr(r.Range(mItemNo).Value & "")
            it.ItemName = CStr(r.Range(mName).Value & "")
            it.Pattern = CLng(Val(r.Range(mPattern).Value & ""))
            ctx.Items.Add it
        End If
    Next

    ' --- T_納期計画 (キャンセル含む全行。UI 側で用途に応じて絞る) ---
    Set lo = wb.Worksheets(SH_PLAN).ListObjects("tbl_plan")
    Dim pLot As Long: pLot = TblCol(lo, PCOL_LOTID)
    Dim pLocal As Long: pLocal = TblCol(lo, PCOL_LOCALID)
    Dim pItemNo As Long: pItemNo = TblCol(lo, PCOL_ITEMNO)
    Dim pWeek As Long: pWeek = TblCol(lo, PCOL_SHIPWEEK)
    Dim pDest As Long: pDest = TblCol(lo, PCOL_DEST)
    Dim pReq As Long: pReq = TblCol(lo, PCOL_REQUIRED)
    Dim pStatus As Long: pStatus = TblCol(lo, PCOL_STATUS)
    For Each r In lo.ListRows
        If Trim$(CStr(r.Range(pLot).Value & "")) <> "" Then
            Dim lt As c_LotInfo: Set lt = New c_LotInfo
            lt.LotId = CStr(r.Range(pLot).Value)
            lt.LocalId = CStr(r.Range(pLocal).Value & "")
            lt.ItemNo = CStr(r.Range(pItemNo).Value & "")
            lt.ShipWeek = CDate(r.Range(pWeek).Value)
            lt.Dest = CStr(r.Range(pDest).Value & "")
            lt.Required = CLng(Val(r.Range(pReq).Value & ""))
            lt.PlanStatus = CStr(r.Range(pStatus).Value & "")
            ctx.Lots.Add lt
        End If
    Next

    ' --- M_リスト (状態リスト・記録者。1 行目のヘッダーで列を引く) ---
    Dim ws As Worksheet: Set ws = wb.Worksheets(SH_LIST)
    Dim listMap As Object: Set listMap = HeaderMap(ws, 1)
    Dim rowNo As Long
    Dim stCol As Long: stCol = MapCol(listMap, ws, 1, LISTCOL_STATE)
    rowNo = 2
    Do While CStr(ws.Cells(rowNo, stCol).Value & "") <> ""
        ctx.States.Add CStr(ws.Cells(rowNo, stCol).Value)
        rowNo = rowNo + 1
    Loop
    Dim recCol As Long: recCol = MapCol(listMap, ws, 1, LISTCOL_RECORDER)
    rowNo = 2
    Do While CStr(ws.Cells(rowNo, recCol).Value & "") <> ""
        ctx.Recorders.Add CStr(ws.Cells(rowNo, recCol).Value)
        rowNo = rowNo + 1
    Loop

    ' --- V_在庫内訳 → c_StockSnapshot (各セクションのヘッダー行で列を引く) ---
    Set ws = wb.Worksheets(SH_STOCK)
    Dim snap As New c_StockSnapshot
    snap.Init ctx.States

    Dim kindCol As Long, lotCol As Long, localCol As Long
    Dim stateCols() As Long
    ResolveStockColumns ws, 1, ctx.States, kindCol, lotCol, localCol, stateCols

    Dim kindCell As String
    Dim bottom As Long: bottom = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    For rowNo = 2 To bottom
        kindCell = CStr(ws.Cells(rowNo, kindCol).Value & "")
        If kindCell = VCOL_KIND Then
            ' 2 つ目 (未割当) のセクションのヘッダー行 → 列を引き直す
            ResolveStockColumns ws, rowNo, ctx.States, kindCol, lotCol, localCol, stateCols
        ElseIf kindCell = VKIND_LOT Or kindCell = VKIND_UNALLOC Then
            Dim stIdx As Long
            Dim q As Long
            For stIdx = 1 To ctx.States.Count
                q = CLng(Val(ws.Cells(rowNo, stateCols(stIdx)).Value & ""))
                If q <> 0 Then
                    If kindCell = VKIND_LOT Then
                        snap.SetLotQty CStr(ws.Cells(rowNo, lotCol).Value & ""), _
                                       ctx.States(stIdx), q
                    Else
                        snap.SetUnallocQty CStr(ws.Cells(rowNo, localCol).Value & ""), _
                                           ctx.States(stIdx), q
                    End If
                End If
            Next
        End If
    Next
    Set ctx.Stock = snap

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
' 列名 → 列位置の解決 (列の追加・並べ替えに耐えるため、位置は名前から都度引く)
'==============================================================================

' テーブル列の位置を列名で引く (無ければ分かりやすいエラー)
Private Function TblCol(ByVal lo As ListObject, ByVal colName As String) As Long
    On Error GoTo Fail
    TblCol = lo.ListColumns(colName).Index
    Exit Function
Fail:
    Err.Raise vbObjectError + 20, , _
        "テーブル " & lo.Name & " に列「" & colName & "」がありません。" & _
        "列名を変更した場合は m_LedgerCore の列名定数も合わせてください。"
End Function

' シートの指定行をヘッダーとして 列名→列番号 の辞書を作る (重複時は左を優先)
Private Function HeaderMap(ByVal ws As Worksheet, ByVal headerRow As Long) As Object
    Dim map As Object: Set map = CreateObject("Scripting.Dictionary")
    Dim lastCol As Long: lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    Dim col As Long
    Dim h As String
    For col = 1 To lastCol
        h = Trim$(CStr(ws.Cells(headerRow, col).Value & ""))
        If h <> "" And Not map.Exists(h) Then map(h) = col
    Next
    Set HeaderMap = map
End Function

Private Function MapCol(ByVal map As Object, ByVal ws As Worksheet, _
                        ByVal headerRow As Long, ByVal colName As String) As Long
    If Not map.Exists(colName) Then
        Err.Raise vbObjectError + 21, , _
            ws.Name & " の " & headerRow & " 行目に列「" & colName & "」が見つかりません。"
    End If
    MapCol = map(colName)
End Function

' V_在庫内訳のセクションヘッダー行から 区分/ロットID/localId/各状態 の列位置を引く
' (未割当セクションにはロットID列が無いため lotCol のみ任意)
Private Sub ResolveStockColumns(ByVal ws As Worksheet, ByVal headerRow As Long, _
                                ByVal states As Collection, ByRef kindCol As Long, _
                                ByRef lotCol As Long, ByRef localCol As Long, _
                                ByRef stateCols() As Long)
    Dim map As Object: Set map = HeaderMap(ws, headerRow)
    kindCol = MapCol(map, ws, headerRow, VCOL_KIND)
    lotCol = 0
    If map.Exists(VCOL_LOTID) Then lotCol = map(VCOL_LOTID)
    localCol = MapCol(map, ws, headerRow, VCOL_LOCALID)
    ReDim stateCols(1 To states.Count)
    Dim i As Long
    For i = 1 To states.Count
        stateCols(i) = MapCol(map, ws, headerRow, CStr(states(i)))
    Next
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
            Set it = ctx.FindItemByLocalId(op.LocalId)
            If it Is Nothing Then
                errs.Add "localId が M_品番 にありません: " & op.LocalId
            Else
                op.ItemNo = it.ItemNo
            End If

        Case LedgerOpKind.opProgress
            Set lt = ctx.FindLot(op.LotId)
            If lt Is Nothing Then
                errs.Add "ロットが T_納期計画 にありません: " & op.LotId
            ElseIf lt.PlanStatus <> "有効" Then
                errs.Add "キャンセルされたロットには工程進行を記録できません" & _
                         " (例外操作で対応してください): " & op.LotId
            Else
                op.LocalId = lt.LocalId
                op.ItemNo = lt.ItemNo
                Set it = ctx.FindItemByLocalId(lt.LocalId)
                If it Is Nothing Then
                    errs.Add "ロットの localId が M_品番 にありません: " & lt.LocalId
                Else
                    op.FromState = m_LedgerCore.DeriveFromState(it.Pattern, op.EventKind)
                    If op.FromState = "" Then
                        errs.Add "イベント「" & op.EventKind & "」は " & it.LocalId & _
                                 " (品番 " & it.ItemNo & ") の工程パターン " & it.Pattern & _
                                 " では使用できません。"
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
                op.LocalId = lt.LocalId
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
                    ElseIf op.LocalId <> "" And lt.LocalId <> op.LocalId Then
                        errs.Add "充当元と充当先の品番 (localId) が一致しません。"
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

' 例外操作の対象 (ロット or 未割当) の存在チェック。LocalId/ItemNo の補完も行う
Private Sub ValidateExceptionTarget(ByVal op As c_LedgerOperation, _
                                    ByVal ctx As c_DialogContext, _
                                    ByRef errs As Collection, _
                                    ByVal requireActiveLot As Boolean)
    If Trim$(op.LotId) <> "" Then
        Dim lt As c_LotInfo: Set lt = ctx.FindLot(op.LotId)
        If lt Is Nothing Then
            errs.Add "ロットが T_納期計画 にありません: " & op.LotId
        Else
            op.LocalId = lt.LocalId
            op.ItemNo = lt.ItemNo
            If requireActiveLot And lt.PlanStatus <> "有効" Then
                errs.Add "キャンセルされたロットは対象にできません: " & op.LotId
            End If
        End If
    Else
        Dim it As c_ItemInfo: Set it = ctx.FindItemByLocalId(op.LocalId)
        If it Is Nothing Then
            errs.Add "未割当在庫の localId が M_品番 にありません: " & op.LocalId
        Else
            op.ItemNo = it.ItemNo
        End If
    End If
End Sub

'==============================================================================
' 操作 → 台帳行への展開
'
' 1 行 = Variant 配列 (0～10):
'   (0)実施日 (1)イベント種類 (2)localId (3)品番 (4)ロットID (5)充当先ロットID
'   (6)元状態 (7)数量 (8)納期先 (9)記録者 (10)備考
' No 列は書込時に採番する。
'==============================================================================
Public Function ExpandOperation(ByVal op As c_LedgerOperation, _
                                ByVal ctx As c_DialogContext) As Collection
    Dim rows As New Collection
    Dim lt As c_LotInfo

    Select Case op.Kind
        Case LedgerOpKind.opCarryIn
            rows.Add MakeRow(op.ActionDate, EV_CARRYIN, op.LocalId, op.ItemNo, "", "", "", _
                             op.Qty, "", op.Recorder, op.Note)

        Case LedgerOpKind.opProgress
            Dim dest As String
            If op.EventKind = EV_SHIP Then
                Set lt = ctx.FindLot(op.LotId)
                If Not lt Is Nothing Then dest = lt.Dest
            End If
            rows.Add MakeRow(op.ActionDate, op.EventKind, op.LocalId, op.ItemNo, op.LotId, "", _
                             op.FromState, op.Qty, dest, op.Recorder, op.Note)

        Case LedgerOpKind.opDefect
            rows.Add MakeRow(op.ActionDate, EV_DEFECT, op.LocalId, op.ItemNo, op.LotId, "", _
                             op.FromState, op.Qty, "", op.Recorder, op.Note)
            If op.Refill Then
                ' 補填 = 未割当 → 当該ロットへの充当 (状態維持)
                rows.Add MakeRow(op.ActionDate, EV_ALLOCATE, op.LocalId, op.ItemNo, "", op.LotId, _
                                 op.RefillFromState, op.RefillQty, "", op.Recorder, _
                                 AppendNote(op.Note, "不良補填"))
            End If

        Case LedgerOpKind.opException
            rows.Add MakeRow(op.ActionDate, op.EventKind, op.LocalId, op.ItemNo, op.LotId, _
                             op.TargetLotId, op.FromState, op.Qty, "", _
                             op.Recorder, op.Note)
    End Select

    Set ExpandOperation = rows
End Function

Private Function MakeRow(ByVal actionDate As Date, ByVal eventKind As String, _
                         ByVal localId As String, ByVal itemNo As String, _
                         ByVal lotId As String, ByVal targetLotId As String, _
                         ByVal fromState As String, ByVal qty As Long, _
                         ByVal dest As String, ByVal recorder As String, _
                         ByVal note As String) As Variant
    MakeRow = Array(actionDate, eventKind, localId, itemNo, lotId, targetLotId, fromState, _
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
        Dim eventKind As String, localId As String, lotId As String
        Dim targetLotId As String, fromState As String
        Dim qty As Long
        eventKind = row(1): localId = row(2): lotId = row(4)
        targetLotId = row(5): fromState = row(6): qty = row(7)

        ' "種別|ID|" まで組み立て、状態を後ろに連結して使う
        Dim bucket As String
        If lotId <> "" Then
            bucket = "L|" & lotId & "|"
        Else
            bucket = "U|" & localId & "|"
        End If

        Select Case eventKind
            Case EV_CARRYIN
                AddQty dict, bucket & ST_RAW, qty
            Case EV_SHIP, EV_DISCARD, EV_RETURN
                ConsumeQty dict, bucket & fromState, qty, DescribeBucket(lotId, localId), _
                           fromState, errs
            Case EV_DEFECT
                ConsumeQty dict, bucket & fromState, qty, DescribeBucket(lotId, localId), _
                           fromState, errs
                AddQty dict, bucket & ST_DEFECT, qty
            Case EV_SURPLUS
                ConsumeQty dict, bucket & fromState, qty, DescribeBucket(lotId, localId), _
                           fromState, errs
                AddQty dict, "U|" & localId & "|" & fromState, qty
            Case EV_ALLOCATE
                ConsumeQty dict, bucket & fromState, qty, DescribeBucket(lotId, localId), _
                           fromState, errs
                AddQty dict, "L|" & targetLotId & "|" & fromState, qty
            Case Else
                ' 工程イベント (開始/完了)
                Dim toState As String: toState = m_LedgerCore.EventToState(eventKind)
                ConsumeQty dict, bucket & fromState, qty, DescribeBucket(lotId, localId), _
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

Private Function DescribeBucket(ByVal lotId As String, ByVal localId As String) As String
    If lotId <> "" Then
        DescribeBucket = "ロット " & lotId
    Else
        DescribeBucket = "未割当 (" & localId & ")"
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

    ' 列位置は列名で解決する (列の追加・並べ替えに耐える)
    ' (変数名に cDate を使うと VBA の変換関数 CDate と衝突するため cActDate)
    Dim cNo As Long: cNo = TblCol(lo, LCOL_NO)
    Dim cActDate As Long: cActDate = TblCol(lo, LCOL_DATE)
    Dim cEvent As Long: cEvent = TblCol(lo, LCOL_EVENT)
    Dim cLocal As Long: cLocal = TblCol(lo, LCOL_LOCALID)
    Dim cItemNo As Long: cItemNo = TblCol(lo, LCOL_ITEMNO)
    Dim cLot As Long: cLot = TblCol(lo, LCOL_LOTID)
    Dim cTarget As Long: cTarget = TblCol(lo, LCOL_TARGETLOT)
    Dim cFrom As Long: cFrom = TblCol(lo, LCOL_FROMSTATE)
    Dim cQty As Long: cQty = TblCol(lo, LCOL_QTY)
    Dim cDest As Long: cDest = TblCol(lo, LCOL_DEST)
    Dim cRec As Long: cRec = TblCol(lo, LCOL_RECORDER)
    Dim cNote As Long: cNote = TblCol(lo, LCOL_NOTE)

    Dim addedRows As New Collection
    Dim row As Variant
    For Each row In rows
        Dim lr As ListRow: Set lr = lo.ListRows.Add
        lr.Range(cNo).Value = nextNo
        lr.Range(cActDate).Value = CDate(row(0))
        SetCell lr, cEvent, row(1)
        SetCell lr, cLocal, row(2)
        SetCell lr, cItemNo, row(3)
        SetCell lr, cLot, row(4)
        SetCell lr, cTarget, row(5)
        SetCell lr, cFrom, row(6)
        lr.Range(cQty).Value = row(7)
        SetCell lr, cDest, row(8)
        SetCell lr, cRec, row(9)
        SetCell lr, cNote, row(10)
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

    ws.Protect Password:=PROTECT_PWD, AllowFiltering:=True
    Set WriteOperation = addedNos
End Function

' 空文字はセルに書かず空セルのまま残す (既存サンプル行と SUMIFS の挙動を揃える)
Private Sub SetCell(ByVal lr As ListRow, ByVal colIdx As Long, ByVal value As Variant)
    If Trim$(CStr(value & "")) <> "" Then lr.Range(colIdx).Value = value
End Sub

Private Function MaxLedgerNo(ByVal lo As ListObject) As Long
    Dim maxNo As Long
    Dim noCol As Long: noCol = TblCol(lo, LCOL_NO)
    Dim r As ListRow
    For Each r In lo.ListRows
        If IsNumeric(r.Range(noCol).Value) Then
            If CLng(r.Range(noCol).Value) > maxNo Then maxNo = CLng(r.Range(noCol).Value)
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
'
' 保護は AllowFiltering 付き (V_ シートのオートフィルターを保護中も使えるように)。
' 保護前に V_ シートの列幅を自動調整する。
'==============================================================================
Public Sub ProtectLedgerBook(ByVal wb As Workbook)
    m_LedgerCore.AutoFitViewSheets wb
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        ws.Protect Password:=PROTECT_PWD, AllowFiltering:=True
    Next
    wb.Protect Password:=PROTECT_PWD, Structure:=True
End Sub

' V_ シートの列幅を自動調整する (非表示の作業列は触らない)
Public Sub AutoFitViewSheets(ByVal wb As Workbook)
    Dim sheetName As Variant
    For Each sheetName In Array(SH_CHECK, SH_STOCK)
        Dim ws As Worksheet: Set ws = wb.Worksheets(CStr(sheetName))
        ws.Unprotect PROTECT_PWD
        Dim col As Range
        For Each col In ws.UsedRange.Columns
            If Not col.EntireColumn.Hidden Then col.EntireColumn.AutoFit
        Next
    Next
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
