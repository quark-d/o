VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_PsLedgerEventDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' c_ILedgerEventDialog 実装: PowerShell (WPF/XAML) 版 (DPI 問題時のフォールバック)
'
' Show-LedgerDialog.ps1 を同期実行し、UTF-8 (BOM 付き) の JSON ファイルで
' データを受け渡す。JSON の生成・解析はこのクラスと m_JsonLite に閉じる。
' スキーマは README.md の「JSON スキーマ」参照 (schemaVersion 1.2)。
'==============================================================================

Implements c_ILedgerEventDialog

Private Const SCHEMA_VERSION As String = "1.2"

' Show-LedgerDialog.ps1 のパス (既定値は Class_Initialize で設定。差し替え可)
Public ScriptPath As String
' 受け渡し用 JSON ファイルのパス (既定は %TEMP%。同時起動が衝突する場合は差し替える)
Public InputPath As String
Public OutputPath As String

Private Sub Class_Initialize()
    ScriptPath = ThisWorkbook.path & "\Show-LedgerDialog.ps1"
    InputPath = Environ$("TEMP") & "\ledger_dialog_in.json"
    OutputPath = Environ$("TEMP") & "\ledger_dialog_out.json"
End Sub

'------------------------------------------------------------------------------
' c_ILedgerEventDialog
'------------------------------------------------------------------------------
Private Function c_ILedgerEventDialog_ShowDialog(ByVal kind As LedgerOpKind, _
        ByVal ctx As c_DialogContext) As c_LedgerDialogResult
    Dim result As New c_LedgerDialogResult
    On Error GoTo Fail

    ' 古い出力ファイルが残っていれば削除 (誤読防止)
    On Error Resume Next
    Kill OutputPath
    On Error GoTo Fail

    m_JsonLite.WriteTextUtf8Bom InputPath, BuildInputJson(kind, ctx)

    Dim cmd As String
    cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -STA" & _
          " -File """ & ScriptPath & """" & _
          " -InputPath """ & InputPath & """" & _
          " -OutputPath """ & OutputPath & """"

    Dim wsh As Object: Set wsh = CreateObject("WScript.Shell")
    wsh.Run cmd, 0, True   ' 0=コンソール非表示, True=終了まで待機

    If Dir(OutputPath) = "" Then
        result.Init "error", Nothing, "出力ファイルが作成されませんでした: " & OutputPath
    Else
        ParseOutputJson m_JsonLite.ReadTextUtf8(OutputPath), result
    End If

    Set c_ILedgerEventDialog_ShowDialog = result
    Exit Function

Fail:
    result.Init "error", Nothing, Err.Description
    Set c_ILedgerEventDialog_ShowDialog = result
End Function

'------------------------------------------------------------------------------
' 入力 JSON の生成 (選択肢スナップショット)
'------------------------------------------------------------------------------
Private Function BuildInputJson(ByVal kind As LedgerOpKind, _
                                ByVal ctx As c_DialogContext) As String
    Dim sb As String
    sb = "{" & vbCrLf
    sb = sb & "  ""schemaVersion"": """ & SCHEMA_VERSION & """," & vbCrLf
    sb = sb & "  ""dialog"": """ & DialogKey(kind) & """," & vbCrLf
    sb = sb & "  ""defaults"": { ""actionDate"": """ & _
         Format$(ctx.DefaultActionDate, "yyyy-mm-dd") & _
         """, ""recorder"": """ & m_JsonLite.JsonEscape(ctx.DefaultRecorder) & """ }," & vbCrLf
    sb = sb & "  ""labels"": { ""warehouseA"": """ & m_JsonLite.JsonEscape(ctx.WarehouseA) & _
         """, ""warehouseB"": """ & m_JsonLite.JsonEscape(ctx.WarehouseB) & """ }," & vbCrLf
    sb = sb & "  ""recorders"": [" & JoinStringArray(ctx.Recorders) & "]," & vbCrLf
    sb = sb & "  ""states"": [" & JoinStringArray(ctx.States) & "]," & vbCrLf
    sb = sb & BuildItemsJson(ctx) & "," & vbCrLf
    sb = sb & BuildLotsJson(ctx) & "," & vbCrLf
    sb = sb & BuildStockJson(ctx) & vbCrLf
    sb = sb & "}"
    BuildInputJson = sb
End Function

Private Function DialogKey(ByVal kind As LedgerOpKind) As String
    Select Case kind
        Case LedgerOpKind.opCarryIn:   DialogKey = "carryIn"
        Case LedgerOpKind.opProgress:  DialogKey = "progress"
        Case LedgerOpKind.opDefect:    DialogKey = "defect"
        Case LedgerOpKind.opException: DialogKey = "exception"
        Case Else: Err.Raise vbObjectError + 1, , "不明なダイアログ種別です: " & kind
    End Select
End Function

Private Function JoinStringArray(ByVal col As Collection) As String
    Dim sb As String
    Dim v As Variant
    For Each v In col
        If sb <> "" Then sb = sb & ", "
        sb = sb & """" & m_JsonLite.JsonEscape(CStr(v)) & """"
    Next
    JoinStringArray = sb
End Function

Private Function BuildItemsJson(ByVal ctx As c_DialogContext) As String
    Dim sb As String: sb = "  ""items"": ["
    Dim it As c_ItemInfo
    For Each it In ctx.Items
        If Right$(sb, 1) = "}" Then sb = sb & ","
        sb = sb & vbCrLf & "    { ""localId"": """ & m_JsonLite.JsonEscape(it.LocalId) & _
             """, ""itemNo"": """ & m_JsonLite.JsonEscape(it.ItemNo) & _
             """, ""itemName"": """ & m_JsonLite.JsonEscape(it.ItemName) & _
             """, ""pattern"": " & it.Pattern & " }"
    Next
    BuildItemsJson = sb & vbCrLf & "  ]"
End Function

Private Function BuildLotsJson(ByVal ctx As c_DialogContext) As String
    Dim sb As String: sb = "  ""lots"": ["
    Dim lt As c_LotInfo
    For Each lt In ctx.Lots
        If Right$(sb, 1) = "}" Then sb = sb & ","
        sb = sb & vbCrLf & "    { ""lotId"": """ & m_JsonLite.JsonEscape(lt.LotId) & _
             """, ""localId"": """ & m_JsonLite.JsonEscape(lt.LocalId) & _
             """, ""itemNo"": """ & m_JsonLite.JsonEscape(lt.ItemNo) & _
             """, ""shipWeek"": """ & Format$(lt.ShipWeek, "yyyy-mm-dd") & _
             """, ""dest"": """ & m_JsonLite.JsonEscape(lt.Dest) & _
             """, ""required"": " & lt.Required & _
             ", ""planStatus"": """ & m_JsonLite.JsonEscape(lt.PlanStatus) & """ }"
    Next
    BuildLotsJson = sb & vbCrLf & "  ]"
End Function

Private Function BuildStockJson(ByVal ctx As c_DialogContext) As String
    Dim sb As String
    Dim entry As String
    Dim lt As c_LotInfo
    Dim it As c_ItemInfo

    sb = "  ""stock"": {" & vbCrLf & "    ""lots"": ["
    For Each lt In ctx.Lots
        entry = ByStateJson(ctx, "L", lt.LotId)
        If entry <> "" Then
            If Right$(sb, 1) = "}" Then sb = sb & ","
            sb = sb & vbCrLf & "      { ""lotId"": """ & m_JsonLite.JsonEscape(lt.LotId) & _
                 """, ""byState"": { " & entry & " } }"
        End If
    Next
    sb = sb & vbCrLf & "    ]," & vbCrLf & "    ""unallocated"": ["
    For Each it In ctx.Items
        entry = ByStateJson(ctx, "U", it.LocalId)
        If entry <> "" Then
            If Right$(sb, 1) = "}" Then sb = sb & ","
            sb = sb & vbCrLf & "      { ""localId"": """ & m_JsonLite.JsonEscape(it.LocalId) & _
                 """, ""byState"": { " & entry & " } }"
        End If
    Next
    BuildStockJson = sb & vbCrLf & "    ]" & vbCrLf & "  }"
End Function

' 非ゼロ状態のみの "状態": 数量 並び。全ゼロなら空文字
Private Function ByStateJson(ByVal ctx As c_DialogContext, ByVal bucket As String, _
                             ByVal id As String) As String
    Dim sb As String
    Dim st As Variant
    Dim q As Long
    For Each st In ctx.States
        q = ctx.Stock.RawQty(bucket, id, CStr(st))
        If q <> 0 Then
            If sb <> "" Then sb = sb & ", "
            sb = sb & """" & m_JsonLite.JsonEscape(CStr(st)) & """: " & q
        End If
    Next
    ByStateJson = sb
End Function

'------------------------------------------------------------------------------
' 出力 JSON の解析
'------------------------------------------------------------------------------
Private Sub ParseOutputJson(ByVal json As String, ByRef result As c_LedgerDialogResult)
    Dim status As String: status = m_JsonLite.JsonExtractString(json, "status")

    Select Case status
        Case "ok"
            Dim op As New c_LedgerOperation
            op.Kind = KindFromKey(m_JsonLite.JsonExtractString(json, "kind"))
            op.EventKind = m_JsonLite.JsonExtractString(json, "eventKind")
            op.LocalId = m_JsonLite.JsonExtractString(json, "localId")
            op.ItemNo = m_JsonLite.JsonExtractString(json, "itemNo")
            op.LotId = m_JsonLite.JsonExtractString(json, "lotId")
            op.FromState = m_JsonLite.JsonExtractString(json, "fromState")
            op.Qty = m_JsonLite.JsonExtractLong(json, "qty")
            op.ActionDate = ParseIsoDate(m_JsonLite.JsonExtractString(json, "actionDate"))
            op.Recorder = m_JsonLite.JsonExtractString(json, "recorder")
            op.Note = m_JsonLite.JsonExtractString(json, "note")
            op.Refill = m_JsonLite.JsonExtractBool(json, "refill")
            op.RefillFromState = m_JsonLite.JsonExtractString(json, "refillFromState")
            op.RefillQty = m_JsonLite.JsonExtractLong(json, "refillQty")
            op.TargetLotId = m_JsonLite.JsonExtractString(json, "targetLotId")
            result.Init "ok", op, ""
        Case "cancel"
            result.Init "cancel", Nothing, ""
        Case Else
            result.Init "error", Nothing, m_JsonLite.JsonExtractString(json, "message")
    End Select
End Sub

Private Function KindFromKey(ByVal key As String) As LedgerOpKind
    Select Case key
        Case "carryIn":   KindFromKey = LedgerOpKind.opCarryIn
        Case "progress":  KindFromKey = LedgerOpKind.opProgress
        Case "defect":    KindFromKey = LedgerOpKind.opDefect
        Case "exception": KindFromKey = LedgerOpKind.opException
        Case Else: Err.Raise vbObjectError + 2, , "出力 JSON の kind が不正です: " & key
    End Select
End Function

Private Function ParseIsoDate(ByVal s As String) As Date
    Dim parts() As String: parts = Split(s, "-")
    If UBound(parts) <> 2 Then Err.Raise vbObjectError + 3, , "日付形式が不正です: " & s
    ParseIsoDate = DateSerial(CLng(parts(0)), CLng(parts(1)), CLng(parts(2)))
End Function
