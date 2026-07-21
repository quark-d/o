Attribute VB_Name = "m_LedgerDialogs"
Option Explicit

'==============================================================================
' m_LedgerDialogs: エントリポイント (操作ボタン用マクロ + ファクトリ + 設定)
'
' 設定はツールブックの「設定」シートの名前付きセルから読む:
'   cfgLedgerPath   台帳.xlsx のフルパス (必須)
'   cfgUiKind       "Forms" / "PowerShell" (既定 Forms)
'   cfgPsScriptPath Show-LedgerDialog.ps1 のパス (空 = ツールと同じフォルダ)
'   cfgJsonDir      受け渡し JSON の置き場所 (空 = %TEMP%)
'==============================================================================

'------------------------------------------------------------------------------
' 操作ボタン用マクロ
'------------------------------------------------------------------------------
Public Sub ShowCarryInDialog()
    m_LedgerDialogs.RunLedgerDialog LedgerOpKind.opCarryIn
End Sub

Public Sub ShowProgressDialog()
    m_LedgerDialogs.RunLedgerDialog LedgerOpKind.opProgress
End Sub

Public Sub ShowDefectDialog()
    m_LedgerDialogs.RunLedgerDialog LedgerOpKind.opDefect
End Sub

Public Sub ShowExceptionDialog()
    m_LedgerDialogs.RunLedgerDialog LedgerOpKind.opException
End Sub

' 編集者向け: 行の直接修正のための保護解除 (修正後は必ず再保護すること)
Public Sub UnprotectLedgerNow()
    On Error GoTo Fail
    Dim wb As Workbook: Set wb = OpenLedgerBook()
    If wb Is Nothing Then Exit Sub
    m_LedgerCore.UnprotectLedgerBook wb
    MsgBox "保護を解除しました。修正が終わったら「再保護」を実行してください。", _
           vbInformation, "保護解除"
    Exit Sub
Fail:
    MsgBox Err.Description, vbExclamation, "保護解除"
End Sub

Public Sub ProtectLedgerNow()
    On Error GoTo Fail
    Dim wb As Workbook: Set wb = OpenLedgerBook()
    If wb Is Nothing Then Exit Sub
    Application.CalculateFullRebuild
    m_LedgerCore.ProtectLedgerBook wb
    m_LedgerCore.SaveLedgerBook wb
    MsgBox "再計算・保護・保存が完了しました。", vbInformation, "再保護"
    Exit Sub
Fail:
    MsgBox Err.Description, vbExclamation, "再保護"
End Sub

'------------------------------------------------------------------------------
' ビルド後の疎通確認用 (build_xlsm.ps1 が Application.Run で呼ぶ)
' フォームのインスタンス化と m_LedgerCore の導出ロジックが通ることを確認する
'------------------------------------------------------------------------------
Public Function SmokeTest() As String
    Dim f1 As New CarryInForm
    Dim f2 As New ProgressForm
    Dim f3 As New DefectForm
    Dim f4 As New ExceptionForm
    Unload f1: Unload f2: Unload f3: Unload f4

    If m_LedgerCore.DeriveFromState(4, "梱包開始") <> "アセンブリ済み" Then
        SmokeTest = "NG: DeriveFromState(4, 梱包開始)"
        Exit Function
    End If
    If m_LedgerCore.DeriveFromState(1, "梱包開始") <> "未処理" Then
        SmokeTest = "NG: DeriveFromState(1, 梱包開始)"
        Exit Function
    End If
    If m_LedgerCore.AllowedEvents(2).Count <> 7 Then
        SmokeTest = "NG: AllowedEvents(2).Count"
        Exit Function
    End If
    Dim dlg As c_ILedgerEventDialog
    Set dlg = New c_FormLedgerEventDialog
    Set dlg = New c_PsLedgerEventDialog
    SmokeTest = "ok"
End Function

'------------------------------------------------------------------------------
' メインフロー
'------------------------------------------------------------------------------
Public Sub RunLedgerDialog(ByVal kind As LedgerOpKind)
    On Error GoTo Fail

    Dim wb As Workbook: Set wb = OpenLedgerBook()
    If wb Is Nothing Then Exit Sub

    ' 選択肢スナップショットは必ず再計算後に取る
    Application.CalculateFullRebuild
    Dim ctx As c_DialogContext: Set ctx = m_LedgerCore.BuildDialogContext(wb)

    Dim dlg As c_ILedgerEventDialog: Set dlg = m_LedgerDialogs.CreateLedgerDialog()

    Dim result As c_LedgerDialogResult: Set result = dlg.ShowDialog(kind, ctx)

    If result.Status = "cancel" Then Exit Sub
    If result.Status <> "ok" Then
        MsgBox "ダイアログでエラーが発生しました:" & vbCrLf & result.Message, _
               vbExclamation, "入力エラー"
        Exit Sub
    End If

    ' 検証 (導出フィールドの設定もここで行われる)
    Dim op As c_LedgerOperation: Set op = result.Operation
    Dim errs As Collection, warns As Collection
    If Not m_LedgerCore.ValidateOperation(op, ctx, errs, warns) Then
        MsgBox "入力内容に問題があります:" & vbCrLf & JoinCollection(errs), _
               vbExclamation, "検証エラー"
        Exit Sub
    End If
    If warns.Count > 0 Then
        If MsgBox("確認してください:" & vbCrLf & JoinCollection(warns) & vbCrLf & _
                  "このまま書き込みますか?", vbYesNo + vbQuestion, "警告") = vbNo Then
            Exit Sub
        End If
    End If

    ' 展開 → 最終確認 → 書込
    Dim rows As Collection: Set rows = m_LedgerCore.ExpandOperation(op, ctx)
    If MsgBox("T_台帳に以下を書き込みます:" & vbCrLf & DescribeRows(rows) & vbCrLf & _
              "よろしいですか?", vbYesNo + vbQuestion, "書込確認") = vbNo Then
        Exit Sub
    End If

    Dim addedNos As Collection: Set addedNos = m_LedgerCore.WriteOperation(wb, rows, ctx)
    If addedNos.Count = 0 Then
        MsgBox "書込は取り消されました。", vbInformation, "取消"
        Exit Sub
    End If

    m_LedgerCore.ProtectLedgerBook wb
    m_LedgerCore.SaveLedgerBook wb
    m_LedgerCore.SaveLastRecorder op.Recorder

    MsgBox "書き込みました (No " & addedNos(1) & _
           IIf(addedNos.Count > 1, "～" & addedNos(addedNos.Count), "") & ")。", _
           vbInformation, "完了"
    Exit Sub

Fail:
    MsgBox "エラーが発生しました:" & vbCrLf & Err.Description, vbCritical, "台帳入力ツール"
End Sub

'------------------------------------------------------------------------------
' ファクトリ (実装の切り替えはここだけ)
'------------------------------------------------------------------------------
Public Function CreateLedgerDialog() As c_ILedgerEventDialog
    If LCase$(GetConfig("cfgUiKind", "Forms")) = "powershell" Then
        Dim ps As New c_PsLedgerEventDialog
        Dim scriptPath As String: scriptPath = GetConfig("cfgPsScriptPath", "")
        If scriptPath <> "" Then ps.ScriptPath = scriptPath
        Dim jsonDir As String: jsonDir = GetConfig("cfgJsonDir", "")
        If jsonDir <> "" Then
            ps.InputPath = jsonDir & "\ledger_dialog_in.json"
            ps.OutputPath = jsonDir & "\ledger_dialog_out.json"
        End If
        Set CreateLedgerDialog = ps
    Else
        Set CreateLedgerDialog = New c_FormLedgerEventDialog
    End If
End Function

'------------------------------------------------------------------------------
' 台帳ブックを開く (ReadOnly 検知のフェイルセーフ込み)
'------------------------------------------------------------------------------
Private Function OpenLedgerBook() As Workbook
    Dim path As String: path = GetConfig("cfgLedgerPath", "")
    If path = "" Or Dir(path) = "" Then
        MsgBox "台帳ファイルが見つかりません。設定シートの「台帳パス」を確認してください:" & _
               vbCrLf & path, vbExclamation, "台帳入力ツール"
        Exit Function
    End If

    Dim fileName As String: fileName = Mid$(path, InStrRev(path, "\") + 1)

    Dim wb As Workbook
    Dim openedByUs As Boolean
    On Error Resume Next
    Set wb = Workbooks(fileName)
    On Error GoTo 0

    If wb Is Nothing Then
        Set wb = Workbooks.Open(fileName:=path, UpdateLinks:=0, _
                                IgnoreReadOnlyRecommended:=True)
        openedByUs = True
    End If

    If wb.ReadOnly Then
        MsgBox "台帳が読み取り専用で開かれています (他の人が編集中の可能性があります)。" & _
               vbCrLf & "このままでは書き込めないため、処理を中止します。", _
               vbExclamation, "台帳入力ツール"
        If openedByUs Then wb.Close SaveChanges:=False
        Exit Function
    End If

    Set OpenLedgerBook = wb
End Function

'------------------------------------------------------------------------------
' ヘルパー
'------------------------------------------------------------------------------
Private Function GetConfig(ByVal name As String, ByVal defaultValue As String) As String
    On Error GoTo Fallback
    GetConfig = Trim$(CStr(ThisWorkbook.Names(name).RefersToRange.Value & ""))
    If GetConfig = "" Then GetConfig = defaultValue
    Exit Function
Fallback:
    GetConfig = defaultValue
End Function

Private Function JoinCollection(ByVal col As Collection) As String
    Dim v As Variant
    For Each v In col
        JoinCollection = JoinCollection & "・" & CStr(v) & vbCrLf
    Next
End Function

Private Function DescribeRows(ByVal rows As Collection) As String
    Dim row As Variant
    Dim line As String
    For Each row In rows
        line = "・" & Format$(CDate(row(0)), "yyyy/mm/dd") & "  " & row(1) & "  " & row(2)
        If CStr(row(3) & "") <> "" Then line = line & "  " & row(3)
        If CStr(row(4) & "") <> "" Then line = line & " → " & row(4)
        If CStr(row(5) & "") <> "" Then line = line & "  [" & row(5) & "]"
        line = line & "  " & row(6) & " 個"
        DescribeRows = DescribeRows & line & vbCrLf
    Next
End Function
