Attribute VB_Name = "m_LedgerFormUtil"
Option Explicit

'==============================================================================
' UserForm 4 画面の共通ヘルパー (フォームコードからのみ使用)
'==============================================================================

' 記録者コンボと実施日テキストに既定値を入れる
Public Sub FillCommon(ByVal ctx As c_DialogContext, _
                      ByVal cmbRecorder As MSForms.ComboBox, _
                      ByVal txtDate As MSForms.TextBox)
    Dim v As Variant
    For Each v In ctx.Recorders
        cmbRecorder.AddItem CStr(v)
    Next
    ' 既定値が空 (初回起動など) のときは未選択のまま (DropDownList は空文字を代入できない)
    If Trim$(ctx.DefaultRecorder) <> "" Then cmbRecorder.Text = ctx.DefaultRecorder
    txtDate.Text = Format$(ctx.DefaultActionDate, "yyyy/mm/dd")
End Sub

' 数量・実施日・記録者の形式チェック (NG ならメッセージを出して False)
Public Function CheckCommon(ByVal txtQty As MSForms.TextBox, _
                            ByVal txtDate As MSForms.TextBox, _
                            ByVal cmbRecorder As MSForms.ComboBox) As Boolean
    If Not IsNumeric(txtQty.Text) Then
        MsgBox "数量を数値で入力してください。", vbExclamation
        Exit Function
    End If
    If CDbl(txtQty.Text) < 1 Or CDbl(txtQty.Text) <> Fix(CDbl(txtQty.Text)) Then
        MsgBox "数量は 1 以上の整数で入力してください。", vbExclamation
        Exit Function
    End If
    If Not IsDate(txtDate.Text) Then
        MsgBox "実施日を日付形式 (yyyy/mm/dd) で入力してください。", vbExclamation
        Exit Function
    End If
    If Trim$(cmbRecorder.Text) = "" Then
        MsgBox "記録者を選択してください。", vbExclamation
        Exit Function
    End If
    CheckCommon = True
End Function

' バケットの状態別在庫を「未処理 9 / アニール済み 2」形式で返す (ゼロは省略)
Public Function DescribeStates(ByVal ctx As c_DialogContext, ByVal bucket As String, _
                               ByVal id As String) As String
    Dim st As Variant
    Dim q As Long
    For Each st In ctx.States
        q = ctx.Stock.AvailableQty(bucket, id, CStr(st))
        If q > 0 Then
            If DescribeStates <> "" Then DescribeStates = DescribeStates & " / "
            DescribeStates = DescribeStates & CStr(st) & " " & q
        End If
    Next
    If DescribeStates = "" Then DescribeStates = "なし"
End Function

' コンボの選択肢を Collection で丸ごと入れ替える
Public Sub FillCombo(ByVal cmb As MSForms.ComboBox, ByVal items As Collection)
    cmb.Clear
    Dim v As Variant
    For Each v In items
        cmb.AddItem CStr(v)
    Next
End Sub

' マスタ行の表示用整形 (localId  品番  品名)
Public Function DescribeItem(ByVal it As c_ItemInfo) As String
    DescribeItem = it.LocalId & "  " & it.ItemNo & "  " & it.ItemName
End Function

' 品名重複時の候補選択 (番号入力)。キャンセル/不正入力は Nothing
Public Function PickItemFromCandidates(ByVal cands As Collection) As c_ItemInfo
    Dim msg As String
    msg = "品名が複数の品番に一致しました。番号で選択してください:" & vbCrLf
    Dim i As Long
    For i = 1 To cands.Count
        msg = msg & i & ": " & m_LedgerFormUtil.DescribeItem(cands(i)) & vbCrLf
    Next
    Dim ans As String: ans = InputBox(msg, "候補の選択", "1")
    If ans = "" Or Not IsNumeric(ans) Then Exit Function
    If CLng(ans) < 1 Or CLng(ans) > cands.Count Then Exit Function
    Set PickItemFromCandidates = cands(CLng(ans))
End Function
