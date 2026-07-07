Attribute VB_Name = "SelectItemsDialog"
Option Explicit

'==============================================================================
' アイテム選択ダイアログ - エントリポイント
'
' CreateItemSelector() で実装 (PowerShell 版 / UserForm 版) を選び、
' IItemSelector.SelectItems() で表示する。呼び出し方は両者共通。
'
' 必要なファイル:
'   IItemSelector.cls / SelectItemsResult.cls
'   PsItemSelector.cls   (+ Select-Items.ps1)
'   FormItemSelector.cls (+ SelectItemsForm.frm)
'==============================================================================

Public Enum SelectorKind
    skPowerShell = 0    ' PowerShell (WPF) 版
    skUserForm = 1      ' VBA UserForm 版
End Enum

' 実装を切り替えるファクトリ
Public Function CreateItemSelector(ByVal kind As SelectorKind) As IItemSelector
    Select Case kind
        Case skPowerShell
            Set CreateItemSelector = New PsItemSelector
        Case skUserForm
            Set CreateItemSelector = New FormItemSelector
        Case Else
            Err.Raise 5, "CreateItemSelector", "未対応の SelectorKind です: " & kind
    End Select
End Function

'------------------------------------------------------------------------------
' デモ: アクティブセルが属するテーブルの列名を選択させる
'------------------------------------------------------------------------------
Public Sub Demo_SelectColumns_PowerShell()
    RunDemo skPowerShell
End Sub

Public Sub Demo_SelectColumns_UserForm()
    RunDemo skUserForm
End Sub

Private Sub RunDemo(ByVal kind As SelectorKind)
    Dim lo As ListObject
    On Error Resume Next
    Set lo = ActiveCell.ListObject
    On Error GoTo 0
    If lo Is Nothing Then
        MsgBox "アクティブセルがテーブル内にありません。", vbExclamation
        Exit Sub
    End If

    ' テーブルの列名を Collection にする
    Dim items As New Collection
    Dim lc As ListColumn
    For Each lc In lo.ListColumns
        items.Add lc.Name
    Next

    ' ダイアログを表示 (実装は kind で切り替え、呼び出し方は共通)
    Dim selector As IItemSelector
    Set selector = CreateItemSelector(kind)

    Dim result As SelectItemsResult
    Set result = selector.SelectItems(items, "列を選択してください", True)

    Select Case result.Status
        Case "ok"
            Dim msg As String
            Dim v As Variant
            For Each v In result.Selected
                msg = msg & v & vbCrLf
            Next
            MsgBox "選択された列 (" & result.Selected.Count & " 件):" & vbCrLf & msg, _
                   vbInformation
        Case "cancel"
            MsgBox "キャンセルされました。", vbInformation
        Case Else
            MsgBox "エラー: " & result.Message, vbExclamation
    End Select
End Sub
