Attribute VB_Name = "SelectItemsDialog"
Option Explicit

'==============================================================================
' アイテム選択ダイアログ - エントリポイント
'
' リストボックス 1 つ版:
'   CreateItemSelector() で実装 (PowerShell 版 / UserForm 版) を選び、
'   IItemSelector.SelectItems() で表示する。呼び出し方は両者共通。
' リストボックス 2 つ (左右) 版:
'   CreateDualItemSelector() で実装を選び、左右それぞれの仕様を
'   ItemListSpec で渡して IDualItemSelector.SelectItems() で表示する。
'
' 必要なファイル:
'   共通            : JsonLite.bas
'   1 リスト版      : IItemSelector.cls / SelectItemsResult.cls
'                     PsItemSelector.cls   (+ Select-Items.ps1)
'                     FormItemSelector.cls (+ SelectItemsForm.frm)
'   2 リスト版      : IDualItemSelector.cls / SelectDualItemsResult.cls / ItemListSpec.cls
'                     PsDualItemSelector.cls   (+ Select-DualItems.ps1)
'                     FormDualItemSelector.cls (+ SelectDualItemsForm.frm)
'==============================================================================

Public Enum SelectorKind
    skPowerShell = 0    ' PowerShell (WPF) 版
    skUserForm = 1      ' VBA UserForm 版
End Enum

' 実装を切り替えるファクトリ (1 リスト版)
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

' 実装を切り替えるファクトリ (2 リスト版)
Public Function CreateDualItemSelector(ByVal kind As SelectorKind) As IDualItemSelector
    Select Case kind
        Case skPowerShell
            Set CreateDualItemSelector = New PsDualItemSelector
        Case skUserForm
            Set CreateDualItemSelector = New FormDualItemSelector
        Case Else
            Err.Raise 5, "CreateDualItemSelector", "未対応の SelectorKind です: " & kind
    End Select
End Function

'------------------------------------------------------------------------------
' デモ (1 リスト版): アクティブセルが属するテーブルの列名を選択させる
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

'------------------------------------------------------------------------------
' デモ (2 リスト版): テーブルの列名 (複数選択) と出力形式 (単数選択) を選ばせる
'------------------------------------------------------------------------------
Public Sub Demo_SelectDualColumns_PowerShell()
    RunDualDemo skPowerShell
End Sub

Public Sub Demo_SelectDualColumns_UserForm()
    RunDualDemo skUserForm
End Sub

Private Sub RunDualDemo(ByVal kind As SelectorKind)
    Dim lo As ListObject
    On Error Resume Next
    Set lo = ActiveCell.ListObject
    On Error GoTo 0
    If lo Is Nothing Then
        MsgBox "アクティブセルがテーブル内にありません。", vbExclamation
        Exit Sub
    End If

    ' 左: テーブルの列名 (複数選択 = 既定のまま)
    Dim leftList As New ItemListSpec
    leftList.Caption = "出力する列"
    Dim lc As ListColumn
    For Each lc In lo.ListColumns
        leftList.Add lc.Name
    Next

    ' 右: 出力形式 (単数選択)
    Dim rightList As New ItemListSpec
    rightList.Caption = "出力形式"
    rightList.MultiSelect = False
    rightList.Add "新規シート"
    rightList.Add "CSV ファイル"
    rightList.Add "クリップボード"

    ' ダイアログを表示 (実装は kind で切り替え、呼び出し方は共通)
    Dim selector As IDualItemSelector
    Set selector = CreateDualItemSelector(kind)

    Dim result As SelectDualItemsResult
    Set result = selector.SelectItems(leftList, rightList, "出力する列と形式を選択してください")

    Select Case result.Status
        Case "ok"
            Dim msg As String
            Dim v As Variant
            msg = "出力する列 (" & result.LeftSelected.Count & " 件):" & vbCrLf
            For Each v In result.LeftSelected
                msg = msg & "  " & v & vbCrLf
            Next
            msg = msg & vbCrLf & "出力形式:" & vbCrLf
            For Each v In result.RightSelected
                msg = msg & "  " & v & vbCrLf
            Next
            MsgBox msg, vbInformation
        Case "cancel"
            MsgBox "キャンセルされました。", vbInformation
        Case Else
            MsgBox "エラー: " & result.Message, vbExclamation
    End Select
End Sub
