VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormItemSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' IItemSelector 実装: VBA UserForm 版
'
' SelectItemsForm をモーダル表示する。同一プロセス内なので JSON は使わず、
' Collection を直接受け渡す。
'==============================================================================

Implements IItemSelector

Private Function IItemSelector_SelectItems(ByVal items As Collection, _
                                           ByVal title As String, _
                                           ByVal multiSelect As Boolean) As SelectItemsResult
    Dim result As New SelectItemsResult
    On Error GoTo Fail

    Dim frm As SelectItemsForm
    Set frm = New SelectItemsForm
    frm.Setup items, title, multiSelect
    frm.Show vbModal

    If frm.Accepted Then
        result.Init "ok", frm.SelectedItems, ""
    Else
        result.Init "cancel", Nothing, ""
    End If
    Unload frm

    Set IItemSelector_SelectItems = result
    Exit Function

Fail:
    result.Init "error", Nothing, Err.Description
    Set IItemSelector_SelectItems = result
End Function
