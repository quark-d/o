VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormDualItemSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' IDualItemSelector 実装: VBA UserForm 版
'
' SelectDualItemsForm をモーダル表示する。同一プロセス内なので JSON は使わず、
' ItemListSpec / Collection を直接受け渡す。
'==============================================================================

Implements IDualItemSelector

Private Function IDualItemSelector_SelectItems(ByVal leftList As ItemListSpec, _
                                               ByVal rightList As ItemListSpec, _
                                               ByVal title As String) As SelectDualItemsResult
    Dim result As New SelectDualItemsResult
    On Error GoTo Fail

    Dim frm As SelectDualItemsForm
    Set frm = New SelectDualItemsForm
    frm.Setup leftList, rightList, title
    frm.Show vbModal

    If frm.Accepted Then
        result.Init "ok", frm.LeftSelectedItems, frm.RightSelectedItems, ""
    Else
        result.Init "cancel", Nothing, Nothing, ""
    End If
    Unload frm

    Set IDualItemSelector_SelectItems = result
    Exit Function

Fail:
    result.Init "error", Nothing, Nothing, Err.Description
    Set IDualItemSelector_SelectItems = result
End Function
