VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_DialogContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' ダイアログへの入力 (台帳ブックの選択肢スナップショット)
'
' m_LedgerCore.BuildDialogContext で生成する。生成時点で再計算済みであること。
'==============================================================================

Public Items As Collection         ' c_ItemInfo の一覧 (M_品番)
Public Lots As Collection          ' c_LotInfo の一覧 (T_納期計画。キャンセル含む)
Public Stock As c_StockSnapshot    ' V_在庫内訳のスナップショット
Public Recorders As Collection     ' 記録者リスト (M_リスト G列)
Public States As Collection        ' 状態リスト (M_リスト B列の順)
Public DefaultActionDate As Date   ' 実施日の既定値 (通常は今日)
Public DefaultRecorder As String   ' 記録者の既定値 (前回値)

Private Sub Class_Initialize()
    Set Items = New Collection
    Set Lots = New Collection
    Set Recorders = New Collection
    Set States = New Collection
End Sub

' 品番からマスタ行を引く (なければ Nothing)
Public Function FindItem(ByVal itemNo As String) As c_ItemInfo
    Dim it As c_ItemInfo
    For Each it In Items
        If it.ItemNo = itemNo Then
            Set FindItem = it
            Exit Function
        End If
    Next
End Function

' ロットIDから計画行を引く (なければ Nothing)
Public Function FindLot(ByVal lotId As String) As c_LotInfo
    Dim lt As c_LotInfo
    For Each lt In Lots
        If lt.LotId = lotId Then
            Set FindLot = lt
            Exit Function
        End If
    Next
End Function
