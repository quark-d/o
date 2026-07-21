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
Public Recorders As Collection     ' 記録者リスト (M_リスト「記録者」列)
Public States As Collection        ' 状態リスト (M_リスト「状態」列の順)
Public DefaultActionDate As Date   ' 実施日の既定値 (通常は今日)
Public DefaultRecorder As String   ' 記録者の既定値 (前回値)
Public WarehouseA As String        ' 倉庫Aの表示名 (ツールの設定シートで変更可)
Public WarehouseB As String        ' 倉庫Bの表示名 (同上)

Private Sub Class_Initialize()
    Set Items = New Collection
    Set Lots = New Collection
    Set Recorders = New Collection
    Set States = New Collection
    WarehouseA = "倉庫A"
    WarehouseB = "倉庫B"
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

' localId からマスタ行を引く (なければ Nothing)
Public Function FindItemByLocalId(ByVal localId As String) As c_ItemInfo
    Dim it As c_ItemInfo
    For Each it In Items
        If it.LocalId = localId Then
            Set FindItemByLocalId = it
            Exit Function
        End If
    Next
End Function

' 品名からマスタ行を引く (完全一致。重複し得るため全候補を返す)
Public Function FindItemsByName(ByVal itemName As String) As Collection
    Dim result As New Collection
    Dim it As c_ItemInfo
    For Each it In Items
        If it.ItemName = itemName Then result.Add it
    Next
    Set FindItemsByName = result
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
