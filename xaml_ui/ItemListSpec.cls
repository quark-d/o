VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ItemListSpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' リストボックス 1 つ分の表示仕様
'   Caption     : リストの上に出す見出し (空 = 見出しなし)
'   MultiSelect : True = 複数選択可 (既定)
'   Items       : 表示するアイテム (文字列の Collection)
'
' IDualItemSelector.SelectItems() に左右 1 つずつ渡す。
'==============================================================================

Public Caption As String
Public MultiSelect As Boolean
Private mItems As Collection

Private Sub Class_Initialize()
    Caption = ""
    MultiSelect = True      ' 既定は複数選択
    Set mItems = New Collection
End Sub

Public Property Get Items() As Collection
    Set Items = mItems
End Property

Public Property Set Items(ByVal value As Collection)
    If value Is Nothing Then
        Set mItems = New Collection
    Else
        Set mItems = value
    End If
End Property

' Items.Add の省略形。spec.Add "氏名" のように使う
Public Sub Add(ByVal item As String)
    mItems.Add item
End Sub
