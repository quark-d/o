VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SelectDualItemsResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' 2 リストボックス版アイテム選択の結果
'   Status        : "ok" / "cancel" / "error"
'   LeftSelected  : 左のリストで選択されたアイテム (文字列の Collection)
'   RightSelected : 右のリストで選択されたアイテム (文字列の Collection)
'   Message       : エラー時の詳細
'==============================================================================

Private mStatus As String
Private mLeftSelected As Collection
Private mRightSelected As Collection
Private mMessage As String

Private Sub Class_Initialize()
    mStatus = "error"
    Set mLeftSelected = New Collection
    Set mRightSelected = New Collection
End Sub

' 実装クラスが結果を設定するために呼ぶ (利用側からは呼ばない)
Public Sub Init(ByVal status As String, ByVal leftSelected As Collection, _
                ByVal rightSelected As Collection, ByVal message As String)
    mStatus = status
    If leftSelected Is Nothing Then
        Set mLeftSelected = New Collection
    Else
        Set mLeftSelected = leftSelected
    End If
    If rightSelected Is Nothing Then
        Set mRightSelected = New Collection
    Else
        Set mRightSelected = rightSelected
    End If
    mMessage = message
End Sub

Public Property Get Status() As String
    Status = mStatus
End Property

Public Property Get LeftSelected() As Collection
    Set LeftSelected = mLeftSelected
End Property

Public Property Get RightSelected() As Collection
    Set RightSelected = mRightSelected
End Property

Public Property Get Message() As String
    Message = mMessage
End Property

Public Property Get IsOk() As Boolean
    IsOk = (mStatus = "ok")
End Property
