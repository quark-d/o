VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SelectItemsResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' アイテム選択の結果
'   Status   : "ok" / "cancel" / "error"
'   Selected : 選択されたアイテム (文字列の Collection)
'   Message  : エラー時の詳細
'==============================================================================

Private mStatus As String
Private mSelected As Collection
Private mMessage As String

Private Sub Class_Initialize()
    mStatus = "error"
    Set mSelected = New Collection
End Sub

' 実装クラスが結果を設定するために呼ぶ (利用側からは呼ばない)
Public Sub Init(ByVal status As String, ByVal selected As Collection, _
                ByVal message As String)
    mStatus = status
    If selected Is Nothing Then
        Set mSelected = New Collection
    Else
        Set mSelected = selected
    End If
    mMessage = message
End Sub

Public Property Get Status() As String
    Status = mStatus
End Property

Public Property Get Selected() As Collection
    Set Selected = mSelected
End Property

Public Property Get Message() As String
    Message = mMessage
End Property

Public Property Get IsOk() As Boolean
    IsOk = (mStatus = "ok")
End Property
