VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_LedgerDialogResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' ダイアログの戻り値 (Status / Message / Operation)
'==============================================================================

Public Status As String              ' "ok" / "cancel" / "error"
Public Message As String             ' エラー詳細 (Status="error" のとき)
Public Operation As c_LedgerOperation ' 入力された操作 (Status="ok" のとき)

Public Sub Init(ByVal aStatus As String, ByVal op As c_LedgerOperation, _
                ByVal aMessage As String)
    Status = aStatus
    Set Operation = op
    Message = aMessage
End Sub

Public Property Get IsOk() As Boolean
    IsOk = (Status = "ok")
End Property
