VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_ItemInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' 品番マスタ 1 行分 (M_品番のスナップショット)
'==============================================================================

Public ItemNo As String        ' 品番
Public ItemName As String      ' 品名
Public Pattern As Long         ' 工程パターン 1～4
