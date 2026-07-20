VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_LotInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' 納期計画 1 行分 (T_納期計画のスナップショット)
'==============================================================================

Public LotId As String         ' ロットID (品番_出荷週_納期先)
Public ItemNo As String        ' 品番
Public ShipWeek As Date        ' 出荷週 (月曜日付)
Public Dest As String          ' 納期先 (空欄あり)
Public Required As Long        ' 必要数
Public PlanStatus As String    ' 計画状態 (有効/キャンセル)
