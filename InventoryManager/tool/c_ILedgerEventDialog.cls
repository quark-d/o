VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_ILedgerEventDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' 台帳イベント入力ダイアログの共通インターフェイス
'
' kind で 4 画面 (搬入/工程進行/不良+補填/例外操作) を選ぶ。
' 実装: c_FormLedgerEventDialog (VBA UserForm 版) /
'       c_PsLedgerEventDialog (PS/XAML 版)。
' 生成は m_LedgerDialogs.CreateLedgerDialog を使うこと。
'==============================================================================

Public Function ShowDialog(ByVal kind As LedgerOpKind, _
                           ByVal ctx As c_DialogContext) As c_LedgerDialogResult
End Function
