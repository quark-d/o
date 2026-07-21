VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_FormLedgerEventDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' c_ILedgerEventDialog 実装: VBA UserForm 版 (基本)
'
' 同一プロセスなのでファイルを挟まず、c_DialogContext / c_LedgerOperation を
' オブジェクトのまま受け渡す (設計合意: JSON はプロセス境界のみ)。
'==============================================================================

Implements c_ILedgerEventDialog

Private Function c_ILedgerEventDialog_ShowDialog(ByVal kind As LedgerOpKind, _
        ByVal ctx As c_DialogContext) As c_LedgerDialogResult
    Dim result As New c_LedgerDialogResult
    On Error GoTo Fail

    Select Case kind
        Case LedgerOpKind.opCarryIn
            Dim frmC As New CarryInForm
            frmC.InitDialog ctx
            frmC.Show vbModal
            Set result = frmC.Result
            Unload frmC
        Case LedgerOpKind.opProgress
            Dim frmP As New ProgressForm
            frmP.InitDialog ctx
            frmP.Show vbModal
            Set result = frmP.Result
            Unload frmP
        Case LedgerOpKind.opDefect
            Dim frmD As New DefectForm
            frmD.InitDialog ctx
            frmD.Show vbModal
            Set result = frmD.Result
            Unload frmD
        Case LedgerOpKind.opException
            Dim frmE As New ExceptionForm
            frmE.InitDialog ctx
            frmE.Show vbModal
            Set result = frmE.Result
            Unload frmE
        Case Else
            result.Init "error", Nothing, "不明なダイアログ種別です: " & kind
    End Select

    Set c_ILedgerEventDialog_ShowDialog = result
    Exit Function

Fail:
    result.Init "error", Nothing, Err.Description
    Set c_ILedgerEventDialog_ShowDialog = result
End Function
