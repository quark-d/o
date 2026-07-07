VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelectItemsForm 
   Caption         =   "SelectItemsForm"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "SelectItemsForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "SelectItemsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' アイテム選択フォーム (UserForm 版)
'
' コントロールは UserForm_Initialize で実行時に生成する。
' 直接使わず、FormItemSelector 経由で利用すること。
'==============================================================================

Private WithEvents mList As MSForms.ListBox
Attribute mList.VB_VarHelpID = -1
Private WithEvents mBtnOK As MSForms.CommandButton
Attribute mBtnOK.VB_VarHelpID = -1
Private WithEvents mBtnCancel As MSForms.CommandButton
Attribute mBtnCancel.VB_VarHelpID = -1
Private mAccepted As Boolean

Private Sub UserForm_Initialize()
    Const MARGIN As Single = 10
    Const BTN_W As Single = 72
    Const BTN_H As Single = 24
    Const GAP As Single = 8

    Me.Width = 250
    Me.Height = 330

    Set mList = Me.Controls.Add("Forms.ListBox.1", "lstItems")
    With mList
        .Left = MARGIN
        .Top = MARGIN
        .Width = Me.InsideWidth - MARGIN * 2
        .Height = Me.InsideHeight - MARGIN * 2 - BTN_H - GAP
    End With

    Set mBtnCancel = Me.Controls.Add("Forms.CommandButton.1", "btnCancel")
    With mBtnCancel
        .Caption = "キャンセル"
        .Width = BTN_W
        .Height = BTN_H
        .Left = Me.InsideWidth - MARGIN - BTN_W
        .Top = Me.InsideHeight - MARGIN - BTN_H
        .Cancel = True    ' Esc キーで押される
    End With

    Set mBtnOK = Me.Controls.Add("Forms.CommandButton.1", "btnOK")
    With mBtnOK
        .Caption = "OK"
        .Width = BTN_W
        .Height = BTN_H
        .Left = mBtnCancel.Left - GAP - BTN_W
        .Top = mBtnCancel.Top
        .Default = True   ' Enter キーで押される
    End With
End Sub

' 表示前に呼んで内容を設定する
Public Sub Setup(ByVal items As Collection, ByVal title As String, _
                 ByVal multiSelect As Boolean)
    Me.Caption = title
    mList.multiSelect = IIf(multiSelect, fmMultiSelectExtended, fmMultiSelectSingle)
    Dim v As Variant
    For Each v In items
        mList.AddItem CStr(v)
    Next
End Sub

Public Property Get Accepted() As Boolean
    Accepted = mAccepted
End Property

Public Property Get SelectedItems() As Collection
    Dim col As New Collection
    Dim i As Long
    For i = 0 To mList.ListCount - 1
        If mList.Selected(i) Then col.Add mList.List(i)
    Next
    Set SelectedItems = col
End Property

Private Sub mBtnOK_Click()
    mAccepted = True
    Me.Hide
End Sub

Private Sub mBtnCancel_Click()
    mAccepted = False
    Me.Hide
End Sub

' ダブルクリックで即決定
Private Sub mList_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If mList.ListIndex >= 0 Then
        mAccepted = True
        Me.Hide
    End If
End Sub

' × で閉じられたときは Unload せず Hide してキャンセル扱いにする
' (Unload されると呼び出し元が Accepted を読めなくなるため)
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        mAccepted = False
        Me.Hide
    End If
End Sub


