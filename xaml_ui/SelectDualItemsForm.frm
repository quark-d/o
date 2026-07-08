VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelectDualItemsForm 
   Caption         =   "SelectDualItemsForm"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "SelectDualItemsForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "SelectDualItemsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' 2 リストボックス (左右) アイテム選択フォーム (UserForm 版)
'
' コントロールは UserForm_Initialize で実行時に生成する。
' 左右のリストボックスは常に同じ幅・同じ高さ。
' 直接使わず、FormDualItemSelector 経由で利用すること。
'==============================================================================

Private mLblLeft As MSForms.Label
Private mLblRight As MSForms.Label
Private mListLeft As MSForms.ListBox
Private mListRight As MSForms.ListBox
Private WithEvents mBtnOK As MSForms.CommandButton
Attribute mBtnOK.VB_VarHelpID = -1
Private WithEvents mBtnCancel As MSForms.CommandButton
Attribute mBtnCancel.VB_VarHelpID = -1
Private mAccepted As Boolean

Private Sub UserForm_Initialize()
    Const MARGIN As Single = 10
    Const GAP As Single = 8
    Const LBL_H As Single = 12
    Const BTN_W As Single = 72
    Const BTN_H As Single = 24

    Me.Width = 470
    Me.Height = 330

    ' 左右のリストは常に同じ幅・同じ高さにする
    Dim listW As Single
    listW = (Me.InsideWidth - MARGIN * 2 - GAP) / 2
    Dim listTop As Single
    listTop = MARGIN + LBL_H + 2
    Dim listH As Single
    listH = Me.InsideHeight - listTop - MARGIN - BTN_H - GAP

    Set mLblLeft = Me.Controls.Add("Forms.Label.1", "lblLeft")
    With mLblLeft
        .Left = MARGIN
        .Top = MARGIN
        .Width = listW
        .Height = LBL_H
    End With

    Set mLblRight = Me.Controls.Add("Forms.Label.1", "lblRight")
    With mLblRight
        .Left = MARGIN + listW + GAP
        .Top = MARGIN
        .Width = listW
        .Height = LBL_H
    End With

    Set mListLeft = Me.Controls.Add("Forms.ListBox.1", "lstLeft")
    With mListLeft
        .Left = MARGIN
        .Top = listTop
        .Width = listW
        .Height = listH
    End With

    Set mListRight = Me.Controls.Add("Forms.ListBox.1", "lstRight")
    With mListRight
        .Left = MARGIN + listW + GAP
        .Top = listTop
        .Width = listW
        .Height = listH
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
Public Sub Setup(ByVal leftList As ItemListSpec, ByVal rightList As ItemListSpec, _
                 ByVal title As String)
    Me.Caption = title
    SetupList mListLeft, mLblLeft, leftList
    SetupList mListRight, mLblRight, rightList
End Sub

Private Sub SetupList(ByVal lst As MSForms.ListBox, ByVal lbl As MSForms.Label, _
                      ByVal spec As ItemListSpec)
    lbl.Caption = spec.Caption
    lst.MultiSelect = IIf(spec.MultiSelect, fmMultiSelectExtended, fmMultiSelectSingle)
    Dim v As Variant
    For Each v In spec.Items
        lst.AddItem CStr(v)
    Next
End Sub

Public Property Get Accepted() As Boolean
    Accepted = mAccepted
End Property

Public Property Get LeftSelectedItems() As Collection
    Set LeftSelectedItems = CollectSelected(mListLeft)
End Property

Public Property Get RightSelectedItems() As Collection
    Set RightSelectedItems = CollectSelected(mListRight)
End Property

Private Function CollectSelected(ByVal lst As MSForms.ListBox) As Collection
    Dim col As New Collection
    Dim i As Long
    For i = 0 To lst.ListCount - 1
        If lst.Selected(i) Then col.Add lst.List(i)
    Next
    Set CollectSelected = col
End Function

Private Sub mBtnOK_Click()
    mAccepted = True
    Me.Hide
End Sub

Private Sub mBtnCancel_Click()
    mAccepted = False
    Me.Hide
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


